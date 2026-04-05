from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
import os
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict

from append_excel_entry import (
    DEFAULT_TIMEZONE,
    EXPECTED_HEADERS,
    STATUS_PENDING,
    append_record_to_excel,
    convert_telegram_timestamp,
    get_default_payment_channel,
    invalidate_record_by_id,
    load_bot_config,
    normalize_record,
    read_record_by_id,
)
from generate_expense_report import refresh_report_workbook
from telegram_codex_bridge import apply_record as bridge_apply_record
from telegram_codex_bridge import emit_prompt as bridge_emit_prompt
from telegram_record_schema import CODEX_OUTPUT_SCHEMA

API_TIMEOUT_SECONDS = 60
RETRY_SLEEP_SECONDS = 5
NETWORK_LOG_INTERVAL_SECONDS = 600
SEND_REPLY_TIMEOUT_SECONDS = 15
SEND_RETRY_ATTEMPTS = 3
SEND_RETRY_BASE_SLEEP_SECONDS = 2
PENDING_REPLY_BATCH_SIZE = 10
BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
RESTART_SCRIPT_PATH = BASE_DIR / "restart_bot.sh"
VOID_MARK = "作废"
CONFIG_FILE_NAME = "telegram_bot_config.json"


@dataclass(frozen=True)
class RuntimeContext:
    token: str
    workdir: Path
    state_path: Path
    allowed_username: str
    bot_config: Dict[str, Any]
    backend: str
    excel_path: Path
    verbose: bool


def classify_network_error(exc: Exception) -> str:
    msg = str(exc).lower()
    if "eof occurred" in msg:
        return "ssl_eof"
    if "handshake operation timed out" in msg:
        return "ssl_handshake_timeout"
    if "read operation timed out" in msg:
        return "timeout"
    if "remote end closed connection" in msg:
        return "remote_close"
    return "other"


def get_monthly_file_path(base_name: str, ext: str) -> Path:
    month_str = time.strftime("%Y-%m", time.localtime())
    folder = BASE_DIR / base_name
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f"{month_str}.{ext}"


def log_json(payload: Dict[str, Any]) -> None:
    ordered_payload: Dict[str, Any] = {
        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),
    }
    if "stage" in payload:
        ordered_payload["stage"] = payload["stage"]
    for key, value in payload.items():
        if key in {"timestamp", "stage"}:
            continue
        ordered_payload[key] = value

    log_file = get_monthly_file_path("logs", "log")
    try:
        with log_file.open("a", encoding="utf-8") as f:
            f.write(json.dumps(ordered_payload, ensure_ascii=False) + "\n")
    except:
        pass


def maybe_log_network_outage(
    verbose: bool,
    outage_started_at: float | None,
    last_logged_at: float | None,
    stage: str,
    exc: Exception,
) -> tuple[float, float]:
    now = time.time()
    if outage_started_at is None:
        outage_started_at = now

    should_log = last_logged_at is None or (now - last_logged_at) >= NETWORK_LOG_INTERVAL_SECONDS
    if should_log and verbose:
        log_json(
            {
                "stage": stage,
                "error_type": classify_network_error(exc),
                "error": str(exc),
                "outage_started_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(outage_started_at)),
                "retry_interval_seconds": RETRY_SLEEP_SECONDS,
                "log_interval_seconds": NETWORK_LOG_INTERVAL_SECONDS,
            }
        )
        last_logged_at = now

    return outage_started_at, last_logged_at


def log_message_event(stage: str, envelope: Dict[str, Any], **extra: Any) -> None:
    payload: Dict[str, Any] = {
        "stage": stage,
        "chat_id": envelope.get("chat_id"),
        "message_id": envelope.get("message_id"),
        "text": envelope.get("text", "")[:80],
    }
    payload.update(extra)
    log_json(payload)


def api_request(token: str, method: str, payload: Dict[str, Any] | None = None, timeout: int = 30) -> Dict[str, Any]:
    url = f"https://api.telegram.org/bot{token}/{method}"
    data = None
    headers = {}
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"

    request = urllib.request.Request(url, data=data, headers=headers, method="POST" if payload is not None else "GET")
    with urllib.request.urlopen(request, timeout=timeout) as response:
        body = response.read().decode("utf-8")
    result = json.loads(body)
    if not result.get("ok"):
        raise RuntimeError(f"Telegram API {method} failed: {result}")
    return result


def load_state(state_path: Path) -> Dict[str, Any]:
    if not state_path.exists():
        return {"offset": 0, "token": ""}
    return json.loads(state_path.read_text(encoding="utf-8"))


def save_state(state_path: Path, state: Dict[str, Any]) -> None:
    state_path.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")

def _state_reply_queue(state: Dict[str, Any]) -> list[Dict[str, Any]]:
    queue = state.get("pending_replies")
    if not isinstance(queue, list):
        return []
    return [item for item in queue if isinstance(item, dict)]


def _reply_queue_identity(entry: Dict[str, Any]) -> tuple[Any, Any, Any, Any]:
    return (
        entry.get("chat_id"),
        entry.get("reply_to_message_id"),
        entry.get("text"),
        entry.get("parse_mode"),
    )


def queue_pending_reply(
    state_path: Path,
    chat_id: int,
    reply_to_message_id: int,
    text: str,
    stage: str,
    parse_mode: str | None = None,
    error: str = "",
) -> None:
    state = load_state(state_path)
    queue = _state_reply_queue(state)
    queued_at = int(time.time())
    new_entry = {
        "chat_id": chat_id,
        "reply_to_message_id": reply_to_message_id,
        "text": text,
        "parse_mode": parse_mode or "",
        "stage": stage,
        "queued_at": queued_at,
        "last_error": error,
        "attempts": 0,
    }

    new_identity = _reply_queue_identity(new_entry)
    updated_queue: list[Dict[str, Any]] = []
    replaced = False
    for entry in queue:
        if _reply_queue_identity(entry) == new_identity:
            updated = dict(entry)
            updated["stage"] = stage
            updated["last_error"] = error
            updated["queued_at"] = entry.get("queued_at", queued_at)
            updated_queue.append(updated)
            replaced = True
            continue
        updated_queue.append(entry)

    if not replaced:
        updated_queue.append(new_entry)

    state["pending_replies"] = updated_queue
    state.pop("pending_restart_notice", None)
    save_state(state_path, state)


def migrate_legacy_pending_restart_notice(state_path: Path) -> None:
    state = load_state(state_path)
    notice = state.get("pending_restart_notice")
    if not isinstance(notice, dict):
        return

    chat_id = notice.get("chat_id")
    reply_to_message_id = notice.get("reply_to_message_id")
    if isinstance(chat_id, int) and isinstance(reply_to_message_id, int):
        queue_pending_reply(
            state_path,
            chat_id,
            reply_to_message_id,
            "机器人已重启完成",
            stage="restart_reply_failed",
        )
        state = load_state(state_path)

    state.pop("pending_restart_notice", None)
    save_state(state_path, state)


def load_message_index() -> Dict[str, Any]:
    index_path = get_monthly_file_path("indexes", "json")
    if not index_path.exists():
        return {}
    return json.loads(index_path.read_text(encoding="utf-8"))


def save_message_index(index: Dict[str, Any]) -> None:
    index_path = get_monthly_file_path("indexes", "json")
    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")


def configured_project_dir(state: Dict[str, Any]) -> Path:
    raw_value = str(state.get("project_dir", "")).strip()
    if raw_value:
        return Path(raw_value).expanduser()
    return PROJECT_DIR


def configured_excel_path(state: Dict[str, Any]) -> Path:
    return configured_project_dir(state) / "expense.xlsx"


def configured_allowed_username(state: Dict[str, Any]) -> str:
    return str(state.get("allowed_username", "")).strip()


def current_report_period() -> str:
    return time.strftime("%Y-%m", time.localtime())


def report_cutoff_date(current_period: str) -> datetime.date:
    current_month_start = datetime.strptime(f"{current_period}-01", "%Y-%m-%d").date()
    return current_month_start - timedelta(days=1)


def refresh_report_if_period_changed(state_path: Path, excel_path: Path, verbose: bool) -> None:
    state = load_state(state_path)
    current_period = current_report_period()
    last_period = str(state.get("report_period", "")).strip()
    report_path = excel_path.with_name("expense_report.xlsx")

    if last_period == current_period and report_path.exists():
        return

    cutoff_date = report_cutoff_date(current_period)
    refresh_report_workbook(excel_path, report_path, cutoff_date=cutoff_date)
    state["report_period"] = current_period
    save_state(state_path, state)
    if verbose:
        log_json({"stage": "report_refreshed", "period": current_period, "cutoff_date": cutoff_date.isoformat(), "report_path": str(report_path)})


def wait_until_telegram_ready(token: str, verbose: bool, startup_timeout_seconds: int = 0) -> Dict[str, Any]:
    outage_started_at: float | None = None
    last_logged_at: float | None = None
    started_at = time.time()
    last_exception: Exception | None = None

    while True:
        try:
            result = api_request(token, "getMe")
            if verbose and outage_started_at is not None:
                log_json(
                    {
                        "stage": "startup_network_recovered",
                        "outage_started_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(outage_started_at)),
                        "outage_duration_seconds": int(time.time() - outage_started_at),
                    }
                )
            return result
        except urllib.error.HTTPError as exc:
            last_exception = exc
            if exc.code == 401:
                raise RuntimeError("Unauthorized bot token") from exc
            outage_started_at, last_logged_at = maybe_log_network_outage(verbose, outage_started_at, last_logged_at, "startup_network_error", exc)
        except urllib.error.URLError as exc:
            last_exception = exc
            outage_started_at, last_logged_at = maybe_log_network_outage(verbose, outage_started_at, last_logged_at, "startup_network_error", exc)
        except Exception as exc:
            last_exception = exc
            outage_started_at, last_logged_at = maybe_log_network_outage(verbose, outage_started_at, last_logged_at, "startup_network_error", exc)

        if startup_timeout_seconds > 0 and (time.time() - started_at) >= startup_timeout_seconds:
            if last_exception is None:
                raise TimeoutError(f"Telegram API still unavailable after {startup_timeout_seconds} seconds")
            raise TimeoutError(
                f"Telegram API still unavailable after {startup_timeout_seconds} seconds; "
                f"last error: {type(last_exception).__name__}: {last_exception}"
            ) from last_exception
        time.sleep(RETRY_SLEEP_SECONDS)


def load_token(state_path: Path, legacy_token_path: Path | None = None) -> str:
    state = load_state(state_path)
    token = str(state.get("token", "")).strip()
    if token:
        return token

    if legacy_token_path is not None and legacy_token_path.exists():
        legacy_token = legacy_token_path.read_text(encoding="utf-8").strip()
        if legacy_token:
            state["token"] = legacy_token
            save_state(state_path, state)
            return legacy_token

    raise ValueError("telegram_bot_config.json 中缺少 token")


def queue_restart_confirmation(state_path: Path, chat_id: int, reply_to_message_id: int) -> None:
    queue_pending_reply(
        state_path,
        chat_id,
        reply_to_message_id,
        "机器人已重启完成",
        stage="restart_reply_failed",
    )


def build_message_reference(message: Dict[str, Any] | None) -> Dict[str, Any] | None:
    if not isinstance(message, dict):
        return None
    chat = message.get("chat", {}) or {}
    return {
        "message_id": message.get("message_id"),
        "chat_id": chat.get("id"),
        "sender": (message.get("from", {}) or {}).get("username") or (message.get("from", {}) or {}).get("first_name") or "",
        "telegram_timestamp": message.get("date"),
        "text": message.get("text") or "",
    }


def build_envelope(message: Dict[str, Any]) -> Dict[str, Any]:
    sender = message.get("from", {})
    chat = message.get("chat", {})
    return {
        "message_id": message.get("message_id"),
        "chat_id": chat.get("id"),
        "sender": sender.get("username") or sender.get("first_name") or "",
        "telegram_timestamp": message.get("date"),
        "text": message.get("text") or "",
        "reply_to_message": build_message_reference(message.get("reply_to_message")),
    }


def should_process(message: Dict[str, Any], allowed_username: str) -> tuple[bool, str]:
    sender = message.get("from", {}) or {}
    username = sender.get("username")
    if not allowed_username:
        return False, "whitelist username is not configured"
    if username != allowed_username:
        return False, f"忽略非白名单用户: {username!r}"

    text = (message.get("text") or "").strip()
    if not text:
        return False, "忽略非文本消息"
    return True, ""


def message_index_key(chat_id: int, message_id: int) -> str:
    return f"{chat_id}:{message_id}"


def normalize_index_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        number = float(value)
        return int(number) if number.is_integer() else number
    return str(value).strip()


def build_record_fingerprint(record: Dict[str, Any]) -> str:
    stable_record = {field: normalize_index_value(record.get(field)) for field in EXPECTED_HEADERS}
    return json.dumps(stable_record, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def register_record_mapping(envelope: Dict[str, Any], result: Dict[str, Any]) -> None:
    chat_id = envelope.get("chat_id")
    message_id = envelope.get("message_id")
    record = result.get("record")
    if not isinstance(chat_id, int) or not isinstance(message_id, int) or not isinstance(record, dict):
        return

    index = load_message_index()
    index[message_index_key(chat_id, message_id)] = {
        "chat_id": chat_id,
        "message_id": message_id,
        "record_id": record.get("ID"),
        "sheet_name": result.get("sheet_name"),
        "record_fingerprint": build_record_fingerprint(record),
        "voided": False,
        "created_at": int(time.time()),
    }
    save_message_index(index)


def _resolve_invalidate_target(index: Dict[str, Any], envelope: Dict[str, Any]) -> tuple[str, Dict[str, Any]]:
    chat_id = envelope.get("chat_id")
    current_message_id = envelope.get("message_id")
    if not isinstance(chat_id, int) or not isinstance(current_message_id, int):
        raise ValueError("当前消息缺少有效的 chat_id 或 message_id")

    reply_to_message = envelope.get("reply_to_message")
    if isinstance(reply_to_message, dict):
        target_chat_id = reply_to_message.get("chat_id")
        target_message_id = reply_to_message.get("message_id")
        if not isinstance(target_chat_id, int) or not isinstance(target_message_id, int):
            raise ValueError("引用消息无效")
        key = message_index_key(target_chat_id, target_message_id)
        entry = index.get(key)
        if not isinstance(entry, dict):
            raise ValueError("这条消息没有对应账本记录")
        return key, entry

    candidates: list[tuple[int, str, Dict[str, Any]]] = []
    for key, entry in index.items():
        if not isinstance(entry, dict):
            continue
        if entry.get("chat_id") != chat_id:
            continue
        message_id = entry.get("message_id")
        if not isinstance(message_id, int):
            continue
        if message_id >= current_message_id:
            continue
        candidates.append((message_id, key, entry))

    if not candidates:
        raise ValueError("当前消息之前没有可作废的账本记录")

    _, key, entry = max(candidates, key=lambda item: item[0])
    return key, entry


def invalidate_target_record(envelope: Dict[str, Any], excel_path: Path, backend: str) -> Dict[str, Any]:
    index = load_message_index()
    key, entry = _resolve_invalidate_target(index, envelope)
    if entry.get("voided") is True:
        raise ValueError("该记录已作废")

    record_id = entry.get("record_id")
    sheet_name = entry.get("sheet_name")
    if not record_id:
        raise ValueError("索引中缺少 Record ID，无法按 ID 定位")

    current_record = read_record_by_id(excel_path, record_id, sheet_name=sheet_name, backend=backend)
    if str(current_record.get("Status", "")).strip() == VOID_MARK:
        entry["voided"] = True
        save_message_index(index)
        raise ValueError("该记录已作废")

    invalidated_row = invalidate_record_by_id(excel_path, record_id, sheet_name=sheet_name, backend=backend)
    entry["voided"] = True
    entry["voided_at"] = int(time.time())
    save_message_index(index)
    return {"row": invalidated_row, "sheet_name": sheet_name, "target_record_id": record_id, "target_message_id": entry.get("message_id")}


def run_bridge_prompt(workdir: Path, envelope: Dict[str, Any]) -> str:
    output = bridge_emit_prompt(dict(envelope)).strip()
    if not output:
        raise RuntimeError("bridge prompt 未返回内容")
    return output


def run_codex(workdir: Path, prompt: str) -> Dict[str, Any]:
    env = os.environ.copy()
    for key in ["GEMINI_CLI_IDE_SERVER_PORT", "GEMINI_CLI_IDE_AUTH_TOKEN"]:
        env.pop(key, None)

    max_retries = 3
    last_error = ""
    for attempt in range(max_retries):
        with tempfile.TemporaryDirectory(prefix="codex-expense-") as temp_dir:
            temp_path = Path(temp_dir)
            schema_path = temp_path / "codex_output_schema.json"
            output_path = temp_path / "codex_output.json"
            schema_path.write_text(json.dumps(CODEX_OUTPUT_SCHEMA, ensure_ascii=False), encoding="utf-8")

            completed = subprocess.run(
                [
                    "codex",
                    "exec",
                    "--skip-git-repo-check",
                    "--sandbox",
                    "read-only",
                    "--color",
                    "never",
                    "--output-schema",
                    str(schema_path),
                    "-o",
                    str(output_path),
                    "-",
                ],
                cwd=workdir,
                capture_output=True,
                text=True,
                env=env,
                input=prompt,
            )
            if completed.returncode == 0:
                raw_output = output_path.read_text(encoding="utf-8").strip() if output_path.exists() else ""
                if raw_output:
                    try:
                        return json.loads(raw_output)
                    except Exception as exc:
                        last_error = f"Parse failed: {exc}"
                else:
                    last_error = "codex output is empty"
            else:
                last_error = (completed.stderr or completed.stdout).strip() or "codex failed"

        if attempt < max_retries - 1:
            time.sleep(2 ** attempt)
    raise RuntimeError(last_error)


def run_bridge_apply(workdir: Path, envelope: Dict[str, Any], codex_output: Dict[str, Any], backend: str, excel_path: Path) -> Dict[str, Any]:
    payload = dict(envelope)
    payload["codex_output"] = codex_output
    return bridge_apply_record(payload, str(excel_path), None, backend, False)


def send_reply(
    token: str,
    chat_id: int,
    reply_to_message_id: int,
    text: str,
    parse_mode: str | None = None,
    timeout: int = SEND_REPLY_TIMEOUT_SECONDS,
) -> Dict[str, Any]:
    payload = {"chat_id": chat_id, "text": text, "reply_to_message_id": reply_to_message_id}
    if parse_mode:
        payload["parse_mode"] = parse_mode
    return api_request(token, "sendMessage", payload, timeout=timeout)


def send_reply_with_retry(
    token: str,
    chat_id: int,
    reply_to_message_id: int,
    text: str,
    parse_mode: str | None = None,
    verbose: bool = False,
    stage: str = "reply_network_error",
) -> Dict[str, Any]:
    last_exc: Exception | None = None
    for attempt in range(1, SEND_RETRY_ATTEMPTS + 1):
        try:
            return send_reply(token, chat_id, reply_to_message_id, text, parse_mode=parse_mode)
        except Exception as exc:
            last_exc = exc
            if verbose and attempt < SEND_RETRY_ATTEMPTS:
                log_json(
                    {
                        "stage": "reply_network_error",
                        "delivery_stage": stage,
                        "chat_id": chat_id,
                        "message_id": reply_to_message_id,
                        "attempt": attempt,
                        "error": str(exc),
                        "will_retry": True,
                    }
                )
            if attempt < SEND_RETRY_ATTEMPTS:
                time.sleep(SEND_RETRY_BASE_SLEEP_SECONDS ** (attempt - 1))

    if last_exc is None:
        raise RuntimeError("send_reply_with_retry failed without exception")
    raise last_exc


def deliver_reply(
    runtime: RuntimeContext,
    chat_id: int,
    reply_to_message_id: int,
    text: str,
    stage: str,
    parse_mode: str | None = None,
) -> bool:
    try:
        send_reply_with_retry(
            runtime.token,
            chat_id,
            reply_to_message_id,
            text,
            parse_mode=parse_mode,
            verbose=runtime.verbose,
            stage=stage,
        )
        if runtime.verbose:
            log_json(
                {
                    "stage": "message_reply_sent",
                    "delivery_stage": stage,
                    "chat_id": chat_id,
                    "message_id": reply_to_message_id,
                }
            )
        return True
    except Exception as exc:
        queue_pending_reply(
            runtime.state_path,
            chat_id,
            reply_to_message_id,
            text,
            stage=stage,
            parse_mode=parse_mode,
            error=str(exc),
        )
        if runtime.verbose:
            log_json(
                {
                    "stage": "message_reply_failed",
                    "delivery_stage": stage,
                    "chat_id": chat_id,
                    "message_id": reply_to_message_id,
                    "error": str(exc),
                    "queued_for_retry": True,
                }
            )
        return False


def flush_pending_replies(runtime: RuntimeContext, limit: int = PENDING_REPLY_BATCH_SIZE) -> None:
    migrate_legacy_pending_restart_notice(runtime.state_path)
    state = load_state(runtime.state_path)
    queue = _state_reply_queue(state)
    if not queue:
        return

    remaining: list[Dict[str, Any]] = []
    sent_count = 0
    changed = False

    for entry in queue:
        if sent_count >= limit:
            remaining.append(entry)
            continue

        chat_id = entry.get("chat_id")
        reply_to_message_id = entry.get("reply_to_message_id")
        text = entry.get("text")
        parse_mode = entry.get("parse_mode") or None
        stage = str(entry.get("stage") or "reply_network_error")
        attempts = int(entry.get("attempts", 0))

        if not isinstance(chat_id, int) or not isinstance(reply_to_message_id, int) or not isinstance(text, str) or not text:
            changed = True
            continue

        try:
            send_reply_with_retry(
                runtime.token,
                chat_id,
                reply_to_message_id,
                text,
                parse_mode=parse_mode,
                verbose=runtime.verbose,
                stage=stage,
            )
            changed = True
            sent_count += 1
            if runtime.verbose:
                log_json(
                    {
                        "stage": "message_reply_sent",
                        "delivery_stage": stage,
                        "chat_id": chat_id,
                        "message_id": reply_to_message_id,
                        "attempts": attempts + 1,
                        "from_pending_queue": True,
                    }
                )
        except Exception as exc:
            updated = dict(entry)
            updated["attempts"] = attempts + 1
            updated["last_error"] = str(exc)
            updated["last_attempt_at"] = int(time.time())
            remaining.append(updated)
            if runtime.verbose:
                log_json(
                    {
                        "stage": "reply_network_error",
                        "delivery_stage": stage,
                        "chat_id": chat_id,
                        "message_id": reply_to_message_id,
                        "attempts": updated["attempts"],
                        "error": str(exc),
                        "from_pending_queue": True,
                    }
                )

    if changed or len(remaining) != len(queue):
        state["pending_replies"] = remaining
        save_state(runtime.state_path, state)


def trigger_bot_restart(verbose: bool) -> None:
    subprocess.Popen(["bash", str(RESTART_SCRIPT_PATH)], cwd=BASE_DIR, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL, start_new_session=True)
    if verbose: log_json({"stage": "restart_triggered"})


def get_fallback_record(envelope: Dict[str, Any], bot_config: Dict[str, Any] | None = None) -> Dict[str, Any]:
    ts = envelope.get("telegram_timestamp")
    fallback_dt = convert_telegram_timestamp(ts, DEFAULT_TIMEZONE)
    date_str, time_str = fallback_dt.strftime("%Y-%m-%d"), fallback_dt.strftime("%H:%M")

    return {
        "ID": f"{envelope.get('chat_id')}:{envelope.get('message_id')}",
        "Date": date_str,
        "Time": time_str,
        "Timezone": DEFAULT_TIMEZONE,
        "Amount": 0,
        "Currency": "CNY",
        "Type": "支出",
        "Category": "未分类",
        "Note": f"[AI失败兜底] {envelope['text']}",
        "PaymentChannel": get_default_payment_channel(bot_config),
        "Status": "待确认",
    }


def handle_restart_command(runtime: RuntimeContext, envelope: Dict[str, Any]) -> None:
    queue_restart_confirmation(runtime.state_path, envelope["chat_id"], envelope["message_id"])
    trigger_bot_restart(runtime.verbose)


def handle_invalidate_command(runtime: RuntimeContext, envelope: Dict[str, Any]) -> None:
    try:
        result = invalidate_target_record(envelope, runtime.excel_path, runtime.backend)
        if runtime.verbose:
            log_message_event(
                "message_applied",
                envelope,
                operation="invalidate",
                sheet_name=result.get("sheet_name"),
                row=result.get("row"),
                target_record_id=result.get("target_record_id"),
            )
        deliver_reply(
            runtime,
            envelope["chat_id"],
            envelope["message_id"],
            f"已作废：第 {result['row']} 行（{result['sheet_name']}）",
            "invalidate_reply_failed",
        )
    except Exception as exc:
        if runtime.verbose:
            log_json({"stage": "invalidate_error", "message_id": envelope["message_id"], "error": str(exc)})
        deliver_reply(
            runtime,
            envelope["chat_id"],
            envelope["message_id"],
            f"作废失败：{exc}",
            "invalidate_error_reply_failed",
        )


def apply_fallback_record(runtime: RuntimeContext, envelope: Dict[str, Any], ai_exc: Exception) -> tuple[Dict[str, Any], bool]:
    if runtime.verbose:
        log_message_event("message_processing_fallback", envelope, error=str(ai_exc))
    record = get_fallback_record(envelope, runtime.bot_config)
    normalized = normalize_record(record)
    sheet_name = record["Date"].split("-")[0]
    row = append_record_to_excel(runtime.excel_path, normalized, sheet_name, runtime.backend)
    return {"ok": True, "record": normalized, "sheet_name": sheet_name, "row": row, "fallback": True}, True


def apply_bookkeeping_message(runtime: RuntimeContext, envelope: Dict[str, Any]) -> tuple[Dict[str, Any], bool]:
    try:
        payload = dict(envelope)
        payload["runtime_config"] = runtime.bot_config
        prompt = run_bridge_prompt(runtime.workdir, payload)
        codex_output = run_codex(runtime.workdir, prompt)
        return run_bridge_apply(runtime.workdir, payload, codex_output, runtime.backend, runtime.excel_path), False
    except Exception as ai_exc:
        return apply_fallback_record(runtime, envelope, ai_exc)


def build_success_reply(apply_result: Dict[str, Any], fallback_used: bool) -> str:
    record = apply_result["record"]
    if fallback_used:
        return (
            "⚠️ AI 处理失败，已为您自动记录原文：\n"
            f"{record['Type']} / {record['Category']} / {record['Amount']}\n"
            f"备注：{record['Note']}\n"
            f"请稍后手动核对（第 {apply_result['row']} 行）"
        )

    confirm_hint = "，待确认" if record.get("Status") == STATUS_PENDING else ""
    reply = f"已记账：{record['Type']} / {record['Category']} / {record['Amount']}"
    if record["Note"]:
        reply += f" / {record['Note']}"
    reply += f"，第 {apply_result['row']} 行{confirm_hint}"
    return reply


def handle_bookkeeping_message(runtime: RuntimeContext, envelope: Dict[str, Any]) -> None:
    apply_result, fallback_used = apply_bookkeeping_message(runtime, envelope)
    if runtime.verbose:
        log_message_event(
            "message_applied",
            envelope,
            operation="bookkeeping",
            result=apply_result,
            fallback_used=fallback_used,
        )
    if apply_result.get("ignored"):
        deliver_reply(
            runtime,
            envelope["chat_id"],
            envelope["message_id"],
            f"已忽略：{apply_result.get('reason', '不是记账相关消息')}",
            "ignored_reply_failed",
        )
        return

    try:
        register_record_mapping(envelope, apply_result)
    except Exception as idx_exc:
        if runtime.verbose:
            log_json({"stage": "index_error", "message_id": envelope["message_id"], "error": str(idx_exc)})

    reply = build_success_reply(apply_result, fallback_used)
    deliver_reply(runtime, envelope["chat_id"], envelope["message_id"], reply, "reply_network_error")


def handle_message(runtime: RuntimeContext, message: Dict[str, Any]) -> None:
    envelope = build_envelope(message)
    operation = "bookkeeping"
    if envelope["text"] == "重启":
        operation = "restart"
    elif envelope["text"] == "作废":
        operation = "invalidate"
    if runtime.verbose:
        log_message_event("message_processing_started", envelope, operation=operation)

    if envelope["text"] == "重启":
        handle_restart_command(runtime, envelope)
        return

    if envelope["text"] == "作废":
        handle_invalidate_command(runtime, envelope)
        return

    try:
        handle_bookkeeping_message(runtime, envelope)
    except Exception as exc:
        if runtime.verbose:
            log_message_event("message_processing_failed", envelope, error=str(exc), operation=operation)
        deliver_reply(
            runtime,
            envelope["chat_id"],
            envelope["message_id"],
            f"系统故障，无法记账：{exc}",
            "critical_error_reply_failed",
        )


def save_offset(state_path: Path, offset: int) -> None:
    state = load_state(state_path)
    state["offset"] = offset
    save_state(state_path, state)


def poll_loop(runtime: RuntimeContext) -> None:
    state = load_state(runtime.state_path)
    offset = int(state.get("offset", 0))
    outage_started_at: float | None = None
    last_logged_at: float | None = None
    while True:
        try:
            refresh_report_if_period_changed(runtime.state_path, runtime.excel_path, runtime.verbose)
            flush_pending_replies(runtime)
            result = api_request(runtime.token, "getUpdates", {"offset": offset, "timeout": API_TIMEOUT_SECONDS, "allowed_updates": ["message"]}, timeout=API_TIMEOUT_SECONDS + 10)
            if outage_started_at is not None and runtime.verbose:
                log_json(
                    {
                        "stage": "poll_network_recovered",
                        "outage_started_at": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(outage_started_at)),
                        "outage_duration_seconds": int(time.time() - outage_started_at),
                    }
                )
            outage_started_at = None
            last_logged_at = None
            for update in result.get("result", []):
                update_id = int(update["update_id"])
                offset = update_id + 1
                save_offset(runtime.state_path, offset)
                message = update.get("message")
                if not message:
                    continue
                accepted, reason = should_process(message, runtime.allowed_username)
                if runtime.verbose:
                    log_message_event(
                        "message_received",
                        build_envelope(message),
                        accepted=accepted,
                        reason=reason,
                    )
                if not accepted:
                    continue
                try:
                    handle_message(runtime, message)
                except Exception as exc:
                    log_json({"stage": "handle_message_crash", "error": str(exc)})
        except urllib.error.HTTPError as exc:
            if exc.code == 401:
                raise RuntimeError("Unauthorized bot token") from exc
            outage_started_at, last_logged_at = maybe_log_network_outage(runtime.verbose, outage_started_at, last_logged_at, "poll_network_error", exc)
            time.sleep(RETRY_SLEEP_SECONDS)
        except urllib.error.URLError as exc:
            outage_started_at, last_logged_at = maybe_log_network_outage(runtime.verbose, outage_started_at, last_logged_at, "poll_network_error", exc)
            time.sleep(RETRY_SLEEP_SECONDS)
        except Exception as exc:
            outage_started_at, last_logged_at = maybe_log_network_outage(runtime.verbose, outage_started_at, last_logged_at, "poll_loop_error", exc)
            time.sleep(RETRY_SLEEP_SECONDS)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--state-file", default=str(BASE_DIR / CONFIG_FILE_NAME))
    parser.add_argument("--legacy-token-file", default=str(BASE_DIR / "bot_token.txt"))
    parser.add_argument("--excel-path", default="")
    parser.add_argument("--backend", choices=["win32com", "openpyxl"], default="openpyxl")
    parser.add_argument("--once", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    parser.add_argument("--startup-timeout-seconds", type=int, default=0)
    args = parser.parse_args()
    workdir = BASE_DIR
    state_path = Path(args.state_file)
    state = load_state(state_path)
    bot_config = load_bot_config(state_path)
    excel_path = Path(args.excel_path) if str(args.excel_path).strip() else configured_excel_path(state)
    token = load_token(state_path, Path(args.legacy_token_file))
    me = wait_until_telegram_ready(token, args.verbose, startup_timeout_seconds=args.startup_timeout_seconds)
    runtime = RuntimeContext(
        token=token,
        workdir=workdir,
        state_path=state_path,
        allowed_username=configured_allowed_username(state),
        bot_config=bot_config,
        backend=args.backend,
        excel_path=excel_path,
        verbose=args.verbose,
    )
    if args.verbose:
        log_json({"stage": "startup_ready", "result": me["result"]})
    if args.once:
        return 0
    migrate_legacy_pending_restart_notice(state_path)
    flush_pending_replies(runtime)
    refresh_report_if_period_changed(state_path, excel_path, args.verbose)
    poll_loop(runtime)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
