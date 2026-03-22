from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict

from append_excel_entry import (
    EXPECTED_HEADERS,
    STATUS_PENDING,
    append_record_to_excel,
    invalidate_last_record_in_excel,
    invalidate_record_by_id,
    normalize_record,
    read_record_by_id,
)
from generate_expense_report import refresh_report_workbook

API_TIMEOUT_SECONDS = 60
BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
RESTART_SCRIPT_PATH = BASE_DIR / "restart_bot.sh"
VOID_MARK = "作废"
CONFIG_FILE_NAME = "telegram_bot_state.json"


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
    payload["timestamp"] = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print(json.dumps(payload, ensure_ascii=False), flush=True)
    log_file = get_monthly_file_path("logs", "log")
    try:
        with log_file.open("a", encoding="utf-8") as f:
            f.write(json.dumps(payload, ensure_ascii=False) + "\n")
    except:
        pass


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


def refresh_report_if_period_changed(state_path: Path, excel_path: Path, verbose: bool) -> None:
    state = load_state(state_path)
    current_period = current_report_period()
    last_period = str(state.get("report_period", "")).strip()
    report_path = excel_path.with_name("expense_report.xlsx")

    if last_period == current_period and report_path.exists():
        return

    refresh_report_workbook(excel_path, report_path)
    state["report_period"] = current_period
    save_state(state_path, state)
    if verbose:
        log_json({"stage": "report_refreshed", "period": current_period, "report_path": str(report_path)})


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

    raise ValueError("telegram_bot_state.json 中缺少 token")


def queue_restart_confirmation(state_path: Path, chat_id: int, reply_to_message_id: int) -> None:
    state = load_state(state_path)
    state["pending_restart_notice"] = {
        "chat_id": chat_id,
        "reply_to_message_id": reply_to_message_id,
        "created_at": int(time.time()),
    }
    save_state(state_path, state)


def send_pending_restart_confirmation(token: str, state_path: Path, verbose: bool) -> None:
    state = load_state(state_path)
    notice = state.get("pending_restart_notice")
    if not isinstance(notice, dict):
        return

    chat_id = notice.get("chat_id")
    reply_to_message_id = notice.get("reply_to_message_id")
    if not isinstance(chat_id, int) or not isinstance(reply_to_message_id, int):
        state.pop("pending_restart_notice", None)
        save_state(state_path, state)
        return

    try:
        send_reply(token, chat_id, reply_to_message_id, "机器人已重启完成")
        state.pop("pending_restart_notice", None)
        save_state(state_path, state)
        if verbose:
            log_json({"stage": "restart_confirmed", "chat_id": chat_id, "reply_to_message_id": reply_to_message_id})
    except Exception as exc:
        if verbose:
            log_json({"stage": "restart_ack_failed", "message_id": reply_to_message_id, "chat_id": chat_id, "error": str(exc)})


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


def invalidate_reply_target(envelope: Dict[str, Any], excel_path: Path, backend: str) -> Dict[str, Any]:
    reply_to_message = envelope.get("reply_to_message")
    if not isinstance(reply_to_message, dict):
        raise ValueError("当前消息没有引用历史消息")

    chat_id = reply_to_message.get("chat_id")
    message_id = reply_to_message.get("message_id")
    if not isinstance(chat_id, int) or not isinstance(message_id, int):
        raise ValueError("引用消息无效")

    index = load_message_index()
    key = message_index_key(chat_id, message_id)
    entry = index.get(key)
    if not isinstance(entry, dict):
        raise ValueError("这条消息没有对应账本记录")
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
    return {"row": invalidated_row, "sheet_name": sheet_name}


def run_bridge_prompt(workdir: Path, envelope: Dict[str, Any]) -> str:
    completed = subprocess.run(
        ["python3", str(workdir / "telegram_codex_bridge.py"), "prompt", "--json", json.dumps(envelope, ensure_ascii=False)],
        cwd=workdir, capture_output=True, text=True, check=True
    )
    output = completed.stdout.strip()
    if not output:
        raise RuntimeError("bridge prompt 未返回内容")
    return output


def run_gemini(workdir: Path, prompt: str) -> Dict[str, Any]:
    env = os.environ.copy()
    for key in ["GEMINI_CLI_IDE_SERVER_PORT", "GEMINI_CLI_IDE_AUTH_TOKEN"]:
        env.pop(key, None)

    max_retries = 3
    last_error = ""
    for attempt in range(max_retries):
        completed = subprocess.run(
            ["gemini", "--prompt", prompt, "--output-format", "json"],
            cwd=workdir, capture_output=True, text=True, env=env
        )
        if completed.returncode == 0:
            try:
                outer_json = json.loads(completed.stdout.strip())
                raw_response = outer_json.get("response", "").strip()
                if raw_response.startswith("```json"): raw_response = raw_response[7:].strip()
                elif raw_response.startswith("```"): raw_response = raw_response[3:].strip()
                if raw_response.endswith("```"): raw_response = raw_response[:-3].strip()
                return json.loads(raw_response)
            except Exception as e:
                import re
                match = re.search(r"(\{.*\})", raw_response, re.DOTALL)
                if match:
                    try: return json.loads(match.group(1))
                    except: pass
                last_error = f"Parse failed: {e}"
        else:
            last_error = (completed.stderr or completed.stdout).strip() or "gemini failed"
        
        if attempt < max_retries - 1:
            time.sleep(2 ** attempt)
    raise RuntimeError(last_error)


def run_bridge_apply(workdir: Path, envelope: Dict[str, Any], gemini_output: Dict[str, Any], backend: str, excel_path: Path) -> Dict[str, Any]:
    payload = dict(envelope)
    payload["gemini_output"] = gemini_output
    completed = subprocess.run(
        ["python3", str(workdir / "telegram_codex_bridge.py"), "apply", "--backend", backend, "--excel-path", str(excel_path), "--json", json.dumps(payload, ensure_ascii=False)],
        cwd=workdir, capture_output=True, text=True, check=True
    )
    return json.loads(completed.stdout.strip())


def send_reply(token: str, chat_id: int, reply_to_message_id: int, text: str, parse_mode: str | None = None) -> Dict[str, Any]:
    payload = {"chat_id": chat_id, "text": text, "reply_to_message_id": reply_to_message_id}
    if parse_mode: payload["parse_mode"] = parse_mode
    return api_request(token, "sendMessage", payload)


def trigger_bot_restart(verbose: bool) -> None:
    subprocess.Popen(["bash", str(RESTART_SCRIPT_PATH)], cwd=BASE_DIR, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, stdin=subprocess.DEVNULL, start_new_session=True)
    if verbose: log_json({"stage": "restart_triggered"})


def get_fallback_record(envelope: Dict[str, Any]) -> Dict[str, Any]:
    ts = envelope.get("telegram_timestamp")
    if ts:
        local_tm = time.localtime(float(ts))
        date_str, time_str = time.strftime("%Y-%m-%d", local_tm), time.strftime("%H:%M", local_tm)
    else:
        date_str, time_str = time.strftime("%Y-%m-%d"), time.strftime("%H:%M")

    return {
        "ID": f"{envelope.get('chat_id')}:{envelope.get('message_id')}",
        "Date": date_str,
        "Time": time_str,
        "Amount": 0,
        "Currency": "CNY",
        "Type": "支出",
        "Category": "未分类",
        "Note": f"[AI失败兜底] {envelope['text']}",
        "Status": "待确认",
    }


def handle_message(token: str, workdir: Path, state_path: Path, message: Dict[str, Any], backend: str, excel_path: Path, verbose: bool) -> None:
    envelope = build_envelope(message)
    if envelope["text"] == "重启":
        queue_restart_confirmation(state_path, envelope["chat_id"], envelope["message_id"])
        try: send_reply(token, envelope["chat_id"], envelope["message_id"], "正在重启机器人")
        except: pass
        trigger_bot_restart(verbose)
        return

    if envelope["text"] == "作废":
        try:
            if envelope.get("reply_to_message"):
                result = invalidate_reply_target(envelope, excel_path, backend)
                send_reply(token, envelope["chat_id"], envelope["message_id"], f"已作废对应消息的记录：第 {result['row']} 行（{result['sheet_name']}）")
            else:
                row = invalidate_last_record_in_excel(excel_path, backend=backend)
                send_reply(token, envelope["chat_id"], envelope["message_id"], f"已作废：第 {row} 行")
        except Exception as exc:
            if verbose: log_json({"stage": "invalidate_error", "message_id": envelope["message_id"], "error": str(exc)})
            try: send_reply(token, envelope["chat_id"], envelope["message_id"], f"作废失败：{exc}")
            except: pass
        return

    try: send_reply(token, envelope["chat_id"], envelope["message_id"], "正在处理...")
    except: pass

    apply_result = None
    fallback_used = False
    try:
        try:
            prompt = run_bridge_prompt(workdir, envelope)
            gemini_output = run_gemini(workdir, prompt)
            apply_result = run_bridge_apply(workdir, envelope, gemini_output, backend, excel_path)
        except Exception as ai_exc:
            if verbose: log_json({"stage": "ai_failed_using_fallback", "message_id": envelope["message_id"], "error": str(ai_exc)})
            record = get_fallback_record(envelope)
            normalized = normalize_record(record)
            sheet_name = record["Date"].split("-")[0]
            row = append_record_to_excel(excel_path, normalized, sheet_name, backend)
            apply_result = {"ok": True, "record": normalized, "sheet_name": sheet_name, "row": row, "fallback": True}
            fallback_used = True

        if verbose: log_json({"stage": "applied", "message_id": envelope["message_id"], "result": apply_result})
        if apply_result.get("ignored"):
            try: send_reply(token, envelope["chat_id"], envelope["message_id"], f"已忽略：{apply_result.get('reason', '不是记账相关消息')}")
            except: pass
            return

        try: register_record_mapping(envelope, apply_result)
        except Exception as idx_exc:
            if verbose: log_json({"stage": "index_error", "message_id": envelope["message_id"], "error": str(idx_exc)})

        r = apply_result["record"]
        if fallback_used:
            reply = f"⚠️ AI 处理失败，已为您自动记录原文：\n{r['Type']} / {r['Category']} / {r['Amount']}\n备注：{r['Note']}\n请稍后手动核对（第 {apply_result['row']} 行）"
        else:
            confirm_hint = "，待确认" if r.get("Status") == STATUS_PENDING else ""
            reply = f"已记账：{r['Type']} / {r['Category']} / {r['Amount']}"
            if r['Note']: reply += f" / {r['Note']}"
            reply += f"，第 {apply_result['row']} 行{confirm_hint}"
        
        try: send_reply(token, envelope["chat_id"], envelope["message_id"], reply)
        except Exception as reply_exc:
            if verbose: log_json({"stage": "reply_network_error", "message_id": envelope["message_id"], "error": str(reply_exc)})
    except Exception as exc:
        if verbose: log_json({"stage": "critical_error", "message_id": envelope.get("message_id"), "error": str(exc)})
        try: send_reply(token, envelope["chat_id"], envelope["message_id"], f"系统故障，无法记账：{exc}")
        except: pass


def poll_loop(token: str, workdir: Path, state_path: Path, allowed_username: str, backend: str, excel_path: Path, verbose: bool) -> None:
    state = load_state(state_path)
    offset = int(state.get("offset", 0))
    while True:
        try:
            refresh_report_if_period_changed(state_path, excel_path, verbose)
            state = load_state(state_path)
            state["offset"] = offset
            result = api_request(token, "getUpdates", {"offset": offset, "timeout": API_TIMEOUT_SECONDS, "allowed_updates": ["message"]}, timeout=API_TIMEOUT_SECONDS + 10)
            for update in result.get("result", []):
                update_id = int(update["update_id"])
                offset = update_id + 1
                state = load_state(state_path)
                state["offset"] = offset
                save_state(state_path, state)
                message = update.get("message")
                if not message: continue
                accepted, reason = should_process(message, allowed_username)
                if verbose: log_json({"stage": "received", "accepted": accepted, "reason": reason, "text": (message.get("text") or "")[:40]})
                if not accepted: continue
                try: handle_message(token, workdir, state_path, message, backend, excel_path, verbose)
                except Exception as exc: log_json({"stage": "handle_message_crash", "error": str(exc)})
        except urllib.error.HTTPError as exc:
            if exc.code == 401: raise RuntimeError("Unauthorized bot token") from exc
            if verbose: log_json({"stage": "network_error", "error_type": classify_network_error(exc), "error": str(exc)})
            time.sleep(5)
        except urllib.error.URLError as exc:
            if verbose: log_json({"stage": "network_error", "error_type": classify_network_error(exc), "error": str(exc)})
            time.sleep(5)
        except Exception as exc:
            if verbose: log_json({"stage": "loop_error", "error_type": classify_network_error(exc), "error": str(exc)})
            time.sleep(5)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--state-file", default=str(BASE_DIR / CONFIG_FILE_NAME))
    parser.add_argument("--legacy-token-file", default=str(BASE_DIR / "bot_token.txt"))
    parser.add_argument("--excel-path", default="")
    parser.add_argument("--backend", choices=["win32com", "openpyxl"], default="openpyxl")
    parser.add_argument("--once", action="store_true")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()
    workdir = BASE_DIR
    state_path = Path(args.state_file)
    state = load_state(state_path)
    excel_path = Path(args.excel_path) if str(args.excel_path).strip() else configured_excel_path(state)
    token = load_token(state_path, Path(args.legacy_token_file))
    me = api_request(token, "getMe")
    if args.verbose: log_json({"stage": "getMe", "result": me["result"]})
    send_pending_restart_confirmation(token, state_path, args.verbose)
    if args.once: return 0
    refresh_report_if_period_changed(state_path, excel_path, args.verbose)
    poll_loop(token, workdir, state_path, configured_allowed_username(state), args.backend, excel_path, args.verbose)
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
