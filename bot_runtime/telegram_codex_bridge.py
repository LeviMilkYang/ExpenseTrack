from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from append_excel_entry import (
    DEFAULT_TIMEZONE,
    convert_telegram_timestamp,
    get_allowed_categories,
    get_default_payment_channel,
    get_payment_channels,
    normalize_record,
    normalize_timezone,
)
from excel_tools import run_tool_payload

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
DEFAULT_EXCEL_PATH = PROJECT_DIR / "expense.xlsx"

PROMPT_OUTPUT_SHAPE: Dict[str, Any] = {
    "ignored": "boolean",
    "reason": "string",
    "tool_call": "object|null",
}

PROMPT_RECORD_SHAPE: Dict[str, Any] = {
    "Date": "YYYY-MM-DD",
    "Time": "HH:MM",
    "Timezone": DEFAULT_TIMEZONE,
    "DateProvided": "boolean",
    "TimeProvided": "boolean",
    "Amount": "number",
    "Currency": "CNY",
    "Type": "allowed type value",
    "Category": "string",
    "Note": "string",
    "PaymentChannel": "configured payment channel",
    "Status": "",
}

PROMPT_TOOL_CALL_SHAPE: Dict[str, Any] = {
    "tool": "append_record",
    "arguments": {
        "record": PROMPT_RECORD_SHAPE,
        "sheet_name": "string|null",
    },
}

PROMPT_TEMPLATE = """You convert one Telegram bookkeeping message into exactly one JSON object.

Rules:
1. Output JSON only. No markdown. No explanation.
2. Always output this exact top-level shape:
{output_shape}
3. If the message is not about income or expense bookkeeping, set:
{{"ignored":true,"reason":"short reason","tool_call":null}}
4. Otherwise set `"ignored": false`, `"reason": ""`, and make `tool_call` use this schema:
{tool_call_shape}
5. `Timezone` must always be present. If the user explicitly provides a UTC/GMT offset, return it normalized as `UTC+HH:MM` or `UTC-HH:MM`; otherwise return `UTC+08:00`.
6. "DateProvided": Set true ONLY if the user explicitly provided a numeric date (like "3月12号", "2026-03-12", "12号").
7. "TimeProvided": Set true ONLY if the user explicitly provided a numeric time (like "14:30", "2点半", "14点").
8. If bookkeeping content is ambiguous, still output valid JSON and set `Status` to `待确认`.
9. Currency defaults to CNY unless specified.
10. `Status` must be one of: `""`, `待确认`, `作废`. Use `""` for normal records.
11. `Note` should capture the purpose or context only. Do not include the specific amount or currency in `Note`, and do not restate numeric details already captured in `Amount` unless absolutely necessary for meaning.
12. For normal income/expense, `Category` must be one of: {categories}.
13. If the message is about transferring money to mother (for example: `给妈妈转账`, `给妈妈`, `转给妈妈`), `Category` must be `给妈妈`.
14. `PaymentChannel` must be one of: {payment_channels}.
15. If the user explicitly mentioned a payment channel, use that exact configured value. Otherwise use the configured default payment channel: {default_payment_channel}.
16. For bookkeeping messages, always set `tool_call.tool` to `append_record`.
17. Leave `tool_call.arguments.sheet_name` as null unless the user explicitly asked to target a specific sheet.

Telegram envelope:
{envelope}
"""


def _load_payload(args: argparse.Namespace) -> Dict[str, Any]:
    if args.json: return json.loads(args.json)
    if args.json_file: return json.loads(Path(args.json_file).read_text(encoding="utf-8"))
    if not sys.stdin.isatty():
        raw = sys.stdin.read().strip()
        if raw: return json.loads(raw)
    raise ValueError("No JSON input.")


def _message_datetime(payload: Dict[str, Any], timezone_text: str) -> datetime:
    for key in ("telegram_timestamp", "message_date", "timestamp"):
        raw_value = payload.get(key)
        try:
            return convert_telegram_timestamp(raw_value, timezone_text)
        except Exception:
            continue
    return convert_telegram_timestamp(None, timezone_text)


def _default_datetime(payload: Dict[str, Any], timezone_text: str) -> tuple[str, str]:
    message_dt = _message_datetime(payload, timezone_text)
    return message_dt.strftime("%Y-%m-%d"), message_dt.strftime("%H:%M")


def _message_text(payload: Dict[str, Any]) -> str:
    for key in ("text", "message", "raw_text"):
        value = payload.get(key)
        if value is not None: return str(value).strip()
    return ""


def _runtime_config(payload: Dict[str, Any]) -> Dict[str, Any] | None:
    runtime_config = payload.get("runtime_config")
    return runtime_config if isinstance(runtime_config, dict) else None


def _coerce_tool_call(payload: Dict[str, Any]) -> Dict[str, Any]:
    if "tool_call" in payload and isinstance(payload["tool_call"], dict):
        return payload["tool_call"]
    for key in ("gemini_output", "codex_output"):
        if key in payload:
            output = payload[key]
            if isinstance(output, str):
                try: output = json.loads(output)
                except: continue
            if isinstance(output, dict):
                tool_call = output.get("tool_call")
                if isinstance(tool_call, dict):
                    return tool_call
                record = output.get("record")
                if isinstance(record, dict):
                    return {"tool": "append_record", "arguments": {"record": record}}
    raise ValueError("No tool_call found in payload.")


def _is_ignored_payload(payload: Dict[str, Any]) -> bool:
    if payload.get("ignored") is True: return True
    for key in ("gemini_output", "codex_output"):
        output = payload.get(key)
        if not output: continue
        if isinstance(output, str):
            try: output = json.loads(output)
            except: continue
        if isinstance(output, dict) and output.get("ignored") is True: return True
    return False


def _fill_defaults(record: Dict[str, Any], payload: Dict[str, Any]) -> Dict[str, Any]:
    merged = dict(record)
    merged["Timezone"] = normalize_timezone(merged.get("Timezone", DEFAULT_TIMEZONE))
    tg_date, tg_time = _default_datetime(payload, merged["Timezone"])
    default_payment_channel = get_default_payment_channel(_runtime_config(payload))

    # Telegram 消息的账本 ID 必须稳定且可回溯，不能信任模型输出的任意 ID。
    chat_id, msg_id = payload.get("chat_id"), payload.get("message_id")
    if chat_id is not None and msg_id is not None:
        merged["ID"] = f"{chat_id}:{msg_id}"
    elif not str(merged.get("ID", "")).strip():
        merged["ID"] = f"manual_{int(datetime.now().timestamp())}"

    if not merged.get("DateProvided") or not str(merged.get("Date", "")).strip():
        merged["Date"] = tg_date

    if not merged.get("TimeProvided") or not str(merged.get("Time", "")).strip():
        merged["Time"] = tg_time

    if not str(merged.get("Currency", "")).strip(): merged["Currency"] = "CNY"
    if not str(merged.get("PaymentChannel", "")).strip():
        merged["PaymentChannel"] = default_payment_channel
    if "Status" not in merged: merged["Status"] = ""
    return merged


def _sheet_name_for_record(record: Dict[str, Any], explicit_sheet_name: str | None) -> str | None:
    if explicit_sheet_name: return explicit_sheet_name
    raw_date = str(record.get("Date", "")).strip()
    try: return str(datetime.strptime(raw_date, "%Y-%m-%d").year)
    except: return None


def _normalize_tool_call(tool_call: Dict[str, Any], payload: Dict[str, Any], explicit_sheet_name: str | None) -> Dict[str, Any]:
    tool_name = str(tool_call.get("tool", "")).strip() or "append_record"
    arguments = tool_call.get("arguments", {})
    if not isinstance(arguments, dict):
        raise ValueError("tool_call.arguments must be an object")

    if tool_name != "append_record":
        raise ValueError(f"Unsupported bookkeeping tool: {tool_name}")

    record = _fill_defaults(arguments.get("record", {}), payload)
    normalized_record = normalize_record(record)
    resolved_sheet_name = _sheet_name_for_record(normalized_record, explicit_sheet_name or arguments.get("sheet_name"))
    return {
        "tool": tool_name,
        "arguments": {
            "record": normalized_record,
            "sheet_name": resolved_sheet_name,
        },
    }


def emit_prompt(payload: Dict[str, Any]) -> str:
    runtime_config = _runtime_config(payload)
    allowed_categories = sorted(get_allowed_categories(config=runtime_config))
    payment_channels = get_payment_channels(config=runtime_config)
    default_payment_channel = get_default_payment_channel(config=runtime_config) or "(none configured)"
    minimal_envelope = {
        "message_id": payload.get("message_id"),
        "chat_id": payload.get("chat_id"),
        "sender": payload.get("sender"),
        "telegram_timestamp": payload.get("telegram_timestamp") or payload.get("message_date") or payload.get("timestamp"),
        "text": _message_text(payload),
    }
    return PROMPT_TEMPLATE.format(
        categories="|".join(allowed_categories),
        payment_channels="|".join(payment_channels) if payment_channels else "(none configured)",
        default_payment_channel=default_payment_channel,
        output_shape=json.dumps(PROMPT_OUTPUT_SHAPE, ensure_ascii=False),
        tool_call_shape=json.dumps(PROMPT_TOOL_CALL_SHAPE, ensure_ascii=False),
        envelope=json.dumps(minimal_envelope, ensure_ascii=False),
    )


def apply_tool_call(payload: Dict[str, Any], excel_path: str, sheet_name: str | None, backend: str, dry_run: bool) -> Dict[str, Any]:
    if _is_ignored_payload(payload):
        reason = "not bookkeeping related"
        for key in ("gemini_output", "codex_output"):
            output = payload.get(key, payload if key == "gemini_output" else {})
            if isinstance(output, str):
                try: output = json.loads(output)
                except: continue
            if isinstance(output, dict) and output.get("ignored"):
                reason = output.get("reason", reason)
                break
        return {"ok": False, "ignored": True, "reason": reason}

    normalized_tool_call = _normalize_tool_call(_coerce_tool_call(payload), payload, sheet_name)
    result = {
        "ok": True,
        "tool_call": normalized_tool_call,
        "record": normalized_tool_call["arguments"]["record"],
        "sheet_name": normalized_tool_call["arguments"]["sheet_name"],
    }
    if dry_run:
        return result

    tool_result = run_tool_payload(normalized_tool_call, excel_path=excel_path, backend=backend)
    result["tool"] = tool_result["tool"]
    result["row"] = tool_result["result"].get("row")
    result["tool_result"] = tool_result["result"]
    return result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=["prompt", "apply"])
    parser.add_argument("--json")
    parser.add_argument("--json-file")
    parser.add_argument("--excel-path", default=str(DEFAULT_EXCEL_PATH))
    parser.add_argument("--sheet-name")
    parser.add_argument("--backend", default="openpyxl")
    parser.add_argument("--dry-run", action="store_true")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    payload = _load_payload(args)
    try:
        if args.mode == "prompt":
            print(emit_prompt(payload))
            return 0
        result = apply_tool_call(payload, args.excel_path, args.sheet_name, args.backend, args.dry_run)
        print(json.dumps(result, ensure_ascii=False))
    except Exception as exc:
        print(json.dumps({"ok": False, "ignored": True, "reason": str(exc)}, ensure_ascii=False))
        return 0
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
