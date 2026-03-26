from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from append_excel_entry import (
    DEFAULT_TIMEZONE,
    append_record_to_excel,
    convert_telegram_timestamp,
    get_allowed_categories,
    normalize_record,
    normalize_timezone,
)

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
DEFAULT_EXCEL_PATH = PROJECT_DIR / "expense.xlsx"

# 更新 Schema：增加显式的日期/时间提供标志
CODEX_RECORD_SCHEMA: Dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "type": "object",
    "additionalProperties": False,
    "required": ["ignored", "reason", "record"],
    "properties": {
        "ignored": {"type": "boolean"},
        "reason": {"type": "string"},
        "record": {
            "type": ["object", "null"],
            "additionalProperties": False,
            "required": ["ID", "Date", "Time", "Timezone", "DateProvided", "TimeProvided", "Amount", "Currency", "Type", "Category", "Note", "Status"],
            "properties": {
                "ID": {"type": "string"},
                "Date": {"type": "string"},
                "Time": {"type": "string"},
                "Timezone": {"type": "string"},
                "DateProvided": {"type": "boolean"},
                "TimeProvided": {"type": "boolean"},
                "Amount": {"type": "number"},
                "Currency": {"type": "string"},
                "Type": {"type": "string", "enum": ["收入", "支出", "借入", "贷出", "收回", "偿还"]},
                "Category": {"type": "string"},
                "Note": {"type": "string"},
                "Status": {"type": "string"},
            },
        },
    },
}

PROMPT_OUTPUT_SHAPE: Dict[str, Any] = {
    "ignored": "boolean",
    "reason": "string",
    "record": "object|null",
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
    "Status": "",
}

PROMPT_TEMPLATE = """You convert one Telegram bookkeeping message into exactly one JSON object.

Rules:
1. Output JSON only. No markdown. No explanation.
2. Always output this exact top-level shape:
{output_shape}
3. If the message is not about income or expense bookkeeping, set:
{{"ignored":true,"reason":"short reason","record":null}}
4. Otherwise set `"ignored": false`, `"reason": ""`, and make `record` use this schema:
{record_shape}
5. `Timezone` must always be present. If the user explicitly provides a UTC offset, return that exact offset normalized as `UTC+HH:MM` or `UTC-HH:MM`. If the user does not provide a timezone, return `UTC+08:00`.
6. Only treat explicit UTC/GMT offsets as provided timezones, such as `UTC+8`, `UTC+08:00`, `UTC-5`, `GMT+0`, `UTC+05:30`, `UTC+05:45`. Valid range is `UTC-12:00` to `UTC+14:00`.
7. "DateProvided": Set true ONLY if the user explicitly provided a numeric date (like "3月12号", "2026-03-12", "12号").
8. "TimeProvided": Set true ONLY if the user explicitly provided a numeric time (like "14:30", "2点半", "14点").
9. If the user provides a timezone but does not provide a numeric time, keep `TimeProvided` as false and do not invent a time in JSON.
10. If bookkeeping content is ambiguous, still output valid JSON and set `Status` to `待确认`.
11. Currency defaults to CNY unless specified.
12. `Status` must be one of: `""`, `待确认`, `作废`. Use `""` for normal records.
13. `Note` should capture the purpose or context only. Do not include the specific amount or currency in `Note`, and do not restate numeric details already captured in `Amount` unless absolutely necessary for meaning.
14. For normal income/expense, `Category` must be one of: {categories}.
15. If the message is about transferring money to mother (for example: `给妈妈转账`, `给妈妈`, `转给妈妈`), `Category` must be `给妈妈`.

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


def _coerce_record(payload: Dict[str, Any]) -> Dict[str, Any]:
    if "record" in payload and isinstance(payload["record"], dict):
        return payload["record"]
    for key in ("gemini_output", "codex_output"):
        if key in payload:
            output = payload[key]
            if isinstance(output, str):
                try: output = json.loads(output)
                except: continue
            if isinstance(output, dict): return output.get("record", output)
    raise ValueError("No record found in payload.")


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
    if "Status" not in merged: merged["Status"] = ""
    return merged


def _sheet_name_for_record(record: Dict[str, Any], explicit_sheet_name: str | None) -> str | None:
    if explicit_sheet_name: return explicit_sheet_name
    raw_date = str(record.get("Date", "")).strip()
    try: return str(datetime.strptime(raw_date, "%Y-%m-%d").year)
    except: return None


def emit_prompt(payload: Dict[str, Any]) -> str:
    minimal_envelope = {
        "message_id": payload.get("message_id"),
        "chat_id": payload.get("chat_id"),
        "sender": payload.get("sender"),
        "telegram_timestamp": payload.get("telegram_timestamp") or payload.get("message_date") or payload.get("timestamp"),
        "text": _message_text(payload),
    }
    return PROMPT_TEMPLATE.format(
        categories="|".join(sorted(get_allowed_categories())),
        output_shape=json.dumps(PROMPT_OUTPUT_SHAPE, ensure_ascii=False),
        record_shape=json.dumps(PROMPT_RECORD_SHAPE, ensure_ascii=False),
        envelope=json.dumps(minimal_envelope, ensure_ascii=False),
    )


def apply_record(payload: Dict[str, Any], excel_path: str, sheet_name: str | None, backend: str, dry_run: bool) -> Dict[str, Any]:
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

    record = _fill_defaults(_coerce_record(payload), payload)
    normalized = normalize_record(record)
    resolved_sheet_name = _sheet_name_for_record(normalized, sheet_name)
    result = {"ok": True, "record": normalized, "sheet_name": resolved_sheet_name}
    if dry_run: return result
    row = append_record_to_excel(excel_path, normalized, resolved_sheet_name, backend)
    result["row"] = row
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
        result = apply_record(payload, args.excel_path, args.sheet_name, args.backend, args.dry_run)
        print(json.dumps(result, ensure_ascii=False))
    except Exception as exc:
        print(json.dumps({"ok": False, "ignored": True, "reason": str(exc)}, ensure_ascii=False))
        return 0
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
