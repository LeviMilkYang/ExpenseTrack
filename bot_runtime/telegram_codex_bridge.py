from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict

from append_excel_entry import ALLOWED_CATEGORIES, append_record_to_excel, normalize_record

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
DEFAULT_EXCEL_PATH = PROJECT_DIR / "expense.xlsx"

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
            "required": ["Date", "Time", "Amount", "Currency", "Type", "Category", "Note", "NeedConfirm"],
            "properties": {
                "Date": {"type": "string"},
                "Time": {"type": "string"},
                "Amount": {"type": "number"},
                "Currency": {"type": "string"},
                "Type": {"type": "string", "enum": ["鏀跺叆", "鏀嚭", "鍊熷叆", "璐峰嚭", "鏀跺洖", "鍋胯繕"]},
                "Category": {"type": "string"},
                "Note": {"type": "string"},
                "NeedConfirm": {"type": "boolean"},
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
    "Amount": "number",
    "Currency": "CNY",
    "Type": "allowed type value",
    "Category": "string",
    "Note": "string",
    "NeedConfirm": False,
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
5. If the bookkeeping content is ambiguous, still output valid JSON and set "NeedConfirm": true.
6. If the message does not explicitly mention a date or time, set `Date` to `""` and `Time` to `""`. Do not copy Telegram metadata timestamps into `record`.
7. Keep "Note" concise and preserve useful source wording, but do not repeat the numeric amount or currency in "Note" unless necessary for meaning.
8. Currency defaults to CNY unless the message clearly says otherwise.
9. For normal income/expense, `Category` must be one of: {categories}.
10. For `Type = "借入"` or `Type = "贷出"` or `Type = "收回"` or `Type = "偿还"`, put the lending/borrowing counterparty in `Category`, and put the remaining context in `Note`.
11. If the message is about transferring money to mother (for example: `给妈妈转账`, `给妈妈`, `转给妈妈`), `Category` must be `给妈妈`, not `人情往来`.

Telegram envelope:
{envelope}
"""


def _load_payload(args: argparse.Namespace) -> Dict[str, Any]:
    if args.json:
        return json.loads(args.json)

    if args.json_file:
        return json.loads(Path(args.json_file).read_text(encoding="utf-8"))

    if not sys.stdin.isatty():
        raw = sys.stdin.read().strip()
        if raw:
            return json.loads(raw)

    raise ValueError("未收到 JSON 输入。请通过 --json、--json-file 或 stdin 传入。")


def _message_datetime(payload: Dict[str, Any]) -> datetime:
    for key in ("telegram_timestamp", "message_date", "timestamp"):
        raw_value = payload.get(key)
        if raw_value in (None, ""):
            continue

        if isinstance(raw_value, (int, float)):
            return datetime.fromtimestamp(float(raw_value), tz=timezone.utc).astimezone()

        text_value = str(raw_value).strip()
        if not text_value:
            continue

        if re.fullmatch(r"\d+(\.\d+)?", text_value):
            return datetime.fromtimestamp(float(text_value), tz=timezone.utc).astimezone()

        iso_value = text_value.replace("Z", "+00:00")
        try:
            parsed = datetime.fromisoformat(iso_value)
        except ValueError:
            continue

        if parsed.tzinfo is None:
            return parsed.astimezone()
        return parsed.astimezone()

    return datetime.now().astimezone()


def _default_datetime(payload: Dict[str, Any]) -> tuple[str, str]:
    message_dt = _message_datetime(payload)
    return message_dt.strftime("%Y-%m-%d"), message_dt.strftime("%H:%M")


def _message_text(payload: Dict[str, Any]) -> str:
    for key in ("text", "message", "raw_text"):
        value = payload.get(key)
        if value is not None:
            return str(value).strip()
    return ""


def _has_explicit_date(text: str) -> bool:
    patterns = (
        r"\b\d{4}[-/]\d{1,2}[-/]\d{1,2}\b",
        r"\b\d{1,2}[-/]\d{1,2}\b",
        r"\d{1,2}月\d{1,2}日",
        r"\d{1,2}月\d{1,2}号",
        r"\d{1,2}日",
        r"\d{1,2}号",
        r"(今天|今日|昨天|前天|大前天|明天|后天|本周|上周|这周|周[一二三四五六日天]|星期[一二三四五六日天])",
    )
    return any(re.search(pattern, text) for pattern in patterns)


def _has_explicit_time(text: str) -> bool:
    patterns = (
        r"\b\d{1,2}:\d{2}\b",
        r"\b\d{1,2}点(?:\d{1,2}分)?\b",
        r"\b\d{1,2}点半\b",
        r"(早上|上午|中午|下午|晚上|傍晚|凌晨|夜里|半夜|今早|今晚)",
    )
    return any(re.search(pattern, text) for pattern in patterns)


def _coerce_record(payload: Dict[str, Any]) -> Dict[str, Any]:
    if "record" in payload and isinstance(payload["record"], dict):
        return payload["record"]

    for key in ("gemini_output", "codex_output"):
        if key in payload:
            output = payload[key]
            if isinstance(output, str):
                output = json.loads(output)
            if isinstance(output, dict):
                return output.get("record", output)

    if all(key in payload for key in ("Date", "Time", "Amount", "Currency", "Type", "Category", "Note")):
        return payload

    raise ValueError("payload 中没有可写入的 record。期望 `record`、`gemini_output` 或扁平 record 字段。")


def _is_ignored_payload(payload: Dict[str, Any]) -> bool:
    if payload.get("ignored") is True:
        return True

    for key in ("gemini_output", "codex_output"):
        output = payload.get(key)
        if output:
            if isinstance(output, str):
                output = json.loads(output)
            if isinstance(output, dict) and output.get("ignored") is True:
                return True
    return False


def _fill_defaults(record: Dict[str, Any], payload: Dict[str, Any]) -> Dict[str, Any]:
    date_value, time_value = _default_datetime(payload)
    merged = dict(record)
    message_text = _message_text(payload)

    if not _has_explicit_date(message_text):
        merged["Date"] = date_value
    elif not str(merged.get("Date", "")).strip():
        merged["Date"] = date_value

    if not _has_explicit_time(message_text):
        merged["Time"] = time_value
    elif not str(merged.get("Time", "")).strip():
        merged["Time"] = time_value

    if not str(merged.get("Currency", "")).strip():
        merged["Currency"] = "CNY"
    if "NeedConfirm" not in merged:
        merged["NeedConfirm"] = False
    return merged


def _sheet_name_for_record(record: Dict[str, Any], explicit_sheet_name: str | None) -> str | None:
    if explicit_sheet_name:
        return explicit_sheet_name

    raw_date = str(record.get("Date", "")).strip()
    if not raw_date:
        return None

    try:
        parsed_date = datetime.strptime(raw_date, "%Y-%m-%d")
    except ValueError:
        return None
    return str(parsed_date.year)


def emit_prompt(payload: Dict[str, Any]) -> str:
    minimal_envelope = {
        "message_id": payload.get("message_id"),
        "chat_id": payload.get("chat_id"),
        "sender": payload.get("sender"),
        "telegram_timestamp": payload.get("telegram_timestamp") or payload.get("message_date") or payload.get("timestamp"),
        "text": _message_text(payload),
    }
    return PROMPT_TEMPLATE.format(
        categories="|".join(sorted(ALLOWED_CATEGORIES)),
        output_shape=json.dumps(PROMPT_OUTPUT_SHAPE, ensure_ascii=False),
        record_shape=json.dumps(PROMPT_RECORD_SHAPE, ensure_ascii=False),
        envelope=json.dumps(minimal_envelope, ensure_ascii=False),
    )


def apply_record(payload: Dict[str, Any], excel_path: str, sheet_name: str | None, backend: str, dry_run: bool) -> Dict[str, Any]:
    if _is_ignored_payload(payload):
        # Try to find reason in gemini_output or codex_output
        reason = "not bookkeeping related"
        for key in ("gemini_output", "codex_output"):
            output = payload.get(key, payload if key == "gemini_output" else {})
            if isinstance(output, str):
                try:
                    output = json.loads(output)
                except:
                    continue
            if isinstance(output, dict) and output.get("ignored"):
                reason = output.get("reason", reason)
                break
        
        return {
            "ok": False,
            "ignored": True,
            "reason": reason,
        }
    record = _fill_defaults(_coerce_record(payload), payload)
    normalized = normalize_record(record)
    resolved_sheet_name = _sheet_name_for_record(normalized, sheet_name)
    result: Dict[str, Any] = {"ok": True, "record": normalized, "sheet_name": resolved_sheet_name}

    if dry_run:
        return result

    row = append_record_to_excel(excel_path, normalized, resolved_sheet_name, backend)
    result["row"] = row
    return result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Bridge between Telegram daemon, Gemini JSON output, and Excel append.")
    parser.add_argument("mode", choices=["prompt", "apply"])
    parser.add_argument("--json", help="Inline Telegram envelope / Gemini payload JSON.")
    parser.add_argument("--json-file", help="Path to a JSON file containing the envelope / payload.")
    parser.add_argument("--excel-path", default=str(DEFAULT_EXCEL_PATH))
    parser.add_argument("--sheet-name")
    parser.add_argument("--backend", choices=["win32com", "openpyxl"], default="openpyxl")
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
    except ValueError as exc:
        print(json.dumps({"ok": False, "ignored": True, "reason": str(exc)}, ensure_ascii=False))
        return 0
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
