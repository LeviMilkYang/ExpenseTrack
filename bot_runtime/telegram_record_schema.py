from __future__ import annotations

from typing import Any, Dict

RECORD_REQUIRED_FIELDS = [
    "ID",
    "Date",
    "Time",
    "Timezone",
    "DateProvided",
    "TimeProvided",
    "Amount",
    "Currency",
    "Type",
    "Category",
    "Note",
    "PaymentChannel",
    "Status",
]

RECORD_PROPERTIES: Dict[str, Any] = {
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
    "PaymentChannel": {"type": "string"},
    "Status": {"type": "string"},
}

TOOL_CALL_ARGUMENTS_PROPERTIES: Dict[str, Any] = {
    "record": {
        "type": "object",
        "additionalProperties": False,
        "required": RECORD_REQUIRED_FIELDS,
        "properties": RECORD_PROPERTIES,
    },
    "sheet_name": {"type": ["string", "null"]},
}

TOOL_CALL_PROPERTIES: Dict[str, Any] = {
    "tool": {"type": "string", "enum": ["append_record"]},
    "arguments": {
        "type": "object",
        "additionalProperties": False,
        "required": ["record", "sheet_name"],
        "properties": TOOL_CALL_ARGUMENTS_PROPERTIES,
    },
}

CODEX_OUTPUT_SCHEMA: Dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "type": "object",
    "additionalProperties": False,
    "required": ["ignored", "reason", "tool_call"],
    "properties": {
        "ignored": {"type": "boolean"},
        "reason": {"type": "string"},
        "tool_call": {
            "type": ["object", "null"],
            "additionalProperties": False,
            "required": ["tool", "arguments"],
            "properties": TOOL_CALL_PROPERTIES,
        },
    },
}
