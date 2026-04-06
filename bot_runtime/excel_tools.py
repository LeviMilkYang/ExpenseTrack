from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Callable, Dict

from append_excel_entry import (
    DEFAULT_EXCEL_PATH,
    append_record_to_excel,
    invalidate_last_record_in_excel,
    invalidate_record_by_id,
    normalize_record,
    read_record_by_id,
)

ToolHandler = Callable[[Dict[str, Any], Path, str], Dict[str, Any]]


TOOL_REGISTRY: Dict[str, Dict[str, Any]] = {
    "append_record": {
        "description": "Append one normalized bookkeeping record into the target workbook.",
        "input_schema": {
            "type": "object",
            "additionalProperties": False,
            "required": ["record"],
            "properties": {
                "record": {"type": "object"},
                "sheet_name": {"type": ["string", "null"]},
            },
        },
    },
    "invalidate_record": {
        "description": "Mark a record as void by its stable record ID.",
        "input_schema": {
            "type": "object",
            "additionalProperties": False,
            "required": ["record_id"],
            "properties": {
                "record_id": {"type": "string"},
                "sheet_name": {"type": ["string", "null"]},
            },
        },
    },
    "read_record": {
        "description": "Read one record by its stable record ID.",
        "input_schema": {
            "type": "object",
            "additionalProperties": False,
            "required": ["record_id"],
            "properties": {
                "record_id": {"type": "string"},
                "sheet_name": {"type": ["string", "null"]},
            },
        },
    },
    "invalidate_last_record": {
        "description": "Mark the last non-empty record in a sheet as void.",
        "input_schema": {
            "type": "object",
            "additionalProperties": False,
            "properties": {
                "sheet_name": {"type": ["string", "null"]},
            },
        },
    },
}


def _as_payload(raw_payload: str) -> Dict[str, Any]:
    payload = json.loads(raw_payload)
    if not isinstance(payload, dict):
        raise ValueError("Tool payload must be a JSON object")
    return payload


def describe_tools() -> Dict[str, Any]:
    return {
        "tools": [
            {
                "name": name,
                "description": spec["description"],
                "input_schema": spec["input_schema"],
            }
            for name, spec in TOOL_REGISTRY.items()
        ]
    }


def _tool_sheet_name(arguments: Dict[str, Any]) -> str | None:
    sheet_name = arguments.get("sheet_name")
    if sheet_name is None:
        return None
    text = str(sheet_name).strip()
    return text or None


def _handle_append_record(arguments: Dict[str, Any], excel_path: Path, backend: str) -> Dict[str, Any]:
    record = arguments.get("record")
    if not isinstance(record, dict):
        raise ValueError("append_record requires object field: record")
    sheet_name = _tool_sheet_name(arguments)
    normalized = normalize_record(record)
    row = append_record_to_excel(excel_path, normalized, sheet_name=sheet_name, backend=backend)
    return {
        "row": row,
        "sheet_name": sheet_name,
        "record": normalized,
        "record_id": normalized["ID"],
    }


def _handle_invalidate_record(arguments: Dict[str, Any], excel_path: Path, backend: str) -> Dict[str, Any]:
    record_id = str(arguments.get("record_id", "")).strip()
    if not record_id:
        raise ValueError("invalidate_record requires string field: record_id")
    sheet_name = _tool_sheet_name(arguments)
    row = invalidate_record_by_id(excel_path, record_id, sheet_name=sheet_name, backend=backend)
    return {"row": row, "sheet_name": sheet_name, "record_id": record_id}


def _handle_read_record(arguments: Dict[str, Any], excel_path: Path, backend: str) -> Dict[str, Any]:
    record_id = str(arguments.get("record_id", "")).strip()
    if not record_id:
        raise ValueError("read_record requires string field: record_id")
    sheet_name = _tool_sheet_name(arguments)
    record = read_record_by_id(excel_path, record_id, sheet_name=sheet_name, backend=backend)
    return {"record": record, "sheet_name": sheet_name, "record_id": record_id}


def _handle_invalidate_last_record(arguments: Dict[str, Any], excel_path: Path, backend: str) -> Dict[str, Any]:
    sheet_name = _tool_sheet_name(arguments)
    row = invalidate_last_record_in_excel(excel_path, sheet_name=sheet_name, backend=backend)
    return {"row": row, "sheet_name": sheet_name}


TOOL_HANDLERS: Dict[str, ToolHandler] = {
    "append_record": _handle_append_record,
    "invalidate_record": _handle_invalidate_record,
    "read_record": _handle_read_record,
    "invalidate_last_record": _handle_invalidate_last_record,
}


def run_tool_payload(payload: Dict[str, Any], excel_path: str | Path = DEFAULT_EXCEL_PATH, backend: str = "openpyxl") -> Dict[str, Any]:
    tool_name = str(payload.get("tool", "")).strip()
    if not tool_name:
        raise ValueError("Tool payload requires field: tool")
    handler = TOOL_HANDLERS.get(tool_name)
    if handler is None:
        raise ValueError(f"Unsupported excel tool: {tool_name}")

    arguments = payload.get("arguments", {})
    if arguments is None:
        arguments = {}
    if not isinstance(arguments, dict):
        raise ValueError("Tool payload field arguments must be an object")

    result = handler(arguments, Path(excel_path).resolve(), backend)
    return {"ok": True, "tool": tool_name, "result": result}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=["run", "describe"], default="run", nargs="?")
    parser.add_argument("--tool-json")
    parser.add_argument("--tool-json-file")
    parser.add_argument("--excel-path", default=str(DEFAULT_EXCEL_PATH))
    parser.add_argument("--backend", default="openpyxl")
    return parser


def _load_tool_payload(args: argparse.Namespace) -> Dict[str, Any]:
    if args.tool_json:
        return _as_payload(args.tool_json)
    if args.tool_json_file:
        return _as_payload(Path(args.tool_json_file).read_text(encoding="utf-8"))
    if not sys.stdin.isatty():
        raw = sys.stdin.read().strip()
        if raw:
            return _as_payload(raw)
    raise ValueError("No tool JSON input.")


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        if args.mode == "describe":
            print(json.dumps(describe_tools(), ensure_ascii=False))
            return 0

        payload = _load_tool_payload(args)
        print(json.dumps(run_tool_payload(payload, excel_path=args.excel_path, backend=args.backend), ensure_ascii=False))
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
