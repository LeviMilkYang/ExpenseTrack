from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict

BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
DEFAULT_EXCEL_PATH = PROJECT_DIR / "expense.xlsx"
HELPER_TEMP_DIR = BASE_DIR / "tmp_excel_helper"
DEFAULT_CONFIG_PATH = BASE_DIR / "telegram_bot_config.json"
FALLBACK_CATEGORY = "未分类"

STATUS_NORMAL = ""
STATUS_PENDING = "待确认"
STATUS_VOID = "作废"
ALLOWED_STATUS = {STATUS_NORMAL, STATUS_PENDING, STATUS_VOID}

EXPECTED_HEADERS = [
    "Date",
    "Time",
    "Amount",
    "Currency",
    "Type",
    "Category",
    "Note",
    "Status",
]

FIELD_ALIASES = {
    "Date": ("Date", "date", "日期"),
    "Time": ("Time", "time", "时间"),
    "Amount": ("Amount", "amount", "金额"),
    "Currency": ("Currency", "currency", "币种"),
    "Type": ("Type", "type", "收支"),
    "Category": ("Category", "category", "分类"),
    "Note": ("Note", "note", "备注"),
    "Status": ("Status", "status", "状态", "NeedConfirm", "need_confirm", "待确认"),
}

TYPE_ALIASES = {
    "收入": "收入",
    "收": "收入",
    "income": "收入",
    "支出": "支出",
    "支": "支出",
    "expense": "支出",
    "借入": "借入",
    "借到": "借入",
    "借款": "借入",
    "borrow": "借入",
    "loanin": "借入",
    "贷出": "贷出",
    "借出": "贷出",
    "出借": "贷出",
    "lend": "贷出",
    "loanout": "贷出",
    "还入": "收回",
    "收回": "收回",
    "收回借款": "收回",
    "收回借出": "收回",
    "repayin": "收回",
    "还出": "偿还",
    "归还": "偿还",
    "还款": "偿还",
    "归还借款": "偿还",
    "repayout": "偿还",
}

ALLOWED_TYPES = {"收入", "支出", "借入", "贷出", "收回", "偿还"}
LOAN_TYPES = {"借入", "贷出", "收回", "偿还"}

DEFAULT_ALLOWED_CATEGORIES = [
    "吃喝",
    "日用消耗",
    "交通",
    "娱乐电子",
    "医疗",
    "人情往来",
    "给妈妈",
    "和女同志",
    "工资",
    "公积金",
    "房租",
]

HELPER_CODE_OPENPYXL = r"""
from __future__ import annotations

import json
import sys
from datetime import date, datetime
from pathlib import Path
from openpyxl import load_workbook

EXPECTED_HEADERS = ["Date", "Time", "Amount", "Currency", "Type", "Category", "Note", "Status"]
STATUS_COL = EXPECTED_HEADERS.index("Status") + 1


def row_values(worksheet, row: int) -> list[object]:
    return [worksheet.cell(row=row, column=col).value for col in range(1, len(EXPECTED_HEADERS) + 1)]


def json_safe(value):
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return value


def find_last_non_empty_row(worksheet) -> int | None:
    for row in range(max(worksheet.max_row, 2), 1, -1):
        values = row_values(worksheet, row)
        if any(value not in (None, "") for value in values):
            return row
    return None


def find_next_row(worksheet) -> int:
    last_non_empty_row = find_last_non_empty_row(worksheet)
    if last_non_empty_row is None:
        return 2
    return last_non_empty_row + 1


def find_last_record_row(worksheet) -> int:
    row = find_last_non_empty_row(worksheet)
    if row is not None:
        return row
    raise ValueError("没有可作废的上一条记录")


def find_record_row(worksheet, row: int) -> int:
    if row < 2:
        raise ValueError(f"记录行号非法: {row}")
    values = row_values(worksheet, row)
    if any(value not in (None, "") for value in values):
        return row
    raise ValueError(f"第 {row} 行不存在有效记录")


def main() -> int:
    excel_path = Path(sys.argv[1])
    payload_path = Path(sys.argv[2])

    payload = json.loads(payload_path.read_text(encoding="utf-8"))
    action = payload.get("action", "append")
    record = payload.get("record")
    sheet_name = payload.get("sheet_name")

    workbook = load_workbook(excel_path)
    worksheet = workbook[sheet_name] if sheet_name else workbook.worksheets[0]

    actual_headers = [worksheet.cell(row=1, column=i).value for i in range(1, len(EXPECTED_HEADERS) + 1)]
    if actual_headers != EXPECTED_HEADERS:
        raise ValueError(f"Header mismatch: {actual_headers}")

    if action == "append":
        next_row = find_next_row(worksheet)
        for col, header in enumerate(EXPECTED_HEADERS, start=1):
            worksheet.cell(row=next_row, column=col, value=record[header])

        worksheet.cell(row=next_row, column=1).number_format = "yyyy-mm-dd"
        worksheet.cell(row=next_row, column=2).number_format = "hh:mm"
        workbook.save(excel_path)

        print(json.dumps({"row": next_row}, ensure_ascii=False))
        return 0

    if action == "invalidate_last":
        target_row = find_last_record_row(worksheet)
        worksheet.cell(row=target_row, column=STATUS_COL, value="作废")
        workbook.save(excel_path)

        print(json.dumps({"row": target_row}, ensure_ascii=False))
        return 0

    if action == "invalidate_row":
        target_row = find_record_row(worksheet, int(payload["row"]))
        worksheet.cell(row=target_row, column=STATUS_COL, value="作废")
        workbook.save(excel_path)

        print(json.dumps({"row": target_row}, ensure_ascii=False))
        return 0

    if action == "read_row":
        target_row = find_record_row(worksheet, int(payload["row"]))
        record = {key: json_safe(value) for key, value in zip(EXPECTED_HEADERS, row_values(worksheet, target_row))}
        print(
            json.dumps(
                {
                    "row": target_row,
                    "record": record,
                },
                ensure_ascii=False,
            )
        )
        return 0

    raise ValueError(f"Unsupported action: {action}")


if __name__ == "__main__":
    raise SystemExit(main())
"""

HELPER_CODE_WIN32COM = r"""
from __future__ import annotations

import json
import pythoncom
import sys
from datetime import date, datetime
from pathlib import Path

import win32com.client as win32

EXPECTED_HEADERS = ["Date", "Time", "Amount", "Currency", "Type", "Category", "Note", "Status"]
STATUS_COL = EXPECTED_HEADERS.index("Status") + 1


def row_values(worksheet, row: int) -> list[object]:
    return [worksheet.Cells(row, col).Value for col in range(1, len(EXPECTED_HEADERS) + 1)]


def json_safe(value):
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    return value


def find_last_non_empty_row(worksheet) -> int | None:
    last_row = worksheet.UsedRange.Rows.Count + worksheet.UsedRange.Row - 1
    for row in range(max(last_row, 2), 1, -1):
        values = row_values(worksheet, row)
        if any(value not in (None, "") for value in values):
            return row
    return None


def find_next_row(worksheet) -> int:
    last_non_empty_row = find_last_non_empty_row(worksheet)
    if last_non_empty_row is None:
        return 2
    return last_non_empty_row + 1


def find_last_record_row(worksheet) -> int:
    row = find_last_non_empty_row(worksheet)
    if row is not None:
        return row
    raise ValueError("没有可作废的上一条记录")


def find_record_row(worksheet, row: int) -> int:
    if row < 2:
        raise ValueError(f"记录行号非法: {row}")
    values = row_values(worksheet, row)
    if any(value not in (None, "") for value in values):
        return row
    raise ValueError(f"第 {row} 行不存在有效记录")


def main() -> int:
    excel_path = Path(sys.argv[1])
    payload_path = Path(sys.argv[2])

    payload = json.loads(payload_path.read_text(encoding="utf-8"))
    action = payload.get("action", "append")
    record = payload.get("record")
    sheet_name = payload.get("sheet_name")

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(str(excel_path))
        worksheet = workbook.Worksheets(sheet_name) if sheet_name else workbook.Worksheets(1)

        actual_headers = [worksheet.Cells(1, i).Value for i in range(1, len(EXPECTED_HEADERS) + 1)]
        if actual_headers != EXPECTED_HEADERS:
            raise ValueError(f"Header mismatch: {actual_headers}")

        if action == "append":
            next_row = find_next_row(worksheet)

            for col, header in enumerate(EXPECTED_HEADERS, start=1):
                worksheet.Cells(next_row, col).Value = record[header]

            worksheet.Cells(next_row, 1).NumberFormat = "yyyy-mm-dd"
            worksheet.Cells(next_row, 2).NumberFormat = "hh:mm"
            workbook.Save()

            print(json.dumps({"row": next_row}, ensure_ascii=False))
            return 0

        if action == "invalidate_last":
            target_row = find_last_record_row(worksheet)
            worksheet.Cells(target_row, STATUS_COL).Value = "作废"
            workbook.Save()

            print(json.dumps({"row": target_row}, ensure_ascii=False))
            return 0

        if action == "invalidate_row":
            target_row = find_record_row(worksheet, int(payload["row"]))
            worksheet.Cells(target_row, STATUS_COL).Value = "作废"
            workbook.Save()

            print(json.dumps({"row": target_row}, ensure_ascii=False))
            return 0

        if action == "read_row":
            target_row = find_record_row(worksheet, int(payload["row"]))
            record = {key: json_safe(value) for key, value in zip(EXPECTED_HEADERS, row_values(worksheet, target_row))}
            print(
                json.dumps(
                    {
                        "row": target_row,
                        "record": record,
                    },
                    ensure_ascii=False,
                )
            )
            return 0

        raise ValueError(f"Unsupported action: {action}")
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
"""


def _pick_value(record: Dict[str, Any], field: str) -> Any:
    for alias in FIELD_ALIASES[field]:
        if alias in record and record[alias] is not None:
            return record[alias]
    return None


def load_bot_config(config_path: Path = DEFAULT_CONFIG_PATH) -> Dict[str, Any]:
    if not config_path.exists():
        return {"allowed_categories": list(DEFAULT_ALLOWED_CATEGORIES)}
    return json.loads(config_path.read_text(encoding="utf-8"))


def get_allowed_categories(config_path: Path = DEFAULT_CONFIG_PATH) -> set[str]:
    config = load_bot_config(config_path)
    raw_categories = config.get("allowed_categories", DEFAULT_ALLOWED_CATEGORIES)
    if not isinstance(raw_categories, list) or not raw_categories:
        raise ValueError("telegram_bot_config.json 中的 allowed_categories 非法")
    categories = {str(item).strip() for item in raw_categories if str(item).strip()}
    if not categories:
        raise ValueError("telegram_bot_config.json 中的 allowed_categories 不能为空")
    return categories


def normalize_status(raw_value: Any) -> str:
    if raw_value in (None, "", False, 0, "0", "false", "False", "否", "normal", "Normal"):
        return STATUS_NORMAL
    if raw_value in (True, 1, "1", "true", "True", "是", "yes", "Yes"):
        return STATUS_PENDING

    text = str(raw_value).strip()
    if text == "":
        return STATUS_NORMAL
    if text not in ALLOWED_STATUS:
        raise ValueError(f"状态字段非法: {raw_value}")
    return text


def normalize_record(record: Dict[str, Any]) -> Dict[str, Any]:
    normalized: Dict[str, Any] = {}
    for field in EXPECTED_HEADERS:
        value = _pick_value(record, field)
        if field != "Status" and value is None:
            raise ValueError(f"缺少必填字段: {field}")
        normalized[field] = value

    normalized_type = str(normalized["Type"]).strip()
    normalized_type = TYPE_ALIASES.get(normalized_type.lower(), TYPE_ALIASES.get(normalized_type, normalized_type))
    if normalized_type not in ALLOWED_TYPES:
        raise ValueError(f"类型字段非法: {normalized['Type']}")

    category = str(normalized["Category"]).strip()
    if not category:
        raise ValueError("分类字段不能为空")
    allowed_categories = get_allowed_categories()
    if normalized_type not in LOAN_TYPES and category not in allowed_categories and category != FALLBACK_CATEGORY:
        raise ValueError(f"分类字段非法: {category}")

    try:
        amount = float(normalized["Amount"])
    except Exception as exc:
        raise ValueError(f"金额无法转换为数字: {normalized['Amount']}") from exc

    status = normalize_status(normalized.get("Status", STATUS_NORMAL))

    return {
        "Date": str(normalized["Date"]).strip(),
        "Time": str(normalized["Time"]).strip(),
        "Amount": amount,
        "Currency": str(normalized["Currency"]).strip(),
        "Type": normalized_type,
        "Category": category,
        "Note": str(normalized["Note"]).strip(),
        "Status": status,
    }


def _windows_path(path: Path) -> str:
    completed = subprocess.run(
        ["wslpath", "-w", str(path)],
        check=True,
        capture_output=True,
        text=True,
    )
    return completed.stdout.strip()


def _run_excel_helper(
    excel_path: Path,
    payload: Dict[str, Any],
    backend: str,
) -> int:
    result = _run_excel_helper_with_result(excel_path, payload, backend)
    return int(result["row"])


def _run_excel_helper_with_result(
    excel_path: Path,
    payload: Dict[str, Any],
    backend: str,
) -> Dict[str, Any]:
    helper_code = HELPER_CODE_WIN32COM if backend == "win32com" else HELPER_CODE_OPENPYXL
    HELPER_TEMP_DIR.mkdir(parents=True, exist_ok=True)
    helper_path = HELPER_TEMP_DIR / "excel_append_helper.py"
    payload_path = HELPER_TEMP_DIR / "record.json"

    helper_path.write_text(helper_code, encoding="utf-8")
    payload_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    helper_win = _windows_path(helper_path)
    payload_win = _windows_path(payload_path)
    excel_win = _windows_path(excel_path)

    command = f"python '{helper_win}' '{excel_win}' '{payload_win}'"
    completed = subprocess.run(
        ["powershell.exe", "-Command", command],
        capture_output=True,
        text=True,
    )

    if completed.returncode != 0:
        message = completed.stderr.strip() or completed.stdout.strip() or "unknown error"
        raise RuntimeError(f"PowerShell 写入 Excel 失败({backend}): {message}")

    stdout = completed.stdout.strip()
    if not stdout:
        raise RuntimeError("PowerShell 写入 Excel 成功但未返回行号")

    return json.loads(stdout.splitlines()[-1])


def _build_payload(record: Dict[str, Any], sheet_name: str | None) -> Dict[str, Any]:
    return {
        "action": "append",
        "record": normalize_record(record),
        "sheet_name": sheet_name,
    }


def append_record_to_excel(
    excel_path: str | Path,
    record: Dict[str, Any],
    sheet_name: str | None = None,
    backend: str = "openpyxl",
) -> int:
    excel_path = Path(excel_path).resolve()
    payload = _build_payload(record, sheet_name)
    return _run_excel_helper(excel_path, payload, backend)


def invalidate_last_record_in_excel(
    excel_path: str | Path,
    sheet_name: str | None = None,
    backend: str = "openpyxl",
) -> int:
    excel_path = Path(excel_path).resolve()
    payload = {
        "action": "invalidate_last",
        "sheet_name": sheet_name,
    }
    return _run_excel_helper(excel_path, payload, backend)


def invalidate_record_in_excel(
    excel_path: str | Path,
    row: int,
    sheet_name: str | None = None,
    backend: str = "openpyxl",
) -> int:
    excel_path = Path(excel_path).resolve()
    payload = {
        "action": "invalidate_row",
        "sheet_name": sheet_name,
        "row": int(row),
    }
    return _run_excel_helper(excel_path, payload, backend)


def read_record_from_excel(
    excel_path: str | Path,
    row: int,
    sheet_name: str | None = None,
    backend: str = "openpyxl",
) -> Dict[str, Any]:
    excel_path = Path(excel_path).resolve()
    payload = {
        "action": "read_row",
        "sheet_name": sheet_name,
        "row": int(row),
    }
    result = _run_excel_helper_with_result(excel_path, payload, backend)
    record = result.get("record")
    if not isinstance(record, dict):
        raise RuntimeError("读取 Excel 记录失败：返回值缺少 record")
    return record


def _record_from_args(args: argparse.Namespace) -> Dict[str, Any]:
    return {
        "Date": args.date,
        "Time": args.time,
        "Amount": args.amount,
        "Currency": args.currency,
        "Type": args.type,
        "Category": args.category,
        "Note": args.note,
        "Status": args.status,
    }


def _load_record(args: argparse.Namespace) -> Dict[str, Any]:
    if args.json:
        return json.loads(args.json)

    if args.json_file:
        return json.loads(Path(args.json_file).read_text(encoding="utf-8"))

    if not sys.stdin.isatty():
        raw = sys.stdin.read().strip()
        if raw:
            return json.loads(raw)

    if all(
        value is not None
        for value in [args.date, args.time, args.amount, args.currency, args.type, args.category, args.note]
    ):
        return _record_from_args(args)

    raise ValueError("未提供有效输入。请使用 --json、--json-file、stdin JSON 或完整的命令行字段参数。")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Append one expense record into expense.xlsx from WSL via PowerShell.")
    parser.add_argument("--excel-path", default=str(DEFAULT_EXCEL_PATH))
    parser.add_argument("--sheet-name")
    parser.add_argument("--json", help="Inline JSON record.")
    parser.add_argument("--json-file", help="Path to a JSON file containing one record.")
    parser.add_argument("--backend", choices=["win32com", "openpyxl"], default="openpyxl")
    parser.add_argument("--date")
    parser.add_argument("--time")
    parser.add_argument("--amount", type=float)
    parser.add_argument("--currency")
    parser.add_argument("--type")
    parser.add_argument("--category")
    parser.add_argument("--note")
    parser.add_argument("--status")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    record = _load_record(args)
    row_num = append_record_to_excel(args.excel_path, record, args.sheet_name, args.backend)
    print(json.dumps({"ok": True, "row": row_num}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
