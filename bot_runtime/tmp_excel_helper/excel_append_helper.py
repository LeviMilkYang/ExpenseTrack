
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
