
from __future__ import annotations

import json
import sys
from datetime import date, datetime, time as dt_time
from pathlib import Path
from openpyxl import load_workbook

EXPECTED_HEADERS = ["ID", "Date", "Time", "Amount", "Currency", "Type", "Category", "Note", "Status"]
ID_COL = EXPECTED_HEADERS.index("ID") + 1
DATE_COL = EXPECTED_HEADERS.index("Date") + 1
TIME_COL = EXPECTED_HEADERS.index("Time") + 1
STATUS_COL = EXPECTED_HEADERS.index("Status") + 1


def row_values(worksheet, row: int) -> list[object]:
    return [worksheet.cell(row=row, column=col).value for col in range(1, len(EXPECTED_HEADERS) + 1)]


def json_safe(value):
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, dt_time):
        return value.isoformat()
    return value


def find_last_non_empty_row(worksheet) -> int | None:
    for row in range(max(worksheet.max_row, 2), 1, -1):
        values = row_values(worksheet, row)
        if any(value not in (None, "") for value in values):
            return row
    return None


def find_row_by_id(worksheet, record_id: str) -> int:
    if not record_id:
        raise ValueError("ID 不能为空")
    
    last_row = worksheet.max_row
    for row in range(2, last_row + 1):
        if str(worksheet.cell(row=row, column=ID_COL).value) == str(record_id):
            return row
    raise ValueError(f"未找到 ID 为 {record_id} 的记录")


def sort_worksheet(worksheet):
    # 简单的冒泡或提取重写排序。对于 Excel 脚本，提取所有数据排序后再写回最稳妥。
    rows = []
    last_row = find_last_non_empty_row(worksheet)
    if not last_row or last_row < 2:
        return

    for r in range(2, last_row + 1):
        rows.append(row_values(worksheet, r))

    def sort_key(row):
        # Date 可能是 datetime.date 或 str
        d = row[DATE_COL-1]
        t = row[TIME_COL-1]
        
        # 统一转为 datetime 进行比较
        if isinstance(d, str):
            try:
                d_obj = datetime.strptime(d, "%Y-%m-%d").date()
            except:
                d_obj = date(1970, 1, 1)
        elif isinstance(d, datetime):
            d_obj = d.date()
        else:
            d_obj = d or date(1970, 1, 1)

        if isinstance(t, str):
            try:
                t_obj = datetime.strptime(t, "%H:%M").time()
            except:
                t_obj = dt_time(0, 0)
        elif isinstance(t, datetime):
            t_obj = t.time()
        else:
            t_obj = t or dt_time(0, 0)
            
        return datetime.combine(d_obj, t_obj)

    rows.sort(key=sort_key)

    for i, row_data in enumerate(rows, start=2):
        for j, val in enumerate(row_data, start=1):
            worksheet.cell(row=i, column=j, value=val)
            if j == DATE_COL:
                worksheet.cell(row=i, column=j).number_format = "yyyy-mm-dd"
            elif j == TIME_COL:
                worksheet.cell(row=i, column=j).number_format = "hh:mm"


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
        # 如果是旧表，尝试升级表头
        if actual_headers[:2] == [None, None] or actual_headers[0] != "ID":
             for col, header in enumerate(EXPECTED_HEADERS, start=1):
                 worksheet.cell(row=1, column=col, value=header)
        else:
            raise ValueError(f"Header mismatch: {actual_headers}")

    if action == "append":
        last_row = find_last_non_empty_row(worksheet) or 1
        next_row = last_row + 1
        for col, header in enumerate(EXPECTED_HEADERS, start=1):
            worksheet.cell(row=next_row, column=col, value=record[header])

        worksheet.cell(row=next_row, column=DATE_COL).number_format = "yyyy-mm-dd"
        worksheet.cell(row=next_row, column=TIME_COL).number_format = "hh:mm"
        
        # 排序
        sort_worksheet(worksheet)
        workbook.save(excel_path)

        # 排序后行号会变，需要重新找回该 ID 的行号返回给用户
        final_row = find_row_by_id(worksheet, record["ID"])
        print(json.dumps({"row": final_row}, ensure_ascii=False))
        return 0

    if action == "invalidate_id":
        target_id = payload.get("id")
        target_row = find_row_by_id(worksheet, target_id)
        worksheet.cell(row=target_row, column=STATUS_COL, value="作废")
        workbook.save(excel_path)
        print(json.dumps({"row": target_row}, ensure_ascii=False))
        return 0

    if action == "read_by_id":
        target_id = payload.get("id")
        target_row = find_row_by_id(worksheet, target_id)
        record_data = {key: json_safe(value) for key, value in zip(EXPECTED_HEADERS, row_values(worksheet, target_row))}
        print(json.dumps({"row": target_row, "record": record_data}, ensure_ascii=False))
        return 0

    if action == "invalidate_last":
        # 依然支持 invalidate_last，逻辑为找最后一行
        target_row = find_last_non_empty_row(worksheet)
        if not target_row or target_row < 2:
            raise ValueError("没有可作废的记录")
        worksheet.cell(row=target_row, column=STATUS_COL, value="作废")
        workbook.save(excel_path)
        print(json.dumps({"row": target_row}, ensure_ascii=False))
        return 0

    raise ValueError(f"Unsupported action: {action}")


if __name__ == "__main__":
    raise SystemExit(main())
