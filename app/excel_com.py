"""Optional Excel COM backend (Windows + installed Microsoft Excel)."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Any


COM_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx", ".xltm"}


def excel_com_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
        import pythoncom  # noqa: F401
    except Exception:
        return False
    return True


def _dispatch_excel():
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    try:
        excel.ScreenUpdating = False
    except Exception:
        pass
    return excel


def create_sheet_com(
    path: Path,
    sheet_name: str,
    headers: list[str],
    data_rows: list[list[Any]] | None = None,
) -> dict[str, Any]:
    if not excel_com_available():
        raise RuntimeError("Excel COM недоступен")
    data_rows = data_rows or []
    excel = _dispatch_excel()
    wb = None
    try:
        wb = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
        existing = {wb.Worksheets(i).Name.lower() for i in range(1, wb.Worksheets.Count + 1)}
        if sheet_name.lower() in existing:
            raise ValueError(f"Лист уже существует: {sheet_name}")
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheet_name
        for idx, header in enumerate(headers, start=1):
            ws.Cells(1, idx).Value = header
        inserted = 0
        for r_idx, row in enumerate(data_rows, start=2):
            empty = True
            for c_idx, value in enumerate(row, start=1):
                if value not in (None, ""):
                    empty = False
                ws.Cells(r_idx, c_idx).Value = value
            if not empty:
                inserted += 1
        wb.Save()
        return {
            "sheet": sheet_name,
            "columns": headers,
            "inserted": inserted,
            "engine": "excel-com",
        }
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()


def apply_values_com(
    path: Path,
    *,
    sheet_name: str,
    header_row: int,
    mode: str,
    rows: list[dict[str, Any]],
    key_columns: list[str],
) -> dict[str, Any]:
    """Write mapped rows into an existing sheet via Excel COM."""
    if not excel_com_available():
        raise RuntimeError("Excel COM недоступен")
    excel = _dispatch_excel()
    wb = None
    try:
        wb = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=False)
        try:
            ws = wb.Worksheets(sheet_name)
        except Exception as exc:  # noqa: BLE001
            raise KeyError(f"Лист не найден: {sheet_name}") from exc

        headers: list[str] = []
        name_to_col: dict[str, int] = {}
        used = ws.UsedRange
        last_col = int(used.Column + used.Columns.Count - 1) if used else 1
        for c in range(1, max(last_col, 1) + 1):
            value = ws.Cells(header_row, c).Value
            if value in (None, ""):
                continue
            name = str(value).strip()
            headers.append(name)
            name_to_col[name] = c
        if not name_to_col:
            raise ValueError("На листе не найдены заголовки колонок")

        stats = {"updated": 0, "inserted": 0, "cleared": 0, "unchanged": 0, "engine": "excel-com"}

        def find_next_row() -> int:
            last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1 if ws.UsedRange else header_row
            return max(last + 1, header_row + 1)

        if mode == "replace_sheet":
            last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1 if ws.UsedRange else header_row
            if last > header_row:
                ws.Range(ws.Cells(header_row + 1, 1), ws.Cells(last, max(last_col, 1))).ClearContents()
                stats["cleared"] = last - header_row
            write_row = header_row + 1
            for item in rows:
                for key, value in item.items():
                    if key in name_to_col:
                        ws.Cells(write_row, name_to_col[key]).Value = value
                write_row += 1
            stats["inserted"] = len(rows)

        elif mode == "append":
            write_row = find_next_row()
            for item in rows:
                for key, value in item.items():
                    if key in name_to_col:
                        ws.Cells(write_row, name_to_col[key]).Value = value
                write_row += 1
            stats["inserted"] = len(rows)

        else:  # upsert
            if not key_columns:
                raise ValueError("Для режима upsert укажите ключевые колонки")
            for key in key_columns:
                if key not in name_to_col:
                    raise ValueError(f"Ключевая колонка отсутствует на листе: {key}")

            def norm(v: Any) -> str:
                return "" if v is None else str(v).strip().lower()

            index: dict[tuple[str, ...], int] = {}
            last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1 if ws.UsedRange else header_row
            for r in range(header_row + 1, last + 1):
                key = tuple(norm(ws.Cells(r, name_to_col[k]).Value) for k in key_columns)
                if any(key):
                    index[key] = r

            next_row = find_next_row()
            for item in rows:
                key = tuple(norm(item.get(k)) for k in key_columns)
                if key in index:
                    r = index[key]
                    changed = False
                    for field, value in item.items():
                        if field in key_columns or field not in name_to_col:
                            continue
                        cell = ws.Cells(r, name_to_col[field])
                        if norm(cell.Value) != norm(value):
                            cell.Value = value
                            changed = True
                    if changed:
                        stats["updated"] += 1
                    else:
                        stats["unchanged"] += 1
                else:
                    for field, value in item.items():
                        if field in name_to_col:
                            ws.Cells(next_row, name_to_col[field]).Value = value
                    index[key] = next_row
                    next_row += 1
                    stats["inserted"] += 1

        wb.Save()
        return stats
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()
