"""Excel load / preview / merge helpers for TARGET.xlsx."""

from __future__ import annotations

import csv
import json
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from . import excel_com
from . import excel_ooxml

ROOT = Path(__file__).resolve().parent.parent
DEFAULT_TARGET = ROOT / "TARGET.xlsx"
DATA_DIR = ROOT / "data"
BACKUP_DIR = ROOT / "backups"
UPLOAD_DIR = ROOT / "uploads"
ACTIVE_META = DATA_DIR / "active.json"
WORKING_NAME = "working.xlsx"

SUPPORTED_WORKBOOK_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx", ".xltm"}
OPENPYXL_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
SOURCE_SUFFIXES = {".xlsx", ".xlsm", ".xls", ".xlsb", ".xltx", ".xltm", ".csv"}

# Backward-compatible alias used by older callers/tests
TARGET_PATH = DEFAULT_TARGET

# Sheets where the real header is not on row 1
HEADER_ROW_OVERRIDES: dict[str, int] = {
    "CF by blend Primary (2)": 2,
}


def ensure_dirs() -> None:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def _safe_stem(name: str) -> str:
    stem = Path(name).stem.strip() or "workbook"
    cleaned = re.sub(r"[^\w\-а-яА-ЯёЁ]+", "_", stem, flags=re.UNICODE).strip("_")
    return (cleaned or "workbook")[:40]


def _read_active_meta() -> dict[str, Any] | None:
    if not ACTIVE_META.exists():
        return None
    try:
        data = json.loads(ACTIVE_META.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict):
        return None
    return data


def _write_active_meta(original_name: str, path: Path, *, inplace: bool) -> None:
    ensure_dirs()
    resolved = path.resolve()
    ACTIVE_META.write_text(
        json.dumps(
            {
                "original_name": original_name,
                "path": str(resolved),
                "inplace": bool(inplace),
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def active_info() -> dict[str, Any]:
    """Return currently selected workbook metadata."""
    ensure_dirs()
    meta = _read_active_meta()
    if meta and meta.get("path"):
        raw = Path(meta["path"])
        path = raw if raw.is_absolute() else (ROOT / raw)
        if path.exists():
            inplace = bool(meta.get("inplace", not _is_working_copy(path)))
            return {
                "file": meta.get("original_name") or path.name,
                "path": str(path.resolve()),
                "selected": True,
                "exists": True,
                "inplace": inplace,
            }

    if DEFAULT_TARGET.exists():
        return {
            "file": DEFAULT_TARGET.name,
            "path": str(DEFAULT_TARGET.resolve()),
            "selected": False,
            "exists": True,
            "inplace": True,
        }

    return {
        "file": None,
        "path": None,
        "selected": False,
        "exists": False,
        "inplace": False,
    }


def _is_working_copy(path: Path) -> bool:
    try:
        return path.resolve() == (DATA_DIR / WORKING_NAME).resolve()
    except OSError:
        return False


def target_path() -> Path:
    info = active_info()
    if not info["exists"] or not info["path"]:
        raise FileNotFoundError(
            "Excel-файл не выбран. Нажмите «Открыть Excel» и выберите файл в проводнике."
        )
    path = Path(info["path"])
    if not path.exists():
        raise FileNotFoundError(f"Не найден файл: {path}")
    return path


def _validate_workbook(path: Path) -> int:
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_WORKBOOK_SUFFIXES:
        raise ValueError(
            "Поддерживаются Excel-файлы: .xlsx .xlsm .xls .xlsb .xltx .xltm"
        )
    if suffix in OPENPYXL_SUFFIXES:
        wb = load_workbook(path, read_only=True, keep_vba=(suffix in {".xlsm", ".xltm"}))
        try:
            sheet_count = len(wb.sheetnames)
            if sheet_count < 1:
                raise ValueError("В файле нет листов")
            return sheet_count
        finally:
            wb.close()
    # .xls / .xlsb require Excel COM on Windows
    if excel_com.excel_com_available():
        excel = excel_com._dispatch_excel()
        wb = None
        try:
            wb = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
            count = int(wb.Worksheets.Count)
            if count < 1:
                raise ValueError("В файле нет листов")
            return count
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            excel.Quit()
    raise ValueError(
        f"Формат {suffix} требует установленный Microsoft Excel (Windows)."
    )


def open_workbook_from_path(path: Path) -> dict[str, Any]:
    """Open a real filesystem path and write all later changes back to it."""
    ensure_dirs()
    resolved = path.expanduser().resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"Не найден файл: {resolved}")
    sheet_count = _validate_workbook(resolved)
    _write_active_meta(resolved.name, resolved, inplace=True)
    return {
        "file": resolved.name,
        "path": str(resolved),
        "sheets": sheet_count,
        "selected": True,
        "exists": True,
        "inplace": True,
    }


def browse_workbook() -> dict[str, Any]:
    """Open native Explorer / file dialog and select a workbook path."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(
            "Не удалось открыть системный проводник. Используйте Python с поддержкой tkinter."
        ) from exc

    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except tk.TclError:
        pass
    root.update()
    chosen = filedialog.askopenfilename(
        parent=root,
        title="Выберите Excel-файл",
        filetypes=[
            ("Excel files", "*.xlsx *.xlsm *.xls *.xlsb *.xltx *.xltm"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()
    if not chosen:
        raise ValueError("Файл не выбран")
    return open_workbook_from_path(Path(chosen))


def open_workbook_file(source: Path, original_name: str) -> dict[str, Any]:
    """Open an uploaded browser copy (changes stay in data/working.* until download)."""
    ensure_dirs()
    sheet_count = _validate_workbook(source)
    suffix = source.suffix.lower() if source.suffix else ".xlsx"
    if suffix not in SUPPORTED_WORKBOOK_SUFFIXES:
        suffix = ".xlsx"

    working = DATA_DIR / f"working{suffix}"
    # remove previous working copies with other extensions
    for old in DATA_DIR.glob("working.*"):
        if old.name != "active.json":
            old.unlink(missing_ok=True)
    shutil.copy2(source, working)
    display_name = Path(original_name).name or working.name
    if Path(display_name).suffix.lower() not in SUPPORTED_WORKBOOK_SUFFIXES:
        display_name = f"{display_name}{suffix}"
    _write_active_meta(display_name, working, inplace=False)
    return {
        "file": display_name,
        "path": str(working.resolve()),
        "sheets": sheet_count,
        "selected": True,
        "exists": True,
        "inplace": False,
    }


def save_workbook(wb: Any) -> Path:
    """Persist workbook atomically to the active path."""
    path = target_path()
    suffix = path.suffix.lower() or ".xlsx"
    tmp = path.with_name(f".{path.stem}.tmp{suffix}")
    try:
        wb.save(tmp)
        tmp.replace(path)
    except PermissionError as exc:
        tmp.unlink(missing_ok=True)
        raise PermissionError(
            f"Не удалось сохранить «{path.name}». Закройте файл в Excel и повторите."
        ) from exc
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    return path


def backup_target() -> Path:
    ensure_dirs()
    info = active_info()
    src = target_path()
    stem = _safe_stem(info.get("file") or src.name)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = BACKUP_DIR / f"{stem}_{stamp}{src.suffix.lower() or '.xlsx'}"
    shutil.copy2(src, dest)
    return dest


def list_backups() -> list[dict[str, Any]]:
    ensure_dirs()
    files = sorted(
        [p for p in BACKUP_DIR.iterdir() if p.suffix.lower() in SUPPORTED_WORKBOOK_SUFFIXES],
        reverse=True,
    )
    return [
        {
            "name": f.name,
            "size": f.stat().st_size,
            "mtime": f.stat().st_mtime,
        }
        for f in files[:50]
    ]


def _norm(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def _header_row_for(sheet_name: str) -> int:
    return HEADER_ROW_OVERRIDES.get(sheet_name, 1)


def _read_headers(ws: Worksheet, header_row: int) -> list[dict[str, Any]]:
    headers: list[dict[str, Any]] = []
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(header_row, col)
        name = _norm(cell.value)
        if name:
            headers.append(
                {
                    "index": col - 1,
                    "col": get_column_letter(col),
                    "name": name,
                }
            )
    return headers


def _used_row_count(ws: Worksheet, header_row: int) -> int:
    count = 0
    for row in range(header_row + 1, ws.max_row + 1):
        if any(ws.cell(row, col).value not in (None, "") for col in range(1, ws.max_column + 1)):
            count += 1
    return count


def _keep_vba_for(path: Path) -> bool:
    return path.suffix.lower() in {".xlsm", ".xltm"}


def _list_sheets_com(path: Path) -> list[dict[str, Any]]:
    excel = excel_com._dispatch_excel()
    wb = None
    try:
        wb = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
        enriched: list[dict[str, Any]] = []
        for i in range(1, wb.Worksheets.Count + 1):
            ws = wb.Worksheets(i)
            name = str(ws.Name)
            header_row = _header_row_for(name)
            headers = []
            used = ws.UsedRange
            last_col = int(used.Column + used.Columns.Count - 1) if used else 1
            last_row = int(used.Row + used.Rows.Count - 1) if used else header_row
            for c in range(1, max(last_col, 1) + 1):
                value = ws.Cells(header_row, c).Value
                if value in (None, ""):
                    continue
                headers.append(
                    {
                        "index": c - 1,
                        "col": get_column_letter(c),
                        "name": _norm(value),
                    }
                )
            row_count = max(0, last_row - header_row) if headers else 0
            enriched.append(
                {
                    "name": name,
                    "header_row": header_row,
                    "columns": headers,
                    "row_count": row_count,
                    "max_row": last_row,
                    "max_col": last_col,
                }
            )
        return enriched
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()


def list_sheets() -> list[dict[str, Any]]:
    path = target_path()
    if path.suffix.lower() not in OPENPYXL_SUFFIXES:
        if not excel_com.excel_com_available():
            raise ValueError("Для этого формата нужен Microsoft Excel на Windows")
        return _list_sheets_com(path)

    wb = load_workbook(path, data_only=False, read_only=True, keep_vba=_keep_vba_for(path))
    result: list[dict[str, Any]] = []
    try:
        for name in wb.sheetnames:
            header_row = _header_row_for(name)
            result.append({"name": name, "header_row": header_row})
    finally:
        wb.close()

    wb = load_workbook(path, data_only=False, keep_vba=_keep_vba_for(path))
    try:
        enriched: list[dict[str, Any]] = []
        for item in result:
            name = item["name"]
            ws = wb[name]
            header_row = item["header_row"]
            headers = _read_headers(ws, header_row)
            enriched.append(
                {
                    "name": name,
                    "header_row": header_row,
                    "columns": headers,
                    "row_count": _used_row_count(ws, header_row),
                    "max_row": ws.max_row,
                    "max_col": ws.max_column,
                }
            )
        return enriched
    finally:
        wb.close()


def preview_sheet(sheet_name: str, limit: int = 30) -> dict[str, Any]:
    path = target_path()
    if path.suffix.lower() not in OPENPYXL_SUFFIXES:
        sheets = {s["name"]: s for s in list_sheets()}
        if sheet_name not in sheets:
            raise KeyError(f"Лист не найден: {sheet_name}")
        meta = sheets[sheet_name]
        excel = excel_com._dispatch_excel()
        wb = None
        try:
            wb = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
            ws = wb.Worksheets(sheet_name)
            header_row = meta["header_row"]
            columns = [c["name"] for c in meta["columns"]]
            col_indexes = [c["index"] + 1 for c in meta["columns"]]
            rows: list[dict[str, Any]] = []
            last = int(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1) if ws.UsedRange else header_row
            for row_idx in range(header_row + 1, last + 1):
                values = [ws.Cells(row_idx, col).Value for col in col_indexes]
                if all(v in (None, "") for v in values):
                    continue
                row_obj = {columns[i]: values[i] for i in range(len(columns))}
                row_obj["__row"] = row_idx
                rows.append(row_obj)
                if len(rows) >= limit:
                    break
            return {
                "sheet": sheet_name,
                "header_row": header_row,
                "columns": columns,
                "rows": rows,
                "shown": len(rows),
                "total_estimate": meta["row_count"],
            }
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            excel.Quit()

    wb = load_workbook(path, data_only=True, keep_vba=_keep_vba_for(path))
    try:
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Лист не найден: {sheet_name}")
        ws = wb[sheet_name]
        header_row = _header_row_for(sheet_name)
        headers = _read_headers(ws, header_row)
        col_indexes = [h["index"] + 1 for h in headers]
        rows = []
        for row_idx in range(header_row + 1, ws.max_row + 1):
            values = [ws.cell(row_idx, col).value for col in col_indexes]
            if all(v in (None, "") for v in values):
                continue
            row_obj = {headers[i]["name"]: values[i] for i in range(len(headers))}
            row_obj["__row"] = row_idx
            rows.append(row_obj)
            if len(rows) >= limit:
                break
        return {
            "sheet": sheet_name,
            "header_row": header_row,
            "columns": [h["name"] for h in headers],
            "rows": rows,
            "shown": len(rows),
            "total_estimate": _used_row_count(ws, header_row),
        }
    finally:
        wb.close()


def _clean_cell(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip()
        return text if text != "" else None
    return value


def _records_from_csv(path: Path) -> tuple[list[str], list[dict[str, Any]]]:
    last_error: Exception | None = None
    for encoding in ("utf-8-sig", "utf-8", "cp1251", "latin-1"):
        try:
            with path.open("r", encoding=encoding, newline="") as handle:
                sample = handle.read(4096)
                handle.seek(0)
                try:
                    dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
                except csv.Error:
                    dialect = csv.excel
                reader = csv.DictReader(handle, dialect=dialect)
                if not reader.fieldnames:
                    return [], []
                columns = [_norm(c) for c in reader.fieldnames]
                rows: list[dict[str, Any]] = []
                for raw in reader:
                    item = {
                        columns[i]: _clean_cell(raw.get(original))
                        for i, original in enumerate(reader.fieldnames)
                        if i < len(columns)
                    }
                    if any(v not in (None, "") for v in item.values()):
                        rows.append(item)
                return columns, rows
        except UnicodeDecodeError as exc:
            last_error = exc
            continue
    raise ValueError(f"Не удалось прочитать CSV: {last_error}")


def _records_from_excel(path: Path, sheet_name: str | None = None) -> tuple[list[str], list[dict[str, Any]]]:
    wb = load_workbook(path, data_only=True, read_only=True, keep_vba=_keep_vba_for(path))
    try:
        if sheet_name is None:
            ws = wb[wb.sheetnames[0]]
        else:
            if sheet_name not in wb.sheetnames:
                raise KeyError(f"Лист источника не найден: {sheet_name}")
            ws = wb[sheet_name]

        rows_iter = ws.iter_rows(values_only=True)
        try:
            header_raw = next(rows_iter)
        except StopIteration:
            return [], []

        columns: list[str] = []
        col_indexes: list[int] = []
        for idx, value in enumerate(header_raw):
            name = _norm(value)
            if name:
                columns.append(name)
                col_indexes.append(idx)

        records: list[dict[str, Any]] = []
        for row in rows_iter:
            item = {
                columns[i]: _clean_cell(row[col_indexes[i]] if col_indexes[i] < len(row) else None)
                for i in range(len(columns))
            }
            if any(v not in (None, "") for v in item.values()):
                records.append(item)
        return columns, records
    finally:
        wb.close()


def _records_from_upload(path: Path, sheet_name: str | None = None) -> tuple[list[str], list[dict[str, Any]]]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return _records_from_csv(path)
    if suffix in OPENPYXL_SUFFIXES:
        return _records_from_excel(path, sheet_name)
    if suffix in {".xls", ".xlsb"} and excel_com.excel_com_available():
        # Save a temporary xlsx snapshot via Excel for reading
        excel = excel_com._dispatch_excel()
        wb = None
        tmp = UPLOAD_DIR / f"conv_{path.stem}.xlsx"
        try:
            wb = excel.Workbooks.Open(str(path.resolve()), UpdateLinks=0, ReadOnly=True)
            # 51 = xlOpenXMLWorkbook (.xlsx)
            wb.SaveAs(str(tmp.resolve()), FileFormat=51)
            wb.Close(SaveChanges=False)
            wb = None
            return _records_from_excel(tmp, sheet_name)
        finally:
            if wb is not None:
                wb.Close(SaveChanges=False)
            excel.Quit()
    raise ValueError("Поддерживаются .xlsx .xlsm .xls .xlsb .xltx .xltm .csv")


def inspect_upload(path: Path) -> dict[str, Any]:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        columns, rows = _records_from_upload(path)
        preview = [{k: ("" if v is None else v) for k, v in row.items()} for row in rows[:15]]
        return {
            "sheets": [
                {
                    "name": path.name,
                    "columns": columns,
                    "row_count": len(rows),
                    "preview": preview,
                }
            ]
        }

    if suffix in {".xls", ".xlsb"}:
        columns, rows = _records_from_upload(path)
        return {
            "sheets": [
                {
                    "name": path.stem,
                    "columns": columns,
                    "row_count": len(rows),
                    "preview": [{k: ("" if v is None else v) for k, v in row.items()} for row in rows[:15]],
                }
            ]
        }

    wb = load_workbook(path, data_only=True, read_only=True, keep_vba=_keep_vba_for(path))
    try:
        sheets: list[dict[str, Any]] = []
        for name in wb.sheetnames:
            columns, rows = _records_from_excel(path, name)
            preview = [{k: ("" if v is None else v) for k, v in row.items()} for row in rows[:15]]
            sheets.append(
                {
                    "name": name,
                    "columns": columns,
                    "row_count": len(rows),
                    "preview": preview,
                }
            )
        return {"sheets": sheets}
    finally:
        wb.close()


def suggest_mapping(source_columns: list[str], target_columns: list[str]) -> dict[str, str | None]:
    source_lower = {c.lower(): c for c in source_columns}
    mapping: dict[str, str | None] = {}
    for target in target_columns:
        key = target.lower()
        if key in source_lower:
            mapping[target] = source_lower[key]
            continue
        # soft match: ignore punctuation / spaces
        compact_t = re.sub(r"[^a-z0-9а-яё]+", "", key)
        found = None
        for src in source_columns:
            compact_s = re.sub(r"[^a-z0-9а-яё]+", "", src.lower())
            if compact_s and compact_s == compact_t:
                found = src
                break
        mapping[target] = found
    return mapping


def _cell_key(values: tuple[Any, ...]) -> tuple[str, ...]:
    return tuple(_norm(v).lower() for v in values)


def apply_load(
    *,
    target_sheet: str,
    upload_path: Path,
    source_sheet: str | None,
    mode: str,
    mapping: dict[str, str | None],
    key_columns: list[str],
) -> dict[str, Any]:
    if mode not in {"replace_sheet", "append", "upsert"}:
        raise ValueError("mode должен быть replace_sheet | append | upsert")

    source_columns, source_rows = _records_from_upload(upload_path, source_sheet)
    source_col_set = set(source_columns)
    active_map = {t: s for t, s in mapping.items() if s}
    if not active_map:
        raise ValueError("Нужно сопоставить хотя бы одну колонку")

    missing_src = [s for s in active_map.values() if s not in source_col_set]
    if missing_src:
        raise ValueError(f"В источнике нет колонок: {', '.join(missing_src)}")

    incoming: list[dict[str, Any]] = []
    for row in source_rows:
        item: dict[str, Any] = {}
        empty = True
        for target_col, source_col in active_map.items():
            value = _clean_cell(row.get(source_col))
            item[target_col] = value
            if value not in (None, ""):
                empty = False
        if not empty:
            incoming.append(item)

    backup = backup_target()
    path = target_path()

    if excel_com.excel_com_available():
        stats = excel_com.apply_values_com(
            path,
            sheet_name=target_sheet,
            header_row=_header_row_for(target_sheet),
            mode=mode,
            rows=incoming,
            key_columns=key_columns,
        )
        stats.update(
            {
                "mode": mode,
                "target_sheet": target_sheet,
                "backup": str(backup.relative_to(ROOT)),
                "incoming_rows": len(incoming),
                "columns_written": list(active_map.keys()),
                "saved_to": str(path),
            }
        )
        return stats

    if path.suffix.lower() not in OPENPYXL_SUFFIXES:
        raise ValueError(
            "Для этого формата нужен Microsoft Excel на Windows (пакет pywin32)."
        )

    wb = load_workbook(path, keep_vba=_keep_vba_for(path))
    try:
        if target_sheet not in wb.sheetnames:
            raise KeyError(f"Лист не найден: {target_sheet}")
        ws = wb[target_sheet]
        header_row = _header_row_for(target_sheet)
        headers = _read_headers(ws, header_row)
        header_names = [h["name"] for h in headers]
        name_to_col = {h["name"]: h["index"] + 1 for h in headers}

        for target_col in active_map:
            if target_col not in name_to_col:
                raise ValueError(f"На целевом листе нет колонки: {target_col}")

        stats = {
            "mode": mode,
            "target_sheet": target_sheet,
            "backup": str(backup.relative_to(ROOT)),
            "incoming_rows": len(incoming),
            "updated": 0,
            "inserted": 0,
            "cleared": 0,
            "unchanged": 0,
        }

        if mode == "replace_sheet":
            # Clear only data rows; keep header and any title rows above header
            cleared = 0
            for row_idx in range(ws.max_row, header_row, -1):
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row_idx, col)
                    if cell.value is not None:
                        cell.value = None
                        cleared += 1
            stats["cleared"] = cleared
            write_row = header_row + 1
            for item in incoming:
                for target_col, value in item.items():
                    ws.cell(write_row, name_to_col[target_col]).value = value
                write_row += 1
            stats["inserted"] = len(incoming)

        elif mode == "append":
            write_row = ws.max_row + 1
            # find first empty data row after header
            for row_idx in range(header_row + 1, ws.max_row + 2):
                if all(
                    ws.cell(row_idx, col).value in (None, "")
                    for col in range(1, ws.max_column + 1)
                ):
                    write_row = row_idx
                    break
            for item in incoming:
                for target_col, value in item.items():
                    ws.cell(write_row, name_to_col[target_col]).value = value
                write_row += 1
            stats["inserted"] = len(incoming)

        else:  # upsert
            if not key_columns:
                raise ValueError("Для режима upsert укажите ключевые колонки")
            for key in key_columns:
                if key not in name_to_col:
                    raise ValueError(f"Ключевая колонка отсутствует на листе: {key}")
                if key not in active_map:
                    raise ValueError(f"Ключевая колонка должна быть в маппинге: {key}")

            # Index existing rows by key
            existing: dict[tuple[str, ...], int] = {}
            for row_idx in range(header_row + 1, ws.max_row + 1):
                key_vals = tuple(ws.cell(row_idx, name_to_col[k]).value for k in key_columns)
                if all(v in (None, "") for v in key_vals):
                    continue
                existing[_cell_key(key_vals)] = row_idx

            next_row = ws.max_row + 1
            for row_idx in range(header_row + 1, ws.max_row + 2):
                if all(
                    ws.cell(row_idx, col).value in (None, "")
                    for col in range(1, ws.max_column + 1)
                ):
                    next_row = row_idx
                    break

            for item in incoming:
                key_vals = tuple(item.get(k) for k in key_columns)
                key = _cell_key(key_vals)
                if key in existing:
                    row_idx = existing[key]
                    changed = False
                    for target_col, value in item.items():
                        if target_col in key_columns:
                            continue
                        cell = ws.cell(row_idx, name_to_col[target_col])
                        # Mapped columns are written intentionally, including over formulas.
                        if _norm(cell.value) != _norm(value):
                            cell.value = value
                            changed = True
                    if changed:
                        stats["updated"] += 1
                    else:
                        stats["unchanged"] += 1
                else:
                    for target_col, value in item.items():
                        ws.cell(next_row, name_to_col[target_col]).value = value
                    existing[key] = next_row
                    next_row += 1
                    stats["inserted"] += 1

        saved_to = save_workbook(wb)
        stats["columns_written"] = list(active_map.keys())
        stats["header_names"] = header_names
        stats["saved_to"] = str(saved_to)
        return stats
    finally:
        wb.close()


def restore_backup(backup_name: str) -> Path:
    src = BACKUP_DIR / Path(backup_name).name
    if not src.exists():
        raise FileNotFoundError("Бэкап не найден")
    shutil.copy2(src, target_path())
    return src


_INVALID_SHEET_CHARS = set(r"[]:*?/\\")


def _validate_sheet_name(name: str) -> str:
    cleaned = _norm(name)
    if not cleaned:
        raise ValueError("Укажите название листа")
    if len(cleaned) > 31:
        raise ValueError("Название листа не длиннее 31 символа")
    if any(ch in _INVALID_SHEET_CHARS for ch in cleaned):
        raise ValueError(r"В названии листа нельзя использовать символы: [ ] : * ? / \ ")
    if cleaned.startswith("'") or cleaned.endswith("'"):
        raise ValueError("Название листа не должно начинаться или заканчиваться кавычкой")
    return cleaned


def _normalize_columns(columns: list[str] | None) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for raw in columns or []:
        name = _norm(raw)
        if not name:
            continue
        key = name.lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(name)
    return result


def create_sheet(
    *,
    name: str,
    columns: list[str] | None = None,
    upload_path: Path | None = None,
    source_sheet: str | None = None,
    fill_data: bool = True,
) -> dict[str, Any]:
    """Create a new worksheet without rewriting/stripping existing workbook parts."""
    sheet_name = _validate_sheet_name(name)
    columns = _normalize_columns(columns)

    source_columns: list[str] = []
    source_rows: list[dict[str, Any]] = []
    if upload_path is not None:
        source_columns, source_rows = _records_from_upload(upload_path, source_sheet)
        if not columns:
            columns = [c for c in source_columns if c]

    if not columns:
        raise ValueError("Укажите колонки или загрузите файл-источник с заголовками")

    data_matrix: list[list[Any]] = []
    inserted = 0
    if fill_data and source_rows:
        source_col_set = set(source_columns)
        for row in source_rows:
            values = []
            empty = True
            for col_name in columns:
                value = _clean_cell(row.get(col_name)) if col_name in source_col_set else None
                values.append(value)
                if value not in (None, ""):
                    empty = False
            if not empty:
                data_matrix.append(values)
                inserted += 1

    backup = backup_target()
    path = target_path()
    suffix = path.suffix.lower()
    engine = "ooxml"

    if suffix in excel_ooxml.OOXML_SUFFIXES:
        # Surgical OOXML edit preserves drawings/media/external links.
        excel_ooxml.create_sheet_ooxml(path, sheet_name, columns, data_matrix)
        engine = "ooxml"
    elif excel_com.excel_com_available():
        excel_com.create_sheet_com(path, sheet_name, columns, data_matrix)
        engine = "excel-com"
    else:
        raise ValueError(
            f"Нельзя безопасно изменить {suffix} без Microsoft Excel на Windows."
        )

    return {
        "sheet": sheet_name,
        "columns": columns,
        "inserted": inserted,
        "backup": str(backup.relative_to(ROOT)),
        "from_upload": upload_path is not None,
        "saved_to": str(path),
        "inplace": bool(active_info().get("inplace")),
        "engine": engine,
    }
