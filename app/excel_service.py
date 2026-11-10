"""Excel load / preview / merge helpers for TARGET.xlsx."""

from __future__ import annotations

import csv
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path(__file__).resolve().parent.parent
TARGET_PATH = ROOT / "TARGET.xlsx"
BACKUP_DIR = ROOT / "backups"
UPLOAD_DIR = ROOT / "uploads"

# Sheets where the real header is not on row 1
HEADER_ROW_OVERRIDES: dict[str, int] = {
    "CF by blend Primary (2)": 2,
}


def ensure_dirs() -> None:
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


def target_path() -> Path:
    if not TARGET_PATH.exists():
        raise FileNotFoundError(f"Не найден файл: {TARGET_PATH}")
    return TARGET_PATH


def backup_target() -> Path:
    ensure_dirs()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = BACKUP_DIR / f"TARGET_{stamp}.xlsx"
    shutil.copy2(target_path(), dest)
    return dest


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


def list_sheets() -> list[dict[str, Any]]:
    wb = load_workbook(target_path(), data_only=False, read_only=True)
    result: list[dict[str, Any]] = []
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            header_row = _header_row_for(name)
            # read_only worksheets need iteration; reopen without read_only for headers
            result.append({"name": name, "header_row": header_row})
    finally:
        wb.close()

    wb = load_workbook(target_path(), data_only=False)
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
    wb = load_workbook(target_path(), data_only=True)
    try:
        if sheet_name not in wb.sheetnames:
            raise KeyError(f"Лист не найден: {sheet_name}")
        ws = wb[sheet_name]
        header_row = _header_row_for(sheet_name)
        headers = _read_headers(ws, header_row)
        col_indexes = [h["index"] + 1 for h in headers]
        rows: list[dict[str, Any]] = []
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
    wb = load_workbook(path, data_only=True, read_only=True)
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
    if suffix in {".xlsx", ".xlsm"}:
        return _records_from_excel(path, sheet_name)
    if suffix == ".xls":
        raise ValueError("Формат .xls не поддерживается. Сохраните файл как .xlsx или .csv")
    raise ValueError("Поддерживаются только .xlsx / .xlsm / .csv")


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

    if suffix == ".xls":
        raise ValueError("Формат .xls не поддерживается. Сохраните файл как .xlsx или .csv")

    wb = load_workbook(path, data_only=True, read_only=True)
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

    backup = backup_target()
    wb = load_workbook(target_path())
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

        # Build incoming rows in target column order for mapped fields only
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

        wb.save(target_path())
        stats["columns_written"] = list(active_map.keys())
        stats["header_names"] = header_names
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
    """Create a new worksheet. Optionally seed headers/data from an uploaded file."""
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

    backup = backup_target()
    wb = load_workbook(target_path())
    try:
        existing_lower = {s.lower(): s for s in wb.sheetnames}
        if sheet_name.lower() in existing_lower:
            raise ValueError(f"Лист уже существует: {existing_lower[sheet_name.lower()]}")

        ws = wb.create_sheet(title=sheet_name)
        for idx, col_name in enumerate(columns, start=1):
            ws.cell(1, idx).value = col_name

        inserted = 0
        if fill_data and source_rows:
            write_row = 2
            source_col_set = set(source_columns)
            for row in source_rows:
                empty = True
                for idx, col_name in enumerate(columns, start=1):
                    if col_name not in source_col_set:
                        continue
                    value = _clean_cell(row.get(col_name))
                    if value not in (None, ""):
                        empty = False
                    ws.cell(write_row, idx).value = value
                if not empty:
                    inserted += 1
                    write_row += 1

        wb.save(target_path())
        return {
            "sheet": sheet_name,
            "columns": columns,
            "inserted": inserted,
            "backup": str(backup.relative_to(ROOT)),
            "from_upload": upload_path is not None,
        }
    finally:
        wb.close()
