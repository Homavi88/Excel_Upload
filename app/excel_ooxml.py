"""Safe OOXML worksheet creation that preserves drawings/media/links."""

from __future__ import annotations

import re
import shutil
import tempfile
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
SHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
SHEET_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

OOXML_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}


def _col_letter(index: int) -> str:
    """1-based column index to Excel letters."""
    result = []
    while index:
        index, rem = divmod(index - 1, 26)
        result.append(chr(65 + rem))
    return "".join(reversed(result))


def _cell_xml(col: int, row: int, value: Any) -> str:
    ref = f"{_col_letter(col)}{row}"
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"><v>{value}</v></c>'
    text = escape(str(value), {"'": "&apos;", '"': "&quot;"})
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def _sheet_xml(headers: list[str], rows: list[list[Any]]) -> str:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<worksheet xmlns="{NS_MAIN}" xmlns:r="{NS_REL}">',
        "<sheetData>",
    ]
    header_cells = "".join(_cell_xml(i, 1, h) for i, h in enumerate(headers, start=1))
    parts.append(f'<row r="1">{header_cells}</row>')
    for r_idx, row in enumerate(rows, start=2):
        cells = "".join(_cell_xml(i, r_idx, v) for i, v in enumerate(row, start=1))
        if cells:
            parts.append(f'<row r="{r_idx}">{cells}</row>')
    parts.append("</sheetData></worksheet>")
    return "".join(parts)


def _next_sheet_part_name(names: list[str]) -> tuple[str, int]:
    nums = []
    for name in names:
        m = re.fullmatch(r"xl/worksheets/sheet(\d+)\.xml", name)
        if m:
            nums.append(int(m.group(1)))
    next_n = (max(nums) + 1) if nums else 1
    return f"xl/worksheets/sheet{next_n}.xml", next_n


def _next_rid(rels_xml: str) -> str:
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels_xml)]
    return f"rId{(max(ids) + 1) if ids else 1}"


def _next_sheet_id(workbook_xml: str) -> int:
    ids = [int(x) for x in re.findall(r'sheetId="(\d+)"', workbook_xml)]
    return (max(ids) + 1) if ids else 1


def _existing_sheet_names(workbook_xml: str) -> list[str]:
    return re.findall(r'<sheet[^>]*name="([^"]+)"', workbook_xml)


def create_sheet_ooxml(
    path: Path,
    sheet_name: str,
    headers: list[str],
    data_rows: list[list[Any]] | None = None,
) -> dict[str, Any]:
    """Append a worksheet to an existing xlsx/xlsm without rewriting other parts."""
    suffix = path.suffix.lower()
    if suffix not in OOXML_SUFFIXES:
        raise ValueError("OOXML create_sheet поддерживает только .xlsx / .xlsm / .xltx / .xltm")

    data_rows = data_rows or []
    with ZipFile(path, "r") as zin:
        names = zin.namelist()
        if "xl/workbook.xml" not in names or "xl/_rels/workbook.xml.rels" not in names:
            raise ValueError("Некорректный Excel-файл: нет workbook.xml")
        workbook_xml = zin.read("xl/workbook.xml").decode("utf-8")
        rels_xml = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        ctypes_xml = zin.read("[Content_Types].xml").decode("utf-8")

        existing = {n.lower(): n for n in _existing_sheet_names(workbook_xml)}
        if sheet_name.lower() in existing:
            raise ValueError(f"Лист уже существует: {existing[sheet_name.lower()]}")

        part_name, _sheet_no = _next_sheet_part_name(names)
        rid = _next_rid(rels_xml)
        sheet_id = _next_sheet_id(workbook_xml)
        sheet_xml = _sheet_xml(headers, data_rows)

        # Update relationships
        if "</Relationships>" not in rels_xml:
            raise ValueError("Повреждён workbook.xml.rels")
        rel_line = (
            f'<Relationship Id="{rid}" Type="{SHEET_REL_TYPE}" '
            f'Target="worksheets/{Path(part_name).name}"/>'
        )
        rels_xml = rels_xml.replace("</Relationships>", f"{rel_line}</Relationships>")

        # Update workbook sheets list (append before closing </sheets>)
        sheet_tag = f'<sheet name="{escape(sheet_name)}" sheetId="{sheet_id}" r:id="{rid}"/>'
        if "</sheets>" not in workbook_xml:
            raise ValueError("В workbook.xml нет секции sheets")
        workbook_xml = workbook_xml.replace("</sheets>", f"{sheet_tag}</sheets>", 1)

        # Content types
        override = f'<Override PartName="/{part_name}" ContentType="{SHEET_CONTENT_TYPE}"/>'
        if override not in ctypes_xml:
            if "</Types>" not in ctypes_xml:
                raise ValueError("Повреждён [Content_Types].xml")
            ctypes_xml = ctypes_xml.replace("</Types>", f"{override}</Types>")

        tmp_fd, tmp_name = tempfile.mkstemp(suffix=suffix, dir=str(path.parent))
        tmp_path = Path(tmp_name)
        try:
            import os

            os.close(tmp_fd)
            with ZipFile(tmp_path, "w") as zout:
                for info in zin.infolist():
                    data = zin.read(info.filename)
                    if info.filename == "xl/workbook.xml":
                        data = workbook_xml.encode("utf-8")
                    elif info.filename == "xl/_rels/workbook.xml.rels":
                        data = rels_xml.encode("utf-8")
                    elif info.filename == "[Content_Types].xml":
                        data = ctypes_xml.encode("utf-8")
                    # Drop stale calc chain; Excel rebuilds it and avoids repair prompts
                    if info.filename == "xl/calcChain.xml":
                        continue
                    zout.writestr(info, data)
                zout.writestr(part_name, sheet_xml.encode("utf-8"))

            shutil.move(str(tmp_path), str(path))
        finally:
            if tmp_path.exists():
                tmp_path.unlink(missing_ok=True)

    return {
        "sheet": sheet_name,
        "part": part_name,
        "rows": len(data_rows),
        "columns": headers,
    }
