"""Web Excel loader for TARGET.xlsx."""

from __future__ import annotations

import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from . import excel_service as xls

ROOT = Path(__file__).resolve().parent.parent
STATIC = Path(__file__).resolve().parent / "static"

app = FastAPI(title="Excel Loader", version="1.0.0")
app.mount("/static", StaticFiles(directory=STATIC), name="static")

xls.ensure_dirs()


class SuggestBody(BaseModel):
    source_columns: list[str] = Field(default_factory=list)
    target_columns: list[str] = Field(default_factory=list)


class ApplyBody(BaseModel):
    upload_id: str
    source_sheet: str | None = None
    target_sheet: str
    mode: str = "upsert"
    mapping: dict[str, str | None] = Field(default_factory=dict)
    key_columns: list[str] = Field(default_factory=list)


class RestoreBody(BaseModel):
    backup: str


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((STATIC / "index.html").read_text(encoding="utf-8"))


@app.get("/api/workbook")
def workbook_info() -> dict:
    try:
        sheets = xls.list_sheets()
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    return {
        "file": xls.TARGET_PATH.name,
        "path": str(xls.TARGET_PATH),
        "sheets": sheets,
    }


@app.get("/api/sheets/{sheet_name}/preview")
def sheet_preview(sheet_name: str, limit: int = 30) -> dict:
    try:
        return xls.preview_sheet(sheet_name, limit=limit)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc


@app.post("/api/upload")
async def upload_source(file: UploadFile = File(...)) -> dict:
    if not file.filename:
        raise HTTPException(status_code=400, detail="Файл не выбран")
    suffix = Path(file.filename).suffix.lower()
    if suffix not in {".xlsx", ".xlsm", ".xls", ".csv"}:
        raise HTTPException(status_code=400, detail="Нужен .xlsx / .csv")

    xls.ensure_dirs()
    upload_id = uuid.uuid4().hex
    dest = xls.UPLOAD_DIR / f"{upload_id}{suffix}"
    content = await file.read()
    dest.write_bytes(content)

    try:
        info = xls.inspect_upload(dest)
    except Exception as exc:  # noqa: BLE001
        dest.unlink(missing_ok=True)
        raise HTTPException(status_code=400, detail=f"Не удалось прочитать файл: {exc}") from exc

    return {
        "upload_id": upload_id,
        "filename": file.filename,
        "path_suffix": suffix,
        **info,
    }


@app.post("/api/suggest-mapping")
def suggest_mapping(body: SuggestBody) -> dict:
    return {"mapping": xls.suggest_mapping(body.source_columns, body.target_columns)}


@app.post("/api/apply")
def apply_load(body: ApplyBody) -> dict:
    matches = list(xls.UPLOAD_DIR.glob(f"{body.upload_id}.*"))
    if not matches:
        raise HTTPException(status_code=404, detail="Загрузка не найдена. Загрузите файл снова.")
    try:
        result = xls.apply_load(
            target_sheet=body.target_sheet,
            upload_path=matches[0],
            source_sheet=body.source_sheet,
            mode=body.mode,
            mapping=body.mapping,
            key_columns=body.key_columns,
        )
    except (ValueError, KeyError, FileNotFoundError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return result


@app.get("/api/download")
def download_target() -> FileResponse:
    path = xls.target_path()
    return FileResponse(
        path,
        filename=path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/backups")
def list_backups() -> dict:
    xls.ensure_dirs()
    files = sorted(xls.BACKUP_DIR.glob("TARGET_*.xlsx"), reverse=True)
    return {
        "backups": [
            {
                "name": f.name,
                "size": f.stat().st_size,
                "mtime": f.stat().st_mtime,
            }
            for f in files[:50]
        ]
    }


@app.post("/api/restore")
def restore(body: RestoreBody) -> dict:
    try:
        src = xls.restore_backup(body.backup)
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    return {"restored_from": src.name}


@app.post("/api/replace-whole-sheet")
async def replace_whole_sheet(
    target_sheet: str = Form(...),
    file: UploadFile = File(...),
    source_sheet: str | None = Form(None),
) -> dict:
    """Replace an entire worksheet from an uploaded workbook sheet (structure copy of values)."""
    if not file.filename:
        raise HTTPException(status_code=400, detail="Файл не выбран")
    suffix = Path(file.filename).suffix.lower()
    if suffix not in {".xlsx", ".xlsm", ".xls", ".csv"}:
        raise HTTPException(status_code=400, detail="Нужен .xlsx / .csv")

    xls.ensure_dirs()
    upload_id = uuid.uuid4().hex
    dest = xls.UPLOAD_DIR / f"{upload_id}{suffix}"
    dest.write_bytes(await file.read())

    try:
        info = xls.inspect_upload(dest)
        src_name = source_sheet or info["sheets"][0]["name"]
        source_cols = next(s["columns"] for s in info["sheets"] if s["name"] == src_name)
        target_info = next(s for s in xls.list_sheets() if s["name"] == target_sheet)
        target_cols = [c["name"] for c in target_info["columns"]]
        mapping = xls.suggest_mapping(source_cols, target_cols)
        result = xls.apply_load(
            target_sheet=target_sheet,
            upload_path=dest,
            source_sheet=None if suffix == ".csv" else src_name,
            mode="replace_sheet",
            mapping=mapping,
            key_columns=[],
        )
        result["mapping_used"] = mapping
        return result
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=str(exc)) from exc
