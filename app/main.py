"""Local web app for payment reconciliation."""

from __future__ import annotations

import os
import uuid
from contextlib import asynccontextmanager
from decimal import Decimal
from pathlib import Path
from tempfile import mkstemp
from urllib.parse import quote

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from app.demo import SAMPLE_DIR, write_sample_files
from app.errors import UserError
from app.excel_export import build_workbook
from app.fa import format_jalali, parse_decimal_input, quantize_money, suggest_role
from app.readers import describe_sheet, detect_header_row, read_tabular
from app.reconcile import job_from_payload, reconcile
from app.state import Export, Uploaded, exports, files

ROOT = Path(__file__).resolve().parent.parent
STATIC = ROOT / "static"
MAX_BYTES = 30 * 1024 * 1024
MAX_FILES = 30

ROLE_LABEL = {"due": "قابل‌پرداخت", "paid": "پرداخت‌شده"}


@asynccontextmanager
async def lifespan(_app: FastAPI):
    print("\nسامانه مغایرت‌گیری آماده است: http://127.0.0.1:8000\n", flush=True)
    yield


app = FastAPI(title="سامانه مغایرت‌گیری", lifespan=lifespan)


@app.exception_handler(UserError)
async def user_error(_request, exc: UserError):
    return JSONResponse({"detail": str(exc)}, status_code=400)


@app.get("/")
def index():
    return FileResponse(STATIC / "index.html")


@app.post("/api/upload")
async def upload(uploaded_files: list[UploadFile] = File(...)):
    if len(files) + len(uploaded_files) > MAX_FILES:
        raise UserError(f"بیش از {MAX_FILES} فایل را با هم نگه نمی‌داریم. چند فایل را حذف کنید.")
    created = []
    for item in uploaded_files:
        payload = await item.read()
        if len(payload) > MAX_BYTES:
            raise UserError(f"«{item.filename}» بزرگ‌تر از ۳۰ مگابایت است.")
        if not item.filename:
            raise UserError("نام یکی از فایل‌ها خالی است.")
        created.append(_store(item.filename, payload))
    return {"files": [public_file(item) for item in created]}


@app.post("/api/demo")
def load_demo():
    paths = sorted(SAMPLE_DIR.glob("*"))
    paths = [path for path in paths if path.suffix.lower() in {".xlsx", ".xls", ".csv"}]
    if len(paths) < 3:
        paths = write_sample_files(SAMPLE_DIR)
    created = []
    for path in paths:
        created.append(_store(path.name, path.read_bytes()))
    return {"files": [public_file(item) for item in created]}


@app.delete("/api/files/{file_id}")
def delete_file(file_id: str):
    files.pop(file_id, None)
    return {"ok": True}


@app.post("/api/files/{file_id}/preview")
def preview(file_id: str, body: dict):
    uploaded = files.get(file_id)
    if uploaded is None:
        raise UserError("این فایل دیگر در دسترس نیست. دوباره بارگذاری کنید.")
    sheet_name = str(body.get("sheet") or "")
    if sheet_name not in uploaded.sheets:
        raise UserError("این برگه در فایل نیست.")
    try:
        header_row = int(body.get("header_row") or 1)
    except (TypeError, ValueError) as exc:
        raise UserError("ردیف عنوان باید عدد باشد.") from exc
    described = describe_sheet(uploaded.sheets[sheet_name], header_row)
    described["name"] = sheet_name
    return described


@app.post("/api/reconcile")
def reconcile_files(body: dict):
    try:
        tolerance = parse_decimal_input(body.get("tolerance", "0"))
    except (TypeError, ValueError) as exc:
        raise UserError("حد چشم‌پوشی باید عدد باشد.") from exc
    jobs = []
    for item in body.get("files") or []:
        file_id = str(item.get("id") or "")
        uploaded = files.get(file_id)
        if uploaded is None:
            raise UserError("یکی از فایل‌ها پیدا نشد. صفحه را تازه کنید و دوباره بارگذاری کنید.")
        jobs.append(job_from_payload(item, uploaded.sheets, uploaded.filename))
    report = reconcile(jobs, tolerance)
    token = uuid.uuid4().hex
    filename = f"گزارش-مغایرت-{report.created_at.strftime('%Y%m%d-%H%M')}.xlsx"
    handle, raw_path = mkstemp(prefix="moghayerat-", suffix=".xlsx")
    os.close(handle)
    path = Path(raw_path)
    path.write_bytes(build_workbook(report))
    exports[token] = Export(path=path, filename=filename)
    due_total = quantize_money(sum((person.due for person in report.people), Decimal("0")))
    paid_total = quantize_money(sum((person.paid for person in report.people), Decimal("0")))
    balance = quantize_money(due_total - paid_total)
    return {
        "token": token,
        "filename": filename,
        "download_url": f"/api/download/{token}",
        "created_at": format_jalali(report.created_at),
        "summary": {
            "people": len(report.people),
            "due_total": format(due_total, "f"),
            "paid_total": format(paid_total, "f"),
            "balance_total": format(balance, "f"),
            "counts": report.counts(),
            "warnings": report.warnings,
            "sources": [
                {
                    "title": source.title,
                    "role": ROLE_LABEL.get(source.role, source.role),
                    "total": format(source.total, "f"),
                    "used_rows": source.used_rows,
                }
                for source in report.sources
            ],
        },
    }


@app.get("/api/download/{token}")
def download(token: str):
    item = exports.get(token)
    if item is None or not item.path.exists():
        raise UserError("فایل گزارش پیدا نشد. یک بار دیگر گزارش را بسازید.")
    quoted = quote(item.filename)
    return FileResponse(
        item.path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="moghayerat.xlsx",
        headers={"Content-Disposition": f"attachment; filename=\"moghayerat.xlsx\"; filename*=UTF-8''{quoted}"},
    )


def _store(filename: str, payload: bytes) -> Uploaded:
    sheets = read_tabular(filename, payload)
    if not any(rows for rows in sheets.values()):
        raise UserError(f"«{filename}» ردیفی ندارد.")
    uploaded = Uploaded(id=uuid.uuid4().hex, filename=filename, sheets=sheets)
    files[uploaded.id] = uploaded
    return uploaded


def public_file(uploaded: Uploaded) -> dict:
    sheets = []
    for name, rows in uploaded.sheets.items():
        described = describe_sheet(rows, detect_header_row(rows))
        described["name"] = name
        sheets.append(described)
    stem = uploaded.filename.rsplit(".", 1)[0]
    return {
        "id": uploaded.id,
        "filename": uploaded.filename,
        "role": suggest_role(uploaded.filename),
        "label": stem,
        "active_sheet": sheets[0]["name"] if sheets else "",
        "sheets": sheets,
    }


app.mount("/static", StaticFiles(directory=STATIC), name="static")
