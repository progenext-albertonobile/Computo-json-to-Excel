#!/usr/bin/env python3
"""
main.py — FastAPI backend for branded XLSX generation.
Imports genera_computo and validate_bundle directly (no subprocess).
"""
from __future__ import annotations

import json
import os
import sys
import tempfile
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

sys.path.insert(0, str(Path(__file__).resolve().parent))
from genera_computo import DEFAULT_TEMPLATE, build, inspect_template_capacity

MAX_JSON_BYTES = 5 * 1024 * 1024   # 5 MB
MAX_XLSX_BYTES = 20 * 1024 * 1024  # 20 MB

app = FastAPI(title="Computo XLSX Generator", docs_url=None, redoc_url=None)

_static = Path(__file__).resolve().parent / "static"
if _static.exists():
    app.mount("/static", StaticFiles(directory=str(_static)), name="static")


@app.get("/", response_class=HTMLResponse)
async def root():
    html_path = Path(__file__).resolve().parent / "static" / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.post("/validate")
async def validate(
    json_file: UploadFile = File(...),
    template_file: UploadFile | None = File(default=None),
):
    _check_json(json_file)
    if template_file and template_file.filename:
        _check_xlsx(template_file)

    raw_json = await json_file.read()
    if len(raw_json) > MAX_JSON_BYTES:
        raise HTTPException(413, "JSON file too large (max 5 MB)")

    try:
        payload = json.loads(raw_json.decode("utf-8"))
    except Exception as exc:
        raise HTTPException(400, f"Invalid JSON: {exc}") from exc

    template_path, _tmp = await _resolve_template(template_file)
    try:
        report = inspect_template_capacity(payload, template_path)
    except Exception as exc:
        raise HTTPException(422, f"Validation error: {exc}") from exc
    finally:
        if _tmp:
            _tmp.cleanup()

    sections_out = []
    has_overflow = False
    for sec in report["sections"]:
        overflow = sec["overflow"]
        if overflow > 0:
            has_overflow = True
        sections_out.append(
            {
                "label": sec["label"],
                "items": sec["item_count"],
                "capacity": sec["capacity"],
                "overflow": overflow,
            }
        )

    ext = report["external"]
    if ext["overflow"] > 0:
        has_overflow = True

    return {
        "ok": not has_overflow,
        "sections": sections_out,
        "external": {
            "count": ext["count"],
            "capacity": ext["supported_rows"],
            "overflow": ext["overflow"],
        },
    }


@app.post("/generate")
async def generate(
    json_file: UploadFile = File(...),
    template_file: UploadFile | None = File(default=None),
):
    _check_json(json_file)
    if template_file and template_file.filename:
        _check_xlsx(template_file)

    raw_json = await json_file.read()
    if len(raw_json) > MAX_JSON_BYTES:
        raise HTTPException(413, "JSON file too large (max 5 MB)")

    try:
        payload = json.loads(raw_json.decode("utf-8"))
    except Exception as exc:
        raise HTTPException(400, f"Invalid JSON: {exc}") from exc

    template_path, _tmp = await _resolve_template(template_file)
    try:
        report = inspect_template_capacity(payload, template_path)
        for sec in report["sections"]:
            if sec["overflow"] > 0:
                raise HTTPException(
                    422,
                    f"Overflow in section '{sec['label']}': "
                    f"{sec['item_count']} items > {sec['capacity']} capacity",
                )
        if report["external"]["overflow"] > 0:
            ext = report["external"]
            raise HTTPException(
                422,
                f"Overflow in external items: "
                f"{ext['count']} items > {ext['supported_rows']} capacity",
            )

        wb = build(payload, template_path)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(500, f"Generation error: {exc}") from exc
    finally:
        if _tmp:
            _tmp.cleanup()

    cliente = payload.get("cliente", "computo")
    filename = f"COMPUTO_{cliente}.xlsx".replace(" ", "_")

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


def _check_json(f: UploadFile) -> None:
    if f is None or not f.filename:
        raise HTTPException(400, "JSON file required")
    if not f.filename.lower().endswith(".json"):
        raise HTTPException(415, "Only .json files accepted")


def _check_xlsx(f: UploadFile) -> None:
    if not f.filename.lower().endswith(".xlsx"):
        raise HTTPException(415, "Only .xlsx files accepted for template")


async def _resolve_template(template_file: UploadFile | None):
    if template_file and template_file.filename:
        raw = await template_file.read()
        if len(raw) > MAX_XLSX_BYTES:
            raise HTTPException(413, "Template file too large (max 20 MB)")
        tmp = tempfile.TemporaryDirectory()
        path = Path(tmp.name) / "template.xlsx"
        path.write_bytes(raw)
        return path, tmp
    return DEFAULT_TEMPLATE, None


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)
