#!/usr/bin/env python3
"""
main.py — FastAPI backend for branded XLSX generation.
Imports the data-driven `genera_computo.build(data)` engine directly.
"""
from __future__ import annotations

import json
import os
import sys
from io import BytesIO
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

sys.path.insert(0, str(Path(__file__).resolve().parent))
from genera_computo import build, build_sintesi

MAX_JSON_BYTES = 5 * 1024 * 1024  # 5 MB

app = FastAPI(title="Computo XLSX Generator", docs_url=None, redoc_url=None)

_static = Path(__file__).resolve().parent / "static"
if _static.exists():
    app.mount("/static", StaticFiles(directory=str(_static)), name="static")


@app.get("/", response_class=HTMLResponse)
async def root():
    html_path = Path(__file__).resolve().parent / "static" / "index.html"
    return HTMLResponse(html_path.read_text(encoding="utf-8"))


@app.post("/generate")
async def generate(json_file: UploadFile = File(...)):
    _check_json(json_file)

    raw_json = await json_file.read()
    if len(raw_json) > MAX_JSON_BYTES:
        raise HTTPException(413, "JSON file too large (max 5 MB)")

    try:
        payload = json.loads(raw_json.decode("utf-8"))
    except Exception as exc:
        raise HTTPException(400, f"Invalid JSON: {exc}") from exc

    if not isinstance(payload, dict) or "rows" not in payload:
        raise HTTPException(422, "JSON must be an object with a 'rows' array")

    try:
        wb = build(payload)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
    except Exception as exc:
        raise HTTPException(500, f"Generation error: {exc}") from exc

    cliente = (payload.get("metadata", {}) or {}).get("cliente", "computo")
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in cliente)
    filename = f"COMPUTO_{safe}.xlsx"

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.post("/generate/sintesi")
async def generate_sintesi(json_file: UploadFile = File(...)):
    _check_json(json_file)

    raw_json = await json_file.read()
    if len(raw_json) > MAX_JSON_BYTES:
        raise HTTPException(413, "JSON file too large (max 5 MB)")

    try:
        payload = json.loads(raw_json.decode("utf-8"))
    except Exception as exc:
        raise HTTPException(400, f"Invalid JSON: {exc}") from exc

    if not isinstance(payload, dict) or "rows" not in payload:
        raise HTTPException(422, "JSON must be an object with a 'rows' array")

    try:
        wb = build_sintesi(payload)
        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
    except Exception as exc:
        raise HTTPException(500, f"Generation error: {exc}") from exc

    cliente = (payload.get("metadata", {}) or {}).get("cliente", "stima")
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in cliente)
    filename = f"STIMA_{safe}.xlsx"

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


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)
