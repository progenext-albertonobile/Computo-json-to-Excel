#!/usr/bin/env python3
from __future__ import annotations

import json
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st

from genera_computo import build, inspect_template_capacity


DEFAULT_TEMPLATE = Path("templates") / "Computo preliminare-V3.xlsx"

st.set_page_config(page_title="Computo XLSX Generator", layout="wide")
st.title("Computo XLSX Generator")
st.caption("Upload JSON -> validate -> generate XLSX")

uploaded_json = st.file_uploader("JSON input", type=["json"])
uploaded_template = st.file_uploader("Template XLSX (optional)", type=["xlsx"])

if uploaded_json is None:
    st.info("Upload a JSON file to start.")
    st.stop()

try:
    payload = json.loads(uploaded_json.getvalue().decode("utf-8"))
except Exception as exc:
    st.error(f"Invalid JSON: {exc}")
    st.stop()


def _run(template_path: Path):
    report = inspect_template_capacity(payload, template_path)
    section_rows = []
    has_overflow = False

    for sec in report["sections"]:
        section_rows.append(
            {
                "section": sec["index"] + 1,
                "label": sec["label"],
                "items": sec["item_count"],
                "capacity": sec["capacity"],
                "overflow": sec["overflow"],
            }
        )
        if sec["overflow"] > 0:
            has_overflow = True

    ext = report["external"]
    if ext["overflow"] > 0:
        has_overflow = True

    wb = build(payload, template_path)
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return section_rows, ext, has_overflow, out


try:
    if uploaded_template is not None:
        with TemporaryDirectory() as td:
            tmp_template = Path(td) / "template.xlsx"
            tmp_template.write_bytes(uploaded_template.getvalue())
            section_rows, ext, has_overflow, out = _run(tmp_template)
    else:
        if not DEFAULT_TEMPLATE.exists():
            st.error(
                f"Default template not found: {DEFAULT_TEMPLATE}. "
                "Upload a template file or add it under templates/."
            )
            st.stop()
        section_rows, ext, has_overflow, out = _run(DEFAULT_TEMPLATE)
except Exception as exc:
    st.error(f"Validation/generation error: {exc}")
    st.stop()

st.subheader("Validation")
st.dataframe(section_rows, use_container_width=True)
st.write(
    {
        "external_count": ext["count"],
        "external_capacity": ext["supported_rows"],
        "external_overflow": ext["overflow"],
        "external_rows": ext.get("rows", []),
    }
)

if has_overflow:
    st.error("Overflow detected. Reduce items or expand template capacity.")
    st.stop()

st.success("XLSX generated successfully.")
st.download_button(
    label="Download COMPUTO_OUTPUT.xlsx",
    data=out.getvalue(),
    file_name="COMPUTO_OUTPUT.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

