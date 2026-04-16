#!/usr/bin/env python3
"""
Valida combinazione JSON + template per la generazione computo.

Usage:
  python validate_bundle.py input.json [template.xlsx]
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

from genera_computo import DEFAULT_TEMPLATE, inspect_template_capacity


def main() -> int:
    if len(sys.argv) < 2:
        print("Usage: python validate_bundle.py input.json [template.xlsx]")
        return 2

    input_json = Path(sys.argv[1])
    template_xlsx = Path(sys.argv[2]) if len(sys.argv) > 2 else DEFAULT_TEMPLATE

    if not input_json.exists():
        print(f"ERROR: input json non trovato: {input_json}")
        return 2
    if not template_xlsx.exists():
        print(f"ERROR: template xlsx non trovato: {template_xlsx}")
        return 2

    with input_json.open(encoding="utf-8") as f:
        data = json.load(f)

    report = inspect_template_capacity(data, template_xlsx)
    errors = []
    warnings = []

    for sec in report["sections"]:
        if sec["overflow"] > 0:
            errors.append(
                f"Sezione {sec['index'] + 1} '{sec['label']}': "
                f"items={sec['item_count']} > capacity={sec['capacity']} (overflow={sec['overflow']})"
            )
        if sec["item_count"] == 0:
            warnings.append(f"Sezione {sec['index'] + 1} '{sec['label']}': nessun item nel JSON")

    ext = report["external"]
    if ext["overflow"] > 0:
        errors.append(
            f"External items={ext['count']} > supported_rows={ext['supported_rows']} "
            f"(overflow={ext['overflow']})"
        )

    print(f"Template: {report['template']}")
    print(f"Sheet: {report['scenario_sheet']}, header_row={report['header_row']}")
    for sec in report["sections"]:
        print(
            f"- S{sec['index'] + 1} row={sec['section_row']} subtotal={sec['subtotal_row']} "
            f"items={sec['item_count']} capacity={sec['capacity']} overflow={sec['overflow']}"
        )
    print(
        f"- External rows={ext['rows']} count={ext['count']} "
        f"capacity={ext['supported_rows']} overflow={ext['overflow']}"
    )

    if warnings:
        print("\nWARNINGS:")
        for w in warnings:
            print(f"- {w}")

    if errors:
        print("\nERRORS:")
        for e in errors:
            print(f"- {e}")
        return 1

    print("\nOK: bundle compatibile con il template.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
