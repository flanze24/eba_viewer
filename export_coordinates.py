"""
export_coordinates.py
Exports all DPM coordinates found in a workbook to a CSV file.

Usage (standalone):
    python export_coordinates.py [path/to/file.xlsx] [path/to/output.csv]

Defaults:
    input  → data/eba_template.xlsx  (relative to this script)
    output → data/coordinates.csv    (relative to this script)

Usage (as module):
    from export_coordinates import export_coordinates
    export_coordinates("path/to/file.xlsx", "path/to/output.csv")
"""

from __future__ import annotations

import csv
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from excel_parser import parse_workbook

_DEFAULT_XLSX = Path(__file__).parent / "data" / "eba_template.xlsx"
_DEFAULT_CSV  = Path(__file__).parent / "data" / "coordinates.csv"


def export_coordinates(
    xlsx_path: str | Path = _DEFAULT_XLSX,
    csv_path:  str | Path = _DEFAULT_CSV,
) -> int:
    """
    Parse *xlsx_path*, collect every cell coordinate and label key,
    and write them to *csv_path*.

    CSV columns:
        key         – unique key: full coordinate OR label key
                      e.g. "C 01.00_0010_0020"  (input cell)
                           "C 01.00_col_0020"    (column header label)
                           "C 01.00_row_0010"    (row label)
        type        – "cell" | "col_label" | "row_label"
        sheet       – sheet name,  e.g. "C 01.00"
        row_code    – four-digit row code  (empty for col_label)
        col_code    – four-digit col code  (empty for row_label)
        annotation  – optional free-text note (manually editable, preserved on re-export)

    Returns the number of rows written.
    """
    xlsx_path = Path(xlsx_path)
    csv_path  = Path(csv_path)

    # Load existing annotations so manual edits are not lost on re-export
    existing_annotations: dict[str, str] = {}
    if csv_path.exists():
        with open(csv_path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            key_field = "key" if (reader.fieldnames and "key" in reader.fieldnames) else "coordinate"
            if reader.fieldnames and "annotation" in reader.fieldnames:
                for r in reader:
                    note = r.get("annotation", "").strip()
                    if note:
                        existing_annotations[r.get(key_field, "")] = note

    sheets = parse_workbook(xlsx_path)

    rows: list[tuple[str, str, str, str, str, str]] = []
    seen: set[str] = set()

    for sheet_name, sheet in sheets.items():
        for row in sheet.rows:
            for cell in row:
                # ── Input cell with full coordinate ──────────────────────
                if cell.coordinate and cell.coordinate not in seen:
                    seen.add(cell.coordinate)
                    parts = cell.coordinate.rsplit("_", 2)
                    row_code = parts[1] if len(parts) == 3 else ""
                    col_code = parts[2] if len(parts) == 3 else ""
                    annotation = existing_annotations.get(cell.coordinate, "")
                    rows.append((
                        cell.coordinate, "cell",
                        sheet_name, row_code, col_code, annotation
                    ))

                # ── Column or row label ───────────────────────────────────
                lk = getattr(cell, "label_key", None)
                if lk and lk not in seen:
                    seen.add(lk)
                    # key format: "<sheet>_col_<code>"  or  "<sheet>_row_<code>"
                    if "_col_" in lk:
                        ltype    = "col_label"
                        row_code = ""
                        col_code = lk.rsplit("_col_", 1)[-1]
                    elif "_row_" in lk:
                        ltype    = "row_label"
                        row_code = lk.rsplit("_row_", 1)[-1]
                        col_code = ""
                    else:
                        ltype = row_code = col_code = ""
                    annotation = existing_annotations.get(lk, "")
                    rows.append((lk, ltype, sheet_name, row_code, col_code, annotation))

    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["key", "type", "sheet", "row_code", "col_code", "annotation"])
        writer.writerows(rows)

    return len(rows)


if __name__ == "__main__":
    xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else _DEFAULT_XLSX
    csv_ = Path(sys.argv[2]) if len(sys.argv) > 2 else _DEFAULT_CSV
    n = export_coordinates(xlsx, csv_)
    print(f"✓ {n} Einträge (Koordinaten + Labels) → {csv_}")