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

# Allow running from any working directory
sys.path.insert(0, str(Path(__file__).parent))

from excel_parser import parse_workbook

_DEFAULT_XLSX = Path(__file__).parent / "data" / "eba_template.xlsx"
_DEFAULT_CSV  = Path(__file__).parent / "data" / "coordinates.csv"


def export_coordinates(
    xlsx_path: str | Path = _DEFAULT_XLSX,
    csv_path:  str | Path = _DEFAULT_CSV,
) -> int:
    """
    Parse *xlsx_path*, collect every cell coordinate and write them to *csv_path*.

    CSV columns:
        coordinate  – full coordinate string, e.g. "C 01.00_0010_0020"
        sheet       – sheet name,             e.g. "C 01.00"
        row_code    – four-digit row code,     e.g. "0010"
        col_code    – four-digit column code,  e.g. "0020"

    Returns the number of coordinates written.
    """
    xlsx_path = Path(xlsx_path)
    csv_path  = Path(csv_path)

    sheets = parse_workbook(xlsx_path)

    rows: list[tuple[str, str, str, str]] = []
    for sheet_name, sheet in sheets.items():
        for row in sheet.rows:
            for cell in row:
                if not cell.coordinate:
                    continue
                # coordinate format: "<sheet>_<row_code>_<col_code>"
                parts = cell.coordinate.rsplit("_", 2)
                if len(parts) == 3:
                    _, row_code, col_code = parts
                else:
                    row_code = col_code = ""
                rows.append((cell.coordinate, sheet_name, row_code, col_code))

    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["coordinate", "sheet", "row_code", "col_code"])
        writer.writerows(rows)

    return len(rows)


if __name__ == "__main__":
    xlsx = Path(sys.argv[1]) if len(sys.argv) > 1 else _DEFAULT_XLSX
    csv_ = Path(sys.argv[2]) if len(sys.argv) > 2 else _DEFAULT_CSV

    n = export_coordinates(xlsx, csv_)
    print(f"✓ {n} Koordinaten → {csv_}")
