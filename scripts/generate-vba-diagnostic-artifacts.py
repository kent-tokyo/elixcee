#!/usr/bin/env python3
"""Generate the small, redistributable VBA diagnostic fixtures."""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "compat" / "vba-diagnostics" / "artifacts"

CASES = {
    "paste-shape-mismatch": (
        'Sub Run()\n    Range("A1:C10").Copy\n    Range("E1:F10").PasteSpecial\nEnd Sub\n',
        {"root_cause": "PASTE_SHAPE_MISMATCH", "exit_success": False},
    ),
    "hidden-range-observation": (
        'Sub Run()\n    Range("A1:B2").Select\nEnd Sub\n',
        {"observation": "RANGE_CONTAINS_HIDDEN_CELLS", "exit_success": True},
    ),
    "protected-sheet-write": (
        'Sub Run()\n    Worksheets("Sheet1").Cells(1,1).Value = 1\nEnd Sub\n',
        {"root_cause": "SHEET_PROTECTED", "exit_success": False},
    ),
    "formula-error-result": (
        'Sub Run()\n    Dim arr(5)\n    arr(9) = 1\nEnd Sub\n',
        {"root_cause": "ARRAY_INDEX_OUT_OF_BOUNDS", "exit_success": False},
    ),
    "vba-entry-resolution": (
        'Sub Run()\n    Cells(1,1).Value = 1\nEnd Sub\n',
        {"diagnostic": "missing_entrypoint", "exit_success": False},
    ),
}


def main() -> None:
    for case_id, (source, expected) in CASES.items():
        directory = OUT / case_id
        directory.mkdir(parents=True, exist_ok=True)
        book = Workbook()
        fixed_time = datetime(2000, 1, 1, 0, 0, 0)
        book.properties.created = fixed_time
        book.properties.modified = fixed_time
        sheet = book.active
        if case_id == "paste-shape-mismatch":
            for row in range(1, 11):
                for col in range(1, 4):
                    sheet.cell(row, col, row * 10 + col)
        elif case_id == "hidden-range-observation":
            sheet.row_dimensions[2].hidden = True
            sheet.column_dimensions["B"].hidden = True
            sheet["A1"] = "visible"
            sheet["B2"] = "hidden"
        elif case_id == "protected-sheet-write":
            sheet.protection.sheet = True
        elif case_id == "formula-error-result":
            sheet["B2"] = 9
        sheet.freeze_panes = None
        book.save(directory / "input.xlsx")
        (directory / "Run.bas").write_text(source, encoding="utf-8")
        (directory / "expected.json").write_text(
            json.dumps({"schema_version": 1, **expected}, indent=2) + "\n", encoding="utf-8"
        )


if __name__ == "__main__":
    main()
