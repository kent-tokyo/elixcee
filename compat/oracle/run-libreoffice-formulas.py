#!/usr/bin/env python3
"""Run a small formula-only oracle without invoking the Basic object model.

This is test infrastructure only. LibreOffice is an independent oracle and
must not be described as Microsoft Excel compatibility evidence.
"""

from __future__ import annotations

import argparse
import datetime
import json
import subprocess
import tempfile
from pathlib import Path

import openpyxl
from openpyxl.utils.datetime import CALENDAR_MAC_1904, to_excel


CASES = {
    "sum": ("=SUM(A1:A3)", 6),
    "average": ("=AVERAGE(A1:A3)", 2),
    "min": ("=MIN(A1:A3)", 1),
    "max": ("=MAX(A1:A3)", 3),
    "count": ("=COUNT(A1:B3)", 3),
    "counta": ("=COUNTA(B1:B3)", 3),
    "sum_mixed": ("=SUM(A1:B3)", 6),
    "if": ('=IF(A1>2,"yes","no")', "no"),
    "iferror": ("=IFERROR(1/0,99)", 99),
    "and": ("=AND(A1=1,A2=2)", True),
    "or": ("=OR(A1=9,A2=2)", True),
    "not": ("=NOT(A1=1)", False),
    "ifna": ('=IFNA(NA(),"missing")', "missing"),
    "iserror": ("=ISERROR(1/0)", True),
    "error_value": ("=1/0", "#DIV/0!"),
    "round": ("=ROUND(1.235,2)", 1.24),
    "roundup": ("=ROUNDUP(1.231,2)", 1.24),
    "rounddown": ("=ROUNDDOWN(1.239,2)", 1.23),
    "date": ("=DATE(2024,2,29)", 45351),
    "date1904": ("=DATE(2024,2,29)", 45351),
    "left": ('=LEFT("elixcee",3)', "eli"),
    "mid": ('=MID("hello",2,3)', "ell"),
    "len": ('=LEN("hello")', 5),
    "concatenate": ('=CONCATENATE("A","B","C")', "ABC"),
    "match": ("=MATCH(2,A1:A3,0)", 2),
    "index": ("=INDEX(A1:B3,2,2)", "two"),
    "countif": ('=COUNTIF(A1:A3,">1")', 2),
    "vlookup": ('=VLOOKUP(2,A1:B3,2,FALSE)', "two"),
}

# LibreOffice 26.2.5 on this host leaves IFNA results as #N/A, including the
# direct NA() form. Keep the case in the generated workbook as a visible
# probe, but do not misclassify this oracle limitation as an engine mismatch.
ORACLE_UNSUPPORTED = {"ifna"}


def serial_or_value(value):
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return to_excel(value)
    return value


def run(soffice: str) -> dict:
    with tempfile.TemporaryDirectory(prefix="elixcee-formula-oracle-") as raw:
        root = Path(raw)
        source = root / "formula-oracle.xlsx"
        input_book = openpyxl.Workbook()
        input_book.epoch = CALENDAR_MAC_1904
        sheet = input_book.active
        sheet.title = "Oracle"
        sheet["A1"], sheet["A2"], sheet["A3"] = 1, 2, 3
        sheet["B1"], sheet["B2"], sheet["B3"] = "one", "two", "three"
        for row, (name, (formula, expected)) in enumerate(CASES.items(), start=5):
            sheet.cell(row=row, column=1, value=name)
            sheet.cell(row=row, column=2, value=formula)
            sheet.cell(row=row, column=3, value=expected)
            if name == "date1904":
                sheet.cell(row=row, column=2).number_format = "yyyy-mm-dd"
        input_book.calculation.fullCalcOnLoad = True
        input_book.calculation.forceFullCalc = True
        input_book.save(source)

        out_dir = root / "recalculated"
        out_dir.mkdir()
        profile = root / "profile"
        subprocess.run(
            [
                soffice,
                "--headless",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(out_dir),
                f"-env:UserInstallation={profile.as_uri()}",
                str(source),
            ],
            check=True,
            capture_output=True,
            text=True,
        )
        result = out_dir / source.name
        if not result.exists():
            raise RuntimeError(f"LibreOffice did not produce {result}")
        values = openpyxl.load_workbook(result, data_only=True).active
        records = []
        skipped = []
        for row, (name, (_, expected)) in enumerate(CASES.items(), start=5):
            if name in ORACLE_UNSUPPORTED:
                skipped.append({"case": name, "reason": "oracle_unsupported"})
                continue
            actual = serial_or_value(values.cell(row=row, column=2).value)
            records.append(
                {
                    "case": name,
                    "formula": CASES[name][0],
                    "expected": expected,
                    "actual": actual,
                    "match": actual == expected,
                }
            )
        return {
            "oracle": "libreoffice",
            "comparable_cases": len(records),
            "matches": sum(item["match"] for item in records),
            "skipped": skipped,
            "records": records,
        }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--soffice", default="soffice")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    payload = run(args.soffice)
    encoded = json.dumps(payload, indent=2) + "\n"
    if args.output:
        args.output.write_text(encoded, encoding="utf-8")
    print(encoded, end="")
    if payload["matches"] != payload["comparable_cases"]:
        raise SystemExit("formula oracle mismatch")


if __name__ == "__main__":
    main()
