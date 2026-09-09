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
    "median": ("=MEDIAN(A1:A3)", 2),
    "product": ("=PRODUCT(A1:A3)", 6),
    "rank": ("=RANK(2,A1:A3,0)", 2),
    "count": ("=COUNT(A1:B3)", 3),
    "counta": ("=COUNTA(B1:B3)", 3),
    "countblank": ("=COUNTBLANK(A4:A4)", 1),
    "isblank": ("=ISBLANK(A4)", True),
    "isnumber": ("=ISNUMBER(A1)", True),
    "istext": ("=ISTEXT(B1)", True),
    "sum_mixed": ("=SUM(A1:B3)", 6),
    "sumproduct": ("=SUMPRODUCT(A1:A3,A1:A3)", 14),
    "sum_sequence": ("=SUM(SEQUENCE(3))", 6),
    "sum_transpose": ("=SUM(TRANSPOSE(A1:A3))", 6),
    "averageif": ('=AVERAGEIF(A1:A3,">1")', 2.5),
    "maxifs": ('=MAXIFS(A1:A3,A1:A3,">1")', 3),
    "minifs": ('=MINIFS(A1:A3,A1:A3,">1")', 2),
    "if": ('=IF(A1>2,"yes","no")', "no"),
    "iferror": ("=IFERROR(1/0,99)", 99),
    "and": ("=AND(A1=1,A2=2)", True),
    "or": ("=OR(A1=9,A2=2)", True),
    "not": ("=NOT(A1=1)", False),
    "ifna": ('=IFNA(NA(),"missing")', "missing"),
    "iserror": ("=ISERROR(1/0)", True),
    "iserr": ("=ISERR(1/0)", True),
    "isna": ('=ISNA(VLOOKUP(9,A1:B3,2,FALSE))', True),
    "isnontext": ("=ISNONTEXT(42)", True),
    "type_number": ("=TYPE(1)", 1),
    "type_text": ('=TYPE("text")', 2),
    "type_boolean": ("=TYPE(TRUE)", 4),
    "error_value": ("=1/0", "#DIV/0!"),
    "error_type": ("=ERROR.TYPE(1/0)", 2),
    "round": ("=ROUND(1.235,2)", 1.24),
    "roundup": ("=ROUNDUP(1.231,2)", 1.24),
    "rounddown": ("=ROUNDDOWN(1.239,2)", 1.23),
    "abs": ("=ABS(-3.5)", 3.5),
    "mod": ("=MOD(10,3)", 1),
    "power": ("=POWER(2,3)", 8),
    "int": ("=INT(3.9)", 3),
    "trunc": ("=TRUNC(-3.9)", -3),
    "sign": ("=SIGN(-3)", -1),
    "sqrt": ("=SQRT(9)", 3),
    "rows": ("=ROWS(A1:B3)", 3),
    "columns": ("=COLUMNS(A1:B3)", 2),
    "islogical": ("=ISLOGICAL(TRUE)", True),
    "n_boolean": ("=N(TRUE)", 1),
    "n_text": ('=N("elixcee")', 0),
    "upper": ('=UPPER("elixcee")', "ELIXCEE"),
    "lower": ('=LOWER("ELIXCEE")', "elixcee"),
    "trim": ('=TRIM("  elixcee  runtime ")', "elixcee runtime"),
    "find": ('=FIND("ix","elixcee")', 3),
    "replace": ('=REPLACE("elixcee",4,3,"X")', "eliXe"),
    "substitute": ('=SUBSTITUTE("elixcee","e","E")', "ElixcEE"),
    "search": ('=SEARCH("IX","elixcee")', 3),
    "exact": ('=EXACT("elixcee","ELIXCEE")', False),
    "proper": ('=PROPER("elixcee runtime")', "Elixcee Runtime"),
    "date": ("=DATE(2024,2,29)", 45351),
    "date1904": ("=DATE(2024,2,29)", 45351),
    "days": ("=DAYS(DATE(2024,3,1),DATE(2024,2,29))", 1),
    "year": ("=YEAR(DATE(2024,2,29))", 2024),
    "month": ("=MONTH(DATE(2024,2,29))", 2),
    "day": ("=DAY(DATE(2024,2,29))", 29),
    "weekday": ("=WEEKDAY(DATE(2024,2,29))", 5),
    "weeknum_monday": ("=WEEKNUM(DATE(2024,2,29),2)", 9),
    "isoweeknum": ("=ISOWEEKNUM(DATE(2024,2,29))", 9),
    "networkdays": ("=NETWORKDAYS(DATE(2024,2,26),DATE(2024,3,1))", 5),
    "datevalue": ('=DATEVALUE("2024-02-29")', 45351),
    "edate": ("=EDATE(DATE(2024,1,31),1)", 45351),
    "eomonth": ("=EOMONTH(DATE(2024,2,1),0)", 45351),
    "hour": ("=HOUR(TIME(14,30,15))", 14),
    "minute": ("=MINUTE(TIME(14,30,15))", 30),
    "second": ("=SECOND(TIME(14,30,15))", 15),
    "floor": ("=FLOOR(5.7,2)", 4),
    "ceiling": ("=CEILING(5.1,2)", 6),
    "mround": ("=MROUND(10,3)", 9),
    "left": ('=LEFT("elixcee",3)', "eli"),
    "mid": ('=MID("hello",2,3)', "ell"),
    "len": ('=LEN("hello")', 5),
    "concatenate": ('=CONCATENATE("A","B","C")', "ABC"),
    "match": ("=MATCH(2,A1:A3,0)", 2),
    "index": ("=INDEX(A1:B3,2,2)", "two"),
    "countif": ('=COUNTIF(A1:A3,">1")', 2),
    "sumif": ('=SUMIF(A1:A3,">1",A1:A3)', 5),
    "sumifs": ('=SUMIFS(A1:A3,A1:A3,">1")', 5),
    "countifs": ('=COUNTIFS(A1:A3,">1")', 2),
    "vlookup": ('=VLOOKUP(2,A1:B3,2,FALSE)', "two"),
    "vlookup_missing": ('=VLOOKUP(9,A1:B3,2,FALSE)', "#N/A"),
    "value": ('=VALUE("12.5")', 12.5),
    "xmatch": ("=XMATCH(2,A1:A3,0)", 2),
    "right": ('=RIGHT("elixcee",3)', "cee"),
    "textjoin": ('=TEXTJOIN("-",TRUE,B1:B3)', "one-two-three"),
}

# LibreOffice 26.2.5 on this host leaves IFNA results as #N/A, including the
# direct NA() form. Keep the case in the generated workbook as a visible
# probe, but do not misclassify this oracle limitation as an engine mismatch.
# LibreOffice 26.2.5 on this host also leaves DAYS(), MAXIFS(), and MINIFS()
# as #NAME? in this generated workbook. Keep them visible as probes, but
# exclude them from the cross-engine count for the same reason as the other
# build-specific gaps.
# This LibreOffice build also leaves ISOWEEKNUM() as #NAME? and reports
# TYPE(TRUE) as 1 rather than Excel's logical-type code 4.
ORACLE_UNSUPPORTED = {
    "days",
    "ifna",
    "maxifs",
    "minifs",
    "xmatch",
    "textjoin",
    "isoweeknum",
    "type_boolean",
    "sum_sequence",
}


def serial_or_value(value):
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return to_excel(value)
    return value


def json_safe_value(value):
    """Keep scalar results unchanged and expose binding-specific errors as text."""
    try:
        json.dumps(value)
    except TypeError:
        return str(value)
    return value


def normalize_binding_value(value, expected):
    """Normalize elixcee's date display value to the fixture's serial convention."""
    value = json_safe_value(value)
    if isinstance(value, str) and isinstance(expected, (int, float)):
        try:
            return to_excel(datetime.date.fromisoformat(value))
        except ValueError:
            pass
    return value


def run(soffice: str, with_elixcee: bool = False) -> dict:
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
            if name in {"date1904", "datevalue", "edate", "eomonth"}:
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
        payload = {
            "oracle": "libreoffice",
            "comparable_cases": len(records),
            "matches": sum(item["match"] for item in records),
            "skipped": skipped,
            "records": records,
        }
        if with_elixcee:
            import elixcee

            vm = elixcee.load_workbook(str(source), sheet="Oracle")
            vm.recalculate()
            elixcee_records = []
            for row, (name, (_, expected)) in enumerate(CASES.items(), start=5):
                if name in ORACLE_UNSUPPORTED:
                    continue
                actual = normalize_binding_value(vm.get_cell(row, 2), expected)
                elixcee_records.append(
                    {
                        "case": name,
                        "expected": expected,
                        "actual": actual,
                        "match": actual == expected,
                    }
                )
            payload["elixcee_comparable_cases"] = len(elixcee_records)
            payload["elixcee_matches"] = sum(item["match"] for item in elixcee_records)
            payload["elixcee_records"] = elixcee_records
        return payload


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--soffice", default="soffice")
    parser.add_argument("--output", type=Path)
    parser.add_argument(
        "--with-elixcee",
        action="store_true",
        help="also recalculate the same fixture with the installed elixcee binding",
    )
    args = parser.parse_args()
    payload = run(args.soffice, with_elixcee=args.with_elixcee)
    encoded = json.dumps(payload, indent=2) + "\n"
    if args.output:
        args.output.write_text(encoded, encoding="utf-8")
    print(encoded, end="")
    if payload["matches"] != payload["comparable_cases"]:
        raise SystemExit("formula oracle mismatch")
    if payload.get("elixcee_matches") != payload.get("elixcee_comparable_cases"):
        raise SystemExit("elixcee formula mismatch")


if __name__ == "__main__":
    main()
