#!/usr/bin/env python3
"""Measure snapshot reader behavior on deterministic large XLSX fixtures.

The fixture generator uses only Python's standard library. The measured child
is the already-built elixcee binary, so build and fixture-generation time are
not included in reader timings.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path


PROFILES = {
    "dense": {"rows": 100_000, "columns": 4, "stride": 1},
    "dense-5col": {"rows": 80_000, "columns": 5, "stride": 1},
    "sparse": {"rows": 1_000_000, "columns": 2, "stride": 20},
    "style-heavy": {"rows": 100_000, "columns": 2, "stride": 1},
    "formula-heavy": {"rows": 80_000, "columns": 3, "stride": 1},
    "mixed": {"rows": 80_000, "columns": 4, "stride": 1},
    "multi-sheet-4": {"rows": 25_000, "columns": 4, "stride": 1, "sheets": 4},
}

STYLE_XML = """<styleSheet><numFmts count="1"><numFmt numFmtId="165" formatCode="0.00"/></numFmts><fonts count="2"><font/><font><b/></font></fonts><fills count="2"><fill/><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills><borders count="1"><border/></borders><cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="165"/><xf fontId="1"/><xf fillId="1"/></cellXfs></styleSheet>"""


def cell_xml(row: int, col: int, profile: str) -> tuple[str, int, int]:
    address = chr(ord("A") + col - 1) + str(row)
    style = 0
    formulas = 0
    if profile in {"style-heavy", "mixed"}:
        style = (row + col) % 4
    if profile in {"formula-heavy", "mixed"} and col == 3:
        formulas = 1
        return f'<c r="{address}" s="{style}"><f>A{row}+B{row}</f><v>{row * 3}</v></c>', style, formulas
    if profile == "mixed" and col == 4:
        value = f"row-{row}"
        return f'<c r="{address}" s="{style}" t="inlineStr"><is><t>{value}</t></is></c>', style, formulas
    return f'<c r="{address}" s="{style}"><v>{row * col}</v></c>', style, formulas


def make_sheet_xml(profile: str, spec: dict[str, int]) -> tuple[bytes, dict[str, int]]:
    sheet = ["<worksheet><sheetData>"]
    cell_count = 0
    formula_count = 0
    styled_count = 0
    for row in range(1, spec["rows"] + 1, spec["stride"]):
        cells = []
        for col in range(1, spec["columns"] + 1):
            rendered, style, formulas = cell_xml(row, col, profile)
            cells.append(rendered)
            cell_count += 1
            formula_count += formulas
            styled_count += style != 0
        sheet.append(f'<row r="{row}">' + "".join(cells) + "</row>")
    sheet.append("</sheetData></worksheet>")
    return "".join(sheet).encode(), {
        "observed_rows": (spec["rows"] - 1) // spec["stride"] + 1,
        "observed_cells": cell_count,
        "observed_formulas": formula_count,
        "observed_styled_cells": styled_count,
    }


def make_fixture(path: Path, profile: str, spec: dict[str, int]) -> dict[str, int]:
    sheet_count = spec.get("sheets", 1)
    sheet_parts = [make_sheet_xml(profile, spec) for _ in range(sheet_count)]
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_STORED) as archive:
        archive.writestr(
            "xl/workbook.xml",
            "<workbook><sheets>"
            + "".join(
                f'<sheet name="Sheet{i}" sheetId="{i}" r:id="rId{i}"/>'
                for i in range(1, sheet_count + 1)
            )
            + "</sheets></workbook>",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            "<Relationships>"
            + "".join(
                f'<Relationship Id="rId{i}" Type="/worksheet" Target="worksheets/sheet{i}.xml"/>'
                for i in range(1, sheet_count + 1)
            )
            + "</Relationships>",
        )
        archive.writestr("xl/styles.xml", STYLE_XML)
        for index, (sheet_bytes, _) in enumerate(sheet_parts, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", sheet_bytes)
    per_sheet = sheet_parts[0][1]
    return {
        "fixture_bytes": path.stat().st_size,
        "sheet_count": sheet_count,
        **{key: value * sheet_count for key, value in per_sheet.items()},
    }


def validate_fixture(path: Path, expected: dict[str, int]) -> None:
    with zipfile.ZipFile(path) as archive:
        styles = archive.read("xl/styles.xml").decode()
        observed = {"sheet_count": 0, "observed_rows": 0, "observed_cells": 0,
                    "observed_formulas": 0, "observed_styled_cells": 0}
        for name in archive.namelist():
            if not name.startswith("xl/worksheets/sheet") or not name.endswith(".xml"):
                continue
            sheet = archive.read(name).decode()
            observed["sheet_count"] += 1
            observed["observed_rows"] += sheet.count("<row ")
            observed["observed_cells"] += sheet.count("<c ")
            observed["observed_formulas"] += sheet.count("<f>")
            observed["observed_styled_cells"] += len(re.findall(r'<c [^>]* s="[1-9][0-9]*"', sheet))
    if styles.count("<xf ") != 4 or observed != {
        key: expected[key]
        for key in observed
    }:
        raise RuntimeError(f"fixture validation failed: expected {expected}, observed {observed}")


def measure(binary: Path, fixture: Path, expected: dict[str, int]) -> dict[str, object]:
    output_path = fixture.with_suffix(".json.out")
    started = time.perf_counter()
    process = subprocess.Popen(
        [str(binary), "snapshot", str(fixture), "--json"],
        stdout=output_path.open("wb"),
        stderr=subprocess.PIPE,
    )
    peak_rss = 0
    while process.poll() is None:
        try:
            rss_kib = int(
                subprocess.check_output(
                    ["ps", "-o", "rss=", "-p", str(process.pid)],
                    text=True,
                ).strip()
            )
            peak_rss = max(peak_rss, rss_kib * 1024)
        except (subprocess.CalledProcessError, ValueError):
            pass
        time.sleep(0.01)
    stderr = process.stderr.read().decode(errors="replace")
    wall_ms = (time.perf_counter() - started) * 1000
    stdout = output_path.read_bytes()
    output_path.unlink(missing_ok=True)
    result: dict[str, object] = {
        "exit_code": process.returncode,
        "wall_ms": round(wall_ms, 3),
        "peak_rss_bytes": peak_rss,
        "stdout_bytes": len(stdout),
    }
    if process.returncode == 0:
        payload = json.loads(stdout)
        sheets = payload["sheets"]
        observed_cells = sum(len(sheet["cells"]) for sheet in sheets)
        result["correct"] = (
            payload["ok"] is True
            and len(sheets) == expected["sheet_count"]
            and observed_cells == expected["observed_cells"]
        )
        result["observed_cells"] = observed_cells
    else:
        payload = json.loads(stdout)
        result["correct"] = payload.get("ok") is False and "error" in payload
        result["error"] = payload.get("error", {}).get("message", stderr.strip())
    return result


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, default=Path("target/debug/elixcee"))
    parser.add_argument("--repetitions", type=int, default=3)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.repetitions < 1:
        parser.error("--repetitions must be positive")
    if not args.binary.is_file():
        parser.error(f"binary does not exist: {args.binary}")

    observations = []
    with tempfile.TemporaryDirectory(prefix="elixcee-reader-measure-") as directory:
        root = Path(directory)
        for profile, spec in PROFILES.items():
            fixture = root / f"profile-{profile}.xlsx"
            expected = make_fixture(fixture, profile, spec)
            validate_fixture(fixture, expected)
            for repetition in range(1, args.repetitions + 1):
                observation = measure(args.binary, fixture, expected)
                observations.append(
                    {
                        "profile": profile,
                        "configured_rows": spec["rows"],
                        "columns": spec["columns"],
                        "stride": spec["stride"],
                        "repetition": repetition,
                        **expected,
                        **observation,
                    }
                )

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(
            {
                "schema_version": 1,
                "date": "2026-09-03",
                "host": "Darwin/aarch64-apple-darwin",
                "toolchain": "rustc 1.97.0",
                "binary": str(args.binary),
                "fixture_format": "XLSX with Stored ZIP entries and deterministic density/style/formula/multi-sheet populations",
                "repetitions": args.repetitions,
                "observations": observations,
            },
            indent=2,
        )
        + "\n"
    )


if __name__ == "__main__":
    main()
