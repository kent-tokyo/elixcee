#!/usr/bin/env python3
"""Measure snapshot reader behavior on deterministic large XLSX fixtures.

The fixture generator uses only Python's standard library. The measured child
is the already-built elixcee binary, so build and fixture-generation time are
not included in reader timings.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import tempfile
import time
import zipfile
from pathlib import Path


SIZES = (10_000, 100_000, 150_000, 1_000_000)


def make_fixture(path: Path, rows: int) -> int:
    sheet = ["<worksheet><sheetData>"]
    for row in range(1, rows + 1):
        sheet.append(f'<row r="{row}"><c r="A{row}"><v>{row}</v></c></row>')
    sheet.append("</sheetData></worksheet>")
    sheet_bytes = "".join(sheet).encode()
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_STORED) as archive:
        archive.writestr(
            "xl/workbook.xml",
            '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<Relationships><Relationship Id="rId1" Type="/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        )
        archive.writestr("xl/worksheets/sheet1.xml", sheet_bytes)
    return path.stat().st_size


def measure(binary: Path, fixture: Path, expected_rows: int) -> dict[str, object]:
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
        cells = payload["sheets"][0]["cells"]
        result["correct"] = payload["ok"] is True and len(cells) == expected_rows
        result["observed_rows"] = len(cells)
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
        for rows in SIZES:
            fixture = root / f"rows-{rows}.xlsx"
            size_bytes = make_fixture(fixture, rows)
            for repetition in range(1, args.repetitions + 1):
                observation = measure(args.binary, fixture, rows)
                observations.append(
                    {
                        "rows": rows,
                        "repetition": repetition,
                        "fixture_bytes": size_bytes,
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
                "fixture_format": "XLSX with Stored ZIP entries and one numeric cell per row",
                "repetitions": args.repetitions,
                "observations": observations,
            },
            indent=2,
        )
        + "\n"
    )


if __name__ == "__main__":
    main()
