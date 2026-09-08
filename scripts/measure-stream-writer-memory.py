#!/usr/bin/env python3
"""Measure writer RSS scaling in isolated child processes.

This deliberately measures the installed Python API, not the parent process.
It reports max RSS, wall time, output size, and row count; it does not claim
constant memory or Excel compatibility. Run with a development wheel that
contains ``create_stream_bounded`` after building it.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import pathlib
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
import math
import xml.etree.ElementTree as ET


CHILD = r'''
import ctypes
import json, pathlib, sys, time
import elixcee

def max_rss_bytes():
    if sys.platform == "win32":
        class Counters(ctypes.Structure):
            _fields_ = [
                ("cb", ctypes.c_ulong),
                ("page_fault_count", ctypes.c_ulong),
                ("peak_working_set_size", ctypes.c_size_t),
                ("working_set_size", ctypes.c_size_t),
                ("quota_peak_paged_pool_usage", ctypes.c_size_t),
                ("quota_paged_pool_usage", ctypes.c_size_t),
                ("quota_peak_non_paged_pool_usage", ctypes.c_size_t),
                ("quota_non_paged_pool_usage", ctypes.c_size_t),
                ("pagefile_usage", ctypes.c_size_t),
                ("peak_pagefile_usage", ctypes.c_size_t),
            ]
        counters = Counters()
        counters.cb = ctypes.sizeof(counters)
        process = ctypes.windll.kernel32.GetCurrentProcess()
        ctypes.windll.psapi.GetProcessMemoryInfo(
            process, ctypes.byref(counters), counters.cb
        )
        return counters.peak_working_set_size
    import resource
    raw = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return raw if sys.platform == "darwin" else raw * 1024

path = pathlib.Path(sys.argv[1])
rows = int(sys.argv[2])
columns = int(sys.argv[3])
mode = sys.argv[4]
value_profile = sys.argv[5]
if value_profile == "plain":
    text_value = "value"
elif value_profile == "escape":
    text_value = ("<&>\"'" * 4096)
else:
    text_value = "x" * (1024 * 1024)
work_bytes = max(rows * columns * max(128, len(text_value) * 2), 1024)
started = time.monotonic()
if mode == "append":
    writer = elixcee.create_stream_bounded(
        str(path),
        max_row_bytes=16 * 1024 * 1024,
        max_work_bytes=work_bytes,
        max_rows=rows,
        max_columns=columns,
    )
    for row_number in range(rows):
        writer.append((row_number, text_value, row_number % 7)[:columns])
    writer.close()
elif mode == "normal":
    source = path.with_name(path.stem + "-source.xlsx")
    writer = elixcee.create_stream_bounded(
        str(source),
        max_row_bytes=16 * 1024 * 1024,
        max_work_bytes=work_bytes,
        max_rows=rows,
        max_columns=columns,
    )
    for row_number in range(rows):
        writer.append((row_number, text_value, row_number % 7)[:columns])
    writer.close()
    workbook = elixcee.load_workbook(str(source))
    workbook.set_cell(1, 1, -1)
    workbook.save_workbook(str(path))
    source.unlink()
else:
    workbook = elixcee.Vm()
    # Bulk the normal-VM input so the measurement covers workbook storage and
    # save, not one full undo snapshot per appended row. This is deliberately
    # distinct from the append API measurement above.
    chunk_size = 4096
    workbook.begin_transaction()
    for start in range(0, rows, chunk_size):
        end = min(rows, start + chunk_size)
        grid = [
            (row_number, text_value, row_number % 7)[:columns]
            for row_number in range(start, end)
        ]
        last_column = "A" if columns == 1 else "B" if columns == 2 else "C"
        workbook.set_range(f"A{start + 1}:{last_column}{end}", grid)
    workbook.commit_transaction()
    workbook.save_workbook(str(path))
print(json.dumps({
    "rows": rows,
    "columns": columns,
    "wall_ms": (time.monotonic() - started) * 1000,
    "max_rss_bytes": max_rss_bytes(),
    "output_bytes": path.stat().st_size,
}))
'''


def verify_semantic_output(
    path: pathlib.Path, rows: int, columns: int, mode: str, value_profile: str
) -> bool:
    """Compare streamed worksheet values with the generated input rows.

    The comparison intentionally ignores ZIP/XML layout, shared-string choice,
    styles, and workbook metadata. It therefore checks semantic cell output
    without loading the whole worksheet tree into memory.
    """
    namespace = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    digest = hashlib.sha256()
    seen = 0
    with zipfile.ZipFile(path) as archive:
        shared_strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            with archive.open("xl/sharedStrings.xml") as source:
                for event, element in ET.iterparse(source, events=("end",)):
                    if element.tag == f"{namespace}si":
                        shared_strings.append(
                            "".join(
                                text.text or ""
                                for text in element.iter(f"{namespace}t")
                            )
                        )
                        element.clear()
        with archive.open("xl/worksheets/sheet1.xml") as source:
            for event, element in ET.iterparse(source, events=("end",)):
                if element.tag != f"{namespace}c":
                    continue
                reference = element.attrib.get("r", "")
                letters = "".join(character for character in reference if character.isalpha())
                row_text = "".join(character for character in reference if character.isdigit())
                if not letters or not row_text:
                    element.clear()
                    continue
                column = 0
                for character in letters.upper():
                    column = column * 26 + ord(character) - ord("A") + 1
                row = int(row_text)
                value = element.find(f"{namespace}v")
                raw = value.text if value is not None and value.text is not None else ""
                if element.attrib.get("t") == "s":
                    raw = shared_strings[int(raw)]
                elif element.attrib.get("t") == "inlineStr":
                    inline = element.find(f"{namespace}is")
                    raw = (
                        "".join(text.text or "" for text in inline.iter(f"{namespace}t"))
                        if inline is not None
                        else ""
                    )
                digest.update(f"{row},{column}\0{raw}\n".encode("utf-8"))
                seen += 1
                element.clear()
    expected = hashlib.sha256()
    text_value = "value" if value_profile == "plain" else (
        "<&>\"'" * 4096 if value_profile == "escape" else "x" * (1024 * 1024)
    )
    for row_number in range(rows):
        for column, raw in enumerate(
            (row_number, text_value, row_number % 7)[:columns], start=1
        ):
            if mode == "normal" and row_number == 0 and column == 1:
                raw = -1
            expected.update(f"{row_number + 1},{column}\0{raw}\n".encode("utf-8"))
    return seen == rows * columns and digest.digest() == expected.digest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--rows", nargs="+", type=int, default=[100_000, 250_000])
    parser.add_argument("--columns", type=int, default=3)
    parser.add_argument("--mode", choices=("append", "normal", "normal-fresh"), default="append")
    parser.add_argument(
        "--value-profile",
        choices=("plain", "escape", "giant"),
        default="plain",
        help="input text shape: ordinary, XML-escape-heavy, or one MiB",
    )
    parser.add_argument("--repetitions", type=int, default=1)
    parser.add_argument("--python", default=sys.executable)
    parser.add_argument(
        "--output",
        type=pathlib.Path,
        help="also write the JSON result to this path",
    )
    args = parser.parse_args()
    if args.repetitions < 1:
        parser.error("--repetitions must be at least 1")
    results = []
    with tempfile.TemporaryDirectory(prefix="elixcee-stream-rss-") as directory:
        for rows in args.rows:
            samples = []
            for repetition in range(args.repetitions):
                output = pathlib.Path(directory) / f"stream-{rows}-{repetition}.xlsx"
                started = time.monotonic()
                stop_monitor = threading.Event()
                peak_temp_bytes = [0]

                def monitor_temp_directory() -> None:
                    while not stop_monitor.is_set():
                        total = 0
                        try:
                            for entry in pathlib.Path(directory).iterdir():
                                try:
                                    if entry.is_file():
                                        total += entry.stat().st_size
                                except OSError:
                                    continue
                        except OSError:
                            pass
                        peak_temp_bytes[0] = max(peak_temp_bytes[0], total)
                        stop_monitor.wait(0.01)

                monitor = threading.Thread(target=monitor_temp_directory, daemon=True)
                monitor.start()
                completed = subprocess.run(
                    [
                        args.python,
                        "-I",
                        "-c",
                        CHILD,
                        str(output),
                        str(rows),
                        str(args.columns),
                        args.mode,
                        args.value_profile,
                    ],
                    check=False,
                    capture_output=True,
                    text=True,
                    env={"PATH": os.environ.get("PATH", "")},
                )
                stop_monitor.set()
                monitor.join()
                if completed.returncode:
                    raise SystemExit(completed.stderr or completed.stdout)
                sample = json.loads(completed.stdout)
                with zipfile.ZipFile(output) as archive:
                    assert archive.testzip() is None
                    worksheet = archive.read("xl/worksheets/sheet1.xml")
                    assert worksheet.startswith(b"<?xml")
                    assert f'<row r="{rows}">'.encode() in worksheet
                semantic_equal = verify_semantic_output(
                    output, rows, args.columns, args.mode, args.value_profile
                )
                sample["output_valid"] = True
                sample["semantic_equal"] = semantic_equal
                sample["mode"] = args.mode
                sample["value_profile"] = args.value_profile
                sample["peak_temp_bytes"] = peak_temp_bytes[0]
                sample["harness_wall_ms"] = (time.monotonic() - started) * 1000
                samples.append(sample)
                output.unlink()
            def percentile(key: str, fraction: float) -> float:
                values = sorted(sample[key] for sample in samples)
                return values[min(len(values) - 1, max(0, math.ceil(len(values) * fraction) - 1))]
            results.append({
                "rows": rows,
                "columns": args.columns,
                "mode": args.mode,
                "value_profile": args.value_profile,
                "repetitions": args.repetitions,
                "wall_ms_p50": percentile("wall_ms", 0.50),
                "wall_ms_p95": percentile("wall_ms", 0.95),
                "max_rss_bytes_p50": percentile("max_rss_bytes", 0.50),
                "max_rss_bytes_p95": percentile("max_rss_bytes", 0.95),
                "peak_temp_bytes_p50": percentile("peak_temp_bytes", 0.50),
                "peak_temp_bytes_p95": percentile("peak_temp_bytes", 0.95),
                "output_bytes": samples[0]["output_bytes"],
                "output_valid": all(sample["output_valid"] for sample in samples),
                "semantic_equal": all(sample["semantic_equal"] for sample in samples),
                "samples": samples,
            })
    report = json.dumps({"schema_version": 1, "cases": results}, indent=2)
    if args.output is not None:
        args.output.write_text(report + "\n", encoding="utf-8")
    print(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
