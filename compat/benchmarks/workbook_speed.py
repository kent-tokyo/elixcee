#!/usr/bin/env python3
"""Historical smoke harness; NOT an equal-durability comparator on macOS.

Use equal_workbook_speed.py for current comparisons. Rust sync_all uses
F_FULLFSYNC on macOS; os.fsync below does not reproduce that contract.
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

import openpyxl


def percentile(values: list[float], p: int) -> float:
    ordered = sorted(values)
    index = min(len(ordered) - 1, max(0, (len(ordered) * p + 99) // 100 - 1))
    return ordered[index]


def summary(values: list[float]) -> dict[str, float]:
    return {
        "p50_ms": percentile(values, 50),
        "p95_ms": percentile(values, 95),
        "min_ms": min(values),
        "max_ms": max(values),
    }


def bench_openpyxl(fixture: Path, iterations: int) -> list[float]:
    samples = []
    with tempfile.TemporaryDirectory(prefix="elixcee-bench-openpyxl-") as directory:
        for iteration in range(1, iterations + 1):
            output = Path(directory) / f"out-{iteration}.xlsx"
            started = time.perf_counter()
            workbook = openpyxl.load_workbook(fixture)
            sheet = workbook.active
            sheet["A1"] = iteration
            sheet["B1"] = "=1+2"
            workbook.save(output)
            # Historical barrier only; this does NOT match sync_all on macOS.
            with output.open("rb") as saved:
                os.fsync(saved.fileno())
            reloaded = openpyxl.load_workbook(output, data_only=False)
            assert reloaded.active["A1"].value == iteration
            samples.append((time.perf_counter() - started) * 1000.0)
    return samples


def bench_libreoffice(soffice: str, fixture: Path, iterations: int) -> list[float]:
    samples = []
    with tempfile.TemporaryDirectory(prefix="elixcee-bench-libreoffice-") as directory:
        root = Path(directory)
        input_path = root / "input.xlsx"
        shutil.copy2(fixture, input_path)
        for iteration in range(1, iterations + 1):
            output_dir = root / f"out-{iteration}"
            output_dir.mkdir()
            profile = root / f"profile-{iteration}"
            started = time.perf_counter()
            subprocess.run(
                [
                    soffice,
                    "--headless",
                    f"-env:UserInstallation={profile.as_uri()}",
                    "--convert-to",
                    "xlsx",
                    "--outdir",
                    str(output_dir),
                    str(input_path),
                ],
                check=True,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                text=True,
            )
            result = output_dir / "input.xlsx"
            assert result.exists()
            samples.append((time.perf_counter() - started) * 1000.0)
    return samples


def bench_elixcee(binary: Path, fixture: Path, iterations: int, fast: bool) -> tuple[list[float], dict]:
    command = [str(binary), str(fixture), str(iterations)]
    if fast:
        command.append("--fast")
    completed = subprocess.run(
        command,
        check=True,
        capture_output=True,
        text=True,
    )
    payload = json.loads(completed.stdout)
    return [item["wall_ms"] for item in payload["observations"]], payload


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("fixture", type=Path)
    parser.add_argument("--elixcee", type=Path, required=True)
    parser.add_argument("--soffice", default="soffice")
    parser.add_argument("--iterations", type=int, default=3)
    parser.add_argument("--elixcee-fast", action="store_true")
    args = parser.parse_args()
    if args.iterations < 3:
        parser.error("--iterations must be at least 3")

    elixcee_samples, elixcee_payload = bench_elixcee(
        args.elixcee, args.fixture, args.iterations, args.elixcee_fast
    )
    openpyxl_samples = bench_openpyxl(args.fixture, args.iterations)
    libreoffice_samples = bench_libreoffice(
        args.soffice, args.fixture, args.iterations
    )
    print(
        json.dumps(
            {
                "fixture": str(args.fixture),
                "iterations": args.iterations,
                "operation": "load, mutate A1/B1, save, reload/verify",
                "elixcee": {"version": "1.0.1", **summary(elixcee_samples)},
                "elixcee_raw": elixcee_payload,
                "openpyxl": {
                    "version": openpyxl.__version__,
                    **summary(openpyxl_samples),
                },
                "libreoffice": {
                    "version": "26.2.5.2",
                    "operation_note": "load/save conversion; no cell mutation",
                    **summary(libreoffice_samples),
                },
            },
            separators=(",", ":"),
        )
    )


if __name__ == "__main__":
    main()
