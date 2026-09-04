#!/usr/bin/env python3
"""Measure repeated CLI reader processes for resource reclamation drift."""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import subprocess
import tempfile
import time
from pathlib import Path


PROFILES = ("dense", "mixed", "multi-sheet-4")


def load_fixture_module():
    source = Path(__file__).with_name("measure-reader-large.py")
    spec = importlib.util.spec_from_file_location("reader_large", source)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load fixture generator: {source}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def rss_bytes(pid: int) -> int | None:
    try:
        value = subprocess.check_output(
            ["ps", "-o", "rss=", "-p", str(pid)], text=True
        ).strip()
        return int(value) * 1024 if value else None
    except (subprocess.CalledProcessError, ValueError):
        return None


def measure(binary: Path, fixture: Path, expected_cells: int) -> dict[str, object]:
    with tempfile.NamedTemporaryFile(prefix="elixcee-reclaim-", suffix=".json") as output:
        started = time.perf_counter()
        process = subprocess.Popen(
            [str(binary), "snapshot", str(fixture), "--json"],
            stdout=output,
            stderr=subprocess.PIPE,
        )
        peak_rss = 0
        while process.poll() is None:
            current = rss_bytes(process.pid)
            if current is not None:
                peak_rss = max(peak_rss, current)
            time.sleep(0.01)
        stderr = process.stderr.read().decode(errors="replace")
        wall_ms = (time.perf_counter() - started) * 1000
        output.seek(0)
        payload = json.load(output)
    cells = sum(len(sheet["cells"]) for sheet in payload.get("sheets", []))
    return {
        "exit_code": process.returncode,
        "wall_ms": round(wall_ms, 3),
        "peak_rss_bytes": peak_rss,
        "correct": process.returncode == 0 and payload.get("ok") is True and cells == expected_cells,
        "observed_cells": cells,
        "error": stderr.strip() if process.returncode else None,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, default=Path("target/debug/elixcee"))
    parser.add_argument("--repetitions", type=int, default=10)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    if args.repetitions < 2:
        parser.error("--repetitions must be at least 2")
    if not args.binary.is_file():
        parser.error(f"binary does not exist: {args.binary}")

    fixture_module = load_fixture_module()
    observations = []
    with tempfile.TemporaryDirectory(prefix="elixcee-reader-reclaim-") as directory:
        root = Path(directory)
        for profile in PROFILES:
            spec = fixture_module.PROFILES[profile]
            fixture = root / f"profile-{profile}.xlsx"
            expected = fixture_module.make_fixture(fixture, profile, spec)
            fixture_module.validate_fixture(fixture, expected)
            baseline_rss = rss_bytes(os.getpid())
            for repetition in range(1, args.repetitions + 1):
                observation = measure(args.binary, fixture, expected["observed_cells"])
                observations.append(
                    {
                        "profile": profile,
                        "repetition": repetition,
                        "sheet_count": expected["sheet_count"],
                        "fixture_bytes": expected["fixture_bytes"],
                        "expected_cells": expected["observed_cells"],
                        "parent_rss_bytes_after": rss_bytes(os.getpid()),
                        **observation,
                    }
                )
            final_rss = rss_bytes(os.getpid())
            for item in observations:
                if item["profile"] == profile:
                    item["parent_rss_baseline_bytes"] = baseline_rss
                    item["parent_rss_final_bytes"] = final_rss
                    item["parent_rss_delta_bytes"] = (
                        final_rss - baseline_rss if baseline_rss is not None and final_rss is not None else None
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
                "protocol": "sequential CLI child processes; parent RSS sampled before/after each profile",
                "repetitions": args.repetitions,
                "observations": observations,
            },
            indent=2,
        )
        + "\n"
    )


if __name__ == "__main__":
    main()
