#!/usr/bin/env python3
"""Measure the versioned VBA diagnostic corpus with reproducible metadata."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import subprocess
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "compat" / "vba-diagnostics" / "corpus.json"
RUNNER = ROOT / "scripts" / "run-vba-diagnostic-corpus.py"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def timed_run(command: list[str]) -> tuple[subprocess.CompletedProcess[str], int | None, float]:
    """Run through the platform's external time utility when available."""
    if sys.platform == "darwin":
        wrapper = ["/usr/bin/time", "-l"]
    elif sys.platform.startswith("linux"):
        wrapper = ["/usr/bin/time", "-v"]
    else:
        wrapper = []
    started = time.perf_counter_ns()
    completed = subprocess.run(
        [*wrapper, *command], cwd=ROOT, text=True, capture_output=True, check=False
    )
    elapsed_ms = (time.perf_counter_ns() - started) / 1_000_000
    rss = None
    for line in completed.stderr.splitlines():
        if "maximum resident set size" in line.lower():
            rss = int(line.split()[0])
        elif "maximum resident set size (kbytes)" in line.lower():
            rss = int(line.split(":", 1)[1].strip()) * 1024
    return completed, rss, elapsed_ms


def runtime_version() -> str:
    result = subprocess.run(
        ["cargo", "pkgid", "-p", "elixcee", "--offline"],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
    )
    return result.stdout.strip() or "unknown"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case", help="measure one manifest case by id")
    parser.add_argument("--repetitions", type=int, default=3)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    if args.repetitions < 1:
        parser.error("--repetitions must be positive")

    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    cases = manifest["cases"]
    if args.case:
        cases = [case for case in cases if case["id"] == args.case]
        if not cases:
            parser.error(f"unknown corpus case: {args.case}")

    samples = []
    rss_values = []
    for _ in range(args.repetitions):
        command = [sys.executable, str(RUNNER), "--json"]
        if args.case:
            command.extend(["--case", args.case])
        completed, rss, elapsed_ms = timed_run(command)
        samples.append({"wall_time_ms": round(elapsed_ms, 3), "exit_code": completed.returncode})
        if rss is not None:
            rss_values.append(rss)

    result = {
        "schema_version": 1,
        "corpus_version": manifest["corpus_version"],
        "manifest_sha256": sha256(MANIFEST),
        "runtime": runtime_version(),
        "git_revision": subprocess.run(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True, capture_output=True, check=True
        ).stdout.strip(),
        "host": {
            "os": platform.platform(),
            "machine": platform.machine(),
            "processor": platform.processor(),
            "cpu_count": os.cpu_count(),
            "python": platform.python_version(),
        },
        "case_ids": [case["id"] for case in cases],
        "repetitions": args.repetitions,
        "samples": samples,
        "peak_rss_bytes": max(rss_values) if rss_values else None,
        "rss_source": "/usr/bin/time -l/-v maximum resident set size for runner and descendants",
        "success": all(sample["exit_code"] == 0 for sample in samples),
        "boundary": "Corpus runner regression timing only; not a cross-library or Excel-oracle benchmark.",
    }
    payload = json.dumps(result, ensure_ascii=False, indent=2) + "\n"
    if args.output:
        args.output.write_text(payload, encoding="utf-8")
    else:
        print(payload, end="")
    return 0 if result["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
