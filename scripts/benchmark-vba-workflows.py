#!/usr/bin/env python3
"""Measure local headless VBA workflow wall time.

This intentionally measures only the elixcee CLI.  It is not an Excel/xlflow
comparison and does not infer compatibility or speedup from missing arms.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import platform
import shlex
import subprocess
import tempfile
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "compat" / "vba-workflows" / "manifest.json"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def percentile(values: list[float], percent: int) -> float:
    ordered = sorted(values)
    index = min(len(ordered) - 1, max(0, (len(ordered) * percent + 99) // 100 - 1))
    return ordered[index]


def rss_bytes(pid: int) -> int | None:
    """Return the process RSS in bytes when the host exposes `ps` RSS."""
    try:
        completed = subprocess.run(
            ["ps", "-o", "rss=", "-p", str(pid)],
            capture_output=True,
            text=True,
            check=False,
        )
        value = completed.stdout.strip()
        return int(value) * 1024 if completed.returncode == 0 and value else None
    except (OSError, ValueError):
        return None


def run_case(
    command: list[str], timeout_seconds: float, measure_rss: bool
) -> tuple[subprocess.CompletedProcess[str], float, int | None]:
    started = time.perf_counter()
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    peak_rss = rss_bytes(process.pid) if measure_rss else None
    while process.poll() is None:
        if measure_rss:
            current = rss_bytes(process.pid)
            if current is not None:
                peak_rss = max(peak_rss or 0, current)
        if time.perf_counter() - started > timeout_seconds:
            process.kill()
            stdout, stderr = process.communicate()
            raise TimeoutError(
                f"timed out after {timeout_seconds:.3f}s: {shlex.join(command)}\n{stderr}"
            )
        time.sleep(0.001)
    stdout, stderr = process.communicate()
    elapsed_ms = (time.perf_counter() - started) * 1000
    return (
        subprocess.CompletedProcess(command, process.returncode, stdout, stderr),
        elapsed_ms,
        peak_rss,
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", required=True, type=Path)
    parser.add_argument("--repetitions", type=int, default=3)
    parser.add_argument("--warmup", type=int, default=5)
    parser.add_argument("--case", action="append", dest="case_ids")
    parser.add_argument("--output", type=Path, help="optional JSON report path")
    parser.add_argument(
        "--measure-rss",
        action="store_true",
        help="poll child RSS with ps; adds host sampling overhead to wall time",
    )
    args = parser.parse_args()
    if args.repetitions < 3:
        parser.error("--repetitions must be at least 3")
    if args.warmup < 0:
        parser.error("--warmup must not be negative")
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    selected = [
        case
        for case in manifest["cases"]
        if case.get("local", {}).get("result_class", "pass") == "pass"
        and (not args.case_ids or case["id"] in args.case_ids)
    ]
    if not selected:
        parser.error("no successful local workflow cases selected")

    observations: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(prefix="elixcee-vba-bench-") as directory:
        output_dir = Path(directory)
        for case in selected:
            samples: list[float] = []
            rss_samples: list[int | None] = []
            output_hashes: list[str] = []
            for run_number in range(-args.warmup, args.repetitions):
                is_warmup = run_number < 0
                label = f"warmup-{run_number + args.warmup}" if is_warmup else str(run_number)
                output_path = output_dir / f"{case['id']}-{label}.xlsx"
                command = [
                    str(args.binary),
                    str(ROOT / case["vba_source"]),
                    case["entrypoint"],
                    "--file",
                    str(ROOT / case["workbook"]),
                    "--output",
                    str(output_path),
                    "--json",
                ]
                completed, elapsed_ms, peak_rss = run_case(
                    command,
                    case["limits"]["timeout_ms"] / 1000,
                    args.measure_rss,
                )
                try:
                    result = json.loads(completed.stdout)
                except json.JSONDecodeError as error:
                    raise RuntimeError(
                        f"{case['id']} run {label}: invalid CLI JSON: {error}"
                    ) from error
                if completed.returncode != 0 or result.get("ok") is not True:
                    raise RuntimeError(
                        f"{case['id']} run {label}: CLI failed: {result}"
                    )
                if not output_path.is_file():
                    raise RuntimeError(f"{case['id']} run {label}: output missing")
                if is_warmup:
                    continue
                samples.append(elapsed_ms)
                rss_samples.append(peak_rss)
                output_hashes.append(sha256(output_path))
            observations.append(
                {
                    "id": case["id"],
                    "operation_boundary": [
                        "read",
                        "vba",
                        "calculate",
                        "save",
                        "process_exit",
                    ],
                    "repetitions": args.repetitions,
                    "warmup": args.warmup,
                    "wall_ms": {
                        "p50": percentile(samples, 50),
                        "p95": percentile(samples, 95),
                        "samples": samples,
                    },
                    "peak_rss_bytes": (
                        {
                            "p50": percentile(
                                [value for value in rss_samples if value is not None], 50
                            ),
                            "p95": percentile(
                                [value for value in rss_samples if value is not None], 95
                            ),
                            "samples": rss_samples,
                        }
                        if rss_samples and all(value is not None for value in rss_samples)
                        else None
                    ),
                    "output_sha256": output_hashes,
                    "output_stable": len(set(output_hashes)) == 1,
                }
            )

    report = {
        "schema": "elixcee.vba-workflow-benchmark.v1",
        "runner": "elixcee-only",
        "measurement_scope": "CLI process: read -> VBA -> calculate -> save -> process exit",
        "warmup_excluded": args.warmup,
        "rss_measured": args.measure_rss,
        "rss_measurement_method": "ps_polling_child_rss" if args.measure_rss else None,
        "external_oracle_status": manifest["comparison"]["measurement_status"],
        "binary": str(args.binary),
        "binary_sha256": sha256(args.binary),
        "manifest_sha256": sha256(MANIFEST),
        "host": {"platform": platform.platform(), "python": platform.python_version()},
        "cases": observations,
    }
    encoded = json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(encoded, encoding="utf-8")
    print(encoded, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
