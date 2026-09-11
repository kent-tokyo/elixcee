#!/usr/bin/env python3
"""Validate JSON reports emitted by measure-stream-writer-memory.py."""

from __future__ import annotations

import json
import pathlib
import sys
import tempfile


REQUIRED_CASE_FIELDS = {
    "rows",
    "columns",
    "mode",
    "value_profile",
    "repetitions",
    "wall_ms_p50",
    "wall_ms_p95",
    "max_rss_bytes_p50",
    "max_rss_bytes_p95",
    "peak_temp_bytes_p50",
    "peak_temp_bytes_p95",
    "output_bytes",
    "output_valid",
    "semantic_equal",
    "samples",
}


def validate(path: pathlib.Path) -> None:
    report = json.loads(path.read_text(encoding="utf-8"))
    if report.get("schema_version") != 1:
        raise ValueError(f"{path}: unsupported schema_version")
    cases = report.get("cases")
    if not isinstance(cases, list) or not cases:
        raise ValueError(f"{path}: cases must be a non-empty list")
    for case in cases:
        missing = REQUIRED_CASE_FIELDS - case.keys()
        if missing:
            raise ValueError(f"{path}: missing case fields: {sorted(missing)}")
        if case["mode"] not in {"append", "normal", "normal-fresh"}:
            raise ValueError(f"{path}: invalid mode")
        if case["value_profile"] not in {"plain", "escape", "giant"}:
            raise ValueError(f"{path}: invalid value_profile")
        if case["rows"] < 1 or case["columns"] < 1:
            raise ValueError(f"{path}: rows and columns must be positive")
        if case["repetitions"] < 1 or len(case["samples"]) != case["repetitions"]:
            raise ValueError(f"{path}: repetitions do not match samples")
        for metric in ("wall_ms", "max_rss_bytes", "peak_temp_bytes"):
            if case[f"{metric}_p95"] < case[f"{metric}_p50"]:
                raise ValueError(f"{path}: {metric} p95 is below p50")
        if case["output_bytes"] <= 0 or not case["output_valid"]:
            raise ValueError(f"{path}: output was not validated")
        if not case["semantic_equal"]:
            raise ValueError(f"{path}: semantic output did not match generated input")
        if any(not sample.get("output_valid") for sample in case["samples"]):
            raise ValueError(f"{path}: an individual sample was not validated")
        if any(not sample.get("semantic_equal") for sample in case["samples"]):
            raise ValueError(f"{path}: an individual semantic output did not match")


def self_test() -> None:
    case = {
        "rows": 1,
        "columns": 1,
        "mode": "append",
        "value_profile": "plain",
        "repetitions": 1,
        "wall_ms_p50": 1.0,
        "wall_ms_p95": 1.0,
        "max_rss_bytes_p50": 1,
        "max_rss_bytes_p95": 1,
        "peak_temp_bytes_p50": 1,
        "peak_temp_bytes_p95": 1,
        "output_bytes": 1,
        "output_valid": True,
        "semantic_equal": True,
        "samples": [{"output_valid": True, "semantic_equal": True}],
    }
    with tempfile.TemporaryDirectory(prefix="elixcee-stream-validator-") as directory:
        path = pathlib.Path(directory) / "valid.json"
        path.write_text(json.dumps({"schema_version": 1, "cases": [case]}), encoding="utf-8")
        validate(path)
        for mutation in (
            {"schema_version": 2},
            {"cases": [{**case, "wall_ms_p95": 0.5}]},
            {"cases": [{**case, "samples": [{"output_valid": False}]}]},
            {"cases": [{**case, "semantic_equal": False}]},
            {"cases": [{**case, "samples": [{"output_valid": True, "semantic_equal": False}]}]},
        ):
            path.write_text(json.dumps({"schema_version": 1, **mutation}), encoding="utf-8")
            try:
                validate(path)
            except ValueError:
                continue
            raise AssertionError(f"validator accepted invalid report: {mutation}")


def main() -> int:
    if sys.argv[1:] == ["--self-test"]:
        self_test()
        print("stream writer measurement validator self-test: OK")
        return 0
    if len(sys.argv) < 2:
        raise SystemExit(
            f"usage: {pathlib.Path(sys.argv[0]).name} REPORT.json ... | --self-test"
        )
    for raw_path in sys.argv[1:]:
        validate(pathlib.Path(raw_path))
    print(f"stream writer measurements: OK ({len(sys.argv) - 1} files)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
