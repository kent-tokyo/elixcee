#!/usr/bin/env python3
"""Run the redistributable VBA diagnostic corpus manifest.

The manifest deliberately names existing Rust regression tests instead of
embedding customer workbooks. Each test constructs its own synthetic input or
uses a checked-in fixture, and the test assertions are the expected outcome.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "compat" / "vba-diagnostics" / "corpus.json"

# Keep the allowed cargo target mapping in code so manifest edits cannot inject
# arbitrary commands. The test name itself is passed as a separate argument.
TARGETS = {
    "range_paste": ("--test", "cli_diagnose"),
    "hidden_rows_columns": ("--lib",),
    "protection": ("--test", "cli_diagnose"),
    "formula_result": ("--test", "cli_diagnose_workbook"),
    "vba_resolution": ("--test", "cli_check"),
}


def load_cases() -> list[dict[str, object]]:
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    if manifest.get("schema_version") != 1:
        raise ValueError("unsupported corpus schema_version")
    cases = manifest.get("cases")
    if not isinstance(cases, list) or not cases:
        raise ValueError("corpus cases must be a non-empty array")
    for case in cases:
        if not isinstance(case, dict):
            raise ValueError("corpus case must be an object")
        case_id = case.get("id")
        case_class = case.get("class")
        test_name = case.get("runner_test")
        if not all(isinstance(value, str) and value for value in (case_id, case_class, test_name)):
            raise ValueError("each case needs non-empty id, class, and runner_test")
        if case_class not in TARGETS:
            raise ValueError(f"no safe cargo target mapping for class {case_class!r}")
        artifact_dir = case.get("artifact_dir")
        if not isinstance(case.get("seed"), int) or not isinstance(artifact_dir, str):
            raise ValueError(f"case {case_id!r} needs an integer seed and artifact_dir")
        artifact = ROOT / "compat" / "vba-diagnostics" / artifact_dir
        for name in ("input.xlsx", "Run.bas", "expected.json"):
            if not (artifact / name).is_file():
                raise ValueError(f"case {case_id!r} is missing artifact {name}")
    return cases


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--case", help="run one manifest case by id")
    parser.add_argument("--json", action="store_true", help="emit one JSON summary")
    args = parser.parse_args()
    try:
        cases = load_cases()
        if args.case:
            cases = [case for case in cases if case["id"] == args.case]
            if not cases:
                raise ValueError(f"unknown corpus case: {args.case}")
    except (OSError, ValueError, json.JSONDecodeError) as error:
        print(f"corpus error: {error}", file=sys.stderr)
        return 2

    results = []
    for case in cases:
        command = ["cargo", "test", "-p", "elixcee", "--offline", *TARGETS[case["class"]], "--", case["runner_test"]]
        completed = subprocess.run(command, cwd=ROOT, text=True, capture_output=True, check=False)
        results.append({"id": case["id"], "command": command, "exit_code": completed.returncode})
        if completed.returncode != 0 and not args.json:
            sys.stderr.write(completed.stdout)
            sys.stderr.write(completed.stderr)
            return completed.returncode

    summary = {"schema_version": 1, "ok": all(result["exit_code"] == 0 for result in results), "cases": results}
    if args.json:
        print(json.dumps(summary, ensure_ascii=False, separators=(",", ":")))
    else:
        print(f"diagnostic corpus: {len(results)} case(s) passed")
    return 0 if summary["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
