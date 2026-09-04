#!/usr/bin/env python3
"""Validate local reader measurement artifact contracts without extra packages."""

from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path


REQUIRED_META = {"schema_version", "date", "host", "toolchain", "binary"}


def fail(path: Path, message: str) -> None:
    raise ValueError(f"{path}: {message}")


def check(path: Path) -> None:
    try:
        document = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as error:
        fail(path, f"invalid JSON: {error}")
    if not isinstance(document, dict):
        fail(path, "top-level JSON value must be an object")
    missing = REQUIRED_META - document.keys()
    if missing:
        fail(path, f"missing metadata: {sorted(missing)}")
    if document["schema_version"] != 1:
        fail(path, "unsupported schema_version")
    observations = document.get("observations")
    fixtures = document.get("fixtures")
    if observations is not None:
        if not observations:
            fail(path, "observations must not be empty")
        for item in observations:
            if item.get("correct") is not True:
                fail(path, f"unsuccessful observation: {item}")
            if item.get("exit_code") != 0 and not item.get("error"):
                fail(path, "rejected observation has no recorded error")
        return
    if not fixtures:
        fail(path, "expected observations or fixtures")
    expected_iterations = 5 if "save_workbook" in document.get("api", "") else 10
    for fixture in fixtures:
        if fixture.get("iterations") != expected_iterations:
            fail(path, f"fixture does not have {expected_iterations} iterations: {fixture}")
        if "successes" in fixture and fixture["successes"] != expected_iterations:
            fail(path, f"fixture does not have all successful iterations: {fixture}")
        for item in fixture.get("observations", []):
            if item.get("cells", 0) <= 0:
                fail(path, f"invalid in-process observation: {item}")


def self_test() -> None:
    valid = {
        "schema_version": 1,
        "date": "2026-09-03",
        "host": "test",
        "toolchain": "test",
        "binary": "test",
        "observations": [{"correct": True, "exit_code": 0}],
    }
    with tempfile.TemporaryDirectory(prefix="elixcee-measurement-check-") as directory:
        root = Path(directory)
        valid_path = root / "valid.json"
        valid_path.write_text(json.dumps(valid))
        check(valid_path)
        cases = [
            {**valid, "schema_version": 2},
            {key: value for key, value in valid.items() if key != "binary"},
            {**valid, "observations": [{"correct": False, "exit_code": 0}]},
        ]
        for index, case in enumerate(cases):
            path = root / f"invalid-{index}.json"
            path.write_text(json.dumps(case))
            try:
                check(path)
            except ValueError:
                continue
            raise AssertionError(f"self-test accepted invalid artifact: {path}")
    print("reader measurement validator self-test: OK")


def main() -> None:
    if sys.argv[1:] == ["--self-test"]:
        self_test()
        return
    paths = [Path(value) for value in sys.argv[1:]]
    if not paths:
        raise SystemExit("usage: check-reader-measurements.py FILE.json ... | --self-test")
    for path in paths:
        check(path)
    print(f"reader measurements: OK ({len(paths)} files)")


if __name__ == "__main__":
    try:
        main()
    except ValueError as error:
        print(error, file=sys.stderr)
        raise SystemExit(1)
