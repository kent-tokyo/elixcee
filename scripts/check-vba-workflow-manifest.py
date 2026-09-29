#!/usr/bin/env python3
"""Validate the V0 headless VBA workflow manifest without executing it."""

from __future__ import annotations

import json
import hashlib
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "compat" / "vba-workflows" / "manifest.json"
RESULT_CLASSES = {
    "pass",
    "mismatch",
    "expected_error",
    "unsupported",
    "policy_blocked",
    "timeout",
    "not_measured",
}
BOUNDARIES = {"read", "vba", "calculate", "save", "reload"}


def fail(message: str) -> None:
    raise ValueError(message)


def main() -> int:
    document = json.loads(MANIFEST.read_text(encoding="utf-8"))
    if document.get("schema_version") != 1:
        fail("schema_version must be 1")
    comparison = document.get("comparison")
    if not isinstance(comparison, dict):
        fail("comparison must be an object")
    for key in ("elixcee_commit", "xlflow_release", "xlflow_commit"):
        if not isinstance(comparison.get(key), str) or not comparison[key]:
            fail(f"comparison.{key} must be non-empty")
    declared = document.get("result_classes")
    if not isinstance(declared, list) or set(declared) != RESULT_CLASSES:
        fail("result_classes must contain the complete sorted result vocabulary")
    cases = document.get("cases")
    if not isinstance(cases, list) or not cases:
        fail("cases must be a non-empty array")
    seen: set[str] = set()
    for case in cases:
        if not isinstance(case, dict):
            fail("each case must be an object")
        case_id = case.get("id")
        if not isinstance(case_id, str) or not case_id or case_id in seen:
            fail(f"invalid or duplicate case id: {case_id!r}")
        seen.add(case_id)
        for key in ("workbook", "vba_source", "entrypoint", "class"):
            if not isinstance(case.get(key), str) or not case[key]:
                fail(f"{case_id}: {key} must be non-empty")
        for key in ("workbook", "vba_source"):
            fixture = ROOT / case[key]
            if not fixture.is_file():
                fail(f"{case_id}: missing {key} fixture {case[key]!r}")
            digest = hashlib.sha256(fixture.read_bytes()).hexdigest()
            if case.get(f"{key}_sha256") != digest:
                fail(f"{case_id}: {key}_sha256 does not match fixture")
        operation = case.get("operation_boundary")
        if not isinstance(operation, list) or not operation or not set(operation) <= BOUNDARIES:
            fail(f"{case_id}: invalid operation_boundary")
        expected = case.get("expected")
        if not isinstance(expected, dict) or expected.get("result_class") not in RESULT_CLASSES:
            fail(f"{case_id}: invalid expected.result_class")
        local = case.get("local", {})
        if not isinstance(local, dict):
            fail(f"{case_id}: local must be an object when present")
        if "result_class" in local and not isinstance(local["result_class"], str):
            fail(f"{case_id}: local.result_class must be a string")
        limits = case.get("limits")
        if not isinstance(limits, dict) or not isinstance(limits.get("timeout_ms"), int) or limits["timeout_ms"] <= 0:
            fail(f"{case_id}: positive limits.timeout_ms is required")
        provenance = case.get("provenance")
        if not isinstance(provenance, dict) or not provenance.get("kind") or not provenance.get("source"):
            fail(f"{case_id}: provenance.kind and provenance.source are required")
        if expected["result_class"] == "pass" and not expected.get("source_unchanged"):
            fail(f"{case_id}: unchanged source is required for pass")
    print(f"validated {len(cases)} headless VBA workflow case(s); measurement remains not_measured")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, json.JSONDecodeError) as error:
        raise SystemExit(f"workflow manifest error: {error}")
