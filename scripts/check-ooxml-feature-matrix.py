#!/usr/bin/env python3
"""Validate the machine-readable OOXML compatibility matrix."""

from __future__ import annotations

import json
import sys
from pathlib import Path


STATUSES = {"supported", "preserved", "warned", "rejected", "unverified"}
AXES = ("read", "preserve", "edit", "recalculate", "excel_reopen")
FEATURES = {"charts", "pivot_tables", "drawings", "external_links"}


def validate(document: dict, root: Path) -> list[str]:
    errors: list[str] = []
    if document.get("schema_version") != 1:
        errors.append("schema_version must be 1")
    if set(document.get("statuses", [])) != STATUSES:
        errors.append("statuses must list the supported matrix vocabulary exactly")
    rows = document.get("features")
    if not isinstance(rows, list):
        return errors + ["features must be a list"]
    seen: set[str] = set()
    for index, row in enumerate(rows):
        prefix = f"features[{index}]"
        feature = row.get("id")
        if feature not in FEATURES:
            errors.append(f"{prefix}.id is not a known OOXML feature")
        if feature in seen:
            errors.append(f"duplicate feature: {feature}")
        seen.add(feature)
        for axis in AXES:
            if row.get(axis) not in STATUSES:
                errors.append(f"{prefix}.{axis} has an invalid status")
        fixture = row.get("fixture")
        if fixture is not None and not isinstance(fixture, str):
            errors.append(f"{prefix}.fixture must be a path or null")
        elif isinstance(fixture, str) and not (root / fixture).is_file():
            errors.append(f"{prefix}.fixture does not exist: {fixture}")
        if not isinstance(row.get("notes"), str) or not row["notes"].strip():
            errors.append(f"{prefix}.notes must be non-empty")
    missing = FEATURES - seen
    errors.extend(f"missing feature: {feature}" for feature in sorted(missing))
    return errors


def main() -> int:
    path = Path(sys.argv[1]) if len(sys.argv) == 2 else Path("compat/ooxml-feature-matrix.json")
    try:
        document = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        print(f"OOXML feature matrix: {exc}", file=sys.stderr)
        return 1
    errors = validate(document, Path.cwd())
    if errors:
        print("\n".join(f"OOXML feature matrix: {error}" for error in errors), file=sys.stderr)
        return 1
    print(f"OOXML feature matrix: ok ({len(document['features'])} features)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
