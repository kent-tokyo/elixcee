#!/usr/bin/env python3
"""Audit the formula evaluator's canonical names and aliases.

This intentionally reads the real `eval_func` dispatch table instead of
maintaining a second hand-written registry. The first name in a grouped arm
is treated as canonical and the remaining names as aliases; duplicate names
are always an error because they make coverage accounting ambiguous.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


ARM_RE = re.compile(r"(?P<names>\"[A-Z][A-Z0-9.]*\"(?:\s*\|\s*\"[A-Z][A-Z0-9.]*\")*)\s*=>")
DOC_ROW_RE = re.compile(r"^\|\s*(.*?)\s*\|")
CONTRACT_NAME_RE = re.compile(r"^[A-Z][A-Z0-9.]*$")


def extract_dispatch(source: str) -> tuple[list[str], list[str]]:
    start = source.index("fn eval_func(")
    end = source.index("// ── Arithmetic", start)
    names: list[str] = []
    for match in ARM_RE.finditer(source[start:end]):
        names.extend(re.findall(r'"([A-Z][A-Z0-9.]*)"', match.group("names")))
    duplicates = sorted({name for name in names if names.count(name) > 1})
    if duplicates:
        raise ValueError("duplicate formula dispatch names: " + ", ".join(duplicates))
    canonical = []
    aliases = []
    for match in ARM_RE.finditer(source[start:end]):
        arm_names = re.findall(r'"([A-Z][A-Z0-9.]*)"', match.group("names"))
        canonical.append(arm_names[0])
        aliases.extend(arm_names[1:])
    return canonical, aliases


def extract_documented_functions(document: str) -> set[str]:
    start = document.index("## Worksheet Functions")
    end = document.index("## Criteria Syntax", start)
    names: set[str] = set()
    for line in document[start:end].splitlines():
        match = DOC_ROW_RE.match(line)
        if not match:
            continue
        names.update(re.findall(r"`([A-Z][A-Z0-9.]*)`", match.group(1)))
    return names


def validate_contracts(contract_path: Path, dispatch: set[str], documented: set[str], source: str) -> int:
    payload = json.loads(contract_path.read_text(encoding="utf-8"))
    if payload.get("schema_version") != 1:
        raise ValueError("formula contract schema_version must be 1")
    rows = payload.get("functions")
    if not isinstance(rows, list) or not rows:
        raise ValueError("formula contract functions must be a non-empty list")
    seen: set[str] = set()
    for row in rows:
        if not isinstance(row, dict):
            raise ValueError("formula contract entries must be objects")
        name = row.get("name")
        if not isinstance(name, str) or not CONTRACT_NAME_RE.fullmatch(name):
            raise ValueError(f"invalid formula contract name: {name!r}")
        if name in seen:
            raise ValueError(f"duplicate formula contract: {name}")
        seen.add(name)
        if name not in dispatch:
            raise ValueError(f"formula contract not dispatched: {name}")
        if name not in documented:
            raise ValueError(f"formula contract not documented: {name}")
        if not isinstance(row.get("argument_shape"), str) or not row["argument_shape"].strip():
            raise ValueError(f"formula contract missing argument_shape: {name}")
        if not isinstance(row.get("supported_modes"), dict) or not row["supported_modes"]:
            raise ValueError(f"formula contract missing supported_modes: {name}")
        if not isinstance(row.get("unsupported_modes"), list) or not row["unsupported_modes"]:
            raise ValueError(f"formula contract missing unsupported_modes: {name}")
        if row.get("oracle_status") != "unverified":
            raise ValueError(f"formula contract oracle_status must remain unverified: {name}")
        tests = row.get("tests")
        if not isinstance(tests, list) or not tests or not all(isinstance(test, str) and test for test in tests):
            raise ValueError(f"formula contract missing tests: {name}")
        for test in tests:
            if test not in source:
                raise ValueError(f"formula contract test not found in source: {name}: {test}")
    return len(rows)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, default=Path("src/formula/eval.rs"))
    parser.add_argument("--docs", type=Path, default=Path("FUNCTIONS.md"))
    parser.add_argument("--check-docs", action="store_true")
    parser.add_argument("--contracts", type=Path, default=Path("compat/formula-contracts.json"))
    parser.add_argument("--check-contracts", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        canonical, aliases = extract_dispatch(
            'fn eval_func(name: &str) { match name { "SUM" => x, "STDEV" | "STDEV.S" => x, } }\n'
            "// ── Arithmetic"
        )
        assert canonical == ["SUM", "STDEV"]
        assert aliases == ["STDEV.S"]
        assert extract_documented_functions(
            "## Worksheet Functions\n| `SUM` / `SUMIFS` | x |\n## Criteria Syntax"
        ) == {"SUM", "SUMIFS"}
        print("formula dispatch self-test: ok")
        return 0
    canonical, aliases = extract_dispatch(args.source.read_text(encoding="utf-8"))
    dispatch = set(canonical) | set(aliases)
    if args.check_docs or args.check_contracts:
        dispatch = set(canonical) | set(aliases)
        documented = extract_documented_functions(args.docs.read_text(encoding="utf-8"))
    if args.check_contracts:
        count = validate_contracts(args.contracts, dispatch, documented, args.source.read_text(encoding="utf-8"))
        print(f"formula contracts: ok ({count} functions)")
    if args.check_docs:
        missing = sorted(dispatch - documented)
        stale = sorted(documented - dispatch)
        if missing or stale:
            if missing:
                print("dispatch not documented: " + ", ".join(missing))
            if stale:
                print("documented but not dispatched: " + ", ".join(stale))
            return 1
        print(f"formula dispatch documentation: ok ({len(dispatch)} names)")
        return 0
    if args.check_contracts:
        return 0
    print(f"canonical={len(canonical)} aliases={len(aliases)} total={len(canonical) + len(aliases)}")
    print("canonical names: " + ", ".join(canonical))
    print("aliases: " + (", ".join(aliases) if aliases else "(none)"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
