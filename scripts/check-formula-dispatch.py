#!/usr/bin/env python3
"""Audit the formula evaluator's canonical names and aliases.

This intentionally reads the real `eval_func` dispatch table instead of
maintaining a second hand-written registry. The first name in a grouped arm
is treated as canonical and the remaining names as aliases; duplicate names
are always an error because they make coverage accounting ambiguous.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path


ARM_RE = re.compile(r"(?P<names>\"[A-Z][A-Z0-9.]*\"(?:\s*\|\s*\"[A-Z][A-Z0-9.]*\")*)\s*=>")
DOC_ROW_RE = re.compile(r"^\|\s*(.*?)\s*\|")


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


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, default=Path("src/formula/eval.rs"))
    parser.add_argument("--docs", type=Path, default=Path("FUNCTIONS.md"))
    parser.add_argument("--check-docs", action="store_true")
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
    if args.check_docs:
        dispatch = set(canonical) | set(aliases)
        documented = extract_documented_functions(args.docs.read_text(encoding="utf-8"))
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
    print(f"canonical={len(canonical)} aliases={len(aliases)} total={len(canonical) + len(aliases)}")
    print("canonical names: " + ", ".join(canonical))
    print("aliases: " + (", ".join(aliases) if aliases else "(none)"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
