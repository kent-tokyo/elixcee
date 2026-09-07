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


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, default=Path("src/formula/eval.rs"))
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()
    if args.self_test:
        canonical, aliases = extract_dispatch(
            'fn eval_func(name: &str) { match name { "SUM" => x, "STDEV" | "STDEV.S" => x, } }\n'
            "// ── Arithmetic"
        )
        assert canonical == ["SUM", "STDEV"]
        assert aliases == ["STDEV.S"]
        print("formula dispatch self-test: ok")
        return 0
    canonical, aliases = extract_dispatch(args.source.read_text(encoding="utf-8"))
    print(f"canonical={len(canonical)} aliases={len(aliases)} total={len(canonical) + len(aliases)}")
    print("canonical names: " + ", ".join(canonical))
    print("aliases: " + (", ".join(aliases) if aliases else "(none)"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
