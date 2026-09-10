#!/usr/bin/env python3
"""Compare the current Microsoft Excel function table with eval_func dispatch.

The HTML is supplied by the caller so this audit remains deterministic and
does not make network access part of a build or test.
"""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


FUNCTION_LINK = re.compile(r"<td><a[^>]*>([A-Z][A-Z0-9.]*)</a>")
DISPATCH_NAME = re.compile(r'"([A-Z][A-Z0-9.]*)"')
EXTERNAL_BOUNDARY = {
    "CALL",
    "CUBEKPIMEMBER",
    "CUBEMEMBER",
    "CUBEMEMBERPROPERTY",
    "CUBERANKEDMEMBER",
    "CUBESET",
    "CUBESETCOUNT",
    "CUBEVALUE",
    "IMAGE",
    "REGISTER.ID",
    "RTD",
    "STOCKHISTORY",
    "TRANSLATE",
    "WEBSERVICE",
}


def audit(html: str, source: str) -> dict[str, object]:
    official = set(FUNCTION_LINK.findall(html))
    dispatch_start = source.index("match name {")
    dispatch_end = source.index("\n    }", dispatch_start)
    dispatch = set(DISPATCH_NAME.findall(source[dispatch_start:dispatch_end]))
    missing = sorted(official - dispatch - EXTERNAL_BOUNDARY)
    return {
        "official_function_names": len(official),
        "dispatch_literals": len(dispatch),
        "external_boundary": sorted(EXTERNAL_BOUNDARY),
        "missing_after_external_boundary": missing,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("html", type=Path, help="saved Microsoft function-list HTML")
    parser.add_argument(
        "--source", type=Path, default=Path("src/formula/eval.rs")
    )
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    result = audit(args.html.read_text(encoding="utf-8"), args.source.read_text())
    print(json.dumps(result, indent=2) + "\n")
    if args.check and result["missing_after_external_boundary"]:
        raise SystemExit("missing safe Excel function dispatch")


if __name__ == "__main__":
    main()
