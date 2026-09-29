#!/usr/bin/env python3
"""Audit XLSM VBA provenance without parsing or executing OLE/VBA content.

The runtime deliberately distinguishes a preserved ``vbaProject.bin`` from
source that can be parsed and executed.  This small, dependency-free auditor
records that distinction in a reproducible JSON document instead of guessing
module names from an opaque OLE container.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


DEFAULT_MAX_TOTAL_BYTES = 256 * 1024 * 1024
MAX_XML_PART_BYTES = 4 * 1024 * 1024


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def xml_attribute(data: bytes, local_name: str) -> str | None:
    """Read one OOXML attribute without inspecting or executing OLE data."""
    if len(data) > MAX_XML_PART_BYTES:
        raise ValueError(f"XML part exceeds audit budget: {len(data)} > {MAX_XML_PART_BYTES}")
    root = ET.fromstring(data)
    for element in root.iter():
        for key, value in element.attrib.items():
            if key.rsplit("}", 1)[-1] == local_name:
                return value
    return None


def audit(path: Path, sources: list[Path], max_total_bytes: int) -> dict[str, object]:
    if not path.is_file():
        raise FileNotFoundError(path)

    source_records = [
        {
            "path": str(source),
            "exists": source.is_file(),
            "sha256": sha256_file(source) if source.is_file() else None,
            "kind": "external-source",
        }
        for source in sources
    ]

    with zipfile.ZipFile(path) as package:
        entries = package.infolist()
        total_uncompressed = sum(entry.file_size for entry in entries)
        if total_uncompressed > max_total_bytes:
            raise ValueError(
                f"uncompressed XLSM package exceeds budget: {total_uncompressed} > {max_total_bytes}"
            )
        names = sorted(entry.filename for entry in entries)
        vba_parts = [name for name in names if name.startswith("xl/vbaProject")]
        source_parts = [
            name
            for name in names
            if name.lower().endswith((".bas", ".cls", ".frm"))
            or ("vba" in name.lower() and "project" not in name.lower())
        ]
        workbook_code_name = None
        sheet_code_names = {}
        if "xl/workbook.xml" in names:
            workbook_code_name = xml_attribute(package.read("xl/workbook.xml"), "codeName")
        for name in names:
            if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                code_name = xml_attribute(package.read(name), "codeName")
                if code_name is not None:
                    sheet_code_names[name] = code_name
        relationship_parts = []
        for name in names:
            if name.endswith(".rels") and b"/vbaProject" in package.read(name):
                relationship_parts.append(name)
        signature_parts = [
            name for name in names if name.lower().endswith("vbaprojectsignature.bin")
        ]

    has_embedded_binary = any(name.endswith("vbaProject.bin") for name in vba_parts)
    existing_sources = [record for record in source_records if record["exists"]]
    return {
        "schema": "elixcee.vba-provenance.v1",
        "workbook": {
            "path": str(path),
            "sha256": sha256_file(path),
            "suffix": path.suffix.lower(),
        },
        "embedded": {
            "vba_project_present": has_embedded_binary,
            "parts": vba_parts,
            "source_parts": source_parts,
            "source_available_in_package": bool(source_parts),
            "workbook_code_name": workbook_code_name,
            "sheet_code_names": sheet_code_names,
            "vba_relationship_parts": relationship_parts,
            "signature_parts": signature_parts,
            "module_identity": "not_available_without_ole_source_parser",
            "execution_source": "not_selected",
        },
        "external_sources": source_records,
        "classification": (
            "embedded-binary-and-external-source"
            if has_embedded_binary and existing_sources
            else "embedded-binary-only"
            if has_embedded_binary
            else "external-source-only"
            if existing_sources
            else "no-executable-vba-source"
        ),
        "limits": {"max_uncompressed_bytes": max_total_bytes},
        "execution": "not_performed",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook", type=Path, help="XLSM/XLSB-like package to inspect")
    parser.add_argument(
        "--source",
        type=Path,
        action="append",
        default=[],
        help="external .bas/.cls/.frm source supplied separately (repeatable)",
    )
    parser.add_argument(
        "--max-uncompressed-bytes",
        type=int,
        default=DEFAULT_MAX_TOTAL_BYTES,
    )
    args = parser.parse_args()
    try:
        result = audit(args.workbook, args.source, args.max_uncompressed_bytes)
    except (OSError, ValueError, ET.ParseError, zipfile.BadZipFile) as error:
        print(f"audit failed: {error}", file=sys.stderr)
        return 2
    print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
