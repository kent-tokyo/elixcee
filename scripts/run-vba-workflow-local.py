#!/usr/bin/env python3
"""Run the declared headless VBA workflow corpus through a local CLI binary.

This is an elixcee self-check only. It does not replace the Excel/xlflow oracle,
so it never changes the manifest's external measurement status.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import subprocess
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MANIFEST = ROOT / "compat" / "vba-workflows" / "manifest.json"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _element_text(element: ET.Element) -> str:
    return "".join(part.text or "" for part in element.iter() if _local_name(part.tag) == "t")


def _readback_cells(path: Path, requested: set[tuple[str, str]]) -> dict[tuple[str, str], object]:
    """Read expected cells from the saved XLSX with only the Python stdlib.

    This is deliberately a small structural readback, not an Excel oracle. It
    catches a broken ZIP, wrong worksheet routing, and lost scalar/string values
    after the CLI's publish boundary without adding a Python XLSX dependency.
    """
    with zipfile.ZipFile(path) as archive:
        if archive.testzip() is not None:
            raise ValueError("saved workbook contains a corrupt ZIP member")
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relationships
            if _local_name(relation.tag) == "Relationship"
        }
        sheets: dict[str, str] = {}
        for sheet in workbook.iter(f"{{{MAIN_NS}}}sheet"):
            relationship_id = sheet.attrib.get(f"{{{REL_NS}}}id")
            target = targets.get(relationship_id or "")
            if target:
                sheets[sheet.attrib["name"]] = posixpath.normpath(
                    target if target.startswith("xl/") else posixpath.join("xl", target)
                )

        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            shared_strings = [
                _element_text(item)
                for item in shared
                if _local_name(item.tag) == "si"
            ]

        result: dict[tuple[str, str], object] = {}
        for sheet_name, address in requested:
            sheet_path = sheets.get(sheet_name)
            if sheet_path is None:
                continue
            worksheet = ET.fromstring(archive.read(sheet_path))
            for cell in worksheet.iter(f"{{{MAIN_NS}}}c"):
                if cell.attrib.get("r") != address:
                    continue
                value = cell.find(f"{{{MAIN_NS}}}v")
                cell_type = cell.attrib.get("t")
                if cell_type == "inlineStr":
                    parsed: object = _element_text(cell)
                elif value is None or value.text is None:
                    parsed = None
                elif cell_type == "s":
                    parsed = shared_strings[int(value.text)]
                elif cell_type == "b":
                    parsed = value.text == "1"
                elif cell_type == "e":
                    parsed = value.text
                else:
                    raw = value.text
                    parsed = float(raw) if any(marker in raw for marker in ".eE") else int(raw)
                result[(sheet_name, address)] = parsed
                break
        return result


def _readback_formulas(path: Path, requested: set[tuple[str, str]]) -> set[tuple[str, str]]:
    """Return requested cells that still contain an OOXML formula element."""
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relation.attrib["Id"]: relation.attrib["Target"]
            for relation in relationships
            if _local_name(relation.tag) == "Relationship"
        }
        sheets = {}
        for sheet in workbook.iter(f"{{{MAIN_NS}}}sheet"):
            relationship_id = sheet.attrib.get(f"{{{REL_NS}}}id")
            target = targets.get(relationship_id or "")
            if target:
                sheets[sheet.attrib["name"]] = posixpath.normpath(
                    target if target.startswith("xl/") else posixpath.join("xl", target)
                )
        formulas: set[tuple[str, str]] = set()
        for sheet_name, address in requested:
            sheet_path = sheets.get(sheet_name)
            if sheet_path is None:
                continue
            worksheet = ET.fromstring(archive.read(sheet_path))
            for cell in worksheet.iter(f"{{{MAIN_NS}}}c"):
                if cell.attrib.get("r") == address and cell.find(f"{{{MAIN_NS}}}f") is not None:
                    formulas.add((sheet_name, address))
                    break
        return formulas


def _sha256(path: Path) -> str | None:
    if not path.is_file():
        return None
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _trace_metadata(result: dict[str, object]) -> tuple[int, str]:
    trace = result.get("trace", [])
    encoded = json.dumps(trace, ensure_ascii=False, sort_keys=True).encode("utf-8")
    return len(trace) if isinstance(trace, list) else 0, hashlib.sha256(encoded).hexdigest()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", required=True, type=Path)
    parser.add_argument(
        "--report",
        type=Path,
        help="write a machine-readable local verification report to this path",
    )
    args = parser.parse_args()
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    failures: list[str] = []
    report_cases: list[dict[str, object]] = []

    with tempfile.TemporaryDirectory(prefix="elixcee-vba-workflows-") as directory:
        output_dir = Path(directory)
        for case in manifest["cases"]:
            expected = case["expected"]
            if expected["result_class"] != "not_measured":
                failures.append(f"{case['id']}: local runner requires not_measured status")
                continue
            output_path = output_dir / f"{case['id']}.xlsx"
            local = case.get("local", {})
            expected_local_class = local.get("result_class", "pass")
            sentinel = local.get("existing_output")
            if sentinel is not None:
                output_path.write_bytes(str(sentinel).encode("utf-8"))
            command = [
                str(args.binary),
                str(ROOT / case["vba_source"]),
                case["entrypoint"],
                "--file",
                str(ROOT / case["workbook"]),
                "--output",
                str(output_path),
                "--trace",
                f"workflow-{case['id']}",
                "--json",
            ]
            completed = subprocess.run(command, capture_output=True, text=True, check=False)
            try:
                result = json.loads(completed.stdout)
            except json.JSONDecodeError:
                failures.append(f"{case['id']}: CLI did not return JSON: {completed.stderr.strip()}")
                continue
            if expected_local_class != "pass":
                actual_error_kind = result.get("error", {}).get("kind")
                actual_termination_class = result.get("termination_class")
                trace_count, trace_sha256 = _trace_metadata(result)
                if completed.returncode == 0 or result.get("ok") is not False:
                    failures.append(f"{case['id']}: expected local failure, got {result}")
                    continue
                if actual_error_kind != expected_local_class:
                    failures.append(
                        f"{case['id']}: expected local error {expected_local_class!r}, "
                        f"got {actual_error_kind!r}"
                    )
                if actual_termination_class != expected_local_class:
                    failures.append(
                        f"{case['id']}: expected local termination {expected_local_class!r}, "
                        f"got {actual_termination_class!r}"
                    )
                if sentinel is not None and output_path.read_bytes() != str(sentinel).encode("utf-8"):
                    failures.append(f"{case['id']}: local failure replaced existing output")
                report_cases.append(
                    {
                        "id": case["id"],
                        "workbook_sha256": case.get("workbook_sha256"),
                        "vba_source_sha256": case.get("vba_source_sha256"),
                        "external_result_class": expected["result_class"],
                        "local_expected_class": expected_local_class,
                        "local_observed_error": actual_error_kind,
                        "local_observed_termination_class": actual_termination_class,
                        "output_sha256": _sha256(output_path),
                        "trace_event_count": trace_count,
                        "trace_sha256": trace_sha256,
                        "status": "pass"
                        if actual_error_kind == expected_local_class
                        and actual_termination_class == expected_local_class
                        else "mismatch",
                    }
                )
                continue
            if completed.returncode != 0 or result.get("ok") is not True:
                failures.append(f"{case['id']}: local CLI failure: {result}")
                continue
            if result.get("termination_class") != "success":
                failures.append(
                    f"{case['id']}: expected local termination 'success', "
                    f"got {result.get('termination_class')!r}"
                )
            cells = {
                cell["address"]: cell.get("value")
                for cell in result.get("cells", [])
                if cell.get("sheet") == "sheet1"
            }
            for qualified_address, expected_value in expected.get("cells", {}).items():
                _, address = qualified_address.split("!", 1)
                if cells.get(address) != expected_value:
                    failures.append(
                        f"{case['id']}: {qualified_address} expected {expected_value!r}, "
                        f"got {cells.get(address)!r}"
                    )
            if not output_path.is_file():
                failures.append(f"{case['id']}: CLI did not create output workbook")
                continue
            requested = {
                tuple(qualified_address.split("!", 1))
                for qualified_address in expected.get("cells", {})
            }
            try:
                saved_cells = _readback_cells(output_path, requested)
            except (OSError, KeyError, ValueError, ET.ParseError, zipfile.BadZipFile) as error:
                failures.append(f"{case['id']}: saved workbook readback failed: {error}")
                continue
            for qualified_address, expected_value in expected.get("cells", {}).items():
                sheet_name, address = qualified_address.split("!", 1)
                actual = saved_cells.get((sheet_name, address))
                if actual != expected_value:
                    failures.append(
                        f"{case['id']}: saved {qualified_address} expected {expected_value!r}, "
                        f"got {actual!r}"
                    )
            formula_addresses = {
                tuple(qualified_address.split("!", 1))
                for qualified_address in expected.get("formulas_absent", [])
            }
            present_formula_addresses = {
                tuple(qualified_address.split("!", 1))
                for qualified_address in expected.get("formulas_present", [])
            }
            try:
                saved_formulas = _readback_formulas(
                    output_path, formula_addresses | present_formula_addresses
                )
            except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile) as error:
                failures.append(f"{case['id']}: saved formula readback failed: {error}")
            else:
                for sheet_name, address in sorted(saved_formulas):
                    if (sheet_name, address) in formula_addresses:
                        failures.append(f"{case['id']}: saved {sheet_name}!{address} unexpectedly retains a formula")
                for sheet_name, address in sorted(present_formula_addresses - saved_formulas):
                    failures.append(f"{case['id']}: saved {sheet_name}!{address} lost its formula")
            report_cases.append(
                {
                    "id": case["id"],
                    "workbook_sha256": case.get("workbook_sha256"),
                    "vba_source_sha256": case.get("vba_source_sha256"),
                    "external_result_class": expected["result_class"],
                    "local_expected_class": expected_local_class,
                    "local_observed_class": "pass",
                    "local_observed_termination_class": result.get("termination_class"),
                    "output_sha256": _sha256(output_path),
                    "trace_event_count": _trace_metadata(result)[0],
                    "trace_sha256": _trace_metadata(result)[1],
                    "status": "pass" if output_path.is_file() else "mismatch",
                }
            )

    if args.report is not None:
        report = {
            "schema": "elixcee.vba-workflow-local-report.v1",
            "runner": "elixcee-only",
            "oracle_measurement_status": manifest["comparison"]["measurement_status"],
            "binary": str(args.binary),
            "binary_sha256": _sha256(args.binary),
            "manifest_sha256": _sha256(MANIFEST),
            "ok": not failures,
            "cases": report_cases,
            "failures": failures,
        }
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    if failures:
        for failure in failures:
            print(failure)
        return 1
    print(f"locally verified {len(manifest['cases'])} headless VBA workflow case(s); oracle measurement remains not_measured")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
