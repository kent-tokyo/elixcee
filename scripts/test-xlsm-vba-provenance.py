#!/usr/bin/env python3
"""Regression checks for the safe XLSM provenance boundary."""

from __future__ import annotations

import importlib.util
import json
import tempfile
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
AUDITOR_PATH = ROOT / "scripts" / "audit-xlsm-vba-provenance.py"


def load_auditor():
    spec = importlib.util.spec_from_file_location("xlsm_provenance_auditor", AUDITOR_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {AUDITOR_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def make_package(path: Path, *, embedded: bool, package_source: bool) -> None:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", "<Types/>")
        package.writestr("xl/workbook.xml", '<workbook><workbookPr codeName="ThisWorkbook"/></workbook>')
        package.writestr(
            "xl/worksheets/sheet1.xml",
            '<worksheet><sheetPr codeName="Sheet1"/></worksheet>',
        )
        if embedded:
            package.writestr("xl/vbaProject.bin", b"opaque fixture bytes")
            package.writestr(
                "xl/_rels/workbook.xml.rels",
                b'<Relationships><Relationship Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"/></Relationships>',
            )
        if package_source:
            package.writestr("xl/vba/Module1.bas", "Sub Main()\nEnd Sub\n")


def main() -> int:
    auditor = load_auditor()
    with tempfile.TemporaryDirectory(prefix=".elixcee-provenance-case-", dir=ROOT) as temporary:
        directory = Path(temporary)
        embedded = directory / "embedded.xlsm"
        package_source = directory / "package-source.xlsm"
        external_only = directory / "external-only.xlsx"
        external_source = directory / "Module1.bas"
        make_package(embedded, embedded=True, package_source=False)
        make_package(package_source, embedded=True, package_source=True)
        make_package(external_only, embedded=False, package_source=False)
        external_source.write_text("Sub Main()\nEnd Sub\n", encoding="utf-8")

        result = auditor.audit(embedded, [], auditor.DEFAULT_MAX_TOTAL_BYTES)
        assert result["classification"] == "embedded-binary-only"
        assert result["embedded"]["module_identity"] == "not_available_without_ole_source_parser"
        assert result["embedded"]["execution_source"] == "not_selected"
        assert result["execution"] == "not_performed"
        assert result["embedded"]["workbook_code_name"] == "ThisWorkbook"
        assert result["embedded"]["sheet_code_names"] == {"xl/worksheets/sheet1.xml": "Sheet1"}
        assert result["embedded"]["vba_relationship_parts"] == ["xl/_rels/workbook.xml.rels"]

        result = auditor.audit(package_source, [external_source], auditor.DEFAULT_MAX_TOTAL_BYTES)
        assert result["classification"] == "embedded-binary-and-external-source"
        assert result["embedded"]["source_available_in_package"] is True
        assert result["embedded"]["source_parts"] == ["xl/vba/Module1.bas"]
        assert result["external_sources"][0]["exists"] is True
        assert len(result["external_sources"][0]["sha256"]) == 64

        result = auditor.audit(external_only, [external_source], auditor.DEFAULT_MAX_TOTAL_BYTES)
        assert result["classification"] == "external-source-only"
        assert result["embedded"]["vba_project_present"] is False

        try:
            auditor.audit(embedded, [], 1)
        except ValueError as error:
            assert "exceeds budget" in str(error)
        else:
            raise AssertionError("the uncompressed-size budget must reject the fixture")

        print(json.dumps({"ok": True, "cases": 4}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
