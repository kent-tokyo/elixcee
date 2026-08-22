#!/usr/bin/env python3
"""Structural OOXML round-trip validator for elixcee's 0.9.0-A round-trip check.

Pure stdlib (zipfile, xml.etree) -- no Excel involved. This is the PRIMARY signal for
0.9.0-A's file-preservation exit criteria (relationship breakage, part loss, vbaProject
loss): it can run on every save, instantly, without a live Excel process. Excel itself is
the secondary/confirming signal, layered on top by excel_bridge.applescript -- see that
file's own header for why (a repair prompt is a modal dialog that hangs AppleScript
automation rather than returning an error, so Excel-side detection is a timeout, not a
clean check).

Every pass criterion in 0.9.0-A's spec (repair warnings 0, vbaProject loss 0, relationship
breakage 0, ...) is a zero. A checker that structurally cannot detect a failure reports all
zeros whether or not elixcee is actually correct. See self_test() below, which deliberately
corrupts two copies and asserts this checker catches both, BEFORE trusting any real result.
Run standalone (`python3 mechanical_check.py --self-test`) to execute just that check.
"""
from __future__ import annotations

import hashlib
import json
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
SML_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

VBA_ROOT_CONTENT_TYPE = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"


def _read_zip_entries(path):
    """{name: bytes} for every non-directory entry."""
    entries = {}
    with zipfile.ZipFile(path) as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            entries[info.filename] = zf.read(info.filename)
    return entries


def _parse_content_types(xml_bytes):
    """-> (defaults: {ext(lowercase, no dot): type}, overrides: {partname: type})"""
    root = ET.fromstring(xml_bytes)
    defaults, overrides = {}, {}
    for el in root.findall(f"{CT_NS}Default"):
        defaults[el.get("Extension", "").lower()] = el.get("ContentType", "")
    for el in root.findall(f"{CT_NS}Override"):
        overrides[el.get("PartName", "")] = el.get("ContentType", "")
    return defaults, overrides


def _resolve_content_type(part_name, defaults, overrides):
    key = "/" + part_name if not part_name.startswith("/") else part_name
    if key in overrides:
        return overrides[key]
    ext = part_name.rsplit(".", 1)[-1].lower() if "." in part_name else ""
    return defaults.get(ext)


def _rels_target_dir(rels_part_name):
    # e.g. "xl/_rels/workbook.xml.rels" governs targets relative to "xl/"
    m = re.match(r"^(.*)/_rels/[^/]+\.rels$", rels_part_name)
    return m.group(1) + "/" if m else ""


def _normalize_part_path(path):
    parts = []
    for seg in path.split("/"):
        if seg in ("", "."):
            continue
        if seg == "..":
            if parts:
                parts.pop()
            continue
        parts.append(seg)
    return "/".join(parts)


_WRITER_OWNED_FIXED = frozenset(
    {
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/sharedStrings.xml",
        "xl/styles.xml",
    }
)


def default_edited_parts(entries):
    """Mirrors is_writer_owned_part() in src/lib.rs: the fixed set of parts elixcee's
    writer always regenerates, plus any xl/worksheets/*.xml (no subdirectory) present in
    `entries` (a dict or iterable of part names). Pass this as `edited_parts` unless a
    fixture's macro also edits some OTHER part directly (none do in 0.9.0-A)."""
    return {
        name
        for name in entries
        if name in _WRITER_OWNED_FIXED
        or (name.startswith("xl/worksheets/") and name.endswith(".xml") and "/" not in name[len("xl/worksheets/") :])
    }


def _referenced_targets(entries):
    """Every part path referenced by ANY internal (non-External) relationship in ANY
    .rels file across `entries` -- used to detect an orphaned part: physically present
    and byte-identical, but nothing points to it any more. This is exactly the bug found
    authoring elixcee's first real-Excel fixture: xl/theme/theme1.xml passed through
    correctly, but the regenerated xl/_rels/workbook.xml.rels dropped its relationship
    (build_xlsx_workbook_rels only ever emitted worksheet/sharedStrings/styles/vbaProject
    relationships, with no mechanism to carry over any other kind) -- real Excel refused
    to open the result outright, not even a repair prompt."""
    targets = set()
    for name, data in entries.items():
        if not name.endswith(".rels"):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError:
            continue
        base = _rels_target_dir(name)
        for rel in root.findall(f"{REL_NS}Relationship"):
            if rel.get("TargetMode") == "External":
                continue
            targets.add(_normalize_part_path(base + rel.get("Target", "")))
    return targets


def check_roundtrip(original_path, output_path, edited_parts=None):
    """Structural check of `output_path`, produced from `original_path` by elixcee.

    `edited_parts` -- zip entry names (e.g. "xl/worksheets/sheet1.xml") that are EXPECTED
    to differ from the original because elixcee's own writer regenerates them. Everything
    else is expected byte-identical passthrough. Defaults to default_edited_parts(original)
    (mirrors is_writer_owned_part in src/lib.rs) -- pass an explicit set only if a fixture's
    macro edits some OTHER part directly, which none do in 0.9.0-A.

    Returns a dict: {"violations": [...], "structural_verdict": one of
    "STRUCTURALLY_CLEAN" | "ELIXCEE_RELATIONSHIP_BREAK" | "ELIXCEE_DATA_LOSS"}.
    Never raises for a malformed *output* file -- that itself is a violation, not a crash.
    """
    violations = []
    try:
        original = _read_zip_entries(original_path)
    except (zipfile.BadZipFile, FileNotFoundError) as e:
        return {"violations": [f"original fixture unreadable: {e}"], "structural_verdict": "ORACLE_FAILURE"}

    try:
        output = _read_zip_entries(output_path)
    except (zipfile.BadZipFile, FileNotFoundError) as e:
        return {"violations": [f"output not a valid zip: {e}"], "structural_verdict": "ELIXCEE_DATA_LOSS"}

    if edited_parts is None:
        edited_parts = default_edited_parts(original)

    if "[Content_Types].xml" not in output:
        violations.append("output has no [Content_Types].xml")
        return {"violations": violations, "structural_verdict": "ELIXCEE_RELATIONSHIP_BREAK"}

    try:
        defaults, overrides = _parse_content_types(output["[Content_Types].xml"])
    except ET.ParseError as e:
        violations.append(f"[Content_Types].xml is not well-formed XML: {e}")
        return {"violations": violations, "structural_verdict": "ELIXCEE_RELATIONSHIP_BREAK"}

    # 1. Every part in the output resolves to a declared content type.
    for name in output:
        if name == "[Content_Types].xml":
            continue
        if _resolve_content_type(name, defaults, overrides) is None:
            violations.append(f"part '{name}' has no resolvable content type (no Override, no Default)")

    # 2. Every internal relationship target actually exists as a zip entry.
    for name, data in output.items():
        if not name.endswith(".rels"):
            continue
        try:
            root = ET.fromstring(data)
        except ET.ParseError as e:
            violations.append(f"'{name}' is not well-formed XML: {e}")
            continue
        base = _rels_target_dir(name)
        for rel in root.findall(f"{REL_NS}Relationship"):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target", "")
            resolved = _normalize_part_path(base + target)
            if resolved not in output:
                violations.append(
                    f"'{name}': relationship Id={rel.get('Id')} Target='{target}' "
                    f"resolves to '{resolved}', which does not exist in the output"
                )

    # 2b. The dual check: a part that was referenced in the original and still exists
    # (byte-identical, i.e. genuinely passed through, not intentionally dropped) must
    # still be referenced by SOMETHING in the output -- an orphaned part (present but
    # unreferenced) is exactly as invalid to Excel as a dangling reference, but #2 above
    # can't see it (it only walks forward from relationships, never asks "is this part
    # pointed to by anything").
    orig_referenced = _referenced_targets(original)
    out_referenced = _referenced_targets(output)
    for name in orig_referenced:
        if name in output and original.get(name) == output.get(name) and name not in out_referenced:
            violations.append(
                f"'{name}' was referenced by a relationship in the original and survived "
                f"byte-identical into the output, but no relationship in the output "
                f"references it any more (orphaned part)"
            )

    # 3. vbaProject.bin: present in original => must survive byte-identical in an .xlsm
    #    output (0.9.0-A never edits VBA source, so ANY diff here is a loss, not an
    #    intentional rewrite).
    original_vba = {n: d for n, d in original.items() if n.startswith("xl/vbaProject")}
    if original_vba:
        is_macro_output = output_path.lower().endswith(".xlsm")
        for name, orig_bytes in original_vba.items():
            if is_macro_output:
                if name not in output:
                    violations.append(f"'{name}' present in original but missing from .xlsm output")
                elif output[name] != orig_bytes:
                    violations.append(
                        f"'{name}' differs from original "
                        f"(orig {len(orig_bytes)}B sha256={hashlib.sha256(orig_bytes).hexdigest()[:12]}, "
                        f"output {len(output.get(name, b''))}B "
                        f"sha256={hashlib.sha256(output.get(name, b'')).hexdigest()[:12]})"
                    )
            elif name in output:
                violations.append(f"'{name}' leaked into non-.xlsm output '{output_path}'")

        if is_macro_output:
            root_override = overrides.get("/xl/workbook.xml")
            if root_override != VBA_ROOT_CONTENT_TYPE:
                violations.append(
                    f"original has a VBA project but output's /xl/workbook.xml content type "
                    f"is '{root_override}', not the macro-enabled type"
                )

    # 4. Every original part not in edited_parts must survive byte-identical
    #    (elixcee's passthrough claim, checked directly rather than trusted).
    for name, orig_bytes in original.items():
        if name in edited_parts or name.startswith("xl/vbaProject"):
            continue
        if name not in output:
            violations.append(f"'{name}' present in original, missing from output (not in edited_parts)")
        elif output[name] != orig_bytes:
            violations.append(f"'{name}' changed but was not in edited_parts (passthrough should be byte-identical)")

    if not violations:
        verdict = "STRUCTURALLY_CLEAN"
    elif any("relationship" in v.lower() or "content type" in v.lower() for v in violations):
        verdict = "ELIXCEE_RELATIONSHIP_BREAK"
    else:
        verdict = "ELIXCEE_DATA_LOSS"

    return {"violations": violations, "structural_verdict": verdict}


def _sheet_name_to_part(entries):
    """{sheet name: worksheet part path} for a workbook's zip entries -- resolved via
    workbook.xml's <sheet name=... r:id=...> and workbook.xml.rels, NOT by assuming a
    fixed part-name pattern, since elixcee's writer renumbers worksheet parts sequentially
    on every save (sheet3.xml in the original can become sheet1.xml in the output) --
    comparing by physical part name would silently skip every sheet after a rename."""
    if "xl/workbook.xml" not in entries or "xl/_rels/workbook.xml.rels" not in entries:
        return {}
    rels_root = ET.fromstring(entries["xl/_rels/workbook.xml.rels"])
    rid_to_target = {
        rel.get("Id"): rel.get("Target") for rel in rels_root.findall(f"{REL_NS}Relationship")
    }
    wb_root = ET.fromstring(entries["xl/workbook.xml"])
    R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    result = {}
    for sheet_el in wb_root.findall(f"{SML_NS}sheets/{SML_NS}sheet"):
        rid = sheet_el.get(f"{R_NS}id")
        target = rid_to_target.get(rid)
        if target is None:
            continue
        part = _normalize_part_path(("xl/" + target) if not target.startswith("/") else target[1:])
        result[sheet_el.get("name")] = part
    return result


def _formula_cells(sheet_xml_bytes):
    """{cell_ref: formula text} for every <c> with inline <f> text in one worksheet part.
    Shared-formula follower cells (<f t="shared" si="N"/>, no inline text) are skipped --
    reader.rs doesn't resolve those either (see WorkbookSheet.formulas' doc comment), so
    they're not something elixcee could have preserved in the first place; flagging them
    here would be a false positive against a pre-existing, documented limitation."""
    root = ET.fromstring(sheet_xml_bytes)
    out = {}
    for c in root.findall(f".//{SML_NS}c"):
        f_el = c.find(f"{SML_NS}f")
        if f_el is not None and f_el.text:
            out[c.get("r")] = f_el.text
    return out


def check_formula_preservation(original_path, output_path):
    """Semantic check the structural pass above CANNOT do: worksheet XML is writer-owned
    (always regenerated), so a formula silently flattened to a bare cached value produces
    zero structural violations -- this is exactly the bug found authoring elixcee's first
    real-Excel round-trip fixture (A2*2 in A3 survived as a stale literal `20` with no <f>
    at all). Returns {"violations": [...], "formula_verdict": "CLEAN"|"ELIXCEE_DATA_LOSS"}.
    """
    violations = []
    try:
        original = _read_zip_entries(original_path)
        output = _read_zip_entries(output_path)
    except (zipfile.BadZipFile, FileNotFoundError) as e:
        return {"violations": [f"unreadable: {e}"], "formula_verdict": "ORACLE_FAILURE"}

    orig_sheets = _sheet_name_to_part(original)
    # elixcee's writer always lowercases sheet names on save (Vm::active_sheet's own
    # documented invariant, see populate_from_sheets) -- match case-insensitively so that
    # normal, expected renaming doesn't read as "sheet missing, formulas lost".
    out_sheets = {name.lower(): part for name, part in _sheet_name_to_part(output).items()}
    for name, orig_part in orig_sheets.items():
        if orig_part not in original:
            continue
        orig_formulas = _formula_cells(original[orig_part])
        if not orig_formulas:
            continue
        out_part = out_sheets.get(name.lower())
        out_formulas = _formula_cells(output[out_part]) if out_part and out_part in output else {}
        for cell_ref, formula in orig_formulas.items():
            if cell_ref not in out_formulas:
                violations.append(
                    f"sheet '{name}' cell {cell_ref}: formula '{formula}' present in original, "
                    f"missing (flattened to a bare value, or cell dropped) in output"
                )

    verdict = "ELIXCEE_DATA_LOSS" if violations else "CLEAN"
    return {"violations": violations, "formula_verdict": verdict}


def self_test():
    """Negative calibration: corrupt two copies, assert this checker actually catches
    both, PLUS assert a genuinely clean pass-through reports clean. Run before trusting
    any real fixture result -- a checker that can't fail is worthless. See module
    docstring."""
    import os
    import shutil
    import tempfile

    with tempfile.TemporaryDirectory() as tmp:
        orig_path = os.path.join(tmp, "orig.xlsm")
        with zipfile.ZipFile(orig_path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                f'<?xml version="1.0"?><Types xmlns="{CT_NS[1:-1]}">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                f'<Override PartName="/xl/workbook.xml" ContentType="{VBA_ROOT_CONTENT_TYPE}"/>'
                '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "_rels/.rels",
                f'<?xml version="1.0"?><Relationships xmlns="{REL_NS[1:-1]}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                "</Relationships>",
            )
            zf.writestr(
                "xl/_rels/workbook.xml.rels",
                f'<?xml version="1.0"?><Relationships xmlns="{REL_NS[1:-1]}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vbaProject" Target="vbaProject.bin"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
                "</Relationships>",
            )
            zf.writestr("xl/theme/theme1.xml", f'<?xml version="1.0"?><theme xmlns="{SML_NS[1:-1]}"/>')
            R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            zf.writestr(
                "xl/workbook.xml",
                f'<?xml version="1.0"?><workbook xmlns="{SML_NS[1:-1]}" xmlns:r="{R_NS}">'
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId2"/></sheets></workbook>',
            )
            zf.writestr(
                "xl/worksheets/sheet1.xml",
                f'<?xml version="1.0"?><worksheet xmlns="{SML_NS[1:-1]}"><sheetData>'
                '<row r="1"><c r="A1"><v>10</v></c>'
                '<c r="A2"><f>A1*2</f><v>20</v></c></row>'
                "</sheetData></worksheet>",
            )
            zf.writestr("xl/vbaProject.bin", b"\xd0\xcf\x11\xe0" + b"\x42" * 200)  # OLE magic + fill

        # --- Case A: clean passthrough copy (only workbook.xml "edited") -> must be clean.
        clean_path = os.path.join(tmp, "clean.xlsm")
        shutil.copyfile(orig_path, clean_path)
        result = check_roundtrip(orig_path, clean_path, edited_parts={"xl/workbook.xml"})
        assert result["structural_verdict"] == "STRUCTURALLY_CLEAN", (
            f"false positive on a genuinely clean copy: {result['violations']}"
        )

        # --- Case B: truncate vbaProject.bin -> must be caught as data loss.
        truncated_path = os.path.join(tmp, "truncated.xlsm")
        with zipfile.ZipFile(orig_path) as src, zipfile.ZipFile(truncated_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/vbaProject.bin":
                    data = data[:20]
                dst.writestr(item, data)
        result = check_roundtrip(orig_path, truncated_path, edited_parts={"xl/workbook.xml"})
        assert result["structural_verdict"] != "STRUCTURALLY_CLEAN", "failed to detect a truncated vbaProject.bin"
        assert any("vbaProject" in v for v in result["violations"]), result["violations"]

        # --- Case C: drop the vbaProject relationship -> must be caught as a rel break.
        broken_rels_path = os.path.join(tmp, "broken_rels.xlsm")
        with zipfile.ZipFile(orig_path) as src, zipfile.ZipFile(broken_rels_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/_rels/workbook.xml.rels":
                    data = f'<?xml version="1.0"?><Relationships xmlns="{REL_NS[1:-1]}"></Relationships>'.encode()
                dst.writestr(item, data)
        result = check_roundtrip(orig_path, broken_rels_path, edited_parts={"xl/workbook.xml"})
        # dropping the rels entry means vbaProject.bin is now an "extra" passthrough part
        # still present and byte-identical, so this specific mutation is caught as data
        # loss on xl/_rels/workbook.xml.rels itself (an unedited part that changed), not a
        # dangling-relationship case -- assert it's caught at all, not the exact bucket.
        assert result["structural_verdict"] != "STRUCTURALLY_CLEAN", "failed to detect a broken relationships file"

        # --- Case D: dangling relationship target -> must be caught as a rel break.
        dangling_path = os.path.join(tmp, "dangling.xlsm")
        with zipfile.ZipFile(orig_path) as src, zipfile.ZipFile(dangling_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/_rels/workbook.xml.rels":
                    data = data.replace(b"vbaProject.bin", b"vbaProjectMISSING.bin")
                dst.writestr(item, data)
        result = check_roundtrip(orig_path, dangling_path, edited_parts={"xl/workbook.xml"})
        assert result["structural_verdict"] == "ELIXCEE_RELATIONSHIP_BREAK", result
        assert any("resolves to" in v for v in result["violations"]), result["violations"]

        # --- Case E: formula preservation. This is the exact bug found authoring elixcee's
        # first real-Excel fixture: worksheet XML is writer-owned, so a stripped <f> is
        # structurally invisible (Case A's own STRUCTURALLY_CLEAN copy has an intact
        # formula, but a naive structural pass wouldn't notice if it didn't). Confirm the
        # clean copy reports CLEAN, then confirm a copy with <f> stripped is caught.
        clean_formula_result = check_formula_preservation(orig_path, clean_path)
        assert clean_formula_result["formula_verdict"] == "CLEAN", clean_formula_result

        flattened_path = os.path.join(tmp, "flattened.xlsm")
        with zipfile.ZipFile(orig_path) as src, zipfile.ZipFile(flattened_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(b"<f>A1*2</f>", b"")
                dst.writestr(item, data)
        result = check_formula_preservation(orig_path, flattened_path)
        assert result["formula_verdict"] == "ELIXCEE_DATA_LOSS", result
        assert any("A2" in v and "formula" in v for v in result["violations"]), result["violations"]

        # --- Case G: orphaned part. theme1.xml stays byte-identical, but its
        # relationship is dropped from workbook.xml.rels -- this is the exact real bug
        # found opening elixcee's first real-Excel round-trip output in actual Excel
        # (which refused to open the file outright, not even a repair prompt).
        orphaned_path = os.path.join(tmp, "orphaned.xlsm")
        with zipfile.ZipFile(orig_path) as src, zipfile.ZipFile(orphaned_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/_rels/workbook.xml.rels":
                    data = data.replace(
                        b'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>',
                        b"",
                    )
                dst.writestr(item, data)
        result = check_roundtrip(orig_path, orphaned_path, edited_parts={"xl/workbook.xml"})
        assert result["structural_verdict"] != "STRUCTURALLY_CLEAN", "failed to detect an orphaned part"
        assert any("orphaned" in v for v in result["violations"]), result["violations"]

    print(
        "self_test: OK (clean pass-through clean; truncated VBA, broken rels, dangling "
        "target, stripped-formula, and orphaned part all caught)"
    )


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        self_test()
        sys.exit(0)
    if len(sys.argv) != 3:
        print("usage: mechanical_check.py <original.xlsm> <output.xlsm>  (or --self-test)", file=sys.stderr)
        sys.exit(2)
    structural = check_roundtrip(sys.argv[1], sys.argv[2])
    formulas = check_formula_preservation(sys.argv[1], sys.argv[2])
    print(json.dumps({"structural": structural, "formulas": formulas}, indent=2))
    ok = structural["structural_verdict"] == "STRUCTURALLY_CLEAN" and formulas["formula_verdict"] == "CLEAN"
    sys.exit(0 if ok else 1)
