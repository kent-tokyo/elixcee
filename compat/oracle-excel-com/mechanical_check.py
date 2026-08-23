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

check_source_references() (0.10.0-A) closes a blind spot found while investigating
0.10.0's design: a worksheet-level relationship's `.rels` entry and target part can both
survive a save byte-identical, while the regenerated worksheet XML no longer contains the
r:id reference that activates it -- check_roundtrip() above cannot see this (its orphan
check only walks the .rels graph, never asks whether worksheet CONTENT still points at a
relationship). Confirmed as a real, not hypothetical, gap: running elixcee's own save
against fixture3/4/5 under `fixtures/pristine/` (every fixture with a worksheet-level
relationship at all) reproduces SOURCE_REFERENCE_LOSS in every one of them, while
check_roundtrip() reports STRUCTURALLY_CLEAN on all three -- see
docs/xlsx-worksheet-preservation-0.10.0-design.md §4/§9 and
fixtures/pristine/INVENTORY.md for the full account.

check_inline_worksheet_elements() (0.10.0-B) covers a different, relationship-FREE gap:
direct children of <worksheet> like <sheetViews> (freeze panes, active-cell selection)
carry no r:id at all, so check_source_references() has nothing to look for -- yet today's
writer (before 0.10.0-B) never emits them either, silently dropping view state on every
save. INLINE_ELEMENT_LOSS is a third, distinct violation category from both
SOURCE_REFERENCE_LOSS (r:id-bearing) and check_roundtrip()'s structural checks (ZIP-part
level, not worksheet-content level).
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


R_ID_ATTR = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

# Relationship types whose worksheet-side activation is a specific r:id-bearing element
# in the SAME worksheet part's own content -- each entry maps a relationship Type URI to
# the XPath (relative to the worksheet's root <worksheet> element) of the element shape
# that carries the reference. Confirmed two ways before being added here: (1) against a
# real fixture's actual XML (see fixtures/pristine/INVENTORY.md) and (2) against the real
# ECMA-376 sml.xsd (OfficeOpenXML-XMLSchema-Transitional/sml.xsd -- CT_TablePart/
# CT_Drawing/CT_Hyperlink/CT_LegacyDrawing all declare r:id directly), not assumed from
# memory -- see docs/xlsx-worksheet-preservation-0.10.0-design.md §8's own account of why
# memory-only schema recall isn't trusted in this project.
#
# Deliberately does NOT include http://.../comments or .../threadedComment: confirmed
# empirically (fixture4_hyperlink_comment_name.xlsm) that neither relationship is
# referenced by any r:id anywhere in xl/worksheets/sheet1.xml -- both are located purely
# by relationship Type within the sheet's own .rels file. A same-shaped check for them
# would either always report nothing (nothing to find) or, worse, have to special-case an
# exception, so their loss is left to the existing ORPHANED_PART check (check_roundtrip's
# #2b above), which already covers it correctly.
#
# printerSettings/oleObject/control are confirmed via the real XSD to need an r:id
# reference too (<pageSetup r:id>, <oleObjects><oleObject r:id>, <controls><control
# r:id>), but no fixture in this repo exercises any of them yet (see INVENTORY.md's
# "confirmed absent" list) -- deliberately left out of this table until one does, rather
# than shipping a checker path that has never run against real data.
_WORKSHEET_RID_ELEMENT_XPATHS = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table": (
        f"{SML_NS}tableParts/{SML_NS}tablePart"
    ),
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing": f"{SML_NS}drawing",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink": (
        f"{SML_NS}hyperlinks/{SML_NS}hyperlink"
    ),
    # legacyDrawingHF (header/footer VML) shares the same CT_LegacyDrawing shape and
    # could in principle carry a vmlDrawing relationship too, but no fixture here
    # exercises it -- only <legacyDrawing> (not <legacyDrawingHF>) is checked until one
    # does, matching this table's own stated policy of fixture-confirmed rows only.
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing": f"{SML_NS}legacyDrawing",
}


def _rid_referenced_as_type(sheet_root, rel_type, rid):
    """True iff `rid` is referenced by the SPECIFIC element shape
    _WORKSHEET_RID_ELEMENT_XPATHS maps `rel_type` to -- not "referenced by any mapped
    shape at all". Type-aware on purpose: rId strings are only unique within one .rels
    file, so a flat union across all types would let a table relationship's rId "count"
    as satisfied merely because some unrelated drawing element happens to reuse the same
    string, mislabeling (or silently swallowing) a real SOURCE_REFERENCE_LOSS. See
    self_test()'s Case H "type confusion" case, which exists specifically to catch a
    regression back to the flat-union form."""
    xpath = _WORKSHEET_RID_ELEMENT_XPATHS.get(rel_type)
    if xpath is None:
        return None  # unmapped type -- caller's job to skip, not this function's
    return any(el.get(R_ID_ATTR) == rid for el in sheet_root.findall(xpath))


def check_source_references(original_path, output_path):
    """Detects SOURCE_REFERENCE_LOSS: a worksheet-level relationship (and its target
    part) survives byte-for-byte in the output's own `_rels/sheetN.xml.rels`, but the
    regenerated worksheet XML no longer contains the r:id reference that activates it.
    Structurally different from ORPHANED_PART (nothing in any .rels graph references the
    part) and DANGLING_RELATIONSHIP (a .rels entry's target doesn't exist as a zip entry)
    -- both of those are covered by check_roundtrip() above. Here the .rels entry AND its
    target both exist and are individually valid; nothing in the CONSUMING XML content
    points at the relationship any more, so the feature it backs (table/hyperlink/
    drawing/legacyDrawing) is silently inert. Confirmed as a real, previously-undetected
    gap by actually running elixcee against fixture3_table_validation_conditional.xlsm and
    observing check_roundtrip() report STRUCTURALLY_CLEAN on a table whose <tableParts>
    reference had vanished -- see docs/xlsx-worksheet-preservation-0.10.0-design.md §4/§9.

    A worksheet-level `.rels` file is only ever checked when it survives BYTE-IDENTICAL
    into the output: elixcee's writer today never emits its own worksheet-level
    relationships (see is_writer_owned_part() in src/lib.rs -- xl/worksheets/_rels/*.rels
    never matches its writer-owned pattern), so any worksheet .rels present in a real
    elixcee output is, by construction, a passthrough copy of the source's own. **A
    mismatch here is itself reported as a violation, never silently skipped** -- an
    earlier version of this function treated a changed `.rels` as merely "out of scope"
    and moved on, which meant the day this assumption stops holding (e.g. a future
    0.10.0-D writer change that regenerates or renumbers worksheet-level relationships),
    this check would go quiet on exactly the files it exists to guard, reporting CLEAN
    because it stopped looking rather than because the bug is fixed -- the same failure
    shape this whole function was written to catch in check_roundtrip(). Flagging it loud
    instead means a genuinely new, deliberate writer behavior here requires this check to
    be updated to understand it, not silently bypassed.

    Returns {"violations": [...], "source_reference_verdict": "CLEAN"|"SOURCE_REFERENCE_LOSS"}.
    """
    violations = []
    try:
        original = _read_zip_entries(original_path)
        output = _read_zip_entries(output_path)
    except (zipfile.BadZipFile, FileNotFoundError) as e:
        return {"violations": [f"unreadable: {e}"], "source_reference_verdict": "ORACLE_FAILURE"}

    for rels_name, rels_bytes in output.items():
        if not re.match(r"^xl/worksheets/_rels/[^/]+\.rels$", rels_name):
            continue
        if original.get(rels_name) != rels_bytes:
            violations.append(
                f"'{rels_name}' differs from the original (or is new) -- elixcee's writer "
                f"has never been observed to touch a worksheet-level .rels file, so this "
                f"check doesn't know how to verify one that changed; treat this as "
                f"requiring investigation, not as passing"
            )
            continue

        sheet_part = re.sub(r"_rels/([^/]+)\.rels$", r"\1", rels_name)
        if sheet_part not in output:
            continue  # a different, already-covered failure (missing worksheet part)

        try:
            sheet_root = ET.fromstring(output[sheet_part])
        except ET.ParseError as e:
            violations.append(f"'{sheet_part}' is not well-formed XML: {e}")
            continue

        try:
            rels_root = ET.fromstring(rels_bytes)
        except ET.ParseError as e:
            violations.append(f"'{rels_name}' is not well-formed XML: {e}")
            continue
        for rel in rels_root.findall(f"{REL_NS}Relationship"):
            rel_type = rel.get("Type", "")
            rid = rel.get("Id")
            referenced = _rid_referenced_as_type(sheet_root, rel_type, rid)
            if referenced is None:
                continue  # unmapped type (comments/threadedComment/unknown) -- not this check's job
            if not referenced:
                violations.append(
                    f"'{rels_name}': relationship Id={rid} Type='{rel_type}' survived "
                    f"into the output, but '{sheet_part}' no longer contains any element "
                    f"referencing it as that type (SOURCE_REFERENCE_LOSS)"
                )

    verdict = "SOURCE_REFERENCE_LOSS" if violations else "CLEAN"
    return {"violations": violations, "source_reference_verdict": verdict}


# Direct children of <worksheet> that carry NO r:id/relationship dependency at all --
# 0.10.0-B's opaque-fragment passthrough targets. Extended one element at a time, in step
# with 0.10.0-B's own slices (see docs/xlsx-worksheet-preservation-0.10.0-design.md §10):
# only listed here once a real fixture demonstrates it and this checker has a negative
# test for its loss (the same hard gate that governs writer code).
#
# slice 1 (sheetViews): fixture1-7 all carry it.
# slice 2 (sheetPr/sheetFormatPr/phoneticPr/dataValidations): sheetFormatPr and
# phoneticPr are present in all 7 fixtures; sheetPr in fixture2/3/4 only;
# dataValidations in fixture3 only -- all real, none synthesized.
# slice 3 (pageMargins): present in all 7 fixtures, same shape as sheetFormatPr/
# phoneticPr (self-closing, no children, no namespace dependency).
#
# autoFilter and conditionalFormatting are deliberately NOT here yet: autoFilter has
# zero fixture evidence as a standalone worksheet element (INVENTORY.md's "confirmed
# absent" list -- fixture3's <autoFilter> lives inside xl/tables/table1.xml, a different
# part), and conditionalFormatting can reference xl/styles.xml's <dxfs> (dxfId) or
# <extLst> extensions -- the design doc flags it as needing separate consideration
# before being treated as a pure relationship-free opaque fragment. <hyperlinks> is ALSO
# deliberately not here, even though relationship-free location-only hyperlinks exist
# (fixture6) -- a present/absent check on the whole container would false-positive on
# fixture4, whose <hyperlinks> is correctly ABSENT from a correct output (its only child
# is r:id-backed, out of this check's scope) despite being present in the source. See
# check_internal_hyperlinks() below, a dedicated per-child check instead.
_INLINE_WORKSHEET_ELEMENTS = [
    "sheetViews",
    "sheetPr",
    "sheetFormatPr",
    "phoneticPr",
    "dataValidations",
    "pageMargins",
]


def check_inline_worksheet_elements(original_path, output_path):
    """Detects INLINE_ELEMENT_LOSS: a direct child of <worksheet> in
    _INLINE_WORKSHEET_ELEMENTS is present in the original worksheet XML but absent from
    the corresponding output worksheet XML. Distinct from check_source_references()
    (r:id-bearing elements only -- <sheetViews> and friends carry no relationship at all,
    so that check has nothing to look for) and from check_formula_preservation()
    (cell-level, not worksheet-level). Sheets are matched by name via _sheet_name_to_part,
    same as check_formula_preservation -- NOT by physical part path, since elixcee's
    writer renumbers worksheet parts sequentially on every save.

    Confirmed as a real, current gap (not hypothetical) before any 0.10.0-B writer code
    was written: running elixcee's own save (load, edit one cell, --output) against every
    one of fixture1-7 under fixtures/pristine/ reproduces INLINE_ELEMENT_LOSS on every
    sheet of every fixture -- <sheetViews> is universally present in the source and
    universally absent from today's writer output (build_xlsx_sheet emits none). This is
    the pre-fix baseline this check exists to close; 0.10.0-B's writer commit should turn
    all of fixture1-7 CLEAN under this check without touching SOURCE_REFERENCE_LOSS, which
    stays unresolved on fixture3/4/5 until 0.10.0-D.

    slice 2 (sheetPr/sheetFormatPr/phoneticPr/dataValidations) re-confirmed the same way
    against the ALREADY-sheetViews-fixed writer (6a4e596): sheetFormatPr/phoneticPr lost
    on all 7 fixtures, sheetPr lost on fixture2/3/4 (the only 3 that have it), and
    dataValidations lost on fixture3 (the only one that has it) -- sheetViews itself
    correctly stayed CLEAN everywhere, confirming slice 1's fix wasn't disturbed.

    slice 3 (pageMargins) re-confirmed the same way against the slice-1+2-fixed writer
    (1494adc): pageMargins lost on all 7 fixtures (present in every one), and it was the
    ONLY element flagged anywhere -- confirming slices 1 and 2 stayed intact.

    Returns {"violations": [...], "inline_element_verdict": "CLEAN"|"INLINE_ELEMENT_LOSS"}.
    """
    violations = []
    try:
        original = _read_zip_entries(original_path)
        output = _read_zip_entries(output_path)
    except (zipfile.BadZipFile, FileNotFoundError) as e:
        return {"violations": [f"unreadable: {e}"], "inline_element_verdict": "ORACLE_FAILURE"}

    orig_sheets = _sheet_name_to_part(original)
    out_sheets = {name.lower(): part for name, part in _sheet_name_to_part(output).items()}
    for name, orig_part in orig_sheets.items():
        if orig_part not in original:
            continue
        try:
            orig_root = ET.fromstring(original[orig_part])
        except ET.ParseError:
            continue  # not this check's job -- check_roundtrip() already covers malformed parts
        present = [tag for tag in _INLINE_WORKSHEET_ELEMENTS if orig_root.find(f"{SML_NS}{tag}") is not None]
        if not present:
            continue

        out_part = out_sheets.get(name.lower())
        out_root = None
        if out_part and out_part in output:
            try:
                out_root = ET.fromstring(output[out_part])
            except ET.ParseError:
                pass
        for tag in present:
            if out_root is None or out_root.find(f"{SML_NS}{tag}") is None:
                violations.append(
                    f"sheet '{name}': <{tag}> present in original worksheet XML, missing "
                    f"from the output (INLINE_ELEMENT_LOSS)"
                )

    verdict = "INLINE_ELEMENT_LOSS" if violations else "CLEAN"
    return {"violations": violations, "inline_element_verdict": verdict}


_WORKBOOK_ELEMENTS = [
    "workbookPr",
    "calcPr",
    "extLst",
    "bookViews",
]


def check_workbook_elements(original_path, output_path):
    """Detects WORKBOOK_ELEMENT_LOSS: a direct child of the root <workbook> in
    _WORKBOOK_ELEMENTS is present in the original xl/workbook.xml but absent from the
    output's. Distinct from check_inline_worksheet_elements() (that one is per-worksheet,
    matched by sheet name since worksheet parts get renumbered on every save; workbook.xml
    is a single, fixed-path part, so no name-matching is needed here) and from
    check_source_references() (r:id-bearing worksheet relationships, not workbook-level
    metadata). 0.10.0-C, slices C1 (workbookPr/calcPr/extLst) and C2 (bookViews).

    <bookViews>'s <workbookView> carries activeTab/firstSheet attributes, which are
    sheet-position indices -- in principle unsafe to carry verbatim if a macro
    add/deletes sheets and they point at a stale position afterward. Checked all 7 real
    fixtures under fixtures/pristine/ before adding <bookViews> here: NONE of them
    actually sets activeTab or firstSheet (both default to 0, confirmed against the real
    CT_BookView XSD) -- so today, a verbatim copy is correct on every fixture that
    exists. Building carry-over/gating logic for a hazard with zero fixture evidence
    would be exactly the speculative, unvalidated machinery this milestone's hard gate
    exists to prevent (see design doc §7's stated policy against pre-built
    abstractions with nothing to validate them). If a future fixture DOES carry a
    non-default activeTab/firstSheet, this check's own "present in original, missing
    from output" logic still degrades safely -- it would only flag a real loss, not miss
    one -- but the *correctness* of a non-default value surviving position changes still
    needs its own design pass at that point; see design doc §10's C2 entry.
    <definedNames> is deliberately excluded -- its <definedName> text can embed a sheet
    name (e.g. "Sheet1!$F$5") and localSheetId is a position index too, needing the C3
    design (a different mechanism than a blind opaque-fragment copy, since a renamed
    sheet's dangling text reference needs its own handling, not just presence).

    Confirmed as a real, current gap before any 0.10.0-C writer code was written: running
    elixcee's own save against every one of fixture1-7 under fixtures/pristine/ reproduces
    WORKBOOK_ELEMENT_LOSS on workbookPr/calcPr/extLst/bookViews on every fixture --
    build_xlsx_workbook emits none of them today, only <sheets>.

    Returns {"violations": [...], "workbook_element_verdict": "CLEAN"|"WORKBOOK_ELEMENT_LOSS"}.
    """
    violations = []
    try:
        original = _read_zip_entries(original_path)
        output = _read_zip_entries(output_path)
    except (zipfile.BadZipFile, FileNotFoundError) as e:
        return {"violations": [f"unreadable: {e}"], "workbook_element_verdict": "ORACLE_FAILURE"}

    if "xl/workbook.xml" not in original:
        return {"violations": [], "workbook_element_verdict": "CLEAN"}
    try:
        orig_root = ET.fromstring(original["xl/workbook.xml"])
    except ET.ParseError:
        return {"violations": [], "workbook_element_verdict": "CLEAN"}
    present = [tag for tag in _WORKBOOK_ELEMENTS if orig_root.find(f"{SML_NS}{tag}") is not None]
    if not present:
        return {"violations": [], "workbook_element_verdict": "CLEAN"}

    out_root = None
    if "xl/workbook.xml" in output:
        try:
            out_root = ET.fromstring(output["xl/workbook.xml"])
        except ET.ParseError:
            pass
    for tag in present:
        if out_root is None or out_root.find(f"{SML_NS}{tag}") is None:
            violations.append(
                f"<{tag}> present in original xl/workbook.xml, missing from the output "
                f"(WORKBOOK_ELEMENT_LOSS)"
            )

    verdict = "WORKBOOK_ELEMENT_LOSS" if violations else "CLEAN"
    return {"violations": violations, "workbook_element_verdict": verdict}


def _sheet_names_in_workbook(root):
    """Lowercased <sheet name="..."> values, in document order, from a parsed
    <workbook> root -- used by check_defined_names() to decide whether every original
    sheet survived (definedNames must be verbatim) or some were deleted (definedNames
    may be legitimately dropped -- see that function's docstring)."""
    sheets = root.find(f"{SML_NS}sheets")
    if sheets is None:
        return []
    return [s.get("name", "").lower() for s in sheets.findall(f"{SML_NS}sheet")]


def check_defined_names(original_path, output_path):
    """Detects WORKBOOK_ELEMENT_LOSS on <definedNames>: NOT a plain presence check like
    check_workbook_elements() (workbookPr/bookViews/calcPr/extLst), because a
    <definedName>'s localSheetId is a 0-based index into <sheets> -- if a sheet was
    deleted since the source loaded, every localSheetId at or past that position is
    stale, and elixcee's writer deliberately drops <definedNames> ENTIRELY rather than
    try to remap or selectively prune individual names (see
    docs/xlsx-worksheet-preservation-0.10.0-design.md §10's C3 entry for why: partial
    remapping needs its own design, and shipping a blindly-verbatim copy would silently
    reattach a print area / named range to the wrong sheet -- worse than dropping it).

    So this check has two cases, both keyed off whether every ORIGINAL sheet name is
    still present in the output (add-only or no-mutation-at-all changes nothing about
    existing positions; a deletion does):
    - No sheet missing: <definedNames> must survive byte-for-byte identical (a name can
      embed a sheet-qualified formula reference like "Sheet1!$F$5" as free text, which
      this deliberately does NOT try to parse/validate -- only presence+exact content).
    - A sheet is missing: <definedNames> must be ABSENT from the output. If it's still
      there, that's a violation too -- the writer failed to apply its own drop rule,
      which is worse than a plain loss (a stale, wrong reference looks valid).

    Returns {"violations": [...], "defined_names_verdict": "CLEAN"|"WORKBOOK_ELEMENT_LOSS"}.
    """
    violations = []
    try:
        original = _read_zip_entries(original_path)
        output = _read_zip_entries(output_path)
    except (zipfile.BadZipFile, FileNotFoundError) as e:
        return {"violations": [f"unreadable: {e}"], "defined_names_verdict": "ORACLE_FAILURE"}

    if "xl/workbook.xml" not in original:
        return {"violations": [], "defined_names_verdict": "CLEAN"}
    try:
        orig_root = ET.fromstring(original["xl/workbook.xml"])
    except ET.ParseError:
        return {"violations": [], "defined_names_verdict": "CLEAN"}
    orig_defined_names = orig_root.find(f"{SML_NS}definedNames")
    if orig_defined_names is None:
        return {"violations": [], "defined_names_verdict": "CLEAN"}

    out_root = None
    if "xl/workbook.xml" in output:
        try:
            out_root = ET.fromstring(output["xl/workbook.xml"])
        except ET.ParseError:
            pass

    orig_sheet_names = set(_sheet_names_in_workbook(orig_root))
    out_sheet_names = set(_sheet_names_in_workbook(out_root)) if out_root is not None else set()
    every_sheet_survived = orig_sheet_names <= out_sheet_names

    out_defined_names = out_root.find(f"{SML_NS}definedNames") if out_root is not None else None
    if every_sheet_survived:
        if out_defined_names is None:
            violations.append(
                "<definedNames> present in original xl/workbook.xml, missing from the "
                "output even though every original sheet survived (WORKBOOK_ELEMENT_LOSS)"
            )
        else:
            orig_names = [ET.tostring(c, encoding="unicode") for c in orig_defined_names]
            out_names = [ET.tostring(c, encoding="unicode") for c in out_defined_names]
            if orig_names != out_names:
                violations.append(
                    f"<definedNames> content changed even though every original sheet "
                    f"survived -- expected byte-identical: {orig_names!r} -> {out_names!r} "
                    f"(WORKBOOK_ELEMENT_LOSS)"
                )
    elif out_defined_names is not None:
        violations.append(
            "a source sheet was deleted, but <definedNames> is still present in the "
            "output -- must be dropped entirely once any localSheetId could be stale "
            "(WORKBOOK_ELEMENT_LOSS)"
        )

    verdict = "WORKBOOK_ELEMENT_LOSS" if violations else "CLEAN"
    return {"violations": violations, "defined_names_verdict": verdict}


def check_internal_hyperlinks(original_path, output_path):
    """Detects INTERNAL_HYPERLINK_LOSS: a relationship-free ("location=", no "r:id")
    <hyperlink> child present in the original worksheet XML is missing from the output,
    or its location= text changed, or it unexpectedly gained an r:id. 0.10.0-B's B4 slice.

    Deliberately NOT folded into check_inline_worksheet_elements()/
    _INLINE_WORKSHEET_ELEMENTS: a whole-container present/absent check on <hyperlinks>
    would false-positive on fixture4_hyperlink_comment_name.xlsm, whose <hyperlinks> is
    correctly ABSENT from a correct output (its only child is r:id-backed, out of scope
    until 0.10.0-D reconnects the relationship graph) despite being present in the
    source. This check instead compares per-<hyperlink>-child, matched by `ref` (the cell
    address each hyperlink is anchored to -- `required` per CT_Hyperlink, unique within
    one sheet's <hyperlinks>), and only for children that have no r:id in the ORIGINAL.

    Also flags an output <hyperlinks> element with zero <hyperlink> children as itself
    invalid -- confirmed via CT_Hyperlinks' own XSD (`minOccurs="1"` on its <hyperlink>
    child, see docs/xlsx-worksheet-preservation-0.10.0-design.md §8/B4): a correct writer
    must omit <hyperlinks> entirely when nothing survives, never emit an empty
    <hyperlinks/>, which a validating consumer (or Excel) would reject.

    Confirmed as a real, current gap before any B4 writer code was written: run against
    fixture4_hyperlink_comment_name.xlsm (all r:id, out of scope) -> CLEAN, no false
    positive. Run against fixture6_internal_hyperlink.xlsm (one location-only hyperlink)
    against today's writer output -> INTERNAL_HYPERLINK_LOSS, since build_xlsx_sheet
    doesn't emit <hyperlinks> at all yet.

    Returns {"violations": [...], "internal_hyperlink_verdict": "CLEAN"|"INTERNAL_HYPERLINK_LOSS"}.
    """
    violations = []
    try:
        original = _read_zip_entries(original_path)
        output = _read_zip_entries(output_path)
    except (zipfile.BadZipFile, FileNotFoundError) as e:
        return {"violations": [f"unreadable: {e}"], "internal_hyperlink_verdict": "ORACLE_FAILURE"}

    orig_sheets = _sheet_name_to_part(original)
    out_sheets = {name.lower(): part for name, part in _sheet_name_to_part(output).items()}
    for name, orig_part in orig_sheets.items():
        if orig_part not in original:
            continue
        try:
            orig_root = ET.fromstring(original[orig_part])
        except ET.ParseError:
            continue
        location_only = [
            hl
            for hl in orig_root.findall(f"{SML_NS}hyperlinks/{SML_NS}hyperlink")
            if hl.get(R_ID_ATTR) is None
        ]
        if not location_only:
            continue

        out_part = out_sheets.get(name.lower())
        out_root = None
        if out_part and out_part in output:
            try:
                out_root = ET.fromstring(output[out_part])
            except ET.ParseError:
                pass

        out_by_ref = {}
        if out_root is not None:
            for hl in out_root.findall(f"{SML_NS}hyperlinks/{SML_NS}hyperlink"):
                out_by_ref[hl.get("ref")] = hl
            container = out_root.find(f"{SML_NS}hyperlinks")
            if container is not None and len(container) == 0:
                violations.append(
                    f"sheet '{name}': output <hyperlinks> has zero <hyperlink> children -- "
                    f"invalid per CT_Hyperlinks (minOccurs=1 on hyperlink); must be omitted "
                    f"entirely, not emitted empty"
                )

        for hl in location_only:
            ref = hl.get("ref")
            out_hl = out_by_ref.get(ref)
            if out_hl is None:
                violations.append(
                    f"sheet '{name}' ref={ref}: relationship-free hyperlink "
                    f"(location={hl.get('location')!r}) present in original, missing from "
                    f"output (INTERNAL_HYPERLINK_LOSS)"
                )
            elif out_hl.get(R_ID_ATTR) is not None:
                violations.append(
                    f"sheet '{name}' ref={ref}: relationship-free hyperlink survived but "
                    f"unexpectedly gained an r:id it didn't originally have"
                )
            elif out_hl.get("location") != hl.get("location"):
                violations.append(
                    f"sheet '{name}' ref={ref}: location changed from "
                    f"{hl.get('location')!r} to {out_hl.get('location')!r}"
                )

    verdict = "INTERNAL_HYPERLINK_LOSS" if violations else "CLEAN"
    return {"violations": violations, "internal_hyperlink_verdict": verdict}


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

        # --- Case H: SOURCE_REFERENCE_LOSS (0.10.0-A). A second, dedicated fixture with a
        # worksheet-level .rels carrying all four r:id-mapped relationship types
        # (table/drawing/hyperlink/vmlDrawing) PLUS a fifth, unmapped comments
        # relationship (rId5) -- deliberately included so this test also locks in that
        # comments is correctly out of scope (see _WORKSHEET_RID_ELEMENT_XPATHS' own
        # comment for why). Built separately from `orig_path` above rather than extending
        # it, so this test can't accidentally interact with the formula/VBA/theme cases.
        rr_path = os.path.join(tmp, "relref_orig.xlsm")
        SHEET1_RELREF_XML = (
            f'<?xml version="1.0"?><worksheet xmlns="{SML_NS[1:-1]}" xmlns:r="{R_NS}">'
            '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
            '<hyperlinks><hyperlink ref="A1" r:id="rId3"/></hyperlinks>'
            '<drawing r:id="rId2"/>'
            '<legacyDrawing r:id="rId4"/>'
            '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'
            "</worksheet>"
        ).encode()
        with zipfile.ZipFile(rr_path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                f'<?xml version="1.0"?><Types xmlns="{CT_NS[1:-1]}">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "xl/workbook.xml",
                f'<?xml version="1.0"?><workbook xmlns="{SML_NS[1:-1]}" xmlns:r="{R_NS}">'
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
            )
            zf.writestr(
                "xl/_rels/workbook.xml.rels",
                f'<?xml version="1.0"?><Relationships xmlns="{REL_NS[1:-1]}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                "</Relationships>",
            )
            zf.writestr("xl/worksheets/sheet1.xml", SHEET1_RELREF_XML)
            zf.writestr(
                "xl/worksheets/_rels/sheet1.xml.rels",
                f'<?xml version="1.0"?><Relationships xmlns="{REL_NS[1:-1]}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>'
                '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
                '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>'
                "</Relationships>",
            )
            zf.writestr("xl/tables/table1.xml", f'<?xml version="1.0"?><table xmlns="{SML_NS[1:-1]}" id="1" name="Table1" displayName="Table1" ref="A1:A1"/>')
            zf.writestr("xl/drawings/drawing1.xml", '<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>')
            zf.writestr("xl/drawings/vmlDrawing1.vml", '<xml/>')
            zf.writestr("xl/comments1.xml", f'<?xml version="1.0"?><comments xmlns="{SML_NS[1:-1]}"><authors/><commentList/></comments>')

        # Clean passthrough (only workbook.xml "edited") -- must report CLEAN, including
        # for the unmapped comments relationship (rId5), which nothing in sheet1.xml ever
        # references by design (see _WORKSHEET_RID_ELEMENT_XPATHS).
        rr_clean_path = os.path.join(tmp, "relref_clean.xlsm")
        shutil.copyfile(rr_path, rr_clean_path)
        result = check_source_references(rr_path, rr_clean_path)
        assert result["source_reference_verdict"] == "CLEAN", result

        # Four independent mutations, one per mapped relationship type: strip ONLY the
        # worksheet-XML-side r:id reference, leave the .rels file and target part
        # untouched -- exactly the shape of the real bug found against fixture3
        # (tableParts stripped from a regenerated sheet1.xml while
        # _rels/sheet1.xml.rels and tables/table1.xml both survive byte-identical).
        _RELREF_MUTATIONS = {
            "table": (b'<tableParts count="1"><tablePart r:id="rId1"/></tableParts>', b""),
            "drawing": (b'<drawing r:id="rId2"/>', b""),
            "hyperlink": (b'<hyperlinks><hyperlink ref="A1" r:id="rId3"/></hyperlinks>', b""),
            "vmlDrawing": (b'<legacyDrawing r:id="rId4"/>', b""),
        }
        for label, (needle, replacement) in _RELREF_MUTATIONS.items():
            assert needle in SHEET1_RELREF_XML, f"self-test fixture bug: {label} needle not found"
            mutated_path = os.path.join(tmp, f"relref_missing_{label}.xlsm")
            with zipfile.ZipFile(rr_path) as src, zipfile.ZipFile(mutated_path, "w") as dst:
                for item in src.infolist():
                    data = src.read(item.filename)
                    if item.filename == "xl/worksheets/sheet1.xml":
                        data = data.replace(needle, replacement)
                    dst.writestr(item, data)
            result = check_source_references(rr_path, mutated_path)
            assert result["source_reference_verdict"] == "SOURCE_REFERENCE_LOSS", (
                f"failed to detect missing {label} reference: {result}"
            )
            assert any(f"Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/{label}'" in v for v in result["violations"]), (
                label,
                result["violations"],
            )

        # An unexpected change to the worksheet-level .rels itself (elixcee's writer has
        # never been observed to do this) must be flagged, not silently treated as
        # out-of-scope -- regression guard for exactly the bug an earlier version of
        # check_source_references had (see that function's own docstring).
        rels_changed_path = os.path.join(tmp, "relref_rels_changed.xlsm")
        with zipfile.ZipFile(rr_path) as src, zipfile.ZipFile(rels_changed_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/worksheets/_rels/sheet1.xml.rels":
                    data = data.replace(b'Id="rId1"', b'Id="rId9"')
                dst.writestr(item, data)
        result = check_source_references(rr_path, rels_changed_path)
        assert result["source_reference_verdict"] == "SOURCE_REFERENCE_LOSS", (
            f"an unexpectedly-mutated worksheet .rels must not be silently treated as clean: {result}"
        )
        assert any("differs from the original" in v for v in result["violations"]), result["violations"]

        # Type confusion: table's rId1 is removed from <tableParts>, but the SAME string
        # "rId1" is then reused on <drawing> (a different relationship type, rId2's own
        # slot in this fixture, left untouched). A flat union of "any r:id referenced
        # anywhere" would see "rId1" in the drawing element and wrongly call the table
        # relationship satisfied. The type-aware check must still catch the table loss.
        confused_path = os.path.join(tmp, "relref_type_confusion.xlsm")
        with zipfile.ZipFile(rr_path) as src, zipfile.ZipFile(confused_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(
                        b'<tableParts count="1"><tablePart r:id="rId1"/></tableParts>', b""
                    ).replace(b'<drawing r:id="rId2"/>', b'<drawing r:id="rId1"/>')
                dst.writestr(item, data)
        result = check_source_references(rr_path, confused_path)
        assert result["source_reference_verdict"] == "SOURCE_REFERENCE_LOSS", (
            f"type confusion must not mask a real table-relationship loss: {result}"
        )
        assert any(
            "Id=rId1" in v and "Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'" in v
            for v in result["violations"]
        ), result["violations"]

        # --- Case I: INLINE_ELEMENT_LOSS (0.10.0-B). orig_path/clean_path (Case A) have no
        # <sheetViews> at all, so re-use them first to confirm "nothing to lose" reports
        # CLEAN, not a false positive -- then build a dedicated fixture that DOES carry
        # <sheetViews> (freeze pane + selection, mirroring fixture7_freeze_pane.xlsm's real
        # shape) to confirm both a clean copy and a stripped one are judged correctly.
        result = check_inline_worksheet_elements(orig_path, clean_path)
        assert result["inline_element_verdict"] == "CLEAN", result
        assert result["violations"] == [], result

        sv_path = os.path.join(tmp, "sheetviews_orig.xlsm")
        # XSD order (CT_Worksheet, design doc §8): sheetPr, dimension, sheetViews,
        # sheetFormatPr, cols, sheetData, ..., mergeCells, phoneticPr,
        # conditionalFormatting, dataValidations, ... -- this fixture carries all 5
        # slice-2 targets (sheetViews from slice 1 plus sheetPr/sheetFormatPr/
        # phoneticPr/dataValidations from slice 2) in that real relative order.
        SHEET1_SHEETVIEWS_XML = (
            f'<?xml version="1.0"?><worksheet xmlns="{SML_NS[1:-1]}" xmlns:r="{R_NS}">'
            '<sheetPr codeName="Sheet1"/>'
            '<sheetViews><sheetView tabSelected="1" workbookViewId="0">'
            '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>'
            '<selection pane="bottomRight" activeCell="B2" sqref="B2"/>'
            "</sheetView></sheetViews>"
            '<sheetFormatPr baseColWidth="10" defaultRowHeight="20"/>'
            '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>'
            '<phoneticPr fontId="1"/>'
            '<dataValidations count="1"><dataValidation type="list" allowBlank="1" '
            'sqref="E1"><formula1>"Yes,No"</formula1></dataValidation></dataValidations>'
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
            "</worksheet>"
        ).encode()
        with zipfile.ZipFile(sv_path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                f'<?xml version="1.0"?><Types xmlns="{CT_NS[1:-1]}">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "xl/workbook.xml",
                f'<?xml version="1.0"?><workbook xmlns="{SML_NS[1:-1]}" xmlns:r="{R_NS}">'
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
            )
            zf.writestr(
                "xl/_rels/workbook.xml.rels",
                f'<?xml version="1.0"?><Relationships xmlns="{REL_NS[1:-1]}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                "</Relationships>",
            )
            zf.writestr("xl/worksheets/sheet1.xml", SHEET1_SHEETVIEWS_XML)

        sv_clean_path = os.path.join(tmp, "sheetviews_clean.xlsm")
        shutil.copyfile(sv_path, sv_clean_path)
        result = check_inline_worksheet_elements(sv_path, sv_clean_path)
        assert result["inline_element_verdict"] == "CLEAN", result

        sv_stripped_path = os.path.join(tmp, "sheetviews_stripped.xlsm")
        with zipfile.ZipFile(sv_path) as src, zipfile.ZipFile(sv_stripped_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    assert b"<sheetViews>" in data, "self-test fixture bug: sheetViews needle not found"
                    data = re.sub(rb"<sheetViews>.*?</sheetViews>", b"", data, flags=re.DOTALL)
                dst.writestr(item, data)
        result = check_inline_worksheet_elements(sv_path, sv_stripped_path)
        assert result["inline_element_verdict"] == "INLINE_ELEMENT_LOSS", result
        assert any("sheetViews" in v for v in result["violations"]), result["violations"]

        # --- Case J: INLINE_ELEMENT_LOSS, slice 2/3 (0.10.0-B). One independent mutation
        # per element -- strip only that element from sv_path's sheet1.xml, leave the
        # others intact, confirm each is individually detected (not just "something is
        # wrong somewhere"). Mirrors Case H's one-mutation-per-relationship-type shape.
        _SLICE2_MUTATIONS = {
            "sheetPr": (b'<sheetPr codeName="Sheet1"/>', b""),
            "sheetFormatPr": (b'<sheetFormatPr baseColWidth="10" defaultRowHeight="20"/>', b""),
            "phoneticPr": (b'<phoneticPr fontId="1"/>', b""),
            "dataValidations": (
                b'<dataValidations count="1"><dataValidation type="list" allowBlank="1" '
                b'sqref="E1"><formula1>"Yes,No"</formula1></dataValidation></dataValidations>',
                b"",
            ),
            "pageMargins": (
                b'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
                b"",
            ),
        }
        for label, (needle, replacement) in _SLICE2_MUTATIONS.items():
            assert needle in SHEET1_SHEETVIEWS_XML, f"self-test fixture bug: {label} needle not found"
            mutated_path = os.path.join(tmp, f"sheetviews_missing_{label}.xlsm")
            with zipfile.ZipFile(sv_path) as src, zipfile.ZipFile(mutated_path, "w") as dst:
                for item in src.infolist():
                    data = src.read(item.filename)
                    if item.filename == "xl/worksheets/sheet1.xml":
                        data = data.replace(needle, replacement)
                    dst.writestr(item, data)
            result = check_inline_worksheet_elements(sv_path, mutated_path)
            assert result["inline_element_verdict"] == "INLINE_ELEMENT_LOSS", (
                f"failed to detect missing {label}: {result}"
            )
            assert any(label in v for v in result["violations"]), (label, result["violations"])
            # The other 4 elements must NOT be flagged -- confirms per-element
            # granularity, not "the whole check fired because something changed".
            for other in _INLINE_WORKSHEET_ELEMENTS:
                if other != label:
                    assert not any(f"<{other}>" in v for v in result["violations"]), (
                        f"stripping {label} must not also flag unrelated {other}: {result['violations']}"
                    )

        # --- Case K: INTERNAL_HYPERLINK_LOSS (0.10.0-B, B4). Mixed <hyperlinks> container
        # -- one r:id-bearing (external) hyperlink, one location-only (internal) hyperlink
        # -- since no real fixture has both together yet (fixture4=all-r:id,
        # fixture6=all-location; see design doc's B4 entry, this mixed shape is synthetic,
        # generalized from the two real endpoints plus CT_Hyperlinks' XSD).
        hl_path = os.path.join(tmp, "hyperlinks_orig.xlsm")
        SHEET1_HYPERLINKS_XML = (
            f'<?xml version="1.0"?><worksheet xmlns="{SML_NS[1:-1]}" xmlns:r="{R_NS}">'
            '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>'
            '<hyperlinks>'
            '<hyperlink ref="A1" r:id="rId2"/>'
            '<hyperlink ref="B1" location="Sheet2!A1" display="Sheet2!A1"/>'
            "</hyperlinks>"
            "</worksheet>"
        ).encode()
        with zipfile.ZipFile(hl_path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                f'<?xml version="1.0"?><Types xmlns="{CT_NS[1:-1]}">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                "</Types>",
            )
            zf.writestr(
                "xl/workbook.xml",
                f'<?xml version="1.0"?><workbook xmlns="{SML_NS[1:-1]}" xmlns:r="{R_NS}">'
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>',
            )
            zf.writestr(
                "xl/_rels/workbook.xml.rels",
                f'<?xml version="1.0"?><Relationships xmlns="{REL_NS[1:-1]}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                "</Relationships>",
            )
            zf.writestr("xl/worksheets/sheet1.xml", SHEET1_HYPERLINKS_XML)
            zf.writestr(
                "xl/worksheets/_rels/sheet1.xml.rels",
                f'<?xml version="1.0"?><Relationships xmlns="{REL_NS[1:-1]}">'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>'
                "</Relationships>",
            )

        # Clean passthrough -- must report CLEAN (the r:id hyperlink is out of this
        # check's scope entirely; the location-only one survives unchanged).
        hl_clean_path = os.path.join(tmp, "hyperlinks_clean.xlsm")
        shutil.copyfile(hl_path, hl_clean_path)
        result = check_internal_hyperlinks(hl_path, hl_clean_path)
        assert result["internal_hyperlink_verdict"] == "CLEAN", result

        # Strip ONLY the location-only hyperlink (B1), leaving the r:id one (A1) and the
        # <hyperlinks> container itself intact -- must be caught, and the r:id hyperlink's
        # disappearance-from-scope must not mask it or vice versa.
        hl_stripped_path = os.path.join(tmp, "hyperlinks_stripped.xlsm")
        with zipfile.ZipFile(hl_path) as src, zipfile.ZipFile(hl_stripped_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    needle = b'<hyperlink ref="B1" location="Sheet2!A1" display="Sheet2!A1"/>'
                    assert needle in data, "self-test fixture bug: location-only hyperlink needle not found"
                    data = data.replace(needle, b"")
                dst.writestr(item, data)
        result = check_internal_hyperlinks(hl_path, hl_stripped_path)
        assert result["internal_hyperlink_verdict"] == "INTERNAL_HYPERLINK_LOSS", result
        assert any("ref=B1" in v for v in result["violations"]), result["violations"]
        assert not any("ref=A1" in v for v in result["violations"]), (
            f"the r:id hyperlink (A1) is out of this check's scope and must never be "
            f"flagged: {result['violations']}"
        )

        # An output <hyperlinks> with zero <hyperlink> children (both stripped, container
        # left as an empty shell) must be flagged as itself invalid -- CT_Hyperlinks'
        # minOccurs=1 on <hyperlink>, not merely "the location-only one is missing".
        hl_empty_container_path = os.path.join(tmp, "hyperlinks_empty_container.xlsm")
        with zipfile.ZipFile(hl_path) as src, zipfile.ZipFile(hl_empty_container_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    data = data.replace(
                        b'<hyperlinks><hyperlink ref="A1" r:id="rId2"/>'
                        b'<hyperlink ref="B1" location="Sheet2!A1" display="Sheet2!A1"/></hyperlinks>',
                        b"<hyperlinks></hyperlinks>",
                    )
                dst.writestr(item, data)
        result = check_internal_hyperlinks(hl_path, hl_empty_container_path)
        assert result["internal_hyperlink_verdict"] == "INTERNAL_HYPERLINK_LOSS", result
        assert any("zero <hyperlink> children" in v for v in result["violations"]), result["violations"]

        # --- Case L: WORKBOOK_ELEMENT_LOSS (0.10.0-C, C1+C2). workbookPr/bookViews/
        # calcPr/extLst as direct children of the root <workbook>, mirroring fixture4's
        # real relative order (workbookPr, bookViews before <sheets>; calcPr/extLst
        # after -- design doc §8's CT_Workbook sequence). definedNames deliberately
        # absent from this fixture -- that's C3, not C1/C2's scope. The <workbookView>
        # here carries no activeTab/firstSheet, matching every real fixture (see
        # check_workbook_elements()'s own docstring for why that's C2's basis for a
        # plain verbatim copy rather than gated logic).
        WORKBOOK_C1_XML = (
            f'<?xml version="1.0"?><workbook xmlns="{SML_NS[1:-1]}" xmlns:r="{R_NS}">'
            '<workbookPr codeName="ThisWorkbook" defaultThemeVersion="202300"/>'
            '<bookViews><workbookView xWindow="540" yWindow="660" windowWidth="28300" '
            'windowHeight="17160"/></bookViews>'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
            '<calcPr calcId="181029"/>'
            '<extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}">'
            '<x15:workbookPr xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" '
            'chartTrackingRefBase="1"/></ext></extLst>'
            "</workbook>"
        ).encode()
        wb_path = os.path.join(tmp, "workbook_elements_orig.xlsm")
        with zipfile.ZipFile(wb_path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                f'<?xml version="1.0"?><Types xmlns="{CT_NS[1:-1]}">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                "</Types>",
            )
            zf.writestr("xl/workbook.xml", WORKBOOK_C1_XML)
            zf.writestr(
                "xl/_rels/workbook.xml.rels",
                f'<?xml version="1.0"?><Relationships xmlns="{REL_NS[1:-1]}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                "</Relationships>",
            )
            zf.writestr(
                "xl/worksheets/sheet1.xml",
                f'<?xml version="1.0"?><worksheet xmlns="{SML_NS[1:-1]}">'
                '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
            )

        wb_clean_path = os.path.join(tmp, "workbook_elements_clean.xlsm")
        shutil.copyfile(wb_path, wb_clean_path)
        result = check_workbook_elements(wb_path, wb_clean_path)
        assert result["workbook_element_verdict"] == "CLEAN", result
        assert result["violations"] == [], result

        _WORKBOOK_ELEMENT_MUTATIONS = {
            "workbookPr": (b'<workbookPr codeName="ThisWorkbook" defaultThemeVersion="202300"/>', b""),
            "bookViews": (
                b'<bookViews><workbookView xWindow="540" yWindow="660" windowWidth="28300" '
                b'windowHeight="17160"/></bookViews>',
                b"",
            ),
            "calcPr": (b'<calcPr calcId="181029"/>', b""),
            "extLst": (
                b'<extLst><ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}">'
                b'<x15:workbookPr xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" '
                b'chartTrackingRefBase="1"/></ext></extLst>',
                b"",
            ),
        }
        for label, (needle, replacement) in _WORKBOOK_ELEMENT_MUTATIONS.items():
            assert needle in WORKBOOK_C1_XML, f"self-test fixture bug: {label} needle not found"
            mutated_path = os.path.join(tmp, f"workbook_elements_missing_{label}.xlsm")
            with zipfile.ZipFile(wb_path) as src, zipfile.ZipFile(mutated_path, "w") as dst:
                for item in src.infolist():
                    data = src.read(item.filename)
                    if item.filename == "xl/workbook.xml":
                        data = data.replace(needle, replacement)
                    dst.writestr(item, data)
            result = check_workbook_elements(wb_path, mutated_path)
            assert result["workbook_element_verdict"] == "WORKBOOK_ELEMENT_LOSS", (
                f"failed to detect missing {label}: {result}"
            )
            assert any(label in v for v in result["violations"]), (label, result["violations"])
            for other in _WORKBOOK_ELEMENTS:
                if other != label:
                    assert not any(f"<{other}>" in v for v in result["violations"]), (
                        f"stripping {label} must not also flag unrelated {other}: {result['violations']}"
                    )

        # --- Case M: WORKBOOK_ELEMENT_LOSS via check_defined_names() (0.10.0-C, C3).
        # Two sheets so a delete is actually representable; one workbook-scoped name
        # (no localSheetId) and one sheet-scoped name (localSheetId="1", pointing at
        # the second sheet) mirroring fixture4 (plain) and fixture5 (_xlnm.Print_Area)
        # respectively.
        dn_path = os.path.join(tmp, "defined_names_orig.xlsm")
        DEFINED_NAMES_XML = (
            '<definedNames>'
            '<definedName name="test" comment="test desu!!!">Sheet1!$F$5</definedName>'
            '<definedName name="_xlnm.Print_Area" localSheetId="1">Sheet2!$E$3</definedName>'
            "</definedNames>"
        )
        WORKBOOK_DN_XML = (
            f'<?xml version="1.0"?><workbook xmlns="{SML_NS[1:-1]}" xmlns:r="{R_NS}">'
            '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/>'
            '<sheet name="Sheet2" sheetId="2" r:id="rId2"/></sheets>'
            + DEFINED_NAMES_XML +
            "</workbook>"
        ).encode()
        with zipfile.ZipFile(dn_path, "w") as zf:
            zf.writestr(
                "[Content_Types].xml",
                f'<?xml version="1.0"?><Types xmlns="{CT_NS[1:-1]}">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                "</Types>",
            )
            zf.writestr("xl/workbook.xml", WORKBOOK_DN_XML)
            zf.writestr(
                "xl/_rels/workbook.xml.rels",
                f'<?xml version="1.0"?><Relationships xmlns="{REL_NS[1:-1]}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
                "</Relationships>",
            )
            for n in (1, 2):
                zf.writestr(
                    f"xl/worksheets/sheet{n}.xml",
                    f'<?xml version="1.0"?><worksheet xmlns="{SML_NS[1:-1]}">'
                    '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
                )

        # (1) No delete: clean passthrough must report CLEAN.
        dn_clean_path = os.path.join(tmp, "defined_names_clean.xlsm")
        shutil.copyfile(dn_path, dn_clean_path)
        result = check_defined_names(dn_path, dn_clean_path)
        assert result["defined_names_verdict"] == "CLEAN", result

        # (2) No delete, but definedNames dropped anyway -- must be flagged.
        dn_dropped_path = os.path.join(tmp, "defined_names_dropped.xlsm")
        with zipfile.ZipFile(dn_path) as src, zipfile.ZipFile(dn_dropped_path, "w") as dst:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "xl/workbook.xml":
                    assert DEFINED_NAMES_XML.encode() in data
                    data = data.replace(DEFINED_NAMES_XML.encode(), b"")
                dst.writestr(item, data)
        result = check_defined_names(dn_path, dn_dropped_path)
        assert result["defined_names_verdict"] == "WORKBOOK_ELEMENT_LOSS", result
        assert any("every original sheet survived" in v for v in result["violations"]), result

        # (3) A sheet is deleted AND definedNames is correctly dropped too -- this is
        # the writer's required behavior, must report CLEAN (not a loss).
        dn_deleted_correct_path = os.path.join(tmp, "defined_names_deleted_correct.xlsm")
        with zipfile.ZipFile(dn_path) as src, zipfile.ZipFile(dn_deleted_correct_path, "w") as dst:
            for item in src.infolist():
                if item.filename == "xl/worksheets/sheet2.xml":
                    continue  # simulate Sheet2 having been deleted
                data = src.read(item.filename)
                if item.filename == "xl/workbook.xml":
                    data = (
                        data.replace(DEFINED_NAMES_XML.encode(), b"").replace(
                            b'<sheet name="Sheet2" sheetId="2" r:id="rId2"/>', b""
                        )
                    )
                elif item.filename == "xl/_rels/workbook.xml.rels":
                    data = data.replace(
                        b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>',
                        b"",
                    )
                dst.writestr(item, data)
        result = check_defined_names(dn_path, dn_deleted_correct_path)
        assert result["defined_names_verdict"] == "CLEAN", result

        # (4) A sheet is deleted but definedNames was left in place anyway -- a stale,
        # dangerous-looking-valid reference is worse than a plain loss, must be flagged.
        dn_deleted_wrong_path = os.path.join(tmp, "defined_names_deleted_wrong.xlsm")
        with zipfile.ZipFile(dn_path) as src, zipfile.ZipFile(dn_deleted_wrong_path, "w") as dst:
            for item in src.infolist():
                if item.filename == "xl/worksheets/sheet2.xml":
                    continue
                data = src.read(item.filename)
                if item.filename == "xl/workbook.xml":
                    data = data.replace(b'<sheet name="Sheet2" sheetId="2" r:id="rId2"/>', b"")
                elif item.filename == "xl/_rels/workbook.xml.rels":
                    data = data.replace(
                        b'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>',
                        b"",
                    )
                dst.writestr(item, data)
        result = check_defined_names(dn_path, dn_deleted_wrong_path)
        assert result["defined_names_verdict"] == "WORKBOOK_ELEMENT_LOSS", result
        assert any("must be dropped entirely" in v for v in result["violations"]), result

    print(
        "self_test: OK (clean pass-through clean; truncated VBA, broken rels, dangling "
        "target, stripped-formula, orphaned part, all 4 SOURCE_REFERENCE_LOSS shapes, "
        "unexpected .rels mutation, cross-type rId confusion, INLINE_ELEMENT_LOSS "
        "(sheetViews + 5 slice-2/3 elements, independently), INTERNAL_HYPERLINK_LOSS "
        "(mixed container: location-only loss detected, r:id sibling never falsely "
        "flagged, empty-container invalidity detected), and WORKBOOK_ELEMENT_LOSS "
        "(workbookPr/bookViews/calcPr/extLst, independently) and defined-names loss "
        "(verbatim required when no sheet deleted, must be dropped entirely when one "
        "was, both directions checked) all caught; comments correctly "
        "out of SOURCE_REFERENCE_LOSS scope)"
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
    source_references = check_source_references(sys.argv[1], sys.argv[2])
    inline_elements = check_inline_worksheet_elements(sys.argv[1], sys.argv[2])
    internal_hyperlinks = check_internal_hyperlinks(sys.argv[1], sys.argv[2])
    workbook_elements = check_workbook_elements(sys.argv[1], sys.argv[2])
    defined_names = check_defined_names(sys.argv[1], sys.argv[2])
    print(json.dumps({
        "structural": structural,
        "formulas": formulas,
        "source_references": source_references,
        "inline_elements": inline_elements,
        "internal_hyperlinks": internal_hyperlinks,
        "workbook_elements": workbook_elements,
        "defined_names": defined_names,
    }, indent=2))
    ok = (
        structural["structural_verdict"] == "STRUCTURALLY_CLEAN"
        and formulas["formula_verdict"] == "CLEAN"
        and source_references["source_reference_verdict"] == "CLEAN"
        and inline_elements["inline_element_verdict"] == "CLEAN"
        and internal_hyperlinks["internal_hyperlink_verdict"] == "CLEAN"
        and workbook_elements["workbook_element_verdict"] == "CLEAN"
        and defined_names["defined_names_verdict"] == "CLEAN"
    )
    sys.exit(0 if ok else 1)
