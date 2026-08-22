/// Safe round-trip: unknown-OOXML-part passthrough + `xl/vbaProject.bin`
/// preservation (see `docs/xlsx-architecture.md`'s "regenerate vs.
/// preserve-and-merge" section and `/Users/k_tanabe/.claude/plans/wise-waddling-fern.md`).
///
/// No real Excel-authored `.xlsm` exists in this repo yet, so the fixture is
/// hand-built in-test via `zip::write::ZipWriter` (already a normal
/// dependency, used identically by `save_xlsx_impl` itself) rather than a
/// committed binary blob or a SheetJS-generated file (SheetJS can't write
/// macro-enabled workbooks at all). See `tests/fixtures/xlsm_roundtrip/README.md`
/// for where a real file slots in later.
use elixcee::{parser, reader, save_workbook, vm::Vm};
use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use zip::write::{SimpleFileOptions, ZipWriter};
use zip::ZipArchive;

const CONTENT_TYPES: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n",
    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n",
    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\n",
    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.ms-excel.sheet.macroEnabled.main+xml\"/>\n",
    "<Override PartName=\"/xl/worksheets/sheet3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n",
    "<Override PartName=\"/xl/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>\n",
    "<Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>\n",
    "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n",
    "</Types>\n",
);

const CONTENT_TYPES_NO_VBA: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n",
    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n",
    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\n",
    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n",
    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n",
    "<Override PartName=\"/xl/tables/table1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>\n",
    "</Types>\n",
);

const ROOT_RELS: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n",
    "</Relationships>\n",
);

/// Two `<cellXfs>` entries: index 0 (default) and index 1 (bold font + red
/// fill, structurally distinct from index 0) -- used to prove a cell's
/// original `s="N"` index, and the style DEFINITION it points at, both
/// survive a save unchanged (Milestone: safe round-trip, style-index
/// preservation).
const STYLES_XML: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
    "<fonts count=\"2\"><font/><font><b/></font></fonts>\n",
    "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>",
    "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFF0000\"/></patternFill></fill></fills>\n",
    "<borders count=\"1\"><border/></borders>\n",
    "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n",
    "<cellXfs count=\"2\">\n",
    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>\n",
    "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"1\" borderId=\"0\" applyFont=\"1\" applyFill=\"1\"/>\n",
    "</cellXfs>\n",
    "</styleSheet>\n",
);

const TABLE_XML: &str = concat!(
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
    "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
    "id=\"1\" name=\"Table1\" displayName=\"Table1\" ref=\"A1:A1\">\n",
    "<autoFilter ref=\"A1:A1\"/>\n",
    "<tableColumns count=\"1\"><tableColumn id=\"1\" name=\"Col1\"/></tableColumns>\n",
    "</table>\n",
);

/// Deterministic non-trivial "VBA project" stand-in: OLE/CFB magic bytes
/// followed by a fixed non-zero fill pattern, so a writer bug that zeroes or
/// truncates the part is actually caught by byte equality (an all-zero blob
/// would not catch either).
fn vba_project_bytes() -> Vec<u8> {
    let mut b = vec![0xD0u8, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    for i in 0..4096u32 {
        b.push((i % 251) as u8);
    }
    b
}

fn workbook_xml() -> String {
    concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n<sheet name=\"Sheet1\" sheetId=\"3\" r:id=\"rId1\"/>\n</sheets>\n",
        "</workbook>\n",
    )
    .to_string()
}

/// Two cells, both styled (`s="1"`, see `STYLES_XML`): `A1 = 1` (which test
/// macros edit -- proves an edited cell's original style survives) and
/// `B1 = 2` (which no macro touches -- proves an untouched cell's original
/// style also survives the sheet's full regeneration). At whatever part
/// name is passed — deliberately `xl/worksheets/sheet3.xml` in the .xlsm
/// fixture (not `sheet1.xml`), simulating a book that once had 3 sheets
/// with two deleted, so the passthrough exclusion logic is proven to be
/// pattern-based, not keyed off this writer's own sequential naming.
fn sheet_xml() -> String {
    concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<sheetData>\n<row r=\"1\">",
        "<c r=\"A1\" s=\"1\"><v>1</v></c>",
        "<c r=\"B1\" s=\"1\"><v>2</v></c>",
        "</row>\n</sheetData>\n",
        "</worksheet>\n",
    )
    .to_string()
}

fn zip_add(zip: &mut ZipWriter<Cursor<Vec<u8>>>, name: &str, bytes: &[u8]) {
    let opts = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    zip.start_file(name, opts).unwrap();
    zip.write_all(bytes).unwrap();
}

/// Builds a minimal `.xlsm`-shaped fixture: one sheet (non-sequential part
/// name `sheet3.xml`) with two styled cells, a real VBA-project part, a
/// `xl/styles.xml` with a distinct non-default cellXf, and a
/// `xl/tables/table1.xml` stub standing in for "some other part elixcee
/// doesn't parse." Returns the fixture bytes plus the vbaProject, table, and
/// styles bytes for later comparison.
fn build_fixture_xlsm() -> (Vec<u8>, Vec<u8>, Vec<u8>, Vec<u8>) {
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(&mut zip, "[Content_Types].xml", CONTENT_TYPES.as_bytes());
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", workbook_xml().as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/>\n",
            "<Relationship Id=\"rId2\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>\n",
            "</Relationships>\n",
        )
        .as_bytes(),
    );
    zip_add(&mut zip, "xl/worksheets/sheet3.xml", sheet_xml().as_bytes());
    let vba_bytes = vba_project_bytes();
    zip_add(&mut zip, "xl/vbaProject.bin", &vba_bytes);
    zip_add(&mut zip, "xl/tables/table1.xml", TABLE_XML.as_bytes());
    zip_add(&mut zip, "xl/styles.xml", STYLES_XML.as_bytes());
    let data = zip.finish().unwrap().into_inner();
    (data, vba_bytes, TABLE_XML.as_bytes().to_vec(), STYLES_XML.as_bytes().to_vec())
}

/// Same shape, `.xlsx` (no VBA project), one unknown part.
fn build_fixture_xlsx() -> Vec<u8> {
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(&mut zip, "[Content_Types].xml", CONTENT_TYPES_NO_VBA.as_bytes());
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(
        &mut zip,
        "xl/workbook.xml",
        concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
            "<sheets>\n<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n</sheets>\n</workbook>\n",
        )
        .as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
            "</Relationships>\n",
        )
        .as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet1.xml",
        concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
            "<sheetData>\n<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>\n</sheetData>\n</worksheet>\n",
        )
        .as_bytes(),
    );
    zip_add(&mut zip, "xl/tables/table1.xml", TABLE_XML.as_bytes());
    zip.finish().unwrap().into_inner()
}

fn read_all_zip_entries(bytes: &[u8]) -> HashMap<String, Vec<u8>> {
    let mut archive = ZipArchive::new(Cursor::new(bytes.to_vec())).unwrap();
    let mut out = HashMap::new();
    for i in 0..archive.len() {
        let mut entry = archive.by_index(i).unwrap();
        if entry.is_dir() {
            continue;
        }
        let name = entry.name().to_string();
        let mut buf = Vec::new();
        entry.read_to_end(&mut buf).unwrap();
        out.insert(name, buf);
    }
    out
}

fn is_writer_owned(name: &str) -> bool {
    matches!(
        name,
        "[Content_Types].xml"
            | "_rels/.rels"
            | "xl/workbook.xml"
            | "xl/_rels/workbook.xml.rels"
            | "xl/sharedStrings.xml"
            | "xl/styles.xml"
    ) || (name.starts_with("xl/worksheets/")
        && name.ends_with(".xml")
        && !name["xl/worksheets/".len()..].contains('/'))
}

fn extract_attr(tag: &str, attr: &str) -> Option<String> {
    let needle = format!("{}=\"", attr);
    let start = tag.find(&needle)? + needle.len();
    let end = tag[start..].find('"')? + start;
    Some(tag[start..end].to_string())
}

/// Resolves `part_name` (no leading `/`) via `[Content_Types].xml`'s own
/// Default/Override declarations (Override wins) — a small independent
/// string-scanning reimplementation, deliberately NOT calling into
/// `reader::content_type_decls` (which is `pub(crate)`-only and not visible
/// from an integration test anyway), so this assertion exercises the
/// output's own self-consistency rather than trusting the writer's own
/// parsing logic.
fn resolve_content_type(content_types_xml: &str, part_name: &str) -> Option<String> {
    let full = format!("/{}", part_name);
    let ext = part_name.rsplit('.').next().unwrap_or("");
    let mut override_ct = None;
    let mut default_ct = None;
    for (tag_start, _) in content_types_xml.match_indices('<') {
        let rest = &content_types_xml[tag_start..];
        let tag_end = rest.find('>').map(|p| p + 1).unwrap_or(rest.len());
        let tag = &rest[..tag_end];
        if tag.starts_with("<Override ") {
            if extract_attr(tag, "PartName").as_deref() == Some(full.as_str()) {
                override_ct = extract_attr(tag, "ContentType");
            }
        } else if tag.starts_with("<Default ") && extract_attr(tag, "Extension").as_deref() == Some(ext) {
            default_ct = extract_attr(tag, "ContentType");
        }
    }
    override_ct.or(default_ct)
}

fn tmp_path(name: &str) -> String {
    std::env::temp_dir()
        .join(format!("elixcee_test_xlsx_roundtrip_{}_{}", std::process::id(), name))
        .to_string_lossy()
        .to_string()
}

#[test]
fn xlsm_roundtrip_preserves_vba_project_and_declares_macro_enabled_content_types() {
    let (fixture_bytes, vba_bytes, table_bytes, styles_bytes) = build_fixture_xlsm();
    let source_path = tmp_path("source.xlsm");
    let output_path = tmp_path("output.xlsm");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path).expect("fixture should load");
    // A1: edit an already-styled cell's value. C1: write a brand-new cell
    // (empty in the source, no original style to inherit).
    let prog = parser::parse(
        "Sub EditCell()\n    Cells(1, 1).Value = 999\n    Cells(1, 3).Value = 5\nEnd Sub\n",
    )
    .unwrap();
    vm.run_sub(&prog, "EditCell").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save should succeed");

    // (i) edited cell round-trips
    let sheets = reader::read_workbook(&output_path).expect("output should be readable");
    let sheet = sheets.iter().find(|s| s.name == "sheet1").expect("sheet1 present");
    match sheet.cells.get(&(1, 1)) {
        Some(reader::SheetCell::Integer(999)) => {}
        other => panic!("expected A1 == 999, got {:?}", other.map(|_| "non-matching cell")),
    }

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let fixture_entries = read_all_zip_entries(&fixture_bytes);

    // (ii) xl/vbaProject.bin byte-identical
    assert_eq!(output_entries.get("xl/vbaProject.bin"), Some(&vba_bytes), "vbaProject.bin must survive byte-identical");

    // (iii) every non-writer-owned original part is byte-identical in the output
    for (name, bytes) in &fixture_entries {
        if is_writer_owned(name) {
            continue;
        }
        assert_eq!(output_entries.get(name), Some(bytes), "passthrough part {name} must be byte-identical");
    }
    assert_eq!(output_entries.get("xl/tables/table1.xml"), Some(&table_bytes));

    // (iv) stale non-sequential worksheet part must NOT survive
    assert!(
        !output_entries.contains_key("xl/worksheets/sheet3.xml"),
        "stale original worksheet part must not appear alongside the regenerated sheet1.xml"
    );
    assert!(output_entries.contains_key("xl/worksheets/sheet1.xml"));

    // (v) + (vi) content-types: macro-enabled root override, vbaProject resolvable,
    // and every part actually present in the output resolves via the output's
    // own [Content_Types].xml (full self-consistency, not a spot check).
    let ct_xml = String::from_utf8(output_entries["[Content_Types].xml"].clone()).unwrap();
    assert!(ct_xml.contains("macroEnabled.main+xml"), "workbook.xml must declare macro-enabled content type");
    assert_eq!(
        resolve_content_type(&ct_xml, "xl/vbaProject.bin").as_deref(),
        Some("application/vnd.ms-office.vbaProject")
    );
    for name in output_entries.keys() {
        if name == "[Content_Types].xml" {
            continue;
        }
        assert!(
            resolve_content_type(&ct_xml, name).is_some(),
            "output part {name} has no resolvable content type in [Content_Types].xml"
        );
    }

    // Style-index preservation (Milestone: safe round-trip, slice 2).
    let sheet_xml = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    let a1_tag = &sheet_xml[sheet_xml.find("<c r=\"A1\"").unwrap()..];
    let a1_tag = &a1_tag[..a1_tag.find('>').unwrap() + 1];
    assert!(a1_tag.contains("s=\"1\""), "edited cell A1 must keep its original style index: {a1_tag}");

    let b1_tag = &sheet_xml[sheet_xml.find("<c r=\"B1\"").unwrap()..];
    let b1_tag = &b1_tag[..b1_tag.find('>').unwrap() + 1];
    assert!(b1_tag.contains("s=\"1\""), "untouched cell B1 must keep its original style index: {b1_tag}");

    let c1_tag = &sheet_xml[sheet_xml.find("<c r=\"C1\"").unwrap()..];
    let c1_tag = &c1_tag[..c1_tag.find('>').unwrap() + 1];
    assert!(!c1_tag.contains("s=\"1\""), "a brand-new cell must not spuriously inherit style 1: {c1_tag}");

    assert_eq!(
        output_entries.get("xl/styles.xml"),
        Some(&styles_bytes),
        "xl/styles.xml must be byte-identical to the source, not the hardcoded minimal stylesheet"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

#[test]
fn xlsx_roundtrip_passes_through_unknown_parts_without_macro_content_type() {
    let fixture_bytes = build_fixture_xlsx();
    let source_path = tmp_path("source.xlsx");
    let output_path = tmp_path("output2.xlsx");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path).expect("fixture should load");
    let prog = parser::parse("Sub EditCell()\n    Cells(1, 1).Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditCell").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);

    assert_eq!(output_entries.get("xl/tables/table1.xml"), Some(&TABLE_XML.as_bytes().to_vec()));
    assert!(!output_entries.contains_key("xl/vbaProject.bin"));

    let ct_xml = String::from_utf8(output_entries["[Content_Types].xml"].clone()).unwrap();
    assert!(!ct_xml.contains("macroEnabled"), "a workbook that never had a VBA project must not declare macro-enabled content type");
    assert!(ct_xml.contains("spreadsheetml.sheet.main+xml"));

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

#[test]
fn xlsm_roundtrip_in_place_save_preserves_vba_project() {
    let (fixture_bytes, vba_bytes, _table_bytes, styles_bytes) = build_fixture_xlsm();
    let path = tmp_path("inplace.xlsm");
    std::fs::write(&path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&path).expect("fixture should load");
    let prog = parser::parse("Sub EditCell()\n    Cells(1, 1).Value = 7\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditCell").expect("macro should run");
    // Realistic `--file foo.xlsm --output foo.xlsm` usage: source == output.
    save_workbook(&vm, &path).expect("in-place save should succeed");

    let output_bytes = std::fs::read(&path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    assert_eq!(
        output_entries.get("xl/vbaProject.bin"),
        Some(&vba_bytes),
        "vbaProject.bin must survive an in-place overwrite byte-identical"
    );
    assert_eq!(
        output_entries.get("xl/styles.xml"),
        Some(&styles_bytes),
        "xl/styles.xml must also survive an in-place overwrite byte-identical"
    );
    let sheet_xml = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    let a1_tag = &sheet_xml[sheet_xml.find("<c r=\"A1\"").unwrap()..];
    let a1_tag = &a1_tag[..a1_tag.find('>').unwrap() + 1];
    assert!(a1_tag.contains("s=\"1\""), "edited cell A1 must keep its style index across an in-place overwrite: {a1_tag}");

    let sheets = reader::read_workbook(&path).unwrap();
    let sheet = sheets.iter().find(|s| s.name == "sheet1").unwrap();
    assert!(matches!(sheet.cells.get(&(1, 1)), Some(reader::SheetCell::Integer(7))));

    let _ = std::fs::remove_file(&path);
}
