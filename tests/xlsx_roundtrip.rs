/// Safe round-trip: unknown-OOXML-part passthrough + `xl/vbaProject.bin`
/// preservation (see `docs/xlsx-architecture.md`'s "regenerate vs.
/// preserve-and-merge" section).
///
/// These fixtures are hand-built in-test via `zip::write::ZipWriter` (already a
/// normal dependency, used identically by `save_xlsx_impl` itself) rather than a
/// committed binary blob or a SheetJS-generated file (SheetJS can't write
/// macro-enabled workbooks at all) -- kept as a synthetic, reviewable-as-source
/// complement to the real Excel-authored fixtures now under
/// `compat/oracle-excel-com/fixtures/pristine/` (see `ROADMAP.md`'s `0.9.0-A`).
use elixcee::{
    parser, reader, save_workbook,
    vm::{SheetState, Vm},
};
use std::collections::HashMap;
use std::io::{Cursor, Read, Write};
use zip::ZipArchive;
use zip::write::{SimpleFileOptions, ZipWriter};

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
/// pattern-based, not keyed off this writer's own sequential naming. Also
/// carries a merged range (`D1:E1`), a hidden column (F, entirely empty --
/// proves a hidden interval with no cell data still gets a `<col>`
/// declaration), and a hidden, cell-less row (2 -- proves a hidden row with
/// no cells still gets its own `<row hidden="1"/>` element, since
/// hidden-ness is a `<row>` attribute an absent element can't carry).
/// `G1 = A1+B1` (formula, cached value 3) is never touched by any test macro --
/// proves a formula elixcee doesn't itself write is still carried through a
/// save as a real `<f>` element, not flattened to a stale literal (found via
/// a real Excel-authored fixture during elixcee's first 0.9.0-A round-trip
/// run: `read_workbook`'s `WorkbookSheet` had no field to carry formula text
/// at all, so every save silently replaced every untouched formula cell with
/// its last cached value). Deliberately not C1 -- the flagship test already
/// uses C1 as "a brand-new cell the macro writes," and not row 2 or D1:E1 --
/// those are the hidden-row-with-no-cells and merge fixtures respectively.
fn sheet_xml() -> String {
    concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<cols><col min=\"6\" max=\"6\" hidden=\"1\"/></cols>\n",
        "<sheetData>\n<row r=\"1\">",
        "<c r=\"A1\" s=\"1\"><v>1</v></c>",
        "<c r=\"B1\" s=\"1\"><v>2</v></c>",
        "<c r=\"G1\"><f>A1+B1</f><v>3</v></c>",
        "</row>\n",
        "<row r=\"2\" hidden=\"1\"/>\n",
        "</sheetData>\n",
        "<mergeCells count=\"1\"><mergeCell ref=\"D1:E1\"/></mergeCells>\n",
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
    (
        data,
        vba_bytes,
        TABLE_XML.as_bytes().to_vec(),
        STYLES_XML.as_bytes().to_vec(),
    )
}

/// Same shape, `.xlsx` (no VBA project), one unknown part.
fn build_fixture_xlsx() -> Vec<u8> {
    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
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
        } else if tag.starts_with("<Default ")
            && extract_attr(tag, "Extension").as_deref() == Some(ext)
        {
            default_ct = extract_attr(tag, "ContentType");
        }
    }
    override_ct.or(default_ct)
}

fn tmp_path(name: &str) -> String {
    std::env::temp_dir()
        .join(format!(
            "elixcee_test_xlsx_roundtrip_{}_{}",
            std::process::id(),
            name
        ))
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
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
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
    let sheet = sheets
        .iter()
        .find(|s| s.name.to_lowercase() == "sheet1")
        .expect("sheet1 present");
    match sheet.cells.get(&(1, 1)) {
        Some(reader::SheetCell::Integer(999)) => {}
        other => panic!(
            "expected A1 == 999, got {:?}",
            other.map(|_| "non-matching cell")
        ),
    }

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let fixture_entries = read_all_zip_entries(&fixture_bytes);

    // (ii) xl/vbaProject.bin byte-identical
    assert_eq!(
        output_entries.get("xl/vbaProject.bin"),
        Some(&vba_bytes),
        "vbaProject.bin must survive byte-identical"
    );

    // (iii) every non-writer-owned original part is byte-identical in the output
    for (name, bytes) in &fixture_entries {
        if is_writer_owned(name) {
            continue;
        }
        assert_eq!(
            output_entries.get(name),
            Some(bytes),
            "passthrough part {name} must be byte-identical"
        );
    }
    assert_eq!(
        output_entries.get("xl/tables/table1.xml"),
        Some(&table_bytes)
    );

    // (iv) 0.10.0-D, D1: the sheet's output part name is its ORIGIN's real part name
    // (sheet3.xml), not a position-derived sheet1.xml -- this fixture's one sheet was
    // deliberately given a non-sequential source part name specifically to prove this.
    assert!(
        output_entries.contains_key("xl/worksheets/sheet3.xml"),
        "existing sheet's output part name must stay sheet3.xml (its own origin), not be \
         renumbered to sheet1.xml by position"
    );
    assert!(
        !output_entries.contains_key("xl/worksheets/sheet1.xml"),
        "no sheet in this fixture originates from sheet1.xml, so it must not appear"
    );

    // (v) + (vi) content-types: macro-enabled root override, vbaProject resolvable,
    // and every part actually present in the output resolves via the output's
    // own [Content_Types].xml (full self-consistency, not a spot check).
    let ct_xml = String::from_utf8(output_entries["[Content_Types].xml"].clone()).unwrap();
    assert!(
        ct_xml.contains("macroEnabled.main+xml"),
        "workbook.xml must declare macro-enabled content type"
    );
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
    let sheet_xml = String::from_utf8(output_entries["xl/worksheets/sheet3.xml"].clone()).unwrap();
    let a1_tag = &sheet_xml[sheet_xml.find("<c r=\"A1\"").unwrap()..];
    let a1_tag = &a1_tag[..a1_tag.find('>').unwrap() + 1];
    assert!(
        a1_tag.contains("s=\"1\""),
        "edited cell A1 must keep its original style index: {a1_tag}"
    );

    let b1_tag = &sheet_xml[sheet_xml.find("<c r=\"B1\"").unwrap()..];
    let b1_tag = &b1_tag[..b1_tag.find('>').unwrap() + 1];
    assert!(
        b1_tag.contains("s=\"1\""),
        "untouched cell B1 must keep its original style index: {b1_tag}"
    );

    let c1_tag = &sheet_xml[sheet_xml.find("<c r=\"C1\"").unwrap()..];
    let c1_tag = &c1_tag[..c1_tag.find('>').unwrap() + 1];
    assert!(
        !c1_tag.contains("s=\"1\""),
        "a brand-new cell must not spuriously inherit style 1: {c1_tag}"
    );

    assert_eq!(
        output_entries.get("xl/styles.xml"),
        Some(&styles_bytes),
        "xl/styles.xml must be byte-identical to the source, not the hardcoded minimal stylesheet"
    );

    // Merge/hidden-row/hidden-column write-back (Milestone: safe round-trip,
    // slice 3) — merges and hidden rows/columns were already threaded into
    // Vm by populate_from_sheets, but build_xlsx_sheet never emitted them.
    assert!(
        sheet_xml.contains("<mergeCells"),
        "output must re-emit the original merge: {sheet_xml}"
    );
    assert!(
        sheet_xml.contains("ref=\"D1:E1\""),
        "merged range must round-trip unchanged: {sheet_xml}"
    );

    assert!(
        sheet_xml.contains("<cols>"),
        "output must re-emit the hidden-column declaration: {sheet_xml}"
    );
    let col_tag = &sheet_xml[sheet_xml.find("<col ").unwrap()..];
    let col_tag = &col_tag[..col_tag.find('/').unwrap() + 1];
    assert_eq!(extract_attr(col_tag, "min").as_deref(), Some("6"));
    assert_eq!(extract_attr(col_tag, "max").as_deref(), Some("6"));
    assert_eq!(extract_attr(col_tag, "hidden").as_deref(), Some("1"));

    // Row 2 is hidden and has no cells at all -- must still appear as its
    // own <row hidden="1"/> element (an absent element is default-visible).
    let row2_tag = &sheet_xml[sheet_xml.find("<row r=\"2\"").unwrap()..];
    let row2_tag = &row2_tag[..row2_tag.find('>').unwrap() + 1];
    assert!(
        row2_tag.contains("hidden=\"1\""),
        "cell-less hidden row 2 must still be marked hidden: {row2_tag}"
    );

    // Formula preservation: an untouched formula cell must keep its <f>, not
    // flatten to a bare stale <v> (found via a real Excel-authored fixture --
    // see sheet_xml()'s own doc comment).
    let g1_tag = &sheet_xml[sheet_xml.find("<c r=\"G1\"").unwrap()..];
    let g1_tag = &g1_tag[..g1_tag.find("</c>").unwrap() + "</c>".len()];
    assert!(
        g1_tag.contains("<f>A1+B1</f>"),
        "untouched formula cell G1 must keep its formula, not just its cached value: {g1_tag}"
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
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    let prog = parser::parse("Sub EditCell()\n    Cells(1, 1).Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditCell").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);

    assert_eq!(
        output_entries.get("xl/tables/table1.xml"),
        Some(&TABLE_XML.as_bytes().to_vec())
    );
    assert!(!output_entries.contains_key("xl/vbaProject.bin"));

    let ct_xml = String::from_utf8(output_entries["[Content_Types].xml"].clone()).unwrap();
    assert!(
        !ct_xml.contains("macroEnabled"),
        "a workbook that never had a VBA project must not declare macro-enabled content type"
    );
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
    // 0.10.0-D, D1: sheet3.xml (this fixture's real origin part name), not sheet1.xml.
    let sheet_xml = String::from_utf8(output_entries["xl/worksheets/sheet3.xml"].clone()).unwrap();
    let a1_tag = &sheet_xml[sheet_xml.find("<c r=\"A1\"").unwrap()..];
    let a1_tag = &a1_tag[..a1_tag.find('>').unwrap() + 1];
    assert!(
        a1_tag.contains("s=\"1\""),
        "edited cell A1 must keep its style index across an in-place overwrite: {a1_tag}"
    );
    assert!(
        sheet_xml.contains("ref=\"D1:E1\""),
        "merged range must survive an in-place overwrite: {sheet_xml}"
    );

    let sheets = reader::read_workbook(&path).unwrap();
    let sheet = sheets
        .iter()
        .find(|s| s.name.to_lowercase() == "sheet1")
        .unwrap();
    assert!(matches!(
        sheet.cells.get(&(1, 1)),
        Some(reader::SheetCell::Integer(7))
    ));

    let _ = std::fs::remove_file(&path);
}

/// A `.xlsm` output must declare workbook.xml as macro-enabled even when the SOURCE has
/// no VBA project at all -- the macro-enabled content type is a property of the file
/// FORMAT (the `.xlsm` extension), not of whether a VBA project happens to be present.
/// Found live: a real Excel-authored `.xlsm` fixture with zero macros still declares
/// `macroEnabled.main+xml`; elixcee's writer previously kept the plain
/// `spreadsheetml.sheet.main+xml` type whenever `has_vba` was false, regardless of output
/// extension -- Excel treats that mismatch as fatal and refuses to open the file at all,
/// not even a repair prompt. `build_fixture_xlsx()` (no VBA project) is reused here purely
/// as "some real .xlsx-shaped source with no VBA," saved to a `.xlsm` output path.
#[test]
fn xlsm_output_declares_macro_enabled_content_type_even_without_a_vba_project() {
    let fixture_bytes = build_fixture_xlsx();
    let source_path = tmp_path("source_novba.xlsx");
    let output_path = tmp_path("output_novba.xlsm");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    let prog = parser::parse("Sub EditCell()\n    Cells(1, 1).Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditCell").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let ct_xml = String::from_utf8(output_entries["[Content_Types].xml"].clone()).unwrap();
    assert_eq!(
        resolve_content_type(&ct_xml, "xl/workbook.xml").as_deref(),
        Some("application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
        ".xlsm output must declare the macro-enabled content type regardless of whether \
         the source had a VBA project: {ct_xml}"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

/// A relationship pointing at a passthrough part (here: a theme, mirroring
/// `xl/theme/theme1.xml`) must survive into the regenerated `.rels` file, not just the
/// part's bytes. Found live: `xl/theme/theme1.xml` passed through byte-identical, but
/// `xl/_rels/workbook.xml.rels` -- entirely regenerated from a fixed template that only
/// ever emitted worksheet/sharedStrings/styles/vbaProject relationships -- silently
/// dropped the theme relationship, orphaning an otherwise-intact part. Real Excel refused
/// to open the result outright. Uses a dedicated minimal fixture (not `build_fixture_xlsm`)
/// since none of the existing fixtures have a theme part or relationship.
#[test]
fn passthrough_part_referenced_only_by_a_non_writer_owned_relationship_type_keeps_its_relationship()
{
    const CONTENT_TYPES_WITH_THEME: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n",
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n",
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>\n",
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n",
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n",
        "<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\n",
        "</Types>\n",
    );
    const WORKBOOK_RELS_WITH_THEME: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>\n",
        "</Relationships>\n",
    );
    const THEME_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<theme xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"t\"/>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_WITH_THEME.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", &workbook_xml().into_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS_WITH_THEME.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet1.xml",
        concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
            "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>\n</worksheet>\n",
        )
        .as_bytes(),
    );
    zip_add(&mut zip, "xl/theme/theme1.xml", THEME_XML.as_bytes());
    let fixture_bytes = zip.finish().unwrap().into_inner();

    let source_path = tmp_path("source_theme.xlsx");
    let output_path = tmp_path("output_theme.xlsx");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    let prog = parser::parse("Sub EditCell()\n    Cells(1, 1).Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditCell").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);

    assert_eq!(
        output_entries.get("xl/theme/theme1.xml"),
        Some(&THEME_XML.as_bytes().to_vec()),
        "theme1.xml itself must still pass through byte-identical"
    );
    let rels_xml = String::from_utf8(output_entries["xl/_rels/workbook.xml.rels"].clone()).unwrap();
    assert!(
        rels_xml.contains("relationships/theme") && rels_xml.contains("theme/theme1.xml"),
        "the theme relationship must survive into the regenerated workbook.xml.rels, not \
         just the theme part's bytes -- an orphaned part is what made real Excel refuse to \
         open the file outright: {rels_xml}"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

/// A workbook's saved sheet order must match its *source* order, not an
/// alphabetical sort, and each tab's original letter case must survive too.
/// Found via a hand-built two-sheet fixture named "Zebra" (first) / "Alpha"
/// (second): `save_xlsx_impl` used to derive its entire sheet-iteration
/// order from `Vm::sheet_names()`, which sorts and lowercases — so every
/// save of this fixture silently swapped the tab order to "Alpha"/"Zebra"
/// and relabeled both tabs lowercase, with no macro touching sheets at all.
/// Every prior round-trip fixture happened to already be alphabetical
/// (Sheet1/2/3), so nothing caught either bug before. Root-caused to
/// `Vm::sheet_order` (new, insertion-ordered) and `WorksheetOrigin`'s new
/// `original_display_name` not existing at all; `sheet_names()` itself is
/// left alphabetical/lowercased on purpose -- see its doc comment.
#[test]
fn sheet_order_and_display_case_survive_a_save_even_when_source_names_are_not_alphabetical() {
    const WORKBOOK_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n",
        "<sheet name=\"Zebra\" sheetId=\"7\" r:id=\"rId1\"/>\n",
        "<sheet name=\"Alpha\" sheetId=\"3\" r:id=\"rId2\"/>\n",
        "</sheets>\n</workbook>\n",
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>\n",
        "</Relationships>\n",
    );
    const MINIMAL_SHEET: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>\n</worksheet>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", WORKBOOK_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet1.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet2.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    let fixture_bytes = zip.finish().unwrap().into_inner();

    let source_path = tmp_path("source_zebra_alpha.xlsx");
    let output_path = tmp_path("output_zebra_alpha.xlsx");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let wb_xml = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();

    let names: Vec<String> = wb_xml
        .match_indices("<sheet ")
        .map(|(start, _)| {
            let tag_end = wb_xml[start..].find("/>").unwrap() + start;
            extract_attr(&wb_xml[start..tag_end], "name").unwrap()
        })
        .collect();
    assert_eq!(
        names,
        vec!["Zebra".to_string(), "Alpha".to_string()],
        "sheet order must match the source file (not an alphabetical sort) and each name's \
         original case must survive (not get lowercased): {wb_xml}"
    );

    let ids: Vec<String> = wb_xml
        .match_indices("<sheet ")
        .map(|(start, _)| {
            let tag_end = wb_xml[start..].find("/>").unwrap() + start;
            extract_attr(&wb_xml[start..tag_end], "sheetId").unwrap()
        })
        .collect();
    assert_eq!(
        ids,
        vec!["7".to_string(), "3".to_string()],
        "each sheet's original sheetId must stay attached to its own name, not get \
         reshuffled along with position: {wb_xml}"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

/// GitHub #2: a sheet created via `Vm::ensure_sheet` (backing both VBA's `Sheets.Add`
/// and Python's `set_sheet()`, which creates on demand) had no `WorksheetOrigin` from a
/// loaded file, so `save_xlsx_impl`'s display-name fallback used the lowercased internal
/// key -- `"NewSheet"` round-tripped as `"newsheet"`. Non-ASCII names (e.g. Japanese)
/// were never affected: `to_lowercase()` is a no-op on them, which is exactly why this
/// went unnoticed until a plain ASCII name was tried. Fixed by `ensure_sheet` itself
/// recording the caller's as-written name into `WorksheetOrigin.original_display_name`.
#[test]
fn a_sheet_created_via_ensure_sheet_keeps_its_original_case_on_save() {
    let source_path = real_fixture("fixture1_values_styles_merge_hidden.xlsm");
    let output_path = tmp_path("new_sheet_case_output.xlsm");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    vm.ensure_sheet("NewSheet");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let wb_xml = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();

    let names: Vec<String> = wb_xml
        .match_indices("<sheet ")
        .map(|(start, _)| {
            let tag_end = wb_xml[start..].find("/>").unwrap() + start;
            extract_attr(&wb_xml[start..tag_end], "name").unwrap()
        })
        .collect();
    assert!(
        names.contains(&"NewSheet".to_string()),
        "a freshly-created sheet's name must survive exactly as passed to ensure_sheet(), \
         not get lowercased: {names:?}"
    );
    assert!(
        !names.contains(&"newsheet".to_string()),
        "must not ALSO carry a lowercased duplicate: {names:?}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// GitHub #4: `get_cell()` returns a date-formatted cell as the raw Excel serial number
/// (e.g. `45366`), with no way for a caller to tell it apart from a plain number --
/// unlike openpyxl, which converts using the cell's number format. Rather than guess and
/// change `get_cell`'s return type (a breaking change), `get_cell_number_format` exposes
/// the resolved format string itself, letting the caller convert if it wants to.
#[test]
fn get_cell_number_format_resolves_a_builtin_date_format_from_styles_xml() {
    const WORKBOOK_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n</sheets>\n</workbook>\n",
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "</Relationships>\n",
    );
    // cellXfs index 0 = General (no numFmtId), index 1 = numFmtId 14 ("m/d/yyyy") --
    // matches a real producer's shape: index 0 always exists as the default style.
    const STYLES_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<cellXfs count=\"2\">\n<xf/>\n<xf numFmtId=\"14\"/>\n</cellXfs>\n</styleSheet>\n",
    );
    const SHEET_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<sheetData><row r=\"1\">",
        "<c r=\"A1\" s=\"1\"><v>45366</v></c>",
        "<c r=\"B1\"><v>42</v></c>",
        "</row></sheetData>\n</worksheet>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", WORKBOOK_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(&mut zip, "xl/styles.xml", STYLES_XML.as_bytes());
    zip_add(&mut zip, "xl/worksheets/sheet1.xml", SHEET_XML.as_bytes());
    let fixture_bytes = zip.finish().unwrap().into_inner();

    let source_path = tmp_path("date_format_source.xlsx");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");

    assert_eq!(vm.get_cell_number_format(1, 1), Some("m/d/yyyy"));
    assert_eq!(
        vm.get_cell_number_format(1, 2),
        None,
        "a cell with the default/General style must report no format"
    );
    assert_eq!(
        vm.get_cell(1, 1),
        elixcee::vm::Variant::Integer(45366),
        "get_cell itself must still return the raw serial number, unchanged -- \
         get_cell_number_format is additive, not a breaking change to get_cell"
    );

    let _ = std::fs::remove_file(&source_path);
}

/// GitHub #5: `Range.AutoFilter` was a silent no-op -- no rows got hidden even with
/// `Field`/`Criteria1` given. The VM-side effect (hiding non-matching rows) reuses the
/// same `Vm.sheet_visibility` a loaded file's own hidden rows already round-trip
/// through -- this proves that round-trip for AutoFilter-driven hides specifically, not
/// just that the VM's in-memory state changes. `<autoFilter ref="...">` itself (the
/// dropdown-arrow element) is deliberately NOT persisted -- no real fixture in this repo
/// has one, and this project's hard gate is no writer code for an OOXML element without
/// fixture evidence.
#[test]
fn range_autofilter_hidden_rows_survive_a_save() {
    let source_path = real_fixture("fixture1_values_styles_merge_hidden.xlsm");
    let output_path = tmp_path("autofilter_output.xlsm");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse(concat!(
        "Sub FilterIt()\n",
        "    Cells(20,1).Value = \"Name\"\n    Cells(20,2).Value = \"Age\"\n",
        "    Cells(21,1).Value = \"Charlie\"\n    Cells(21,2).Value = 25\n",
        "    Cells(22,1).Value = \"Alice\"\n    Cells(22,2).Value = 40\n",
        "    Range(\"A20:B22\").AutoFilter Field:=2, Criteria1:=\"25\"\n",
        "End Sub\n",
    ))
    .unwrap();
    vm.run_sub(&prog, "FilterIt").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    let row22 = &out_sheet1[out_sheet1.find("r=\"22\"").unwrap() - 5..];
    let row22_tag = &row22[..row22.find('>').unwrap()];
    assert!(
        row22_tag.contains("hidden=\"1\""),
        "row 22 (Alice/40, doesn't match Criteria1) must be hidden in the saved file: {row22_tag}"
    );
    let row21 = &out_sheet1[out_sheet1.find("r=\"21\"").unwrap() - 5..];
    let row21_tag = &row21[..row21.find('>').unwrap()];
    assert!(
        !row21_tag.contains("hidden=\"1\""),
        "row 21 (Charlie/25, matches Criteria1) must stay visible: {row21_tag}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// R1 (bulk worksheet range/row API, see docs/openpyxl-gap-audit.md): the new
/// `Vm::write_rect`/`read_rect` are exercised directly here (no PyO3 needed --
/// the test crate links `elixcee` as a lib), on a real fixture that already
/// has a merge, a hidden column, and a hidden row, to prove the new write
/// path doesn't disturb any of that pre-existing state on save.
#[test]
fn write_rect_on_a_real_fixture_survives_a_save_without_disturbing_existing_state() {
    let source_path = real_fixture("fixture1_values_styles_merge_hidden.xlsm");
    let output_path = tmp_path("write_rect_output.xlsm");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let key = vm.resolve_sheet_key(None).unwrap();
    vm.write_rect(
        &key,
        (10, 1),
        &[vec![
            elixcee::vm::Variant::Str("R1WriteRect".to_string()),
            elixcee::vm::Variant::Integer(42),
        ]],
    );
    assert_eq!(
        vm.read_rect(&key, 10, 1, 10, 2),
        vec![vec![
            elixcee::vm::Variant::Str("R1WriteRect".to_string()),
            elixcee::vm::Variant::Integer(42)
        ]]
    );
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(r#"<c r="B10">"#) || out_sheet1.contains(r#"<c r="B10" "#),
        "the written B10=42 cell must appear in the saved sheet: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains(r#"<mergeCell ref="B1:C1"/>"#),
        "the fixture's pre-existing B1:C1 merge must survive: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains(r#"min="4" max="4" hidden="1""#),
        "the fixture's pre-existing hidden column D must survive: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains(r#"<row r="5" hidden="1">"#)
            || out_sheet1.contains(r#"<row r="5" hidden="1"/>"#),
        "the fixture's pre-existing hidden row 5 must survive: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// P1 core 3: `rename_sheet`'s atomic re-key exercised end-to-end -- the exact
/// regression a broken re-key would produce is losing the merge/hidden-column/
/// hidden-row state that lived under the OLD lowercased key.
#[test]
fn rename_sheet_round_trips_merge_and_hidden_metadata_on_the_real_fixture() {
    let source_path = real_fixture("fixture1_values_styles_merge_hidden.xlsm");
    let output_path = tmp_path("rename_sheet_output.xlsm");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    vm.rename_sheet("Sheet1", "Renamed").unwrap();
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_wb = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();
    assert!(
        out_wb.contains(r#"<sheet name="Renamed""#),
        "the sheet must be renamed in xl/workbook.xml: {out_wb}"
    );

    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        out_sheet1.contains(r#"<mergeCell ref="B1:C1"/>"#),
        "the fixture's pre-existing B1:C1 merge must survive a rename: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains(r#"min="4" max="4" hidden="1""#),
        "the fixture's pre-existing hidden column D must survive a rename: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains(r#"<row r="5" hidden="1">"#)
            || out_sheet1.contains(r#"<row r="5" hidden="1"/>"#),
        "the fixture's pre-existing hidden row 5 must survive a rename: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// P1 core 3 follow-up fix: a real, reviewer-caught bug -- `rename_sheet` re-keys
/// `worksheet_origins`/`sheet_order` in lockstep, so the deletion-only guard on
/// `<definedNames>` passthrough (`no_sheet_was_deleted`) stayed true through a
/// rename, and an earlier version of the `move_sheet` fix only set
/// `defined_names_may_be_stale` from `move_sheet`, not `rename_sheet` -- so a
/// renamed sheet's `<definedNames>` (e.g. `<definedName>Sheet1!$F$5</definedName>`)
/// survived a save verbatim, now dangling: pointing at a sheet name that no longer
/// exists in `<sheets>`. fixture4 is the one real fixture with genuine
/// `<definedNames>` content, confirmed by
/// `real_excel_workbook_metadata_survives_a_save`'s own sibling test file.
#[test]
fn rename_sheet_drops_defined_names_that_would_reference_the_old_name() {
    let source_path = real_fixture("fixture4_hyperlink_comment_name.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_wb = String::from_utf8(fixture_entries["xl/workbook.xml"].clone()).unwrap();
    assert!(
        source_wb.contains("<definedNames>") && source_wb.contains("Sheet1!$F$5"),
        "fixture no longer contains the expected definedName -- test needs updating: {source_wb}"
    );

    let output_path = tmp_path("rename_sheet_defined_names_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    vm.rename_sheet("Sheet1", "Renamed").unwrap();
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_wb = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();
    assert!(
        !out_wb.contains("<definedNames>"),
        "definedNames must be dropped entirely once a sheet is renamed, not carried \
         through referencing a sheet name that no longer exists: {out_wb}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// P2: `defined_names` exercised against the one real fixture with genuine
/// `<definedNames>` content -- confirms the reader correctly parses a real
/// Excel-authored `<definedName name="..." comment="...">TEXT</definedName>`
/// element (ignoring the unrelated `comment` attribute) into `{name: text}`.
#[test]
fn defined_names_reads_the_real_fixtures_defined_name() {
    let source_path = real_fixture("fixture4_hyperlink_comment_name.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");

    let names = vm.defined_names().expect("should read defined names");
    assert_eq!(
        names.get("test").map(|s| s.as_str()),
        Some("Sheet1!$F$5"),
        "{:?}",
        names
    );
}

/// 0.14.0-B Phase 2/3 updated this: merges and same-axis hidden-row/col
/// markers now shift on insert/delete (`shift_merged_ranges_for_structural_edit`,
/// `shift_hidden_intervals_for_structural_edit`), so B1:C1 (row 1) moves to
/// B2:C2 and hidden row 5 moves to row 6 when a row is inserted at row 1 --
/// previously this test pinned the opposite (neither shifted) as the
/// disclosed gap. Hidden COLUMN D is unaffected -- this is a row-axis
/// insert, and column-hidden state only ever shifts on a column-axis edit.
#[test]
fn insert_rows_on_a_merged_and_hidden_row_sheet_shifts_the_merge_and_same_axis_hidden_marker() {
    let source_path = real_fixture("fixture1_values_styles_merge_hidden.xlsm");
    let output_path = tmp_path("insert_rows_on_sheet_output.xlsm");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let key = vm.resolve_sheet_key(None).unwrap();
    vm.insert_rows_on_sheet(&key, 1, 1);
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(r#"<mergeCell ref="B2:C2"/>"#),
        "merge ref must shift down to the post-insert row: {out_sheet1}"
    );
    assert!(
        !out_sheet1.contains(r#"<mergeCell ref="B1:C1"/>"#),
        "the stale, pre-shift merge ref must not also still be present: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains(r#"min="4" max="4" hidden="1""#),
        "hidden column D marker must be unchanged -- column shifting is not this axis: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains(r#"<row r="6" hidden="1">"#)
            || out_sheet1.contains(r#"<row r="6" hidden="1"/>"#),
        "hidden row marker must shift from row 5 to row 6: {out_sheet1}"
    );
    assert!(
        !out_sheet1.contains(r#"<row r="5" hidden="1""#),
        "the stale, pre-shift hidden-row marker must not also still be present: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// P1 remainder: `merge_cells`/`unmerge_cells` exercised end-to-end against the
/// one real fixture with a pre-existing merge (B1:C1). A new non-overlapping
/// merge (D1:E1) is added and must round-trip alongside B1:C1; then, continuing
/// the same session, B1:C1 is removed and must actually disappear from the saved
/// XML rather than merely being absent from a freshly-added set.
#[test]
fn merge_cells_and_unmerge_cells_round_trip_on_the_real_fixture() {
    let source_path = real_fixture("fixture1_values_styles_merge_hidden.xlsm");
    let output_path_1 = tmp_path("merge_cells_output_1.xlsm");
    let output_path_2 = tmp_path("merge_cells_output_2.xlsm");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let key = vm.resolve_sheet_key(None).unwrap();

    vm.merge_cells(&key, 1, 4, 1, 5).unwrap(); // D1:E1, non-overlapping with B1:C1
    save_workbook(&vm, &output_path_1).expect("save-as should succeed");

    let output_bytes_1 = std::fs::read(&output_path_1).unwrap();
    let output_entries_1 = read_all_zip_entries(&output_bytes_1);
    let out_sheet1_1 =
        String::from_utf8(output_entries_1["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        out_sheet1_1.contains(r#"<mergeCell ref="B1:C1"/>"#),
        "the fixture's pre-existing B1:C1 merge must survive: {out_sheet1_1}"
    );
    assert!(
        out_sheet1_1.contains(r#"<mergeCell ref="D1:E1"/>"#),
        "the newly-added D1:E1 merge must be saved: {out_sheet1_1}"
    );

    vm.unmerge_cells(&key, 1, 2, 1, 3).unwrap(); // remove B1:C1
    save_workbook(&vm, &output_path_2).expect("save-as should succeed");

    let output_bytes_2 = std::fs::read(&output_path_2).unwrap();
    let output_entries_2 = read_all_zip_entries(&output_bytes_2);
    let out_sheet1_2 =
        String::from_utf8(output_entries_2["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        !out_sheet1_2.contains(r#"<mergeCell ref="B1:C1"/>"#),
        "B1:C1 must actually be gone after unmerge_cells, not just absent from a fresh set: {out_sheet1_2}"
    );
    assert!(
        out_sheet1_2.contains(r#"<mergeCell ref="D1:E1"/>"#),
        "D1:E1 must remain after removing the unrelated B1:C1 merge: {out_sheet1_2}"
    );

    let _ = std::fs::remove_file(&output_path_1);
    let _ = std::fs::remove_file(&output_path_2);
}

/// P1 remainder: `sort_range_on_sheet` exercised end-to-end -- a bulk value
/// rewrite must not disturb the fixture's pre-existing merge/hidden-row/
/// hidden-column state, matching the same 3-assertion pattern as
/// `write_rect_on_a_real_fixture_survives_a_save_without_disturbing_existing_state`.
#[test]
fn sort_range_on_sheet_survives_a_save_and_does_not_disturb_unrelated_state() {
    let source_path = real_fixture("fixture1_values_styles_merge_hidden.xlsm");
    let output_path = tmp_path("sort_range_output.xlsm");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let key = vm.resolve_sheet_key(None).unwrap();

    // Rows 20-22, column A -- far from the fixture's merge (row 1) and hidden
    // row/column markers (row 5, column D).
    vm.write_rect(
        &key,
        (20, 1),
        &[
            vec![elixcee::vm::Variant::Integer(3)],
            vec![elixcee::vm::Variant::Integer(1)],
            vec![elixcee::vm::Variant::Integer(2)],
        ],
    );
    vm.sort_range_on_sheet(&key, 20, 1, 22, 1, 1, false, false);
    assert_eq!(
        vm.read_rect(&key, 20, 1, 22, 1),
        vec![
            vec![elixcee::vm::Variant::Integer(1)],
            vec![elixcee::vm::Variant::Integer(2)],
            vec![elixcee::vm::Variant::Integer(3)],
        ]
    );
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(r#"<mergeCell ref="B1:C1"/>"#),
        "the fixture's pre-existing B1:C1 merge must survive a sort elsewhere on the sheet: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains(r#"min="4" max="4" hidden="1""#),
        "the fixture's pre-existing hidden column D must survive a sort elsewhere on the sheet: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains(r#"<row r="5" hidden="1">"#)
            || out_sheet1.contains(r#"<row r="5" hidden="1"/>"#),
        "the fixture's pre-existing hidden row 5 must survive a sort elsewhere on the sheet: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// P2: `set_row_hidden`/`set_column_hidden` exercised end-to-end -- hiding a
/// new row/column must round-trip alongside the fixture's pre-existing hidden
/// row 5/column D, and unhiding that pre-existing row must actually remove its
/// `hidden="1"` marker (not just fail to add a redundant one) while leaving
/// unrelated state (the B1:C1 merge, the still-hidden column D) untouched.
#[test]
fn set_row_hidden_and_set_column_hidden_round_trip_on_the_real_fixture() {
    let source_path = real_fixture("fixture1_values_styles_merge_hidden.xlsm");
    let output_path_1 = tmp_path("hidden_output_1.xlsm");
    let output_path_2 = tmp_path("hidden_output_2.xlsm");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let key = vm.resolve_sheet_key(None).unwrap();

    vm.set_row_hidden_on_sheet(&key, 20, true);
    vm.set_column_hidden_on_sheet(&key, 6, true); // column F
    save_workbook(&vm, &output_path_1).expect("save-as should succeed");

    let output_bytes_1 = std::fs::read(&output_path_1).unwrap();
    let output_entries_1 = read_all_zip_entries(&output_bytes_1);
    let out_sheet1_1 =
        String::from_utf8(output_entries_1["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        out_sheet1_1.contains(r#"<row r="5" hidden="1">"#)
            || out_sheet1_1.contains(r#"<row r="5" hidden="1"/>"#),
        "the fixture's pre-existing hidden row 5 must survive: {out_sheet1_1}"
    );
    assert!(
        out_sheet1_1.contains(r#"<row r="20" hidden="1">"#)
            || out_sheet1_1.contains(r#"<row r="20" hidden="1"/>"#),
        "the newly-hidden row 20 must be saved: {out_sheet1_1}"
    );
    assert!(
        out_sheet1_1.contains(r#"min="4" max="4" hidden="1""#),
        "the fixture's pre-existing hidden column D must survive: {out_sheet1_1}"
    );
    assert!(
        out_sheet1_1.contains(r#"min="6" max="6" hidden="1""#),
        "the newly-hidden column F must be saved: {out_sheet1_1}"
    );

    vm.set_row_hidden_on_sheet(&key, 5, false); // unhide the fixture's pre-existing hidden row
    save_workbook(&vm, &output_path_2).expect("save-as should succeed");

    let output_bytes_2 = std::fs::read(&output_path_2).unwrap();
    let output_entries_2 = read_all_zip_entries(&output_bytes_2);
    let out_sheet1_2 =
        String::from_utf8(output_entries_2["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        !out_sheet1_2.contains(r#"<row r="5" hidden="1">"#)
            && !out_sheet1_2.contains(r#"<row r="5" hidden="1"/>"#),
        "row 5 must actually be unhidden, not just fail to gain a duplicate marker: {out_sheet1_2}"
    );
    assert!(
        out_sheet1_2.contains(r#"<row r="20" hidden="1">"#)
            || out_sheet1_2.contains(r#"<row r="20" hidden="1"/>"#),
        "the unrelated newly-hidden row 20 must remain hidden: {out_sheet1_2}"
    );
    assert!(
        out_sheet1_2.contains(r#"<mergeCell ref="B1:C1"/>"#),
        "the fixture's pre-existing B1:C1 merge must survive unhiding an unrelated row: {out_sheet1_2}"
    );
    assert!(
        out_sheet1_2.contains(r#"min="4" max="4" hidden="1""#),
        "the fixture's pre-existing hidden column D must survive unhiding an unrelated row: {out_sheet1_2}"
    );

    let _ = std::fs::remove_file(&output_path_1);
    let _ = std::fs::remove_file(&output_path_2);
}

/// P2: `copy_sheet` exercised end-to-end -- the copy must carry the source's
/// merge/hidden-row/hidden-column state and cell values into its own,
/// independent worksheet part, while the original sheet's own part is left
/// completely untouched.
#[test]
fn copy_sheet_round_trips_merge_and_hidden_metadata_on_the_real_fixture() {
    let source_path = real_fixture("fixture1_values_styles_merge_hidden.xlsm");
    let output_path = tmp_path("copy_sheet_output.xlsm");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    vm.copy_sheet("Sheet1", "Copy").unwrap();
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);

    let out_wb = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();
    assert!(
        out_wb.contains(r#"<sheet name="Sheet1""#),
        "the original sheet must still be listed: {out_wb}"
    );
    assert!(
        out_wb.contains(r#"<sheet name="Copy""#),
        "the copy must be listed in xl/workbook.xml: {out_wb}"
    );

    // The original's own worksheet part must be completely untouched.
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        out_sheet1.contains(r#"<mergeCell ref="B1:C1"/>"#),
        "the original's merge must survive being copied from: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains(r#"min="4" max="4" hidden="1""#),
        "the original's hidden column must survive being copied from: {out_sheet1}"
    );

    // The copy must land in its OWN worksheet part (not overwrite sheet1.xml)
    // and carry the same merge/hidden-row/hidden-column state and values.
    let copy_part = output_entries
        .keys()
        .find(|k| k.starts_with("xl/worksheets/sheet") && k.as_str() != "xl/worksheets/sheet1.xml")
        .unwrap_or_else(|| {
            panic!(
                "expected a second worksheet part, got: {:?}",
                output_entries.keys().collect::<Vec<_>>()
            )
        });
    let out_copy = String::from_utf8(output_entries[copy_part].clone()).unwrap();
    assert!(
        out_copy.contains(r#"<mergeCell ref="B1:C1"/>"#),
        "the copy must carry the source's merge: {out_copy}"
    );
    assert!(
        out_copy.contains(r#"min="4" max="4" hidden="1""#),
        "the copy must carry the source's hidden column: {out_copy}"
    );
    assert!(
        out_copy.contains(r#"<row r="5" hidden="1">"#)
            || out_copy.contains(r#"<row r="5" hidden="1"/>"#),
        "the copy must carry the source's hidden row: {out_copy}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// Real report against the released `0.10.0`: a source that binds the relationships
/// namespace to a prefix OTHER than the conventional `r:` (here `rel:`) is fully valid
/// OOXML on its own -- XML namespace binding is about the URI, not the prefix spelling.
/// Carrying such a source's root `<workbook>` attrs through unchanged, while
/// `build_xlsx_workbook` still hardcodes the literal `r:` prefix on `<sheet r:id="...">`,
/// used to produce output where `r:` was never bound to anything -- a real XML error
/// (`lxml`/openpyxl reject it as "unbound prefix"), not just a lossy passthrough. Fixed
/// by `reader::ensure_r_prefix_bound`.
#[test]
fn workbook_xml_still_binds_the_r_prefix_when_the_source_used_a_different_one() {
    const WORKBOOK_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:rel=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n",
        "<sheet name=\"Sheet1\" sheetId=\"1\" rel:id=\"rId1\"/>\n",
        "</sheets>\n</workbook>\n",
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "</Relationships>\n",
    );
    const MINIMAL_SHEET: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>\n</worksheet>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", WORKBOOK_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet1.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    let fixture_bytes = zip.finish().unwrap().into_inner();

    let source_path = tmp_path("source_alt_rel_prefix.xlsx");
    let output_path = tmp_path("output_alt_rel_prefix.xlsx");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let wb_xml = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();

    assert!(
        wb_xml.contains(
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        ),
        "the writer's own hardcoded r:id usage below must have a real xmlns:r binding, \
         regardless of what prefix the source used for the same namespace: {wb_xml}"
    );
    assert!(
        wb_xml.contains("r:id=\"rId1\""),
        "the sheet must still reference its own relationship: {wb_xml}"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

/// A genuine Microsoft-Excel-for-Mac-authored `.xlsm` (real VBA project, real
/// `xr:uid`/`calcChain.xml`/`theme1.xml`), not a hand-built stand-in -- see
/// `compat/oracle-excel-com/results/0.9.0-A_summary.md` for the full real-Excel
/// validation this fixture is part of (open in real Excel, edit via elixcee,
/// save, reopen -- no repair warning, formulas/relationships/vbaProject
/// preserved). This function reads it directly rather than duplicating the
/// bytes into `tests/fixtures/xlsm_roundtrip/` -- see that directory's own
/// README for why.
fn real_fixture(name: &str) -> String {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("compat/oracle-excel-com/fixtures/pristine")
        .join(name)
        .to_str()
        .unwrap()
        .to_string()
}

/// The real-fixture counterpart to the synthetic flagship test above: same
/// assertions (edited cell, unedited cell, `xl/vbaProject.bin` byte-identity,
/// every other original part byte-identical, content-types self-consistency),
/// but against a real Microsoft-Excel-authored `.xlsm` instead of a hand-built
/// one -- and additionally checks the two relationship-carry-over bugs found
/// via this exact fixture (theme/docProps relationships surviving into the
/// regenerated `.rels` files, not just the parts' bytes) don't regress.
/// Covers both save modes, mirroring `xlsm_roundtrip_in_place_save_preserves_vba_project`.
#[test]
fn real_excel_xlsm_roundtrip_preserves_vba_project_and_relationships() {
    let source_path = real_fixture("fixture2_vba_macro.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let vba_bytes = fixture_entries["xl/vbaProject.bin"].clone();

    // --- save-as ---
    let output_path = tmp_path("real_fixture_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditB1()\n    Cells(1, 2).Value = 999\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditB1").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");
    check_real_fixture_output(&fixture_entries, &vba_bytes, &output_path, "save-as");

    // --- in-place: source_path == output_path, on a scratch copy so the committed
    // fixture under compat/oracle-excel-com/fixtures/pristine/ is never touched ---
    let inplace_path = tmp_path("real_fixture_inplace.xlsm");
    std::fs::copy(&source_path, &inplace_path).unwrap();
    let mut vm2 = Vm::new();
    vm2.load_workbook_file(&inplace_path)
        .expect("copied real fixture should load");
    vm2.run_sub(&prog, "EditB1").expect("macro should run");
    save_workbook(&vm2, &inplace_path).expect("in-place save should succeed");
    check_real_fixture_output(&fixture_entries, &vba_bytes, &inplace_path, "in-place");

    let _ = std::fs::remove_file(&output_path);
    let _ = std::fs::remove_file(&inplace_path);
}

/// 0.10.0-B slice 1: `<sheetViews>` (freeze panes via `<pane>`, active-cell
/// `<selection>`) opaque-fragment passthrough. Before this slice, every real
/// fixture lost `<sheetViews>` entirely on a save -- confirmed via
/// `compat/oracle-excel-com/mechanical_check.py`'s `check_inline_worksheet_elements()`
/// against all 7 fixtures (see that commit). This is the positive regression
/// guard: load the real freeze-pane fixture, edit a cell, save, and assert the
/// exact `<pane .../>` the source contains survives byte-for-byte, plus the
/// root `<worksheet>` tag now carries the source's own namespace declarations
/// (needed for other, later 0.10.0-B slices that use prefixed attributes)
/// instead of the old hardcoded minimal one. Also covers slice 3
/// (`<pageMargins>`), present in this same fixture.
#[test]
fn real_excel_freeze_pane_sheetviews_survive_a_save() {
    let source_path = real_fixture("fixture7_freeze_pane.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_sheet1 =
        String::from_utf8(fixture_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        source_sheet1.contains(r#"<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>"#),
        "fixture no longer contains the expected freeze-pane shape -- test needs updating: {source_sheet1}"
    );

    let output_path = tmp_path("freeze_pane_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditA1()\n    Range(\"A1\").Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditA1").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(r#"<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>"#),
        "freeze pane must survive a save verbatim: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains("<sheetViews>") && out_sheet1.contains("</sheetViews>"),
        "sheetViews container must survive: {out_sheet1}"
    );
    assert!(
        out_sheet1
            .contains("xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\""),
        "root <worksheet> tag should carry the source's own namespace declarations, not the \
         old hardcoded minimal one: {out_sheet1}"
    );
    // Edited cell still round-trips correctly alongside the passthrough fragment.
    assert!(
        out_sheet1.contains("<c r=\"A1\""),
        "edited cell A1 must still be present: {out_sheet1}"
    );
    // 0.10.0-B slice 3: pageMargins.
    assert!(
        out_sheet1.contains(
            r#"<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"#
        ),
        "pageMargins must survive a save verbatim: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// 0.10.0-B slice 2: sheetPr/sheetFormatPr/phoneticPr/dataValidations opaque-fragment
/// passthrough. fixture3_table_validation_conditional.xlsm is the only real fixture that
/// carries all four (see INVENTORY.md) -- fixture7's freeze-pane test above only proves
/// sheetViews/sheetFormatPr/phoneticPr survive, not sheetPr/dataValidations, since fixture7
/// has neither.
#[test]
fn real_excel_sheetpr_and_data_validations_survive_a_save() {
    let source_path = real_fixture("fixture3_table_validation_conditional.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_sheet1 =
        String::from_utf8(fixture_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    for needle in [
        r#"<sheetPr codeName="Sheet1"/>"#,
        r#"<sheetFormatPr baseColWidth="10" defaultRowHeight="20"/>"#,
        r#"<phoneticPr fontId="1"/>"#,
    ] {
        assert!(
            source_sheet1.contains(needle),
            "fixture no longer contains {needle:?} -- test needs updating: {source_sheet1}"
        );
    }

    let output_path = tmp_path("sheetpr_dataval_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditB2()\n    Cells(2, 2).Value = 999\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditB2").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    for needle in [
        r#"<sheetPr codeName="Sheet1"/>"#,
        r#"<sheetFormatPr baseColWidth="10" defaultRowHeight="20"/>"#,
        r#"<phoneticPr fontId="1"/>"#,
    ] {
        assert!(
            out_sheet1.contains(needle),
            "{needle:?} must survive a save verbatim: {out_sheet1}"
        );
    }
    // dataValidations' <dataValidation> child carries an xr:uid attribute -- only valid
    // if the root <worksheet> tag also declares the xr namespace, which root_attrs
    // passthrough (slice 1) already provides. Assert the whole thing survives together,
    // not just the container tag, so a namespace-declaration regression would fail here.
    assert!(
        out_sheet1.contains(
            r#"<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E1" xr:uid="{BF4C2CDE-5B18-5247-880B-6E29EFBEE104}"><formula1>"Yes,No,Maybe"</formula1></dataValidation></dataValidations>"#
        ),
        "dataValidations (with its xr:uid-bearing child) must survive verbatim: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains("<c r=\"B2\""),
        "edited cell B2 must still be present: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// 0.10.0-D: `<tableParts>`, the first relationship-backed element restored --
/// fixture3_table_validation_conditional.xlsm's `<tableParts count="1">
/// <tablePart r:id="rId1"/></tableParts>` used to survive structurally
/// (its `.rels` and `xl/tables/table1.xml` both passed through byte-identical)
/// while going completely unreferenced from the regenerated `sheet1.xml`,
/// confirmed as `SOURCE_REFERENCE_LOSS` by `check_source_references()` before
/// this test existed. Also asserts the sheet's own `.rels` and target table
/// part still survive byte-identical alongside it -- restoring the reference
/// without both of those would just move the loss, not fix it.
#[test]
fn real_excel_table_parts_survive_a_save() {
    let source_path = real_fixture("fixture3_table_validation_conditional.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_sheet1 =
        String::from_utf8(fixture_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        source_sheet1.contains(r#"<tableParts count="1"><tablePart r:id="rId1"/></tableParts>"#),
        "fixture no longer contains the expected tableParts shape -- test needs updating: {source_sheet1}"
    );

    let output_path = tmp_path("table_parts_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditB3()\n    Cells(3, 2).Value = 999\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditB3").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(r#"<tableParts count="1"><tablePart r:id="rId1"/></tableParts>"#),
        "tableParts must survive a save verbatim, referencing the same rId as the source: {out_sheet1}"
    );
    assert_eq!(
        output_entries.get("xl/worksheets/_rels/sheet1.xml.rels"),
        fixture_entries.get("xl/worksheets/_rels/sheet1.xml.rels"),
        "the worksheet .rels the restored r:id points at must still be byte-identical"
    );
    assert_eq!(
        output_entries.get("xl/tables/table1.xml"),
        fixture_entries.get("xl/tables/table1.xml"),
        "the table part itself must still be byte-identical"
    );
    assert!(
        out_sheet1.contains("<c r=\"B3\""),
        "edited cell B3 must still be present: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// 0.10.0-D: the `rels_survived` safety gate. A sheet whose source worksheet XML
/// contains `<tableParts>` but whose own `.rels` file is NOT among the passthrough
/// parts (a shape that shouldn't happen with a well-formed real fixture, but must
/// never be trusted blindly) must NOT get `<tableParts>` spliced back -- doing so
/// would emit a dangling `r:id`, a real Excel repair warning and strictly worse
/// than the pre-0.10.0-D silent inertness this whole milestone exists to fix.
#[test]
fn table_parts_is_not_restored_when_its_own_rels_file_did_not_survive() {
    const WORKBOOK_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n</sheets>\n</workbook>\n",
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "</Relationships>\n",
    );
    // Carries a real <tableParts r:id> reference, but this fixture deliberately
    // never adds a xl/worksheets/_rels/sheet1.xml.rels entry at all -- simulating
    // whatever future state could make output_rels_name absent from `passthrough`.
    const SHEET1_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>\n",
        "<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>\n</worksheet>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", WORKBOOK_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(&mut zip, "xl/worksheets/sheet1.xml", SHEET1_XML.as_bytes());
    let fixture_bytes = zip.finish().unwrap().into_inner();

    let source_path = tmp_path("source_table_parts_no_rels.xlsx");
    let output_path = tmp_path("output_table_parts_no_rels.xlsx");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    let prog = parser::parse("Sub NoOp()\n    n = 1\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "NoOp").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        !out_sheet1.contains("tableParts"),
        "must never emit a tableParts r:id reference whose own .rels file didn't survive: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

/// 0.10.0-D: `<drawing r:id>` (chart/image anchor). fixture5's only worksheet-level
/// relationship is its drawing -- `check_source_references()` goes fully `CLEAN` for this
/// fixture once this restores, the second real fixture (after fixture3's tableParts) to
/// reach that state.
#[test]
fn real_excel_drawing_survives_a_save() {
    let source_path = real_fixture("fixture5_chart_image_freeze_print.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_sheet1 =
        String::from_utf8(fixture_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        source_sheet1.contains(r#"<drawing r:id="rId1"/>"#),
        "fixture no longer contains the expected drawing shape -- test needs updating: {source_sheet1}"
    );

    let output_path = tmp_path("drawing_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditA1()\n    Range(\"A1\").Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditA1").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(r#"<drawing r:id="rId1"/>"#),
        "drawing must survive a save verbatim, referencing the same rId as the source: {out_sheet1}"
    );
    assert_eq!(
        output_entries.get("xl/worksheets/_rels/sheet1.xml.rels"),
        fixture_entries.get("xl/worksheets/_rels/sheet1.xml.rels"),
        "the worksheet .rels the restored r:id points at must still be byte-identical"
    );
    assert_eq!(
        output_entries.get("xl/drawings/drawing1.xml"),
        fixture_entries.get("xl/drawings/drawing1.xml"),
        "the drawing part itself must still be byte-identical"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// Plain (relationship-free) `<pageSetup>` -- fixture5's real shape (`paperSize`/
/// `orientation`/`horizontalDpi`/`verticalDpi`, no `r:id`). Unlike every other 0.10.0-D
/// element in this file, this one is NOT relationship-backed and needs no `rels_survived`
/// gate to restore safely -- `CT_PageSetup` genuinely CAN carry an `r:id` per the real
/// XSD, but this fixture's copy doesn't, so plain opaque-fragment passthrough (same
/// mechanism as `<pageMargins>`) is correct and safe here.
#[test]
fn real_excel_page_setup_survives_a_save() {
    let source_path = real_fixture("fixture5_chart_image_freeze_print.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_sheet1 =
        String::from_utf8(fixture_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        source_sheet1.contains(
            r#"<pageSetup paperSize="9" orientation="portrait" horizontalDpi="0" verticalDpi="0"/>"#
        ),
        "fixture no longer contains the expected pageSetup shape -- test needs updating: {source_sheet1}"
    );

    let output_path = tmp_path("page_setup_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditA1()\n    Range(\"A1\").Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditA1").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(
            r#"<pageSetup paperSize="9" orientation="portrait" horizontalDpi="0" verticalDpi="0"/>"#
        ),
        "pageSetup must survive a save verbatim: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// A real Excel-authored error-typed cell (`t="e"`, e.g. `#VALUE!`) round-trips as an
/// actual error, not a plain string -- ROADMAP.md Known gaps item 14, found live during
/// 0.10.0-C's real-Excel verification (fixture5's D8), fixed by threading
/// `SheetCell::Error`/`ExcelError` through the reader/`Vm`/writer the same way
/// `Variant::Error` already is at the VBA-runtime level. Before this fix, `xlsx_parse_cell`
/// treated `t="e"` identically to `t="str"` (both became `SheetCell::Str`), so the cell
/// round-tripped as `t="s"` with the error text as an ordinary shared string -- readable in
/// Excel, but no longer an error-typed cell underneath.
#[test]
fn real_excel_error_cell_survives_a_save_as_a_real_error_not_a_string() {
    let source_path = real_fixture("fixture5_chart_image_freeze_print.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_sheet1 =
        String::from_utf8(fixture_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        source_sheet1.contains(r#"<c r="D8" t="e" vm="1"><v>#VALUE!</v></c>"#),
        "fixture no longer contains the expected D8 error-cell shape -- test needs updating: {source_sheet1}"
    );

    let output_path = tmp_path("error_cell_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditA1()\n    Range(\"A1\").Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditA1").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(r#"<c r="D8" t="e"><v>#VALUE!</v></c>"#),
        "D8 must round-trip as a real t=\"e\" error cell, literal value (not shared-string \
         indexed, matching real Excel's own shape -- vm=\"1\", value metadata, is a \
         disclosed, accepted loss, not restored by this fix): {out_sheet1}"
    );

    let out_shared_strings = output_entries
        .get("xl/sharedStrings.xml")
        .map(|b| String::from_utf8(b.clone()).unwrap())
        .unwrap_or_default();
    assert!(
        !out_shared_strings.contains("#VALUE!"),
        "an error cell's text must never be shared-string indexed, matching real Excel's \
         own sharedStrings.xml (confirmed empty of \"#VALUE!\" even in the source fixture): \
         {out_shared_strings}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// Synthetic negative guard: `CT_PageSetup` genuinely can carry an `r:id` per the real
/// XSD even though no real fixture in this repo shows it -- an r:id-backed `<pageSetup>`
/// must NOT be restored (no `rels_survived` gate is wired up for it yet, so restoring one
/// would risk a dangling reference the moment a real fixture with this shape exists).
#[test]
fn page_setup_with_an_rid_is_not_restored() {
    const WORKBOOK_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n</sheets>\n</workbook>\n",
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "</Relationships>\n",
    );
    const SHEET1_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>\n",
        "<pageSetup paperSize=\"9\" r:id=\"rId1\"/>\n</worksheet>\n",
    );
    const SHEET1_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings\" Target=\"../printerSettings/printerSettings1.bin\"/>\n",
        "</Relationships>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", WORKBOOK_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(&mut zip, "xl/worksheets/sheet1.xml", SHEET1_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/worksheets/_rels/sheet1.xml.rels",
        SHEET1_RELS.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/printerSettings/printerSettings1.bin",
        b"not-real-printer-settings",
    );
    let fixture_bytes = zip.finish().unwrap().into_inner();

    let source_path = tmp_path("source_page_setup_rid.xlsx");
    let output_path = tmp_path("output_page_setup_rid.xlsx");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    let prog = parser::parse("Sub NoOp()\n    n = 1\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "NoOp").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        !out_sheet1.contains("pageSetup"),
        "must never emit an r:id-backed pageSetup reference -- no rels_survived gate is \
         wired up for it yet: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

/// 0.10.0-D: `<legacyDrawing r:id>` (VML comment shapes). fixture4's `.rels` also
/// carries an r:id-backed hyperlink -- deliberately left un-asserted here, still out of
/// scope (0.10.0-D's hyperlinks slice, not yet done).
#[test]
fn real_excel_legacy_drawing_survives_a_save() {
    let source_path = real_fixture("fixture4_hyperlink_comment_name.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_sheet1 =
        String::from_utf8(fixture_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        source_sheet1.contains(r#"<legacyDrawing r:id="rId2"/>"#),
        "fixture no longer contains the expected legacyDrawing shape -- test needs updating: {source_sheet1}"
    );

    let output_path = tmp_path("legacy_drawing_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditB1()\n    Cells(1, 2).Value = 999\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditB1").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(r#"<legacyDrawing r:id="rId2"/>"#),
        "legacyDrawing must survive a save verbatim, referencing the same rId as the source: {out_sheet1}"
    );
    assert_eq!(
        output_entries.get("xl/worksheets/_rels/sheet1.xml.rels"),
        fixture_entries.get("xl/worksheets/_rels/sheet1.xml.rels"),
        "the worksheet .rels the restored r:id points at must still be byte-identical"
    );
    assert_eq!(
        output_entries.get("xl/drawings/vmlDrawing1.vml"),
        fixture_entries.get("xl/drawings/vmlDrawing1.vml"),
        "the VML drawing part itself must still be byte-identical"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// 0.10.0-C slices C1+C2: workbook-level `<workbookPr>`/`<bookViews>`/`<calcPr>`/
/// `<extLst>`, plus the root `<workbook>` tag's own namespace declarations. `bookViews`'
/// `<workbookView>` carries `xr2:uid`, which genuinely needs the root's `xmlns:xr2`
/// declaration (unlike `extLst`'s `x15:`/`xcalcf:`/`xlwcv:` children, which carry their
/// own inline `xmlns:` redeclarations) -- see
/// `real_excel_sheetpr_and_data_validations_survive_a_save`'s `xr:uid` case for the same
/// shape on a worksheet element. `<definedNames>`/`<xr:revisionPtr>` are deliberately
/// NOT asserted here -- still lost, correctly out of scope (C3, or never-in-scope for
/// revisionPtr).
#[test]
fn real_excel_workbook_metadata_survives_a_save() {
    let source_path = real_fixture("fixture4_hyperlink_comment_name.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_wb = String::from_utf8(fixture_entries["xl/workbook.xml"].clone()).unwrap();
    for needle in [
        r#"<workbookPr codeName="ThisWorkbook" defaultThemeVersion="202300"/>"#,
        r#"<calcPr calcId="181029"/>"#,
        "<extLst>",
        "x15:workbookPr chartTrackingRefBase=\"1\"",
        r#"xr2:uid="{61125CCB-7611-4B43-B4CF-1525CE3D0920}""#,
    ] {
        assert!(
            source_wb.contains(needle),
            "fixture no longer contains {needle:?} -- test needs updating: {source_wb}"
        );
    }

    let output_path = tmp_path("workbook_metadata_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditCell()\n    Cells(1, 1).Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditCell").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_wb = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();

    for needle in [
        r#"<workbookPr codeName="ThisWorkbook" defaultThemeVersion="202300"/>"#,
        r#"<calcPr calcId="181029"/>"#,
        "x15:workbookPr chartTrackingRefBase=\"1\"",
        "xcalcf:calcFeatures",
        "xlwcv:version setVersion=\"2\"",
        r#"xr2:uid="{61125CCB-7611-4B43-B4CF-1525CE3D0920}""#,
    ] {
        assert!(
            out_wb.contains(needle),
            "{needle:?} must survive a save verbatim: {out_wb}"
        );
    }
    assert!(
        out_wb.contains(
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
        ),
        "root <workbook> tag must still declare xmlns:r -- <sheet r:id=...> depends on it: {out_wb}"
    );
    assert!(
        out_wb.contains(
            "xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\""
        ),
        "root <workbook> tag must still declare xmlns:xr2 -- bookViews' xr2:uid depends on it: {out_wb}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// 0.10.0-C slice 3 (C3): `<definedNames>` survives verbatim when a plain edit touches
/// no sheet. fixture5's real print-area defined name (`_xlnm.Print_Area`,
/// `localSheetId="0"`) is the fixture evidence -- single-sheet, so this test only
/// exercises the "no delete happened" branch; the delete-triggers-a-drop branch is
/// covered by `defined_names_are_dropped_entirely_once_a_sheet_is_deleted` below using
/// a synthetic multi-sheet fixture (fixture5 can't exercise it: deleting its only sheet
/// isn't representable the same way).
#[test]
fn real_excel_defined_names_survive_a_save_when_no_sheet_is_deleted() {
    let source_path = real_fixture("fixture5_chart_image_freeze_print.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_wb = String::from_utf8(fixture_entries["xl/workbook.xml"].clone()).unwrap();
    let needle =
        r#"<definedName name="_xlnm.Print_Area" localSheetId="0">Sheet1!$E$3</definedName>"#;
    assert!(
        source_wb.contains(needle),
        "fixture no longer contains {needle:?} -- test needs updating: {source_wb}"
    );

    let output_path = tmp_path("defined_names_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditCell()\n    Cells(1, 1).Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditCell").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_wb = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();
    assert!(
        out_wb.contains(needle),
        "definedNames must survive verbatim when no sheet was deleted: {out_wb}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// The delete-triggers-a-drop branch: a `<definedName>`'s `localSheetId` is a 0-based
/// index into `<sheets>`, so once `Sheets(...).Delete` runs, every remaining
/// localSheetId could point at a different sheet than it originally meant --
/// `<definedNames>` must be dropped entirely rather than carried through stale. No real
/// fixture has more than one sheet with a defined name, so this is a hand-built
/// synthetic fixture (two sheets, one workbook-scoped name, one sheet-scoped name
/// pointing at the sheet that gets deleted), matching this file's established pattern
/// for shapes no real fixture demonstrates (see e.g.
/// `passthrough_part_referenced_only_by_a_non_writer_owned_relationship_type_keeps_its_relationship`).
///
/// Deletes "Sheet2", not "Sheet1" -- `load_workbook_file` sets the *first* sheet
/// (Sheet1) active, and `Stmt::SheetsDelete`'s handler silently no-ops when the target
/// is the active sheet (`if key != self.active_sheet`). Deleting Sheet1 here would
/// still make this test pass (no-op delete -> defined_names correctly survives -- a
/// coincidentally-right result reached for the wrong reason), so it would silently stop
/// exercising the actual drop path this test exists to cover. Keep the target off the
/// active sheet if this ever changes.
#[test]
fn defined_names_are_dropped_entirely_once_a_sheet_is_deleted() {
    const WORKBOOK_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n",
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n",
        "<sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>\n",
        "</sheets>\n",
        "<definedNames>",
        "<definedName name=\"test\">Sheet1!$F$5</definedName>",
        "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"1\">Sheet2!$E$3</definedName>",
        "</definedNames>\n</workbook>\n",
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>\n",
        "</Relationships>\n",
    );
    const MINIMAL_SHEET: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>\n</worksheet>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", WORKBOOK_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet1.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet2.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    let fixture_bytes = zip.finish().unwrap().into_inner();

    let source_path = tmp_path("source_defined_names_delete.xlsx");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    // (1) No delete: definedNames survives verbatim.
    let noop_output_path = tmp_path("output_defined_names_noop.xlsx");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    save_workbook(&vm, &noop_output_path).expect("save should succeed");
    let noop_bytes = std::fs::read(&noop_output_path).unwrap();
    let noop_entries = read_all_zip_entries(&noop_bytes);
    let noop_wb = String::from_utf8(noop_entries["xl/workbook.xml"].clone()).unwrap();
    assert!(
        noop_wb.contains("<definedNames>") && noop_wb.contains("_xlnm.Print_Area"),
        "definedNames must survive verbatim with no delete: {noop_wb}"
    );

    // (2) Sheet2 (the one the sheet-scoped defined name points at) gets deleted --
    // definedNames must be entirely absent from the output, not partially pruned or
    // carried through stale.
    let delete_output_path = tmp_path("output_defined_names_deleted.xlsx");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    let prog = parser::parse("Sub DeleteIt()\n    Sheets(\"Sheet2\").Delete\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "DeleteIt").expect("macro should run");
    save_workbook(&vm, &delete_output_path).expect("save should succeed");
    let delete_bytes = std::fs::read(&delete_output_path).unwrap();
    let delete_entries = read_all_zip_entries(&delete_bytes);
    let delete_wb = String::from_utf8(delete_entries["xl/workbook.xml"].clone()).unwrap();
    assert!(
        !delete_wb.contains("<definedNames>"),
        "definedNames must be dropped entirely once a sheet is deleted, not carried \
         through with a stale localSheetId: {delete_wb}"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&noop_output_path);
    let _ = std::fs::remove_file(&delete_output_path);
}

/// P1 core 3: builds a minimal synthetic 3-sheet workbook, matching this file's
/// established pattern for shapes no real fixture demonstrates (real fixtures are
/// all single-sheet or lack a name distinct enough to test reordering safely).
fn synthetic_three_sheet_workbook(source_name: &str, defined_names_xml: &str) -> String {
    let workbook_xml = format!(
        concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
            "<sheets>\n",
            "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n",
            "<sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>\n",
            "<sheet name=\"Sheet3\" sheetId=\"3\" r:id=\"rId3\"/>\n",
            "</sheets>\n",
            "{}",
            "</workbook>\n",
        ),
        defined_names_xml
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>\n",
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/>\n",
        "</Relationships>\n",
    );
    const MINIMAL_SHEET: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>\n</worksheet>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", workbook_xml.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet1.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet2.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet3.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    let bytes = zip.finish().unwrap().into_inner();
    let path = tmp_path(source_name);
    std::fs::write(&path, &bytes).unwrap();
    path
}

#[test]
fn rename_sheet_preserves_tab_position_in_a_synthetic_three_sheet_workbook() {
    let source_path =
        synthetic_three_sheet_workbook("synthetic_three_sheet_source_rename.xlsx", "");
    let output_path = tmp_path("rename_sheet_synthetic_output.xlsx");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("synthetic fixture should load");
    vm.rename_sheet("Sheet2", "Renamed").unwrap();
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_wb = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();
    let sheet1_pos = out_wb.find("<sheet name=\"Sheet1\"").unwrap();
    let renamed_pos = out_wb.find("<sheet name=\"Renamed\"").unwrap();
    let sheet3_pos = out_wb.find("<sheet name=\"Sheet3\"").unwrap();
    assert!(
        sheet1_pos < renamed_pos && renamed_pos < sheet3_pos,
        "the renamed sheet must stay in the MIDDLE tab position, not move to the end: {out_wb}"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

#[test]
fn move_sheet_reorders_tabs_in_a_synthetic_three_sheet_workbook() {
    let source_path = synthetic_three_sheet_workbook("synthetic_three_sheet_source_move.xlsx", "");
    let output_path = tmp_path("move_sheet_synthetic_output.xlsx");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("synthetic fixture should load");
    vm.move_sheet("Sheet3", 0).unwrap();
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_wb = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();
    let sheet3_pos = out_wb.find("<sheet name=\"Sheet3\"").unwrap();
    let sheet1_pos = out_wb.find("<sheet name=\"Sheet1\"").unwrap();
    let sheet2_pos = out_wb.find("<sheet name=\"Sheet2\"").unwrap();
    assert!(
        sheet3_pos < sheet1_pos && sheet1_pos < sheet2_pos,
        "Sheet3 must now come first, with Sheet1/Sheet2 following in their original \
         relative order: {out_wb}"
    );

    // Each sheet's own worksheet part is unaffected by a pure reorder -- move_sheet
    // only touches sheet_order, never worksheet_origins/part naming.
    assert!(output_entries.contains_key("xl/worksheets/sheet1.xml"));
    assert!(output_entries.contains_key("xl/worksheets/sheet2.xml"));
    assert!(output_entries.contains_key("xl/worksheets/sheet3.xml"));

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

/// Pins the src/lib.rs `<definedNames>`-gate fix: a `<definedName localSheetId="N">`
/// is positional, so `move_sheet` reordering `sheet_order` must drop any surviving
/// `<definedNames>` passthrough exactly like a sheet deletion already does --
/// otherwise a saved workbook could carry a defined name silently pointing at the
/// wrong sheet after a reorder.
#[test]
fn move_sheet_drops_defined_names_that_would_have_stale_positional_indices() {
    let source_path = synthetic_three_sheet_workbook(
        "synthetic_three_sheet_source_move_defined_names.xlsx",
        concat!(
            "<definedNames>",
            "<definedName name=\"test\">Sheet1!$F$5</definedName>",
            "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"2\">Sheet3!$E$3</definedName>",
            "</definedNames>\n",
        ),
    );
    let output_path = tmp_path("move_sheet_defined_names_output.xlsx");

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("synthetic fixture should load");
    vm.move_sheet("Sheet3", 0).unwrap();
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_wb = String::from_utf8(output_entries["xl/workbook.xml"].clone()).unwrap();
    assert!(
        !out_wb.contains("<definedNames>"),
        "definedNames must be dropped entirely once move_sheet has reordered sheet_order, \
         not carried through with a stale localSheetId: {out_wb}"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

/// P2: sheet visibility -- same synthetic 3-sheet shape as
/// `synthetic_three_sheet_workbook`, but with each sheet's `<sheet state="...">`
/// attribute settable. No real fixture in this repo has a hidden/veryHidden sheet
/// (see docs/openpyxl-gap-audit.md), so this is the only way to exercise the reader
/// against real `state="hidden"`/`state="veryHidden"` XML shapes -- a separate
/// helper rather than adding a parameter to `synthetic_three_sheet_workbook`, which
/// already has 4+ unrelated call sites that would all need updating for a shape they
/// don't care about.
fn synthetic_three_sheet_workbook_with_states(
    source_name: &str,
    states: [Option<&str>; 3],
) -> String {
    let attr = |s: Option<&str>| match s {
        Some(v) => format!(" state=\"{}\"", v),
        None => String::new(),
    };
    let workbook_xml = format!(
        concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
            "<sheets>\n",
            "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"{}/>\n",
            "<sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"{}/>\n",
            "<sheet name=\"Sheet3\" sheetId=\"3\" r:id=\"rId3\"{}/>\n",
            "</sheets>\n",
            "</workbook>\n",
        ),
        attr(states[0]),
        attr(states[1]),
        attr(states[2]),
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>\n",
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/>\n",
        "</Relationships>\n",
    );
    const MINIMAL_SHEET: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>\n</worksheet>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", workbook_xml.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet1.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet2.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    zip_add(
        &mut zip,
        "xl/worksheets/sheet3.xml",
        MINIMAL_SHEET.as_bytes(),
    );
    let bytes = zip.finish().unwrap().into_inner();
    let path = tmp_path(source_name);
    std::fs::write(&path, &bytes).unwrap();
    path
}

#[test]
fn sheet_state_reads_hidden_and_very_hidden_from_a_synthetic_fixture() {
    let source_path = synthetic_three_sheet_workbook_with_states(
        "synthetic_three_sheet_source_states.xlsx",
        [None, Some("hidden"), Some("veryHidden")],
    );

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("synthetic fixture should load");

    assert_eq!(vm.sheet_state("Sheet1").unwrap(), SheetState::Visible);
    assert_eq!(vm.sheet_state("Sheet2").unwrap(), SheetState::Hidden);
    assert_eq!(vm.sheet_state("Sheet3").unwrap(), SheetState::VeryHidden);

    let _ = std::fs::remove_file(&source_path);
}

#[test]
fn rename_sheet_preserves_hidden_state_on_a_synthetic_fixture() {
    let source_path = synthetic_three_sheet_workbook_with_states(
        "synthetic_three_sheet_source_states_rename.xlsx",
        [None, Some("hidden"), None],
    );

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("synthetic fixture should load");
    vm.rename_sheet("Sheet2", "Renamed").unwrap();

    assert_eq!(vm.sheet_state("Renamed").unwrap(), SheetState::Hidden);

    let _ = std::fs::remove_file(&source_path);
}

/// P2: row height / column width -- single-sheet synthetic fixture with a custom
/// row height and column width, matching this file's own established pattern for
/// shapes no real fixture demonstrates (see `synthetic_three_sheet_workbook_with_states`
/// above). `<row r="5" ht="30.5" customHeight="1">` and `<col min="2" max="4"
/// width="12.5" customWidth="1"/>`.
fn synthetic_sheet_with_row_heights_and_column_widths(source_name: &str) -> String {
    const WORKBOOK_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n",
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n",
        "</sheets>\n",
        "</workbook>\n",
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "</Relationships>\n",
    );
    const SHEET_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<cols><col min=\"2\" max=\"4\" width=\"12.5\" customWidth=\"1\"/></cols>\n",
        "<sheetData>\n",
        "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>\n",
        "<row r=\"5\" ht=\"30.5\" customHeight=\"1\"><c r=\"A5\"><v>2</v></c></row>\n",
        "</sheetData>\n</worksheet>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", WORKBOOK_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(&mut zip, "xl/worksheets/sheet1.xml", SHEET_XML.as_bytes());
    let bytes = zip.finish().unwrap().into_inner();
    let path = tmp_path(source_name);
    std::fs::write(&path, &bytes).unwrap();
    path
}

#[test]
fn row_height_and_column_width_read_from_a_synthetic_fixture() {
    let source_path = synthetic_sheet_with_row_heights_and_column_widths(
        "synthetic_row_height_column_width.xlsx",
    );

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("synthetic fixture should load");

    assert_eq!(vm.row_height_on_sheet("sheet1", 5), Some(30.5));
    assert_eq!(vm.row_height_on_sheet("sheet1", 1), None);
    assert_eq!(vm.column_width_on_sheet("sheet1", 3), Some(12.5));
    assert_eq!(vm.column_width_on_sheet("sheet1", 1), None);

    let _ = std::fs::remove_file(&source_path);
}

#[test]
fn copy_sheet_preserves_row_height_and_column_width_on_a_synthetic_fixture() {
    let source_path = synthetic_sheet_with_row_heights_and_column_widths(
        "synthetic_row_height_column_width_copy.xlsx",
    );

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("synthetic fixture should load");
    vm.copy_sheet("Sheet1", "Copy").unwrap();

    assert_eq!(vm.row_height_on_sheet("copy", 5), Some(30.5));
    assert_eq!(vm.column_width_on_sheet("copy", 3), Some(12.5));

    let _ = std::fs::remove_file(&source_path);
}

/// 0.10.0-D, slice D1: a surviving sheet's output part name stays its own origin
/// (`WorksheetOrigin.original_part_name`), not renumbered by output position. Three
/// sheets, no VBA; Sheet3 (last, with a real worksheet-level relationship) is the one
/// that must NOT get renumbered when Sheet2 -- an earlier, unrelated, relationship-free
/// sheet -- is deleted, shifting Sheet3 from position 3 to position 2.
///
/// This is a real, previously-reproduced bug, not a hypothetical: before D1, the
/// surviving worksheet content was written to the position-derived `sheet2.xml`, while
/// `xl/worksheets/_rels/sheet3.xml.rels` (which passes through keyed by its ORIGINAL
/// path, untouched by this fix) stayed at `sheet3.xml` -- orphaning the `.rels` file and
/// leaving the real `sheet2.xml` content with no relationship file at all, even though
/// its original content had one. Confirmed by running this exact scenario against the
/// pre-D1 code before writing this test.
#[test]
fn surviving_sheets_keep_their_own_origin_part_name_after_an_earlier_sheet_is_deleted() {
    const WORKBOOK_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n",
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n",
        "<sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>\n",
        "<sheet name=\"Sheet3\" sheetId=\"3\" r:id=\"rId3\"/>\n",
        "</sheets>\n</workbook>\n",
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>\n",
        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet3.xml\"/>\n",
        "</Relationships>\n",
    );
    const PLAIN_SHEET: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>\n</worksheet>\n",
    );
    const SHEET3_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>3</v></c></row></sheetData>\n",
        "<hyperlinks><hyperlink ref=\"A1\" r:id=\"hlink1\"/></hyperlinks>\n</worksheet>\n",
    );
    const SHEET3_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"hlink1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/\" TargetMode=\"External\"/>\n",
        "</Relationships>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", WORKBOOK_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(&mut zip, "xl/worksheets/sheet1.xml", PLAIN_SHEET.as_bytes());
    zip_add(&mut zip, "xl/worksheets/sheet2.xml", PLAIN_SHEET.as_bytes());
    zip_add(&mut zip, "xl/worksheets/sheet3.xml", SHEET3_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/worksheets/_rels/sheet3.xml.rels",
        SHEET3_RELS.as_bytes(),
    );
    let fixture_bytes = zip.finish().unwrap().into_inner();

    let source_path = tmp_path("source_d1_reorder.xlsx");
    let output_path = tmp_path("output_d1_reorder.xlsx");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    let prog =
        parser::parse("Sub DeleteSheet2()\n    Sheets(\"Sheet2\").Delete\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "DeleteSheet2").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);

    assert!(
        output_entries.contains_key("xl/worksheets/sheet3.xml"),
        "Sheet3's content must stay at its own origin part name, sheet3.xml, even though \
         it's now the second (not third) sheet in output order"
    );
    assert!(
        !output_entries.contains_key("xl/worksheets/sheet2.xml"),
        "sheet2.xml must not exist -- Sheet3's content must not be renumbered into it"
    );
    assert!(
        output_entries.contains_key("xl/worksheets/_rels/sheet3.xml.rels"),
        "the passthrough .rels file must still be at sheet3.xml's own path"
    );
    let sheet3_xml = String::from_utf8(output_entries["xl/worksheets/sheet3.xml"].clone()).unwrap();
    assert!(
        sheet3_xml.contains("<c r=\"A1\"><v>3</v></c>"),
        "sheet3.xml must actually contain Sheet3's own cell data, not Sheet1's or an \
         empty regenerated sheet: {sheet3_xml}"
    );

    let wb_rels = String::from_utf8(output_entries["xl/_rels/workbook.xml.rels"].clone()).unwrap();
    assert!(
        wb_rels.contains("Target=\"worksheets/sheet3.xml\""),
        "workbook.xml.rels must point the surviving sheet's relationship at sheet3.xml, \
         not a stale/renumbered target: {wb_rels}"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

/// 0.10.0-D4: deleting a sheet must prune its EXCLUSIVELY-reachable target parts (its own
/// worksheet `.rels`, and whatever only IT points at) while leaving anything SHARED with a
/// surviving sheet alone. Sheet1 (survives) and Sheet2 (deleted) both reference the same
/// `xl/tables/shared_table.xml` via their own `<tableParts>`; Sheet2 additionally
/// references `xl/drawings/exclusive_drawing1.xml`, which nothing else points at.
#[test]
fn deleting_a_sheet_prunes_its_exclusive_targets_but_keeps_shared_ones() {
    const WORKBOOK_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheets>\n",
        "<sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n",
        "<sheet name=\"Sheet2\" sheetId=\"2\" r:id=\"rId2\"/>\n",
        "</sheets>\n</workbook>\n",
    );
    const WORKBOOK_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n",
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>\n",
        "</Relationships>\n",
    );
    const SHEET1_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>\n",
        "<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>\n</worksheet>\n",
    );
    const SHEET1_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/shared_table.xml\"/>\n",
        "</Relationships>\n",
    );
    const SHEET2_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n",
        "<sheetData><row r=\"1\"><c r=\"A1\"><v>2</v></c></row></sheetData>\n",
        "<tableParts count=\"1\"><tablePart r:id=\"rId1\"/></tableParts>\n",
        "<drawing r:id=\"rId2\"/>\n</worksheet>\n",
    );
    const SHEET2_RELS: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/shared_table.xml\"/>\n",
        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/exclusive_drawing1.xml\"/>\n",
        "</Relationships>\n",
    );
    const DRAWING_XML: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
        "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"/>\n",
    );

    let cursor = Cursor::new(Vec::<u8>::new());
    let mut zip = ZipWriter::new(cursor);
    zip_add(
        &mut zip,
        "[Content_Types].xml",
        CONTENT_TYPES_NO_VBA.as_bytes(),
    );
    zip_add(&mut zip, "_rels/.rels", ROOT_RELS.as_bytes());
    zip_add(&mut zip, "xl/workbook.xml", WORKBOOK_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/_rels/workbook.xml.rels",
        WORKBOOK_RELS.as_bytes(),
    );
    zip_add(&mut zip, "xl/worksheets/sheet1.xml", SHEET1_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/worksheets/_rels/sheet1.xml.rels",
        SHEET1_RELS.as_bytes(),
    );
    zip_add(&mut zip, "xl/worksheets/sheet2.xml", SHEET2_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/worksheets/_rels/sheet2.xml.rels",
        SHEET2_RELS.as_bytes(),
    );
    zip_add(&mut zip, "xl/tables/shared_table.xml", TABLE_XML.as_bytes());
    zip_add(
        &mut zip,
        "xl/drawings/exclusive_drawing1.xml",
        DRAWING_XML.as_bytes(),
    );
    let fixture_bytes = zip.finish().unwrap().into_inner();

    let source_path = tmp_path("source_d4_prune.xlsx");
    let output_path = tmp_path("output_d4_prune.xlsx");
    std::fs::write(&source_path, &fixture_bytes).unwrap();

    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("fixture should load");
    let prog =
        parser::parse("Sub DeleteSheet2()\n    Sheets(\"Sheet2\").Delete\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "DeleteSheet2").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);

    assert!(
        !output_entries.contains_key("xl/worksheets/sheet2.xml"),
        "sheet2.xml must not exist -- Sheet2 was deleted"
    );
    assert!(
        !output_entries.contains_key("xl/worksheets/_rels/sheet2.xml.rels"),
        "Sheet2's own .rels is exclusively reachable from the deleted sheet and must be \
         pruned, not left behind as an orphan (ROADMAP.md Known gaps item 15)"
    );
    assert!(
        !output_entries.contains_key("xl/drawings/exclusive_drawing1.xml"),
        "exclusive_drawing1.xml is reachable ONLY via Sheet2's own .rels and must be \
         pruned along with it"
    );
    assert!(
        output_entries.contains_key("xl/tables/shared_table.xml"),
        "shared_table.xml is also referenced by surviving Sheet1 and must NOT be pruned \
         just because one of its two referencing sheets is gone"
    );
    assert!(
        output_entries.contains_key("xl/worksheets/_rels/sheet1.xml.rels"),
        "Sheet1's own .rels must survive untouched -- it isn't reachable from the \
         deleted sheet at all"
    );

    let _ = std::fs::remove_file(&source_path);
    let _ = std::fs::remove_file(&output_path);
}

/// 0.10.0-B slice 4 (B4): internal (location=, relationship-free) hyperlinks.
/// fixture6_internal_hyperlink.xlsm has exactly one, no r:id.
#[test]
fn real_excel_internal_hyperlink_survives_a_save() {
    let source_path = real_fixture("fixture6_internal_hyperlink.xlsm");
    let output_path = tmp_path("internal_hyperlink_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditA1()\n    Range(\"A1\").Value = 42\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditA1").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(
            r#"<hyperlink ref="A1" location="Sheet2!B2" display="Sheet2!B2" xr:uid="{7239724E-8623-EB4C-A548-F5CFD578FC11}"/>"#
        ),
        "internal hyperlink must survive a save verbatim: {out_sheet1}"
    );
    assert!(
        out_sheet1.contains("<hyperlinks>") && out_sheet1.contains("</hyperlinks>"),
        "hyperlinks container must be present: {out_sheet1}"
    );

    let _ = std::fs::remove_file(&output_path);
}

/// 0.10.0-D hyperlinks slice: supersedes the old 0.10.0-B4 behavior of this exact
/// fixture. fixture4's only hyperlink is r:id-backed (external URL) -- B4 deliberately
/// excluded it (no relationship-graph reconnection existed yet); now that
/// `rels_survived` gates it, it must round-trip like every other restored r:id
/// reference. This also clears fixture4's LAST `SOURCE_REFERENCE_LOSS` violation
/// (legacyDrawing was fixed by the previous commit) -- fixture4 is now fully CLEAN.
#[test]
fn real_excel_external_hyperlink_survives_a_save() {
    let source_path = real_fixture("fixture4_hyperlink_comment_name.xlsm");
    let fixture_bytes = std::fs::read(&source_path).expect("real fixture must exist");
    let fixture_entries = read_all_zip_entries(&fixture_bytes);
    let source_sheet1 =
        String::from_utf8(fixture_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        source_sheet1.contains(
            r#"<hyperlink ref="D6" r:id="rId1" xr:uid="{CEF97160-724D-564B-8AAF-8C73BDCBFE82}"/>"#
        ),
        "fixture no longer contains the expected hyperlink shape -- test needs updating: {source_sheet1}"
    );

    let output_path = tmp_path("external_hyperlink_output.xlsm");
    let mut vm = Vm::new();
    vm.load_workbook_file(&source_path)
        .expect("real fixture should load");
    let prog = parser::parse("Sub EditB1()\n    Cells(1, 2).Value = 999\nEnd Sub\n").unwrap();
    vm.run_sub(&prog, "EditB1").expect("macro should run");
    save_workbook(&vm, &output_path).expect("save-as should succeed");

    let output_bytes = std::fs::read(&output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);
    let out_sheet1 = String::from_utf8(output_entries["xl/worksheets/sheet1.xml"].clone()).unwrap();

    assert!(
        out_sheet1.contains(
            r#"<hyperlink ref="D6" r:id="rId1" xr:uid="{CEF97160-724D-564B-8AAF-8C73BDCBFE82}"/>"#
        ),
        "the r:id-backed hyperlink must survive a save verbatim, referencing the same \
         rId as the source: {out_sheet1}"
    );
    assert_eq!(
        output_entries.get("xl/worksheets/_rels/sheet1.xml.rels"),
        fixture_entries.get("xl/worksheets/_rels/sheet1.xml.rels"),
        "the worksheet .rels the restored r:id points at must still be byte-identical \
         (TargetMode=\"External\", so no target part to check -- just the relationship \
         entry itself)"
    );

    let _ = std::fs::remove_file(&output_path);
}

fn check_real_fixture_output(
    fixture_entries: &HashMap<String, Vec<u8>>,
    vba_bytes: &[u8],
    output_path: &str,
    mode: &str,
) {
    let output_bytes = std::fs::read(output_path).unwrap();
    let output_entries = read_all_zip_entries(&output_bytes);

    // (i) edited and unedited cells
    let sheets = reader::read_workbook(output_path).expect("output should be readable");
    let sheet = sheets
        .iter()
        .find(|s| s.name.to_lowercase() == "sheet1")
        .unwrap();
    assert!(
        matches!(
            sheet.cells.get(&(1, 2)),
            Some(reader::SheetCell::Integer(999))
        ),
        "[{mode}] B1 should be edited to 999"
    );
    assert!(
        matches!(sheet.cells.get(&(1, 1)), Some(reader::SheetCell::Str(s)) if s == "Counter"),
        "[{mode}] unedited A1 ('Counter') must survive"
    );

    // (ii) xl/vbaProject.bin byte-identical
    assert_eq!(
        output_entries.get("xl/vbaProject.bin"),
        Some(&vba_bytes.to_vec()),
        "[{mode}] vbaProject.bin must survive byte-identical from a real Excel-authored file"
    );

    // (iii) every non-writer-owned original part is byte-identical
    for (name, bytes) in fixture_entries {
        if is_writer_owned(name) {
            continue;
        }
        assert_eq!(
            output_entries.get(name),
            Some(bytes),
            "[{mode}] passthrough part {name} must be byte-identical"
        );
    }

    // (iv) content-types self-consistency + macro-enabled root type
    let ct_xml = String::from_utf8(output_entries["[Content_Types].xml"].clone()).unwrap();
    assert!(
        ct_xml.contains("macroEnabled.main+xml"),
        "[{mode}] workbook.xml must declare the macro-enabled content type"
    );
    for name in output_entries.keys() {
        if name == "[Content_Types].xml" {
            continue;
        }
        assert!(
            resolve_content_type(&ct_xml, name).is_some(),
            "[{mode}] output part {name} has no resolvable content type"
        );
    }

    // (v) relationship carry-over regression guard: theme and docProps
    // relationships (not just the parts' bytes) must survive -- the exact
    // bugs found and fixed via this fixture (see CHANGELOG.md's [0.9.0]).
    let workbook_rels =
        String::from_utf8(output_entries["xl/_rels/workbook.xml.rels"].clone()).unwrap();
    assert!(
        workbook_rels.contains("relationships/theme") && workbook_rels.contains("theme1.xml"),
        "[{mode}] theme relationship must survive: {workbook_rels}"
    );
    let root_rels = String::from_utf8(output_entries["_rels/.rels"].clone()).unwrap();
    assert!(
        root_rels.contains("core.xml") && root_rels.contains("app.xml"),
        "[{mode}] docProps relationships must survive: {root_rels}"
    );
}

/// 0.14.0-A2 integration: a cross-sheet formula reference survives a real
/// save/reload round trip after a structural edit on the sheet it targets,
/// both save-as and in-place, and a second (no-op) save does NOT shift it
/// again -- `insert_rows_on_sheet` is the only thing that ever calls the
/// rewriter, and it's called exactly once here.
///
/// The formula is inserted directly into the cell map, cached value
/// `Variant::Empty` (an unevaluated cross-sheet formula, matching what
/// `evaluate()` actually leaves it at -- see `eval::references_another_sheet`)
/// rather than through `set_cell_formula`, which also evaluates immediately
/// and would reject this formula outright. This exercises the actual
/// save/reload machinery a real Excel-authored file with such a formula
/// would go through -- `reader.rs` loads `<f>`/`<v>` independently of each
/// other and without evaluating anything, so a real file with this formula
/// would load exactly the same way. `Variant::Empty` here also exercises the
/// formula-with-empty-cached-value writer/reader fix (see
/// `formula_cell_with_empty_cached_value_survives_save_and_reload` below) --
/// this test would have needed a non-empty placeholder value to work around
/// that bug before it was fixed.
#[test]
fn cross_sheet_formula_reference_survives_a_real_save_reload_round_trip() {
    use elixcee::vm::{CellContent, Variant};

    let mut vm = Vm::new(); // default/active sheet key is "sheet1"
    vm.ensure_sheet("Sheet2");
    vm.set_active_sheet("Sheet2").unwrap();
    vm.cells_mut().insert(
        (1, 1),
        CellContent {
            formula: Some("=sheet1!A10+1".to_string()),
            value: Variant::Empty,
        },
    );
    vm.set_active_sheet("Sheet1").unwrap();

    // Insert 1 row before row 5 on sheet1 -- A10 (>= 5) must become A11.
    vm.insert_rows_on_sheet("sheet1", 5, 1);
    assert_eq!(
        vm.get_sheet_cells("sheet2")
            .unwrap()
            .get(&(1, 1))
            .unwrap()
            .formula,
        Some("=sheet1!A11+1".to_string()),
        "in-memory rewrite must have already happened before any save"
    );

    let save_as_path = tmp_path("cross_sheet_ref_save_as.xlsx");
    save_workbook(&vm, &save_as_path).expect("save-as should succeed");

    // Reload into a FRESH Vm (as a real user re-opening the file would) and
    // confirm the rewritten formula text persisted through the XLSX <f>
    // element, not just the in-memory struct.
    let mut reloaded = Vm::new();
    reloaded
        .load_workbook_file(&save_as_path)
        .expect("save-as output should reload");
    assert_eq!(
        reloaded
            .get_sheet_cells("sheet2")
            .unwrap()
            .get(&(1, 1))
            .unwrap()
            .formula,
        Some("sheet1!A11+1".to_string()), // no leading '=' -- matches <f> element text
        "rewritten formula must survive a save-as + reload round trip"
    );

    // A second, no-op save (no structural edit in between) must NOT shift
    // the reference again -- insert_rows_on_sheet is the only thing that
    // calls the rewriter, and it isn't called here.
    save_workbook(&reloaded, &save_as_path).expect("second save should succeed");
    let mut reloaded_again = Vm::new();
    reloaded_again
        .load_workbook_file(&save_as_path)
        .expect("twice-saved output should reload");
    assert_eq!(
        reloaded_again
            .get_sheet_cells("sheet2")
            .unwrap()
            .get(&(1, 1))
            .unwrap()
            .formula,
        Some("sheet1!A11+1".to_string()),
        "a plain re-save with no structural edit must not shift the reference again"
    );

    // In-place save (source == output), the other realistic CLI usage.
    let inplace_path = tmp_path("cross_sheet_ref_inplace.xlsx");
    std::fs::copy(&save_as_path, &inplace_path).unwrap();
    let mut inplace_vm = Vm::new();
    inplace_vm
        .load_workbook_file(&inplace_path)
        .expect("copy should reload");
    inplace_vm.insert_rows_on_sheet("sheet1", 1, 1); // A11 (>=1) -> A12
    save_workbook(&inplace_vm, &inplace_path).expect("in-place save should succeed");
    let mut reloaded_inplace = Vm::new();
    reloaded_inplace
        .load_workbook_file(&inplace_path)
        .expect("in-place output should reload");
    assert_eq!(
        reloaded_inplace
            .get_sheet_cells("sheet2")
            .unwrap()
            .get(&(1, 1))
            .unwrap()
            .formula,
        Some("sheet1!A12+1".to_string()),
        "in-place save must also persist the rewrite correctly"
    );

    let _ = std::fs::remove_file(&save_as_path);
    let _ = std::fs::remove_file(&inplace_path);
}

/// Sheet-rename follow-up integration: a qualified formula reference survives
/// a real save/reload round trip after the sheet it names is renamed, and
/// picks up the new sheet's exact display casing/quoting -- same rationale
/// for direct cell-map insertion (bypassing `set_cell_formula`) as the
/// structural-edit integration test above.
#[test]
fn sheet_rename_qualifier_rewrite_survives_a_real_save_reload_round_trip() {
    use elixcee::vm::{CellContent, Variant};

    let mut vm = Vm::new(); // default/active sheet key is "sheet1"
    vm.ensure_sheet("Sheet2");
    vm.set_active_sheet("Sheet2").unwrap();
    vm.cells_mut().insert(
        (1, 1),
        CellContent {
            formula: Some("=Sheet1!A10+1".to_string()),
            value: Variant::Empty,
        },
    );
    vm.set_active_sheet("Sheet1").unwrap();

    vm.rename_sheet("Sheet1", "Sales 2026").unwrap();
    assert_eq!(
        vm.get_sheet_cells("sheet2")
            .unwrap()
            .get(&(1, 1))
            .unwrap()
            .formula,
        Some("='Sales 2026'!A10+1".to_string()),
        "in-memory rewrite must have already happened before any save"
    );

    let path = tmp_path("sheet_rename_qualifier_rewrite.xlsx");
    save_workbook(&vm, &path).expect("save should succeed");

    let mut reloaded = Vm::new();
    reloaded
        .load_workbook_file(&path)
        .expect("save output should reload");
    assert_eq!(
        reloaded
            .get_sheet_cells("sheet2")
            .unwrap()
            .get(&(1, 1))
            .unwrap()
            .formula,
        Some("'Sales 2026'!A10+1".to_string()), // no leading '=' -- matches <f> element text
        "the renamed qualifier must survive a save + reload round trip"
    );
    // The renamed sheet's own tab name persisted too (not just the formula).
    assert!(
        reloaded.sheet_names().iter().any(|n| n == "sales 2026"),
        "the renamed sheet's key must still resolve after reload: {:?}",
        reloaded.sheet_names()
    );

    let _ = std::fs::remove_file(&path);
}

/// Pre-existing writer correctness bug, discovered while writing the 0.14.0-A2
/// integration test above (unrelated to reference rewriting -- reproduces for
/// an ordinary same-sheet formula with no cross-sheet reference at all).
/// `xlsx_cell_xml` (`src/lib.rs`) treated `Variant::Empty` as "nothing worth
/// writing" and skipped the WHOLE `<c>` element, formula text included, any
/// time a formula cell's cached value happened to be `Variant::Empty` --
/// silently dropping the formula on save. "No cached result" and "no formula"
/// are different things; a formula cell must never be silently dropped just
/// because it hasn't been (or can't be) evaluated.
///
/// Each case: build a formula-only cell (cached value `Variant::Empty`, as a
/// freshly-typed/not-yet-recalculated cell, or an unevaluable cross-sheet
/// reference, would have), save, confirm the raw XML still has a `<f>`
/// element (the cell wasn't dropped), then reload and confirm the formula
/// text survived exactly.
#[test]
fn formula_cell_with_empty_cached_value_survives_save_and_reload() {
    use elixcee::vm::{CellContent, Variant};

    let cases: &[&str] = &[
        "IF(FALSE,1)",     // the exact case found during 0.14.0-A2 review
        "IF(TRUE,\"\",1)", // string literal + escaping in the formula text
        "A2",              // a plain same-sheet reference
        "Sheet2!A1",       // cross-sheet: unevaluable, but the TEXT must still save
    ];

    for (idx, formula_body) in cases.iter().enumerate() {
        let mut vm = Vm::new();
        vm.ensure_sheet("Sheet2"); // somewhere for the Sheet2!A1 case to name
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: Some(format!("={formula_body}")),
                value: Variant::Empty,
            },
        );

        let path = tmp_path(&format!("formula_empty_cached_value_{idx}.xlsx"));
        save_workbook(&vm, &path).expect("save should succeed");

        let bytes = std::fs::read(&path).unwrap();
        let entries = read_all_zip_entries(&bytes);
        let sheet_xml = String::from_utf8(entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
        assert!(
            sheet_xml.contains("<f>"),
            "case {idx} (={formula_body}): the <c> element (and its <f>) must not be \
             dropped just because the cached value is Empty: {sheet_xml}"
        );

        let mut reloaded = Vm::new();
        reloaded
            .load_workbook_file(&path)
            .expect("save output should reload");
        assert_eq!(
            reloaded
                .get_sheet_cells("sheet1")
                .unwrap()
                .get(&(1, 1))
                .and_then(|c| c.formula.clone()),
            Some(formula_body.to_string()), // no leading '=' -- matches <f> element text
            "case {idx}: formula text must survive a save + reload round trip"
        );

        let _ = std::fs::remove_file(&path);
    }
}

/// Regression guard, matrix cases A/C/D/E/F from the bug above: a formula
/// cell with a REAL cached value (any `Variant` the writer already handled
/// correctly) must keep working exactly as before, and a plain empty cell
/// with NO formula must still be omitted entirely -- the fix must only add
/// the missing "formula present" case, not start emitting every empty cell.
#[test]
fn formula_cell_emission_matrix_around_the_empty_value_fix() {
    use elixcee::vm::{CellContent, ExcelError, Variant};

    let mut vm = Vm::new();
    // A: no formula, Empty value -- must still be omitted (no regression).
    vm.cells_mut().insert(
        (1, 1),
        CellContent {
            formula: None,
            value: Variant::Empty,
        },
    );
    // C: formula + Integer cached value.
    vm.cells_mut().insert(
        (2, 1),
        CellContent {
            formula: Some("=1+1".to_string()),
            value: Variant::Integer(2),
        },
    );
    // D: formula + String cached value.
    vm.cells_mut().insert(
        (3, 1),
        CellContent {
            formula: Some("=\"hi\"".to_string()),
            value: Variant::Str("hi".to_string()),
        },
    );
    // E: formula + Boolean cached value.
    vm.cells_mut().insert(
        (4, 1),
        CellContent {
            formula: Some("=TRUE".to_string()),
            value: Variant::Boolean(true),
        },
    );
    // F: formula + Error cached value.
    vm.cells_mut().insert(
        (5, 1),
        CellContent {
            formula: Some("=1/0".to_string()),
            value: Variant::Error(ExcelError::DivZero),
        },
    );

    let path = tmp_path("formula_emission_matrix.xlsx");
    save_workbook(&vm, &path).expect("save should succeed");

    let bytes = std::fs::read(&path).unwrap();
    let entries = read_all_zip_entries(&bytes);
    let sheet_xml = String::from_utf8(entries["xl/worksheets/sheet1.xml"].clone()).unwrap();
    assert!(
        !sheet_xml.contains("r=\"A1\""),
        "A: an empty, formula-less cell must still be omitted entirely: {sheet_xml}"
    );

    let mut reloaded = Vm::new();
    reloaded
        .load_workbook_file(&path)
        .expect("save output should reload");
    let cells = reloaded.get_sheet_cells("sheet1").unwrap();
    assert!(
        cells.get(&(1, 1)).is_none(),
        "A: still omitted after reload"
    );
    assert_eq!(
        cells.get(&(2, 1)).unwrap().value,
        Variant::Integer(2),
        "C: cached Integer value must round-trip unchanged"
    );
    assert_eq!(
        cells.get(&(3, 1)).unwrap().value,
        Variant::Str("hi".to_string()),
        "D: cached String value must round-trip unchanged"
    );
    assert_eq!(
        cells.get(&(4, 1)).unwrap().value,
        Variant::Boolean(true),
        "E: cached Boolean value must round-trip unchanged"
    );
    assert_eq!(
        cells.get(&(5, 1)).unwrap().value,
        Variant::Error(ExcelError::DivZero),
        "F: cached Error value must round-trip unchanged"
    );
    for (row, expected_formula) in [(2, "1+1"), (3, "\"hi\""), (4, "TRUE"), (5, "1/0")] {
        assert_eq!(
            cells.get(&(row, 1)).unwrap().formula,
            Some(expected_formula.to_string()),
            "row {row}: formula text must still round-trip unchanged"
        );
    }

    let _ = std::fs::remove_file(&path);
}

/// A formula cell with an empty cached value must survive not just one
/// save→reload, but repeated ones -- save, reload, save again (in place, the
/// realistic `--file foo.xlsx --output foo.xlsx` CLI usage), reload again.
/// Nothing here evaluates the formula, so there's no risk of the SECOND save
/// somehow computing a real value and papering over a regression in the fix
/// -- the cached value stays `Variant::Empty` throughout, on purpose.
#[test]
fn formula_cell_with_empty_cached_value_survives_two_consecutive_saves() {
    use elixcee::vm::{CellContent, Variant};

    let mut vm = Vm::new();
    vm.cells_mut().insert(
        (1, 1),
        CellContent {
            formula: Some("=IF(FALSE,1)".to_string()),
            value: Variant::Empty,
        },
    );

    let path = tmp_path("formula_empty_value_double_save.xlsx");
    save_workbook(&vm, &path).expect("first save should succeed");

    let mut reloaded_once = Vm::new();
    reloaded_once
        .load_workbook_file(&path)
        .expect("first save output should reload");
    assert_eq!(
        reloaded_once
            .get_sheet_cells("sheet1")
            .unwrap()
            .get(&(1, 1))
            .and_then(|c| c.formula.clone()),
        Some("IF(FALSE,1)".to_string()),
        "formula must survive the first save + reload"
    );

    save_workbook(&reloaded_once, &path).expect("second (in-place) save should succeed");

    let mut reloaded_twice = Vm::new();
    reloaded_twice
        .load_workbook_file(&path)
        .expect("second save output should reload");
    assert_eq!(
        reloaded_twice
            .get_sheet_cells("sheet1")
            .unwrap()
            .get(&(1, 1))
            .and_then(|c| c.formula.clone()),
        Some("IF(FALSE,1)".to_string()),
        "formula must still survive after a second consecutive save"
    );

    let _ = std::fs::remove_file(&path);
}
