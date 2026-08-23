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
use elixcee::{parser, reader, save_workbook, vm::Vm};
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

/// Negative guard, same slice: fixture4's only hyperlink is r:id-backed (external URL,
/// out of scope until 0.10.0-D) -- the output must NOT contain any `<hyperlinks>` element
/// at all, not an empty `<hyperlinks/>` (CT_Hyperlinks' <hyperlink> child is
/// minOccurs="1" -- an empty container would be invalid XML).
#[test]
fn real_excel_external_only_hyperlink_omits_the_hyperlinks_container_entirely() {
    let source_path = real_fixture("fixture4_hyperlink_comment_name.xlsm");
    let output_path = tmp_path("external_only_hyperlink_output.xlsm");
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
        !out_sheet1.contains("<hyperlinks"),
        "an all-r:id source must omit <hyperlinks> entirely, not emit an empty \
         container: {out_sheet1}"
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
