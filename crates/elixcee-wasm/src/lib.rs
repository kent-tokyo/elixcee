//! wasm-bindgen bridge for `@elixcee/xlsx`'s `XLSX.read(bytes)` — a thin JSON-shaping
//! layer over `elixcee::reader::read_workbook_from_bytes`. See
//! `docs/xlsx-architecture.md`'s "Phase 2B-0: sync WASM/read() bridge feasibility" and
//! "Target workspace shape" sections.
//!
//! Ships from two separate `wasm-pack build` invocations (see `packages/xlsx`'s build
//! tooling): `--target nodejs` for the Node entry point (fully synchronous glue — no
//! `await init()`, confirmed in the Phase 2B-0 spike) and `--target web` for the browser
//! entry point, where the caller inlines the compiled `.wasm` bytes into the shipped JS
//! and calls wasm-bindgen's `initSync` itself rather than depending on a bundler resolving
//! a bare `.wasm` import (bundlers don't do that by default — also confirmed in 2B-0). The
//! sync-vs-async difference lives entirely in how each target's glue loads the module, not
//! in this crate's Rust code — both entry points call the exact same export below.

use elixcee::diagnostics::json_string;
use elixcee::reader::{BufferSheet, BufferWorkbook, SheetCell};
use wasm_bindgen::prelude::*;

/// Read an in-memory XLSX/XLSM buffer, returning a JSON string shaped like xlsx@0.18.5's
/// `WorkBook` (`{SheetNames, Sheets}`; each `WorkSheet` a sparse `{"A1": {t,v,f,fmtId}, ...,
/// "!ref": "A1:C3", "!merges": [...], "!hiddenRows": [...], "!hiddenCols": [...] }` object,
/// plus workbook-level `"!numFmts"`/`"!date1904"` — see
/// `packages/xlsx/src/index.d.ts`'s `WorkBook`/`WorkSheet` types). The JS side
/// (`packages/xlsx/src/index.cjs`'s `read()`) does `JSON.parse` on the result — no
/// `serde`/`serde_json` dependency needed for a shape this small; reuses
/// `elixcee::diagnostics::json_string`'s existing hand-rolled escaper (src/diagnostics.rs)
/// rather than duplicating a JSON writer or adding a dependency.
///
/// `!hiddenRows`/`!hiddenCols`/per-cell `fmtId`/`!numFmts`/`!date1904` are NOT the oracle's
/// own `read()` shapes — they're `reader.rs`'s raw parsed data (1-based `[start,end]`
/// intervals; a numFmtId integer; the workbook's custom numFmt table; a bool), passed
/// through as-is. The JS layer resolves all of this into the oracle's real shapes —
/// `!rows`/`!cols` (0-based sparse `{hidden:true}` arrays, gated behind `opts.cellStyles` —
/// confirmed live the oracle never emits them without it), `.w`/`.z` (via the real `ssf`
/// engine, `.z` gated behind `opts.cellNF`/`opts.cellStyles` and always a resolved format
/// STRING, never the raw `fmtId` integer), and `t:'d'`-typed cells (gated behind
/// `opts.cellDates`) — see `packages/xlsx/src/internal/read-shape.cjs`. Keeping that
/// SheetJS-shape-specific (0-based/sparse/option-gated/SSF-backed) work in JS
/// matches how every other xlsx-shape decision already lives in `index.cjs`, not here —
/// and avoids porting SSF's own format-code-to-date heuristic into Rust as a second,
/// unverified implementation of logic already proven correct across 1831 cases
/// (compat/differential/ssf-format.test.mjs).
#[wasm_bindgen(js_name = readWorkbook)]
pub fn read_workbook(bytes: &[u8]) -> Result<String, JsValue> {
    let wb = elixcee::reader::read_workbook_from_bytes(bytes).map_err(|e| JsValue::from_str(&e))?;
    Ok(workbook_json(&wb))
}

fn workbook_json(wb: &BufferWorkbook) -> String {
    let mut names = String::from("[");
    let mut body = String::from("{");
    for (i, bs) in wb.sheets.iter().enumerate() {
        if i > 0 {
            names.push(',');
            body.push(',');
        }
        names.push_str(&json_string(&bs.sheet.name));
        body.push_str(&json_string(&bs.sheet.name));
        body.push(':');
        body.push_str(&worksheet_json(bs));
    }
    names.push(']');
    body.push('}');

    let mut out = format!("{{\"SheetNames\":{},\"Sheets\":{}", names, body);
    if !wb.number_formats.is_empty() {
        // Deterministic key order — same rationale as worksheet_json's cell sort.
        let mut ids: Vec<_> = wb.number_formats.keys().collect();
        ids.sort();
        out.push_str(",\"!numFmts\":{");
        for (i, id) in ids.iter().enumerate() {
            if i > 0 {
                out.push(',');
            }
            out.push_str(&json_string(&id.to_string()));
            out.push(':');
            out.push_str(&json_string(&wb.number_formats[*id]));
        }
        out.push('}');
    }
    out.push_str(&format!(",\"!date1904\":{}", wb.date1904));
    out.push('}');
    out
}

fn worksheet_json(bs: &BufferSheet) -> String {
    let sheet = &bs.sheet;
    let mut out = String::from("{");
    let mut first = true;
    let (mut min_r, mut min_c, mut max_r, mut max_c) = (u32::MAX, u32::MAX, 0u32, 0u32);

    // Deterministic (row, col) order — a HashMap has no defined iteration order, and a
    // stable cell order in the emitted JSON matters for anything downstream that diffs or
    // snapshots the output (e.g. the differential test added alongside this bridge).
    let mut refs: Vec<_> = sheet.cells.iter().collect();
    refs.sort_by_key(|((r, c), _)| (*r, *c));

    for (&(row, col), cell) in refs {
        if !first {
            out.push(',');
        }
        first = false;
        min_r = min_r.min(row);
        max_r = max_r.max(row);
        min_c = min_c.min(col);
        max_c = max_c.max(col);
        out.push_str(&json_string(&cell_ref(row, col)));
        out.push(':');
        out.push_str(&cell_json(
            cell,
            bs.formulas.get(&(row, col)),
            bs.style_ids.get(&(row, col)),
        ));
    }

    // <dimension>, when present and trusted (reader.rs's parse_dimension_ref already
    // replicates the oracle's own colon-required/non-reversed quirks), wins over the
    // populated-cell bounding box — matching the oracle's own parse_ws_xml_dim, which
    // never falls back to a bounding box once a valid <dimension> set !ref. Only fall
    // back to the bounding box (and only when at least one cell exists) when no
    // dimension was trusted, same as this bridge's pre-existing behavior.
    let ref_range = bs
        .dimension
        .or_else(|| (!first).then_some(((min_r, min_c), (max_r, max_c))));
    if let Some(((r1, c1), (r2, c2))) = ref_range {
        out.push_str(",\"!ref\":");
        // A single-cell range collapses to just the cell ref, no colon — matching the
        // oracle's own encode_range (`start === end ? start : start + ':' + end`),
        // ALWAYS used to build !ref regardless of source (bounding box or a trusted
        // <dimension>, even one written as "A1:A1" in the XML — confirmed live: the
        // oracle's own !ref is never the raw <dimension> text echoed back verbatim, it's
        // always re-encoded through encode_range). Found via a real divergence (a
        // single-populated-cell sheet reading back as "A1:A1" here vs the oracle's "A1"),
        // not assumed.
        let start = cell_ref(r1, c1);
        if r1 == r2 && c1 == c2 {
            out.push_str(&json_string(&start));
        } else {
            out.push_str(&json_string(&format!("{}:{}", start, cell_ref(r2, c2))));
        }
    }

    if !sheet.merged_ranges.is_empty() {
        out.push_str(",\"!merges\":[");
        for (i, ((r1, c1), (r2, c2))) in sheet.merged_ranges.iter().enumerate() {
            if i > 0 {
                out.push(',');
            }
            // reader.rs's merged_ranges are 1-based inclusive; WorkSheet's !merges uses
            // 0-based CellAddress (matching encode_cell/decode_cell's own convention) —
            // saturating_sub guards a theoretical 0 from a malformed <mergeCell ref>
            // rather than underflowing (see reader.rs's parse_merge_ref, which does no
            // bounds validation of its own on the crafted-file input it parses).
            out.push_str(&format!(
                "{{\"s\":{{\"r\":{},\"c\":{}}},\"e\":{{\"r\":{},\"c\":{}}}}}",
                r1.saturating_sub(1),
                c1.saturating_sub(1),
                r2.saturating_sub(1),
                c2.saturating_sub(1)
            ));
        }
        out.push(']');
    }

    write_hidden_intervals(&mut out, "!hiddenRows", &sheet.hidden_rows);
    write_hidden_intervals(&mut out, "!hiddenCols", &sheet.hidden_columns);

    out.push('}');
    out
}

/// Serializes `reader.rs`'s native 1-based inclusive `(start, end)` intervals as a raw
/// `[[start,end], ...]` JSON array under `key` — see `worksheet_json`'s doc comment for why
/// this is an internal wire shape, not the oracle's own `!rows`/`!cols`.
fn write_hidden_intervals(out: &mut String, key: &str, intervals: &[(u32, u32)]) {
    if intervals.is_empty() {
        return;
    }
    out.push_str(",\"");
    out.push_str(key);
    out.push_str("\":[");
    for (i, (start, end)) in intervals.iter().enumerate() {
        if i > 0 {
            out.push(',');
        }
        out.push_str(&format!("[{},{}]", start, end));
    }
    out.push(']');
}

fn cell_json(cell: &SheetCell, formula: Option<&String>, fmt_id: Option<&u32>) -> String {
    let mut out = match cell {
        SheetCell::Integer(v) => format!("{{\"t\":\"n\",\"v\":{}", v),
        SheetCell::Float(v) => format!("{{\"t\":\"n\",\"v\":{}", json_number(*v)),
        SheetCell::Str(v) => format!("{{\"t\":\"s\",\"v\":{}", json_string(v)),
        SheetCell::Bool(v) => format!("{{\"t\":\"b\",\"v\":{}", v),
        // Numeric BIFF error code, matching the real oracle's own {t:"e", v:<code>} shape
        // exactly (see ExcelError::biff_code's doc comment) -- shapeCell (read-shape.cjs)
        // derives .w from it, the same layering already used for .w on every other type.
        SheetCell::Error(e) => format!("{{\"t\":\"e\",\"v\":{}", e.biff_code()),
    };
    if let Some(f) = formula {
        out.push_str(",\"f\":");
        out.push_str(&json_string(f));
    }
    if let Some(id) = fmt_id {
        // "fmtId", not the oracle's own "z" — an internal wire key holding a raw
        // numFmtId integer, not yet the resolved format-code STRING the oracle's real
        // `.z` always is (even "General" is a string, never a number, on the oracle —
        // confirmed live). Matches the !hiddenRows/!hiddenCols wire-vs-real-shape
        // convention above rather than overloading `.z`'s two different meanings under
        // one key name. See read-shape.cjs, which resolves this into the real `.z`/`.w`.
        out.push_str(",\"fmtId\":");
        out.push_str(&id.to_string());
    }
    out.push('}');
    out
}

/// A crafted `<v>NaN</v>`/`<v>inf</v>` on a numeric-typed cell parses fine as an f64 (Rust's
/// `FromStr` accepts those literals) but isn't valid JSON — guard at this trust boundary
/// (untrusted file bytes) rather than emit syntactically broken output for a caller that
/// then fails confusingly at `JSON.parse`.
fn json_number(v: f64) -> String {
    if v.is_finite() {
        format!("{}", v)
    } else {
        "null".to_string()
    }
}

/// 1-based (row, col) -> an "A1"-style reference. `reader.rs`'s own module-doc comment
/// on `MergeRect` establishes this codebase's convention of a small per-module
/// col-letter helper rather than a cross-module `utils` dependency — followed here rather
/// than introduced fresh.
fn col_letters(mut col: u32) -> String {
    let mut s = String::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        s.insert(0, (b'A' + rem as u8) as char);
        col = (col - 1) / 26;
    }
    s
}

fn cell_ref(row: u32, col: u32) -> String {
    format!("{}{}", col_letters(col), row)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    fn sheet(name: &str, cells: Vec<((u32, u32), SheetCell)>) -> BufferSheet {
        BufferSheet {
            sheet: elixcee::reader::WorkbookSheet {
                name: name.to_string(),
                cells: cells.into_iter().collect::<HashMap<_, _>>(),
                sheet_id: None,
                workbook_rel_id: None,
                source_part_name: None,
                merged_ranges: vec![],
                hidden_rows: vec![],
                hidden_columns: vec![],
                raw_style_indices: HashMap::new(),
                formulas: HashMap::new(),
                cell_number_formats: HashMap::new(),
                sheet_state: None,
            },
            formulas: HashMap::new(),
            dimension: None,
            style_ids: HashMap::new(),
        }
    }

    // Wraps a single BufferSheet into the BufferWorkbook workbook_json now takes — every
    // test here exercises one sheet at a time, so this is the common case; number_formats/
    // date1904-specific tests build a BufferWorkbook directly instead of through this.
    fn wb1(s: BufferSheet) -> BufferWorkbook {
        BufferWorkbook {
            sheets: vec![s],
            number_formats: HashMap::new(),
            date1904: false,
        }
    }

    #[test]
    fn col_letters_matches_the_usual_a1_z1_aa1_examples() {
        assert_eq!(col_letters(1), "A");
        assert_eq!(col_letters(26), "Z");
        assert_eq!(col_letters(27), "AA");
        assert_eq!(col_letters(702), "ZZ");
    }

    #[test]
    fn workbook_json_shapes_an_empty_sheet_with_no_ref() {
        let json = workbook_json(&wb1(sheet("Sheet1", vec![])));
        assert_eq!(
            json,
            r#"{"SheetNames":["Sheet1"],"Sheets":{"Sheet1":{}},"!date1904":false}"#
        );
    }

    #[test]
    fn workbook_json_computes_ref_and_cell_types_from_mixed_cells() {
        let json = workbook_json(&wb1(sheet(
            "Sheet1",
            vec![
                ((1, 1), SheetCell::Integer(1)),
                ((2, 2), SheetCell::Str("hi".to_string())),
                ((3, 1), SheetCell::Bool(true)),
            ],
        )));
        assert!(json.contains(r#""A1":{"t":"n","v":1}"#));
        assert!(json.contains(r#""B2":{"t":"s","v":"hi"}"#));
        assert!(json.contains(r#""A3":{"t":"b","v":true}"#));
        assert!(json.contains(r#""!ref":"A1:B3""#));
    }

    #[test]
    fn workbook_json_includes_merges_as_zero_based_ranges() {
        let mut s = sheet("Sheet1", vec![((1, 1), SheetCell::Integer(1))]);
        s.sheet.merged_ranges.push(((1, 1), (1, 3)));
        let json = workbook_json(&wb1(s));
        assert!(json.contains(r#""!merges":[{"s":{"r":0,"c":0},"e":{"r":0,"c":2}}]"#));
    }

    #[test]
    fn json_number_guards_non_finite_floats() {
        assert_eq!(json_number(1.5), "1.5");
        assert_eq!(json_number(f64::NAN), "null");
        assert_eq!(json_number(f64::INFINITY), "null");
    }

    // ── read() item 2: <dimension> preference ───────────────────────────────

    #[test]
    fn worksheet_json_prefers_dimension_over_the_populated_bounding_box() {
        let mut s = sheet("Sheet1", vec![((1, 1), SheetCell::Integer(1))]);
        s.dimension = Some(((1, 1), (10, 5)));
        let json = workbook_json(&wb1(s));
        assert!(json.contains(r#""!ref":"A1:E10""#));
    }

    #[test]
    fn worksheet_json_uses_dimension_even_when_no_cells_are_populated() {
        let mut s = sheet("Sheet1", vec![]);
        s.dimension = Some(((1, 1), (3, 3)));
        let json = workbook_json(&wb1(s));
        assert!(json.contains(r#""!ref":"A1:C3""#));
    }

    #[test]
    fn worksheet_json_falls_back_to_the_bounding_box_when_dimension_is_absent() {
        let json = workbook_json(&wb1(sheet(
            "Sheet1",
            vec![
                ((2, 2), SheetCell::Integer(1)),
                ((3, 4), SheetCell::Integer(2)),
            ],
        )));
        assert!(json.contains(r#""!ref":"B2:D3""#));
    }

    // A single populated cell (or a single-cell <dimension>) must collapse !ref to just
    // the cell ref, no colon — matching the oracle's own encode_range convention. Found
    // via a real divergence (see this section's own commit), not assumed.
    #[test]
    fn worksheet_json_collapses_a_single_cell_bounding_box_ref_no_colon() {
        let json = workbook_json(&wb1(sheet("Sheet1", vec![((2, 2), SheetCell::Integer(1))])));
        assert!(json.contains(r#""!ref":"B2""#));
        assert!(!json.contains("\"!ref\":\"B2:B2\""));
    }

    #[test]
    fn worksheet_json_collapses_a_single_cell_dimension_ref_no_colon() {
        let mut s = sheet("Sheet1", vec![((1, 1), SheetCell::Integer(1))]);
        s.dimension = Some(((1, 1), (1, 1)));
        let json = workbook_json(&wb1(s));
        assert!(json.contains(r#""!ref":"A1""#));
        assert!(!json.contains("\"!ref\":\"A1:A1\""));
    }

    // ── read() item 4: formula (.f) ──────────────────────────────────────────

    #[test]
    fn cell_json_includes_f_when_a_formula_is_present() {
        let mut s = sheet("Sheet1", vec![((1, 1), SheetCell::Integer(3))]);
        s.formulas.insert((1, 1), "SUM(B1:B2)".to_string());
        let json = workbook_json(&wb1(s));
        assert!(json.contains(r#""A1":{"t":"n","v":3,"f":"SUM(B1:B2)"}"#));
    }

    #[test]
    fn cell_json_omits_f_when_no_formula_is_present() {
        let json = workbook_json(&wb1(sheet("Sheet1", vec![((1, 1), SheetCell::Integer(3))])));
        assert!(json.contains(r#""A1":{"t":"n","v":3}"#));
        assert!(!json.contains("\"f\":"));
    }

    // ── read() item 3: hidden row/col intervals ─────────────────────────────

    #[test]
    fn worksheet_json_includes_hidden_row_and_col_intervals_when_present() {
        let mut s = sheet("Sheet1", vec![((1, 1), SheetCell::Integer(1))]);
        s.sheet.hidden_rows.push((11, 14));
        s.sheet.hidden_columns.push((2, 2));
        let json = workbook_json(&wb1(s));
        assert!(json.contains(r#""!hiddenRows":[[11,14]]"#));
        assert!(json.contains(r#""!hiddenCols":[[2,2]]"#));
    }

    #[test]
    fn worksheet_json_omits_hidden_interval_keys_when_none_are_hidden() {
        let json = workbook_json(&wb1(sheet("Sheet1", vec![((1, 1), SheetCell::Integer(1))])));
        assert!(!json.contains("!hiddenRows"));
        assert!(!json.contains("!hiddenCols"));
    }

    // ── read() item 6: per-cell fmtId, workbook !numFmts/!date1904 ──────────

    #[test]
    fn cell_json_includes_fmt_id_when_a_non_zero_style_id_is_present() {
        let mut s = sheet("Sheet1", vec![((1, 1), SheetCell::Integer(3))]);
        s.style_ids.insert((1, 1), 14);
        let json = workbook_json(&wb1(s));
        assert!(json.contains(r#""A1":{"t":"n","v":3,"fmtId":14}"#));
    }

    #[test]
    fn cell_json_omits_fmt_id_when_no_style_id_is_present() {
        let json = workbook_json(&wb1(sheet("Sheet1", vec![((1, 1), SheetCell::Integer(3))])));
        assert!(!json.contains("\"fmtId\":"));
    }

    #[test]
    fn workbook_json_includes_num_fmts_when_present() {
        let mut number_formats = HashMap::new();
        number_formats.insert(164u32, "0.00\"kg\"".to_string());
        let wb = BufferWorkbook {
            sheets: vec![sheet("Sheet1", vec![])],
            number_formats,
            date1904: false,
        };
        let json = workbook_json(&wb);
        assert!(json.contains(r#""!numFmts":{"164":"0.00\"kg\""}"#));
    }

    #[test]
    fn workbook_json_omits_num_fmts_when_empty() {
        let json = workbook_json(&wb1(sheet("Sheet1", vec![])));
        assert!(!json.contains("!numFmts"));
    }

    #[test]
    fn workbook_json_always_includes_date1904() {
        let mut wb = wb1(sheet("Sheet1", vec![]));
        wb.date1904 = true;
        let json = workbook_json(&wb);
        assert!(json.contains(r#""!date1904":true"#));
    }
}
