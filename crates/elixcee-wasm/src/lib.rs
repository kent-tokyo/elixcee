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
use elixcee::reader::{SheetCell, WorkbookSheet};
use wasm_bindgen::prelude::*;

/// Read an in-memory XLSX/XLSM buffer, returning a JSON string shaped like xlsx@0.18.5's
/// `WorkBook` (`{SheetNames, Sheets}`; each `WorkSheet` a sparse `{"A1": {t,v}, ...,
/// "!ref": "A1:C3", "!merges": [...] }` object — see `packages/xlsx/src/index.d.ts`'s
/// `WorkBook`/`WorkSheet` types). The JS side (`packages/xlsx/src/index.cjs`'s `read()`)
/// does `JSON.parse` on the result — no `serde`/`serde_json` dependency needed for a shape
/// this small; reuses `elixcee::diagnostics::json_string`'s existing hand-rolled escaper
/// (src/diagnostics.rs) rather than duplicating a JSON writer or adding a dependency.
///
/// Deliberately not feature-complete with the oracle's `read()`: no cell formulas (`.f` —
/// `reader.rs` never parses `<f>`), no formatted display text (`.w` — would need
/// `styles.xml` number-format parsing `reader.rs` doesn't do), no date-typed cells (`t:
/// 'd'` — same `styles.xml` gap, so a date serial reads back as a plain number, matching
/// every other `reader.rs` consumer today), no hidden-row/col `!rows`/`!cols` mapping
/// (`reader.rs` has that as coalesced intervals; `WorkSheet` wants a per-index array —
/// left for a follow-up once a caller needs `skipHidden`). Merged ranges ARE mapped
/// (`!merges`) since `reader.rs` already parses them and `sheet_to_html` already consumes
/// that shape.
#[wasm_bindgen(js_name = readWorkbook)]
pub fn read_workbook(bytes: &[u8]) -> Result<String, JsValue> {
    let sheets = elixcee::reader::read_workbook_from_bytes(bytes).map_err(|e| JsValue::from_str(&e))?;
    Ok(workbook_json(&sheets))
}

fn workbook_json(sheets: &[WorkbookSheet]) -> String {
    let mut names = String::from("[");
    let mut body = String::from("{");
    for (i, sheet) in sheets.iter().enumerate() {
        if i > 0 {
            names.push(',');
            body.push(',');
        }
        names.push_str(&json_string(&sheet.name));
        body.push_str(&json_string(&sheet.name));
        body.push(':');
        body.push_str(&worksheet_json(sheet));
    }
    names.push(']');
    body.push('}');
    format!("{{\"SheetNames\":{},\"Sheets\":{}}}", names, body)
}

fn worksheet_json(sheet: &WorkbookSheet) -> String {
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
        out.push_str(&cell_json(cell));
    }

    if !first {
        out.push_str(",\"!ref\":");
        out.push_str(&json_string(&format!(
            "{}:{}",
            cell_ref(min_r, min_c),
            cell_ref(max_r, max_c)
        )));
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

    out.push('}');
    out
}

fn cell_json(cell: &SheetCell) -> String {
    match cell {
        SheetCell::Integer(v) => format!("{{\"t\":\"n\",\"v\":{}}}", v),
        SheetCell::Float(v) => format!("{{\"t\":\"n\",\"v\":{}}}", json_number(*v)),
        SheetCell::Str(v) => format!("{{\"t\":\"s\",\"v\":{}}}", json_string(v)),
        SheetCell::Bool(v) => format!("{{\"t\":\"b\",\"v\":{}}}", v),
    }
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

    fn sheet(name: &str, cells: Vec<((u32, u32), SheetCell)>) -> WorkbookSheet {
        WorkbookSheet {
            name: name.to_string(),
            cells: cells.into_iter().collect::<HashMap<_, _>>(),
            sheet_id: None,
            merged_ranges: vec![],
            hidden_rows: vec![],
            hidden_columns: vec![],
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
        let json = workbook_json(&[sheet("Sheet1", vec![])]);
        assert_eq!(json, r#"{"SheetNames":["Sheet1"],"Sheets":{"Sheet1":{}}}"#);
    }

    #[test]
    fn workbook_json_computes_ref_and_cell_types_from_mixed_cells() {
        let json = workbook_json(&[sheet(
            "Sheet1",
            vec![
                ((1, 1), SheetCell::Integer(1)),
                ((2, 2), SheetCell::Str("hi".to_string())),
                ((3, 1), SheetCell::Bool(true)),
            ],
        )]);
        assert!(json.contains(r#""A1":{"t":"n","v":1}"#));
        assert!(json.contains(r#""B2":{"t":"s","v":"hi"}"#));
        assert!(json.contains(r#""A3":{"t":"b","v":true}"#));
        assert!(json.contains(r#""!ref":"A1:B3""#));
    }

    #[test]
    fn workbook_json_includes_merges_as_zero_based_ranges() {
        let mut s = sheet("Sheet1", vec![((1, 1), SheetCell::Integer(1))]);
        s.merged_ranges.push(((1, 1), (1, 3)));
        let json = workbook_json(&[s]);
        assert!(json.contains(r#""!merges":[{"s":{"r":0,"c":0},"e":{"r":0,"c":2}}]"#));
    }

    #[test]
    fn json_number_guards_non_finite_floats() {
        assert_eq!(json_number(1.5), "1.5");
        assert_eq!(json_number(f64::NAN), "null");
        assert_eq!(json_number(f64::INFINITY), "null");
    }
}
