//! Workbook snapshot (Milestone B4): reads a `.xlsx`/`.ods` file directly
//! (via `reader::read_workbook`, no VBA execution) and renders it as either
//! JSON (authoritative, for machine consumers) or Markdown (default, for
//! human display) — the `elixcee snapshot` subcommand's own output shapes,
//! distinct from run-mode's `cells`/`entrypoint` and `check`'s `diagnostics`.

use crate::diagnostics::json_string;
use crate::reader::{SheetCell, WorkbookSheet};

/// `stable_id` is derived from the workbook's own `sheetId` attribute
/// (`WorkbookSheet::sheet_id`) when present — real, and not necessarily
/// contiguous, for a genuine external `.xlsx`. Falls back to a synthetic
/// 1-based positional id otherwise (always the case for `.ods`, and for any
/// elixcee-written `.xlsx`, since this repo's own writer regenerates
/// `sheetId` sequentially from current sheet order on every save).
///
/// Deliberately **not** named `code_name`: that name would suggest VBA's
/// real `CodeName` property (assigned in the VBA IDE, stored in the binary
/// `vbaProject.bin` OLE stream — a format this reader doesn't parse), which
/// is a different, stronger guarantee than "the file's own sheetId, or a
/// synthetic fallback." Keeping `sheet_id`/`stable_id` here leaves
/// `code_name`/`vba_code_name` free to name that real property if it's ever
/// implemented, without a collision.
fn stable_ids(sheets: &[WorkbookSheet]) -> Vec<String> {
    sheets
        .iter()
        .enumerate()
        .map(|(i, s)| match &s.sheet_id {
            Some(id) => format!("sheet{}", id),
            None => format!("sheet{}", i + 1),
        })
        .collect()
}

fn col_to_letters(mut col: u32) -> String {
    let mut bytes = Vec::new();
    while col > 0 {
        col -= 1;
        bytes.push(b'A' + (col % 26) as u8);
        col /= 26;
    }
    bytes.reverse();
    String::from_utf8(bytes).unwrap()
}

fn sorted_cells(sheet: &WorkbookSheet) -> Vec<(&(u32, u32), &SheetCell)> {
    let mut cells: Vec<_> = sheet.cells.iter().collect();
    cells.sort_by_key(|&(&(r, c), _)| (r, c));
    cells
}

fn cell_value_json(c: &SheetCell) -> String {
    match c {
        SheetCell::Integer(n) => n.to_string(),
        SheetCell::Float(f) if f.is_finite() => f.to_string(),
        // NaN/Infinity guard, mirroring diagnostics::variant_to_json — cheap
        // insurance against a malformed source file, not a known-reachable case.
        SheetCell::Float(f) if f.is_nan() => json_string("NaN"),
        SheetCell::Float(f) if *f > 0.0 => json_string("Infinity"),
        SheetCell::Float(_) => json_string("-Infinity"),
        SheetCell::Str(s) => json_string(s),
        SheetCell::Bool(b) => {
            if *b {
                "true".into()
            } else {
                "false".into()
            }
        }
    }
}

fn cell_value_display(c: &SheetCell) -> String {
    match c {
        SheetCell::Integer(n) => n.to_string(),
        SheetCell::Float(f) => f.to_string(),
        SheetCell::Str(s) => s.clone(),
        SheetCell::Bool(b) => {
            if *b {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
    }
}

/// `{"schema_version":1,"ok":true,"file":...,"sheets":[{"name":...,"sheet_id":...,"stable_id":...,"cells":[{"address":...,"value":...}]}]}`
///
/// `sheet_id` is the raw XLSX `sheetId` attribute (`null` for `.ods`, or if
/// missing) — exposed so a consumer can tell whether `stable_id` is backed
/// by a real file attribute or a synthetic positional fallback.
pub fn to_json(file: &str, sheets: &[WorkbookSheet]) -> String {
    let stable_ids = stable_ids(sheets);
    let sheet_items: Vec<String> = sheets
        .iter()
        .zip(stable_ids.iter())
        .map(|(sheet, stable_id)| {
            let cell_items: Vec<String> = sorted_cells(sheet)
                .iter()
                .map(|&(&(row, col), content)| {
                    format!(
                        "{{\"address\":{},\"value\":{}}}",
                        json_string(&format!("{}{}", col_to_letters(col), row)),
                        cell_value_json(content),
                    )
                })
                .collect();
            let sheet_id_json = match &sheet.sheet_id {
                Some(id) => json_string(id),
                None => "null".to_string(),
            };
            format!(
                "{{\"name\":{},\"sheet_id\":{},\"stable_id\":{},\"cells\":[{}]}}",
                json_string(&sheet.name),
                sheet_id_json,
                json_string(stable_id),
                cell_items.join(","),
            )
        })
        .collect();
    format!(
        "{{\"schema_version\":1,\"ok\":true,\"file\":{},\"sheets\":[{}]}}",
        json_string(file),
        sheet_items.join(","),
    )
}

/// Markdown table cells can't contain a literal `|` or newline — this is
/// display-only escaping, not meant to round-trip (unlike `to_json`).
fn md_escape(s: &str) -> String {
    s.replace('|', "\\|").replace(['\n', '\r'], " ")
}

/// A top-level sheet index table followed by one cells table per sheet.
/// Shows only `stable_id` (not the raw `sheet_id`) — Markdown here is a
/// simplified display view, not a full-parity mirror of `to_json`.
pub fn to_markdown(file: &str, sheets: &[WorkbookSheet]) -> String {
    let stable_ids = stable_ids(sheets);
    let mut out = format!("# Workbook Snapshot: {}\n\n", md_escape(file));

    out.push_str("| Sheet | stable_id | Cells |\n|---|---|---|\n");
    for (sheet, stable_id) in sheets.iter().zip(stable_ids.iter()) {
        out.push_str(&format!(
            "| {} | {} | {} |\n",
            md_escape(&sheet.name),
            stable_id,
            sheet.cells.len(),
        ));
    }

    for (sheet, stable_id) in sheets.iter().zip(stable_ids.iter()) {
        out.push_str(&format!(
            "\n## {} (stable_id: {})\n\n",
            md_escape(&sheet.name),
            stable_id
        ));
        out.push_str("| Address | Value |\n|---|---|\n");
        for &(&(row, col), content) in &sorted_cells(sheet) {
            out.push_str(&format!(
                "| {}{} | {} |\n",
                col_to_letters(col),
                row,
                md_escape(&cell_value_display(content)),
            ));
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::collections::HashMap;

    fn sheet(
        name: &str,
        sheet_id: Option<&str>,
        cells: Vec<((u32, u32), SheetCell)>,
    ) -> WorkbookSheet {
        WorkbookSheet {
            name: name.to_string(),
            cells: cells.into_iter().collect::<HashMap<_, _>>(),
            sheet_id: sheet_id.map(|s| s.to_string()),
        }
    }

    #[test]
    fn stable_id_uses_real_sheet_id_when_present() {
        let sheets = vec![sheet("Sheet1", Some("5"), vec![])];
        assert_eq!(stable_ids(&sheets), vec!["sheet5"]);
    }

    #[test]
    fn stable_id_falls_back_to_position_when_absent() {
        let sheets = vec![sheet("Sheet1", None, vec![]), sheet("Sheet2", None, vec![])];
        assert_eq!(stable_ids(&sheets), vec!["sheet1", "sheet2"]);
    }

    #[test]
    fn stable_id_mixes_real_and_synthetic_per_sheet() {
        let sheets = vec![
            sheet("Sheet1", Some("3"), vec![]),
            sheet("Sheet2", None, vec![]),
        ];
        // Position 2 (index 1) falls back to "sheet2" even though a real
        // sheet_id of "3" was used for the sheet before it — each sheet's
        // fallback is independent, not offset by neighbors' real ids.
        assert_eq!(stable_ids(&sheets), vec!["sheet3", "sheet2"]);
    }

    #[test]
    fn to_json_escapes_quotes_in_names_and_values() {
        let sheets = vec![sheet(
            "Sh\"eet",
            Some("1"),
            vec![((1, 1), SheetCell::Str("say \"hi\"".to_string()))],
        )];
        let json = to_json("book.xlsx", &sheets);
        assert!(json.contains(r#""name":"Sh\"eet""#));
        assert!(json.contains(r#""value":"say \"hi\"""#));
        assert!(json.contains(r#""address":"A1""#));
        assert!(json.contains(r#""sheet_id":"1""#));
        assert!(json.contains(r#""stable_id":"sheet1""#));
    }

    #[test]
    fn to_json_sheet_id_is_null_when_synthetic() {
        let sheets = vec![sheet("Sheet1", None, vec![])];
        let json = to_json("book.xlsx", &sheets);
        assert!(json.contains(r#""sheet_id":null"#));
        assert!(json.contains(r#""stable_id":"sheet1""#));
    }

    #[test]
    fn to_json_cells_are_sorted_by_row_then_column() {
        let sheets = vec![sheet(
            "Sheet1",
            Some("1"),
            vec![
                ((2, 1), SheetCell::Integer(2)),
                ((1, 2), SheetCell::Integer(12)),
                ((1, 1), SheetCell::Integer(11)),
            ],
        )];
        let json = to_json("book.xlsx", &sheets);
        let a1 = json.find("\"A1\"").unwrap();
        let b1 = json.find("\"B1\"").unwrap();
        let a2 = json.find("\"A2\"").unwrap();
        assert!(a1 < b1 && b1 < a2, "expected A1, B1, A2 order: {}", json);
    }

    #[test]
    fn to_markdown_escapes_pipes_and_newlines() {
        let sheets = vec![sheet(
            "Sheet|One",
            None,
            vec![((1, 1), SheetCell::Str("line1\nline2".to_string()))],
        )];
        let md = to_markdown("book.xlsx", &sheets);
        assert!(md.contains("Sheet\\|One"));
        assert!(md.contains("line1 line2"));
        assert!(!md.contains("line1\nline2"));
    }

    #[test]
    fn to_markdown_contains_header_and_cell_table() {
        let sheets = vec![sheet(
            "Sheet1",
            Some("1"),
            vec![((1, 1), SheetCell::Integer(42))],
        )];
        let md = to_markdown("book.xlsx", &sheets);
        assert!(md.contains("# Workbook Snapshot: book.xlsx"));
        assert!(md.contains("| Sheet1 | sheet1 | 1 |"));
        assert!(md.contains("| A1 | 42 |"));
    }
}
