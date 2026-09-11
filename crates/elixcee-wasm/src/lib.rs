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
use elixcee::reader::{BufferSheet, BufferWorkbook, FilterCriteria, SheetCell};
use elixcee::types::{ArrayShape, CellContent, ExcelError, Variant};
use std::collections::{HashMap, HashSet};
use wasm_bindgen::prelude::*;

const MAX_EDITOR_HISTORY: usize = 128;
const MAX_WORKSHEET_ROW: u32 = 1_048_576;
const MAX_WORKSHEET_COLUMN: u32 = 16_384;

#[derive(Clone)]
struct EditorState {
    sheets: HashMap<String, HashMap<(u32, u32), CellContent>>,
}

#[derive(Clone)]
struct EditorTransaction {
    state: EditorState,
    undo_len: usize,
    redo: Vec<EditorState>,
}

/// Stateful JS/WASM workbook editor for calculation-oriented workflows.
/// OOXML writing remains the responsibility of the package's existing JS writer.
#[wasm_bindgen]
pub struct WorkbookEditor {
    workbook: BufferWorkbook,
    sheets: HashMap<String, HashMap<(u32, u32), CellContent>>,
    undo: Vec<EditorState>,
    redo: Vec<EditorState>,
    transaction: Option<EditorTransaction>,
    transaction_dirty: bool,
}

#[wasm_bindgen]
impl WorkbookEditor {
    #[wasm_bindgen(constructor)]
    pub fn new(bytes: &[u8]) -> Result<WorkbookEditor, JsValue> {
        let workbook =
            elixcee::reader::read_workbook_from_bytes(bytes).map_err(|e| JsValue::from_str(&e))?;
        let sheets = calculation_sheets(&workbook);
        Ok(Self {
            workbook,
            sheets,
            undo: Vec::new(),
            redo: Vec::new(),
            transaction: None,
            transaction_dirty: false,
        })
    }

    #[wasm_bindgen(js_name = setNumber)]
    pub fn set_number(
        &mut self,
        sheet: &str,
        row: u32,
        col: u32,
        value: f64,
    ) -> Result<(), JsValue> {
        let key = self.validate_cell_target(sheet, row, col)?;
        if !value.is_finite() {
            return Err(JsValue::from_str("cell value must be finite"));
        }
        self.record_edit();
        let value = if value.fract() == 0.0 {
            Variant::Integer(value as i64)
        } else {
            Variant::Float(value)
        };
        self.sheets.get_mut(&key).expect("checked above").insert(
            (row, col),
            CellContent {
                formula: None,
                value,
            },
        );
        Ok(())
    }

    #[wasm_bindgen(js_name = setString)]
    pub fn set_string(
        &mut self,
        sheet: &str,
        row: u32,
        col: u32,
        value: &str,
    ) -> Result<(), JsValue> {
        let key = self.validate_cell_target(sheet, row, col)?;
        self.record_edit();
        self.sheets.get_mut(&key).expect("checked above").insert(
            (row, col),
            CellContent {
                formula: None,
                value: Variant::Str(value.to_string()),
            },
        );
        Ok(())
    }

    #[wasm_bindgen(js_name = setBoolean)]
    pub fn set_boolean(
        &mut self,
        sheet: &str,
        row: u32,
        col: u32,
        value: bool,
    ) -> Result<(), JsValue> {
        let key = self.validate_cell_target(sheet, row, col)?;
        self.record_edit();
        self.sheets.get_mut(&key).expect("checked above").insert(
            (row, col),
            CellContent {
                formula: None,
                value: Variant::Boolean(value),
            },
        );
        Ok(())
    }

    pub fn recalculate(&mut self) -> Result<String, JsValue> {
        elixcee::formula::calculate_workbook(&mut self.sheets, &HashMap::new())
            .map_err(|e| JsValue::from_str(&e))?;
        Ok(self.json_snapshot())
    }

    pub fn snapshot(&mut self) -> String {
        self.json_snapshot()
    }

    pub fn undo(&mut self) -> bool {
        let Some(previous) = self.undo.pop() else {
            return false;
        };
        self.redo.push(self.capture_state());
        self.sheets = previous.sheets;
        true
    }

    pub fn redo(&mut self) -> bool {
        let Some(next) = self.redo.pop() else {
            return false;
        };
        self.undo.push(self.capture_state());
        self.sheets = next.sheets;
        true
    }

    #[wasm_bindgen(js_name = beginTransaction)]
    pub fn begin_transaction(&mut self) -> Result<(), JsValue> {
        if self.transaction.is_some() {
            return Err(JsValue::from_str("an edit transaction is already active"));
        }
        self.transaction = Some(EditorTransaction {
            state: self.capture_state(),
            undo_len: self.undo.len(),
            redo: self.redo.clone(),
        });
        self.transaction_dirty = false;
        Ok(())
    }

    #[wasm_bindgen(js_name = commitTransaction)]
    pub fn commit_transaction(&mut self) -> bool {
        let Some(transaction) = self.transaction.take() else {
            return false;
        };
        if self.transaction_dirty {
            self.undo.push(transaction.state);
            self.undo.truncate(MAX_EDITOR_HISTORY);
            self.redo.clear();
        }
        self.transaction_dirty = false;
        true
    }

    #[wasm_bindgen(js_name = abortTransaction)]
    pub fn abort_transaction(&mut self) -> bool {
        let Some(transaction) = self.transaction.take() else {
            return false;
        };
        self.sheets = transaction.state.sheets;
        self.undo.truncate(transaction.undo_len);
        self.redo = transaction.redo;
        self.transaction_dirty = false;
        true
    }

    #[wasm_bindgen(js_name = canUndo)]
    pub fn can_undo(&self) -> bool {
        !self.undo.is_empty()
    }
    #[wasm_bindgen(js_name = canRedo)]
    pub fn can_redo(&self) -> bool {
        !self.redo.is_empty()
    }

    fn capture_state(&self) -> EditorState {
        EditorState {
            sheets: self.sheets.clone(),
        }
    }

    fn validate_cell_target(&self, sheet: &str, row: u32, col: u32) -> Result<String, JsValue> {
        let key = sheet.to_ascii_lowercase();
        if !self.sheets.contains_key(&key) {
            return Err(JsValue::from_str("unknown worksheet"));
        }
        if row == 0 || row > MAX_WORKSHEET_ROW || col == 0 || col > MAX_WORKSHEET_COLUMN {
            return Err(JsValue::from_str(
                "cell coordinates are outside the worksheet bounds",
            ));
        }
        Ok(key)
    }

    fn record_edit(&mut self) {
        if self.transaction.is_some() {
            self.transaction_dirty = true;
            return;
        }
        self.undo.push(self.capture_state());
        self.undo.truncate(MAX_EDITOR_HISTORY);
        self.redo.clear();
    }

    fn json_snapshot(&mut self) -> String {
        sync_workbook_values(&mut self.workbook, &self.sheets);
        workbook_json(&self.workbook)
    }
}

/// Read an in-memory XLSX/XLSM buffer, returning a JSON string shaped like xlsx@0.18.5's
/// `WorkBook` (`{SheetNames, Sheets}`; each `WorkSheet` a sparse `{"A1": {t,v,f,fmtId}, ...,
/// "!ref": "A1:C3", "!merges": [...], "!hiddenRows": [...], "!hiddenCols": [...] }` object,
/// plus workbook-level `"!numFmts"`/`"!date1904"` and, when present, `Workbook.Names` — see
/// `packages/xlsx/src/index.d.ts`'s `WorkBook`/`WorkSheet` types). The JS side
/// (`packages/xlsx/src/index.cjs`'s `read()`) does `JSON.parse` on the result — no
/// `serde`/`serde_json` dependency needed for a shape this small; reuses
/// `elixcee::diagnostics::json_string`'s existing hand-rolled escaper (src/diagnostics.rs)
/// rather than duplicating a JSON writer or adding a dependency.
///
/// `!hiddenRows`/`!hiddenCols`/per-cell `fmtId`/`!numFmts`/`!date1904`/`!dataValidations` are NOT the oracle's
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

/// Recalculate formula cells in an in-memory XLSX/XLSM buffer using the same
/// Rust workbook runtime as the VM, then return the normal raw workbook JSON.
///
/// This deliberately keeps writing in JavaScript: the export is a calculation
/// bridge, not an OOXML writer. Files without qualified formulas are also
/// accepted and returned unchanged because the current shared workbook path is
/// only needed when at least one formula requires workbook context.
#[wasm_bindgen(js_name = calculateWorkbook)]
pub fn calculate_workbook(bytes: &[u8]) -> Result<String, JsValue> {
    let mut wb =
        elixcee::reader::read_workbook_from_bytes(bytes).map_err(|e| JsValue::from_str(&e))?;
    let mut sheets = HashMap::new();
    for bs in &wb.sheets {
        let mut cells = HashMap::with_capacity(bs.sheet.cells.len().max(bs.formulas.len()));
        for (&position, cell) in &bs.sheet.cells {
            cells.insert(
                position,
                CellContent {
                    formula: bs.formulas.get(&position).map(|f| format!("={f}")),
                    value: sheet_cell_to_variant(cell),
                },
            );
        }
        for (&position, formula) in &bs.formulas {
            cells.entry(position).or_insert_with(|| CellContent {
                formula: Some(format!("={formula}")),
                value: Variant::Empty,
            });
        }
        sheets.insert(bs.sheet.name.to_ascii_lowercase(), cells);
    }
    // Feed defined names into the shared workbook evaluator, preserving local
    // worksheet scope instead of silently treating local names as global.
    let named_ranges = wb
        .defined_names
        .iter()
        .filter(|defined_name| defined_name.local_sheet_id.is_none())
        .map(|defined_name| {
            (
                defined_name.name.to_ascii_lowercase(),
                normalize_defined_name_ref(&defined_name.raw_text),
            )
        })
        .collect::<HashMap<_, _>>();
    let mut scoped_named_ranges = HashMap::<String, HashMap<String, String>>::new();
    for defined_name in wb.defined_names.iter().filter_map(|defined_name| {
        defined_name
            .local_sheet_id
            .map(|index| (index, defined_name))
    }) {
        let Some(sheet) = wb
            .sheets
            .get(defined_name.0)
            .map(|buffer| buffer.sheet.name.clone())
        else {
            continue;
        };
        scoped_named_ranges
            .entry(sheet.to_ascii_lowercase())
            .or_default()
            .insert(
                defined_name.1.name.to_ascii_lowercase(),
                normalize_defined_name_ref(&defined_name.1.raw_text),
            );
    }
    // The workbook evaluator's structured-reference map is global, while a table
    // reference is resolved relative to the formula's host sheet. Rewrite the bounded
    // `Table[Column]` form to an A1 range per host before entering the shared evaluator;
    // this also keeps same-sheet references on the fast, unqualified path.
    rewrite_table_structured_formulas(&mut sheets, &wb);
    elixcee::formula::calculate_workbook_with_context(
        &mut sheets,
        &named_ranges,
        &scoped_named_ranges,
        &HashMap::new(),
    )
    .map_err(|e| JsValue::from_str(&e))?;
    materialize_formula_array_spills(&mut sheets).map_err(|e| JsValue::from_str(&e))?;
    for bs in &mut wb.sheets {
        let Some(cells) = sheets.get(&bs.sheet.name.to_ascii_lowercase()) else {
            continue;
        };
        // Project every calculated cell, not only the original formula cells.
        // Dynamic-array evaluation can materialize new spill cells after the
        // OOXML reader has built its initial sparse sheet model.
        for (&position, cell) in cells {
            bs.sheet
                .cells
                .insert(position, variant_to_sheet_cell(&cell.value));
        }
    }
    Ok(workbook_json(&wb))
}

/// Materialize the shared evaluator's flat formula arrays as bounded worksheet spills.
/// Shape-aware functions recover a rectangular footprint from literal arguments; unknown
/// flat arrays retain the legacy one-row fallback. Never overwrite an occupied cell.
fn materialize_formula_array_spills(
    sheets: &mut HashMap<String, HashMap<(u32, u32), CellContent>>,
) -> Result<(), String> {
    let mut pending = Vec::new();
    let mut spill_errors = Vec::new();
    let mut planned = HashSet::new();
    for (sheet, cells) in sheets.iter() {
        for (&(row, col), cell) in cells {
            let Variant::Array(values) = &cell.value else {
                continue;
            };
            let shape = formula_array_shape(cell.formula.as_deref(), values.len());
            if shape.is_empty() {
                continue;
            }
            let mut collision = false;
            let mut targets = Vec::new();
            for offset in 1..shape.cell_count() {
                let value = values
                    .get(offset)
                    .cloned()
                    .unwrap_or(Variant::Error(ExcelError::NA));
                let row_offset = offset / shape.cols;
                let col_offset = offset % shape.cols;
                let target_row = row.checked_add(row_offset as u32).ok_or_else(|| {
                    format!("dynamic array spill exceeds worksheet bounds on {sheet}")
                })?;
                let target_col = col.checked_add(col_offset as u32).ok_or_else(|| {
                    format!("dynamic array spill exceeds worksheet bounds on {sheet}")
                })?;
                let key = (sheet.clone(), (target_row, target_col));
                if cells.contains_key(&(target_row, target_col)) || planned.contains(&key) {
                    collision = true;
                    break;
                }
                targets.push(((target_row, target_col), value));
            }
            if collision {
                spill_errors.push((sheet.clone(), (row, col)));
            } else {
                planned.extend(
                    targets
                        .iter()
                        .map(|(position, _)| (sheet.clone(), *position)),
                );
                pending.extend(
                    targets
                        .into_iter()
                        .map(|(position, value)| (sheet.clone(), position, value)),
                );
            }
        }
    }
    for (sheet, position, value) in pending {
        sheets.entry(sheet).or_default().insert(
            position,
            CellContent {
                formula: None,
                value,
            },
        );
    }
    for (sheet, position) in spill_errors {
        if let Some(cell) = sheets
            .get_mut(&sheet)
            .and_then(|cells| cells.get_mut(&position))
        {
            cell.value = Variant::Str("#SPILL!".to_string());
        }
    }
    Ok(())
}

fn formula_array_shape(formula: Option<&str>, len: usize) -> ArrayShape {
    let Some(formula) = formula else {
        return ArrayShape::new(1, len);
    };
    let expression = formula
        .trim()
        .trim_start_matches('=')
        .trim()
        .to_ascii_uppercase();
    let Some(open) = expression.find('(') else {
        return ArrayShape::new(1, len);
    };
    if !expression.ends_with(')') {
        return ArrayShape::new(1, len);
    }
    let name = &expression[..open];
    let arguments = &expression[open + 1..expression.len() - 1];
    let values = arguments.split(',').map(str::trim).collect::<Vec<_>>();
    let literal = |index: usize, fallback: usize| {
        values
            .get(index)
            .and_then(|value| value.parse::<usize>().ok())
            .unwrap_or(fallback)
    };
    let shape = match name {
        "SEQUENCE" | "RANDARRAY" => ArrayShape::new(literal(0, 0), literal(1, 1)),
        "WRAPROWS" => {
            let cols = literal(1, 0);
            ArrayShape::new(
                if cols == 0 { 0 } else { len.div_ceil(cols) },
                cols.min(len),
            )
        }
        "WRAPCOLS" => {
            let rows = literal(1, 0);
            ArrayShape::new(
                rows.min(len),
                if rows == 0 { 0 } else { len.div_ceil(rows) },
            )
        }
        _ => ArrayShape::new(1, len),
    };
    if shape.cell_count() >= len && !shape.is_empty() {
        shape
    } else {
        ArrayShape::new(1, len)
    }
}

fn normalize_defined_name_ref(raw: &str) -> String {
    let Some(bang) = raw.rfind('!') else {
        // The workbook formula parser accepts absolute markers on qualified
        // references, while the local-range expansion path consumes the
        // compact A1 address directly. Normalize only this unqualified form.
        return raw.replace('$', "");
    };
    let qualifier = &raw[..bang];
    if qualifier.starts_with('\'') && qualifier.ends_with('\'') && qualifier.len() >= 2 {
        return format!(
            "{}!{}",
            qualifier[1..qualifier.len() - 1].replace("''", "'"),
            &raw[bang + 1..]
        );
    }
    raw.to_string()
}

fn wasm_column_label(mut column: u32) -> String {
    let mut label = String::new();
    loop {
        label.insert(0, (b'A' + (column % 26) as u8) as char);
        if column < 26 {
            break;
        }
        column = column / 26 - 1;
    }
    label
}

fn replace_table_ref_case_insensitive(source: &str, needle: &str, replacement: &str) -> String {
    let source_lower = source.to_ascii_lowercase();
    let needle_lower = needle.to_ascii_lowercase();
    let mut out = String::with_capacity(source.len());
    let mut cursor = 0usize;
    while let Some(relative) = source_lower[cursor..].find(&needle_lower) {
        let start = cursor + relative;
        let end = start + needle.len();
        let inside_string = {
            let mut in_string = false;
            let mut chars = source[..start].chars().peekable();
            while let Some(ch) = chars.next() {
                if ch != '"' {
                    continue;
                }
                if chars.peek() == Some(&'"') {
                    chars.next();
                } else {
                    in_string = !in_string;
                }
            }
            in_string
        };
        let preceded_by_identifier = source[..start]
            .chars()
            .next_back()
            .is_some_and(|ch| ch.is_ascii_alphanumeric() || ch == '_');
        let followed_by_identifier = source[end..]
            .chars()
            .next()
            .is_some_and(|ch| ch.is_ascii_alphanumeric() || ch == '_');
        out.push_str(&source[cursor..start]);
        if inside_string || preceded_by_identifier || followed_by_identifier {
            out.push_str(&source[start..end]);
        } else {
            out.push_str(replacement);
        }
        cursor = end;
    }
    out.push_str(&source[cursor..]);
    out
}

fn rewrite_table_structured_formulas(
    sheets: &mut HashMap<String, HashMap<(u32, u32), CellContent>>,
    wb: &BufferWorkbook,
) {
    let mut patterns = Vec::<(String, String, String, bool)>::new();
    for buffer in &wb.sheets {
        let table_sheet = &buffer.sheet.name;
        for table in &buffer.sheet.tables {
            let (table_start, table_left) = table.ref_range.0;
            let (table_end, table_right) = table.ref_range.1;
            let data_start = table.ref_range.0.0 + table.header_row_count;
            let data_end = table.ref_range.1.0.saturating_sub(table.totals_row_count);
            let headers_end = table_start + table.header_row_count - 1;
            let table_address = |start: u32, end: u32, left: u32, right: u32| {
                format!(
                    "{}!{}{}:{}{}",
                    table_sheet,
                    wasm_column_label(left.saturating_sub(1)),
                    start,
                    wasm_column_label(right.saturating_sub(1)),
                    end
                )
            };
            for table_name in [&table.name, &table.display_name] {
                patterns.push((
                    format!("{}[#Headers]", table_name),
                    table_sheet.clone(),
                    table_address(table_start, headers_end, table_left, table_right),
                    false,
                ));
                patterns.push((
                    format!("{}[#All]", table_name),
                    table_sheet.clone(),
                    table_address(table_start, table_end, table_left, table_right),
                    false,
                ));
                if data_start <= data_end {
                    patterns.push((
                        format!("{}[#Data]", table_name),
                        table_sheet.clone(),
                        table_address(data_start, data_end, table_left, table_right),
                        false,
                    ));
                }
            }
            if data_start > data_end {
                continue;
            }
            for (index, column) in table.columns.iter().enumerate() {
                let col = table.ref_range.0.1 + index as u32;
                let qualified = format!(
                    "{}!{}{}:{}{}",
                    table_sheet,
                    wasm_column_label(col.saturating_sub(1)),
                    data_start,
                    wasm_column_label(col.saturating_sub(1)),
                    data_end
                );
                for table_name in [&table.name, &table.display_name] {
                    let column_pattern = format!("{}[{}]", table_name, column.name);
                    patterns.push((
                        column_pattern.clone(),
                        table_sheet.clone(),
                        qualified.clone(),
                        false,
                    ));
                    patterns.push((
                        format!("{}[[#Data],[{}]]", table_name, column.name),
                        table_sheet.clone(),
                        qualified.clone(),
                        false,
                    ));
                    patterns.push((
                        format!("{}[[#Headers],[{}]]", table_name, column.name),
                        table_sheet.clone(),
                        table_address(table_start, headers_end, col, col),
                        false,
                    ));
                    patterns.push((
                        format!("{}[[#All],[{}]]", table_name, column.name),
                        table_sheet.clone(),
                        table_address(table_start, table_end, col, col),
                        false,
                    ));
                    patterns.push((
                        format!("{}[@{}]", table_name, column.name),
                        table_sheet.clone(),
                        qualified.clone(),
                        true,
                    ));
                    patterns.push((
                        format!("{}[[#This Row],[{}]]", table_name, column.name),
                        table_sheet.clone(),
                        qualified.clone(),
                        true,
                    ));
                }
            }
        }
    }
    patterns.sort_by_key(|(pattern, _, _, _)| std::cmp::Reverse(pattern.len()));
    for (host, cells) in sheets.iter_mut() {
        for (&(row, _), cell) in cells.iter_mut() {
            let Some(formula) = cell.formula.as_mut() else {
                continue;
            };
            for (pattern, table_sheet, qualified, this_row) in &patterns {
                if *this_row && !host.eq_ignore_ascii_case(table_sheet) {
                    continue;
                }
                let replacement = if *this_row {
                    let Some((_, data_range)) = qualified.split_once('!') else {
                        continue;
                    };
                    let Some((start, end)) = data_range.split_once(':') else {
                        continue;
                    };
                    let col = start.trim_end_matches(|ch: char| ch.is_ascii_digit());
                    let start_row = start
                        .trim_start_matches(|ch: char| ch.is_ascii_alphabetic())
                        .parse::<u32>()
                        .ok();
                    let end_col = end.trim_end_matches(|ch: char| ch.is_ascii_digit());
                    if start_row.is_none() || end_col != col {
                        continue;
                    }
                    format!("{}{}", col, row)
                } else if host.eq_ignore_ascii_case(table_sheet) {
                    qualified
                        .split_once('!')
                        .map_or(qualified.as_str(), |(_, address)| address)
                        .to_string()
                } else {
                    qualified.clone()
                };
                let rewritten = replace_table_ref_case_insensitive(formula, pattern, &replacement);
                if rewritten != *formula {
                    *formula = rewritten;
                }
            }
        }
    }
}

/// Return a small, deterministic diagnostic summary without materializing a
/// calculated workbook. This gives Node/browser callers a structured preflight
/// signal while keeping read/calculation failures as ordinary `Result` errors.
#[wasm_bindgen(js_name = diagnoseWorkbook)]
pub fn diagnose_workbook(bytes: &[u8]) -> String {
    let wb = match elixcee::reader::read_workbook_from_bytes(bytes) {
        Ok(wb) => wb,
        Err(error) => {
            return format!("{{\"ok\":false,\"error\":{}}}", json_string(&error));
        }
    };
    let mut formula_count = 0usize;
    let mut qualified_formula_count = 0usize;
    let mut formula_parse_errors = 0usize;
    let calculation_sheets = calculation_sheets(&wb);
    for bs in &wb.sheets {
        for formula in bs.formulas.values() {
            formula_count += 1;
            match elixcee::formula::parse(formula) {
                Ok(expr) => {
                    if contains_qualified_reference(&expr) {
                        qualified_formula_count += 1;
                    }
                }
                Err(_) => formula_parse_errors += 1,
            }
        }
    }
    format!(
        "{{\"ok\":true,\"sheetCount\":{},\"formulaCount\":{},\"qualifiedFormulaCount\":{},\"formulaParseErrors\":{},\"hasFormulaCycle\":{}}}",
        wb.sheets.len(),
        formula_count,
        qualified_formula_count,
        formula_parse_errors,
        if elixcee::formula::workbook_has_formula_cycle(&calculation_sheets) {
            "true"
        } else {
            "false"
        }
    )
}

fn contains_qualified_reference(expr: &elixcee::formula::FormulaExpr) -> bool {
    use elixcee::formula::FormulaExpr;
    match expr {
        FormulaExpr::CellRef { sheet, .. } | FormulaExpr::Range { sheet, .. } => sheet.is_some(),
        FormulaExpr::BinOp { lhs, rhs, .. } => {
            contains_qualified_reference(lhs) || contains_qualified_reference(rhs)
        }
        FormulaExpr::UnaryMinus(inner) => contains_qualified_reference(inner),
        FormulaExpr::FuncCall { args, .. } => args.iter().any(contains_qualified_reference),
        FormulaExpr::Call { callee, args } => {
            contains_qualified_reference(callee) || args.iter().any(contains_qualified_reference)
        }
        FormulaExpr::Number(_)
        | FormulaExpr::Str(_)
        | FormulaExpr::Bool(_)
        | FormulaExpr::Omitted => false,
    }
}

fn sheet_cell_to_variant(cell: &SheetCell) -> Variant {
    match cell {
        SheetCell::Integer(value) => Variant::Integer(*value),
        SheetCell::Float(value) => Variant::Float(*value),
        SheetCell::Str(value) => Variant::Str(value.clone()),
        SheetCell::Bool(value) => Variant::Boolean(*value),
        SheetCell::Error(value) => Variant::Error(value.clone()),
    }
}

fn calculation_sheets(
    workbook: &BufferWorkbook,
) -> HashMap<String, HashMap<(u32, u32), CellContent>> {
    let mut sheets = HashMap::new();
    for bs in &workbook.sheets {
        let mut cells = HashMap::with_capacity(bs.sheet.cells.len().max(bs.formulas.len()));
        for (&position, cell) in &bs.sheet.cells {
            cells.insert(
                position,
                CellContent {
                    formula: bs.formulas.get(&position).map(|f| format!("={f}")),
                    value: sheet_cell_to_variant(cell),
                },
            );
        }
        for (&position, formula) in &bs.formulas {
            cells.entry(position).or_insert_with(|| CellContent {
                formula: Some(format!("={formula}")),
                value: Variant::Empty,
            });
        }
        sheets.insert(bs.sheet.name.to_ascii_lowercase(), cells);
    }
    sheets
}

fn sync_workbook_values(
    workbook: &mut BufferWorkbook,
    sheets: &HashMap<String, HashMap<(u32, u32), CellContent>>,
) {
    for bs in &mut workbook.sheets {
        let Some(cells) = sheets.get(&bs.sheet.name.to_ascii_lowercase()) else {
            continue;
        };
        for &position in bs.formulas.keys() {
            if let Some(cell) = cells.get(&position) {
                bs.sheet
                    .cells
                    .insert(position, variant_to_sheet_cell(&cell.value));
            }
        }
        for (&position, cell) in cells {
            if cell.formula.is_none() {
                bs.sheet
                    .cells
                    .insert(position, variant_to_sheet_cell(&cell.value));
            }
        }
    }
}

fn variant_to_sheet_cell(value: &Variant) -> SheetCell {
    match value {
        Variant::Integer(value) => SheetCell::Integer(*value),
        Variant::Float(value) => SheetCell::Float(*value),
        Variant::Str(value) => SheetCell::Str(value.clone()),
        Variant::Boolean(value) => SheetCell::Bool(*value),
        Variant::Date(value) => SheetCell::Integer(*value),
        Variant::Error(value) => SheetCell::Error(value.clone()),
        Variant::Empty | Variant::Null => SheetCell::Str(String::new()),
        Variant::Array(values) => values
            .first()
            .map(variant_to_sheet_cell)
            .unwrap_or_else(|| SheetCell::Str(String::new())),
        Variant::VbaArray(_) | Variant::Record(_) => SheetCell::Str(value.to_string()),
    }
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
    if !wb.defined_names.is_empty() {
        out.push_str(",\"Workbook\":{\"Names\":[");
        for (i, defined_name) in wb.defined_names.iter().enumerate() {
            if i > 0 {
                out.push(',');
            }
            out.push_str("{\"Name\":");
            out.push_str(&json_string(&defined_name.name));
            out.push_str(",\"Ref\":");
            out.push_str(&json_string(&defined_name.raw_text));
            if let Some(sheet) = defined_name.local_sheet_id {
                out.push_str(",\"Sheet\":");
                out.push_str(&sheet.to_string());
            }
            out.push('}');
        }
        out.push_str("]}");
    }
    out.push('}');
    out
}

fn table_filter_column_json(column: &elixcee::reader::FilterColumn) -> String {
    let mut out = format!(
        "{{\"colId\":{},\"hiddenButton\":{},\"showButton\":{},\"criteria\":",
        column.col_offset, column.hidden_button, column.show_button
    );
    match &column.criteria {
        FilterCriteria::Values(values) => {
            out.push_str("{\"kind\":\"values\",\"values\":[");
            for (index, value) in values.iter().enumerate() {
                if index > 0 {
                    out.push(',');
                }
                out.push_str(&json_string(value));
            }
            out.push_str("]}");
        }
        FilterCriteria::Blank => out.push_str("{\"kind\":\"blank\"}"),
        FilterCriteria::Custom {
            op1,
            val1,
            and,
            op2,
            val2,
        } => {
            out.push_str(&format!(
                "{{\"kind\":\"custom\",\"op1\":{},\"val1\":{},\"and\":{}",
                json_string(op1),
                json_string(val1),
                and
            ));
            if let (Some(op2), Some(val2)) = (op2, val2) {
                out.push_str(&format!(
                    ",\"op2\":{},\"val2\":{}",
                    json_string(op2),
                    json_string(val2)
                ));
            }
            out.push('}');
        }
        FilterCriteria::Top10 { top, percent, val } => {
            out.push_str(&format!(
                "{{\"kind\":\"top10\",\"top\":{},\"percent\":{},\"val\":{}}}",
                top, percent, val
            ));
        }
        FilterCriteria::DateGroup(_) => out.push_str("{\"kind\":\"dateGroup\"}"),
    }
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
            bs.cell_styles.get(&(row, col)),
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

    if !sheet.data_validations.is_empty() {
        out.push_str(",\"!dataValidations\":[");
        for (index, validation) in sheet.data_validations.iter().enumerate() {
            if index > 0 {
                out.push(',');
            }
            out.push_str("{\"type\":");
            out.push_str(&json_string(&validation.validation_type));
            out.push_str(",\"sqref\":[");
            for (range_index, range) in validation.sqref.iter().enumerate() {
                if range_index > 0 {
                    out.push(',');
                }
                out.push_str(&json_string(&format_rect(range)));
            }
            out.push_str("]}");
        }
        out.push(']');
    }

    if !sheet.tables.is_empty() {
        out.push_str(",\"!tables\":[");
        for (index, table) in sheet.tables.iter().enumerate() {
            if index > 0 {
                out.push(',');
            }
            out.push_str("{\"name\":");
            out.push_str(&json_string(&table.name));
            out.push_str(",\"displayName\":");
            out.push_str(&json_string(&table.display_name));
            out.push_str(",\"ref\":");
            out.push_str(&json_string(&format_rect(&table.ref_range)));
            if let Some(auto_filter_ref) = &table.auto_filter_ref {
                out.push_str(",\"autoFilterRef\":");
                out.push_str(&json_string(&format_rect(auto_filter_ref)));
                if !table.autofilter_columns.is_empty() {
                    out.push_str(",\"autoFilterColumns\":[");
                    for (column_index, column) in table.autofilter_columns.iter().enumerate() {
                        if column_index > 0 {
                            out.push(',');
                        }
                        out.push_str(&table_filter_column_json(column));
                    }
                    out.push(']');
                }
            }
            if !table.columns.is_empty() {
                out.push_str(",\"columns\":[");
                for (column_index, column) in table.columns.iter().enumerate() {
                    if column_index > 0 {
                        out.push(',');
                    }
                    out.push_str("{\"name\":");
                    out.push_str(&json_string(&column.name));
                    out.push('}');
                }
                out.push(']');
            }
            if let Some(style_name) = &table.style_name {
                out.push_str(",\"styleName\":");
                out.push_str(&json_string(style_name));
            }
            out.push('}');
        }
        out.push(']');
    }

    if !bs.charts.is_empty() {
        out.push_str(",\"!charts\":[");
        for (index, chart) in bs.charts.iter().enumerate() {
            if index > 0 {
                out.push(',');
            }
            out.push_str("{\"ref\":");
            out.push_str(&json_string(&format_rect(&chart.ref_range)));
            out.push_str(",\"type\":");
            out.push_str(&json_string(&chart.chart_type));
            out.push_str(",\"title\":");
            out.push_str(&json_string(&chart.title));
            if let Some(title) = &chart.x_axis_title {
                out.push_str(",\"xAxisTitle\":");
                out.push_str(&json_string(title));
            }
            if let Some(title) = &chart.y_axis_title {
                out.push_str(",\"yAxisTitle\":");
                out.push_str(&json_string(title));
            }
            out.push_str(",\"legend\":");
            out.push_str(if chart.legend { "true" } else { "false" });
            if !chart.series_colors.is_empty() {
                out.push_str(",\"colors\":[");
                for (color_index, color) in chart.series_colors.iter().enumerate() {
                    if color_index > 0 {
                        out.push(',');
                    }
                    out.push_str(&json_string(color));
                }
                out.push(']');
            }
            out.push_str(&format!(
                ",\"widthCols\":{},\"heightRows\":{}",
                chart.width_cols, chart.height_rows
            ));
            out.push('}');
        }
        out.push(']');
    }

    if !bs.comment_notes.is_empty() {
        out.push_str(",\"!comments\":[");
        for (index, comment) in bs.comment_notes.iter().enumerate() {
            if index > 0 {
                out.push(',');
            }
            out.push_str("{\"ref\":");
            out.push_str(&json_string(&cell_ref(comment.cell.0, comment.cell.1)));
            out.push_str(",\"author\":");
            out.push_str(&json_string(&comment.author));
            out.push_str(",\"text\":");
            out.push_str(&json_string(&comment.text));
            out.push('}');
        }
        out.push(']');
    }

    if !bs.conditional_formats.is_empty() {
        out.push_str(",\"!conditionalFormats\":[");
        for (index, rule) in bs.conditional_formats.iter().enumerate() {
            if index > 0 {
                out.push(',');
            }
            out.push_str("{\"type\":");
            out.push_str(&json_string(&rule.rule_type));
            if let Some(operator) = &rule.operator {
                out.push_str(",\"operator\":");
                out.push_str(&json_string(operator));
            }
            out.push_str(",\"sqref\":[");
            for (range_index, range) in rule.sqref.iter().enumerate() {
                if range_index > 0 {
                    out.push(',');
                }
                out.push_str(&json_string(&format_rect(range)));
            }
            out.push_str("],\"formula\":");
            out.push_str(&json_string(&rule.formula1));
            if let Some(formula2) = &rule.formula2 {
                out.push_str(",\"formula2\":");
                out.push_str(&json_string(formula2));
            }
            if let Some(priority) = rule.priority {
                out.push_str(&format!(",\"priority\":{}", priority));
            }
            if rule.stop_if_true {
                out.push_str(",\"stopIfTrue\":true");
            }
            if let Some(dxf) = &rule.dxf {
                out.push_str(",\"dxf\":{");
                let mut first = true;
                if dxf.bold || dxf.italic || dxf.underline || dxf.font_color.is_some() {
                    out.push_str("\"font\":{");
                    let mut font_first = true;
                    if dxf.bold {
                        out.push_str("\"bold\":true");
                        font_first = false;
                    }
                    if dxf.italic {
                        if !font_first {
                            out.push(',');
                        }
                        out.push_str("\"italic\":true");
                        font_first = false;
                    }
                    if dxf.underline {
                        if !font_first {
                            out.push(',');
                        }
                        out.push_str("\"underline\":true");
                        font_first = false;
                    }
                    if let Some(color) = &dxf.font_color {
                        if !font_first {
                            out.push(',');
                        }
                        out.push_str("\"color\":{\"rgb\":");
                        out.push_str(&json_string(color));
                        out.push('}');
                    }
                    out.push('}');
                    first = false;
                }
                if let Some(color) = &dxf.fill_color {
                    if !first {
                        out.push(',');
                    }
                    out.push_str("\"fill\":{\"fgColor\":{\"rgb\":");
                    out.push_str(&json_string(color));
                    out.push_str("}}");
                }
                out.push('}');
            }
            out.push('}');
        }
        out.push(']');
    }

    if let Some(pane) = &bs.freeze_pane {
        out.push_str(",\"!freezePane\":{\"rows\":");
        out.push_str(&pane.rows.to_string());
        out.push_str(",\"cols\":");
        out.push_str(&pane.cols.to_string());
        out.push('}');
    }

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

fn cell_json(
    cell: &SheetCell,
    formula: Option<&String>,
    fmt_id: Option<&u32>,
    style: Option<&elixcee::reader::CellStyleDef>,
) -> String {
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
    if let Some(style) = style {
        let mut parts = Vec::new();
        if style.bold || style.italic || style.underline || style.font_color.is_some() {
            let mut font = String::new();
            if style.bold {
                font.push_str("\"bold\":true");
            }
            if style.italic {
                if !font.is_empty() {
                    font.push(',');
                }
                font.push_str("\"italic\":true");
            }
            if style.underline {
                if !font.is_empty() {
                    font.push(',');
                }
                font.push_str("\"underline\":true");
            }
            if let Some(color) = &style.font_color {
                if !font.is_empty() {
                    font.push(',');
                }
                font.push_str("\"color\":{\"rgb\":");
                font.push_str(&json_string(color));
                font.push('}');
            }
            parts.push(format!("\"font\":{{{}}}", font));
        }
        if let Some(color) = &style.fill_color {
            parts.push(format!(
                "\"fill\":{{\"fgColor\":{{\"rgb\":{}}}}}",
                json_string(color)
            ));
        }
        let border_specs = [
            (
                "bottom",
                style.border_bottom.as_ref(),
                style.border_bottom_color.as_ref(),
            ),
            (
                "left",
                style.border_left.as_ref(),
                style.border_left_color.as_ref(),
            ),
            (
                "right",
                style.border_right.as_ref(),
                style.border_right_color.as_ref(),
            ),
            (
                "top",
                style.border_top.as_ref(),
                style.border_top_color.as_ref(),
            ),
        ];
        let border_parts: Vec<_> = border_specs
            .into_iter()
            .filter_map(|(name, border, color)| {
                border.map(|value| {
                    let color = color
                        .map(|v| format!(",\"color\":{{\"rgb\":{}}}", json_string(v)))
                        .unwrap_or_default();
                    format!(
                        "\"{}\":{{\"style\":{}{}{}}}",
                        name,
                        json_string(value),
                        color,
                        ""
                    )
                })
            })
            .collect();
        if !border_parts.is_empty() {
            parts.push(format!("\"border\":{{{}}}", border_parts.join(",")));
        }
        if style.horizontal.is_some() || style.vertical.is_some() || style.wrap_text.is_some() {
            let mut alignment = String::new();
            if let Some(value) = &style.horizontal {
                alignment.push_str("\"horizontal\":");
                alignment.push_str(&json_string(value));
            }
            if let Some(value) = &style.vertical {
                if !alignment.is_empty() {
                    alignment.push(',');
                }
                alignment.push_str("\"vertical\":");
                alignment.push_str(&json_string(value));
            }
            if let Some(value) = style.wrap_text {
                if !alignment.is_empty() {
                    alignment.push(',');
                }
                alignment.push_str(&format!("\"wrapText\":{}", value));
            }
            parts.push(format!("\"alignment\":{{{}}}", alignment));
        }
        if !parts.is_empty() {
            out.push_str(",\"s\":{");
            out.push_str(&parts.join(","));
            out.push('}');
        }
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

fn format_rect(rect: &((u32, u32), (u32, u32))) -> String {
    let ((r1, c1), (r2, c2)) = *rect;
    let start = format!("{}{}", col_letters(c1), r1);
    if r1 == r2 && c1 == c2 {
        start
    } else {
        format!("{}:{}{}", start, col_letters(c2), r2)
    }
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
                row_heights: HashMap::new(),
                column_widths: Vec::new(),
                row_styles: HashMap::new(),
                column_styles: Vec::new(),
                tables: Vec::new(),
                data_validations: Vec::new(),
                conditional_format_ranges: Vec::new(),
                comment_cells: Vec::new(),
                autofilter: None,
            },
            formulas: HashMap::new(),
            dimension: None,
            style_ids: HashMap::new(),
            charts: Vec::new(),
            comment_notes: Vec::new(),
            conditional_formats: Vec::new(),
            cell_styles: HashMap::new(),
            freeze_pane: None,
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
            defined_names: Vec::new(),
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
    fn structured_reference_rewrite_respects_formula_token_boundaries() {
        assert_eq!(
            replace_table_ref_case_insensitive("=SUM(Sales[Amount])", "Sales[Amount]", "B2:B3"),
            "=SUM(B2:B3)"
        );
        assert_eq!(
            replace_table_ref_case_insensitive("=\"Sales[Amount]\"", "Sales[Amount]", "B2:B3"),
            "=\"Sales[Amount]\""
        );
        assert_eq!(
            replace_table_ref_case_insensitive("=MySales[Amount]", "Sales[Amount]", "B2:B3"),
            "=MySales[Amount]"
        );
        assert_eq!(
            replace_table_ref_case_insensitive(
                "=\"Sales[\"\"Amount\"\"]\"",
                "Sales[Amount]",
                "B2:B3"
            ),
            "=\"Sales[\"\"Amount\"\"]\""
        );
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

    #[test]
    fn worksheet_json_projects_data_validation_type_and_ranges() {
        let mut s = sheet("Sheet1", vec![]);
        s.sheet
            .data_validations
            .push(elixcee::reader::DataValidationRule {
                validation_type: "list".to_string(),
                operator: None,
                formula1: Some("Yes,No".to_string()),
                formula2: None,
                allow_blank: true,
                show_input_message: false,
                prompt_title: None,
                prompt: None,
                show_error_message: true,
                error_style: None,
                error_title: None,
                error: None,
                sqref: vec![((1, 5), (1, 5)), ((2, 5), (4, 5))],
                dirty: false,
                raw_span: String::new(),
            });
        let json = workbook_json(&wb1(s));
        assert!(json.contains(r#""!dataValidations":[{"type":"list","sqref":["E1","E2:E4"]}]"#));
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
            defined_names: Vec::new(),
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
    fn workbook_json_exposes_defined_names_with_scope() {
        let mut wb = wb1(sheet("Sheet1", vec![]));
        wb.defined_names.push(elixcee::reader::XlsxDefinedName {
            name: "SalesRange".to_string(),
            local_sheet_id: Some(0),
            raw_text: "Sheet1!$A$1:$A$2".to_string(),
        });
        let json = workbook_json(&wb);
        assert!(json.contains(
            r#""Workbook":{"Names":[{"Name":"SalesRange","Ref":"Sheet1!$A$1:$A$2","Sheet":0}]}"#
        ));
    }

    #[test]
    fn normalize_defined_name_ref_supports_quoted_and_local_absolute_refs() {
        assert_eq!(
            normalize_defined_name_ref("'Sheet 1'!$A$1:$B$2"),
            "Sheet 1!$A$1:$B$2"
        );
        assert_eq!(normalize_defined_name_ref("$A$1"), "A1");
    }

    #[test]
    fn workbook_json_always_includes_date1904() {
        let mut wb = wb1(sheet("Sheet1", vec![]));
        wb.date1904 = true;
        let json = workbook_json(&wb);
        assert!(json.contains(r#""!date1904":true"#));
    }

    #[test]
    fn editor_transaction_coalesces_multiple_typed_writes_into_one_undo() {
        let workbook = wb1(sheet("Sheet1", vec![]));
        let sheets = calculation_sheets(&workbook);
        let mut editor = WorkbookEditor {
            workbook,
            sheets,
            undo: Vec::new(),
            redo: Vec::new(),
            transaction: None,
            transaction_dirty: false,
        };

        editor.begin_transaction().unwrap();
        editor.set_string("Sheet1", 1, 1, "planned").unwrap();
        editor.set_boolean("Sheet1", 1, 2, true).unwrap();
        assert!(editor.commit_transaction());
        assert!(editor.can_undo());
        assert!(editor.undo());
        let snapshot = editor.snapshot();
        assert!(!snapshot.contains("planned"));
        assert!(!snapshot.contains("\"B1\""));
        assert!(!editor.can_undo());
        assert!(editor.redo());
        let snapshot = editor.snapshot();
        assert!(snapshot.contains("planned"));
        assert!(snapshot.contains("\"B1\""));
    }

    #[test]
    fn formula_array_spill_materializes_empty_horizontal_targets_without_overwrite() {
        let mut sheets = HashMap::from([(
            "sheet1".to_string(),
            HashMap::from([
                (
                    (0, 0),
                    CellContent {
                        formula: Some("=SEQUENCE(1,3)".to_string()),
                        value: Variant::Array(vec![
                            Variant::Integer(1),
                            Variant::Integer(2),
                            Variant::Integer(3),
                        ]),
                    },
                ),
                (
                    (0, 2),
                    CellContent {
                        formula: None,
                        value: Variant::Str("kept".to_string()),
                    },
                ),
            ]),
        )]);
        materialize_formula_array_spills(&mut sheets).unwrap();
        assert!(!sheets["sheet1"].contains_key(&(0, 1)));
        assert_eq!(
            sheets["sheet1"][&(0, 0)].value,
            Variant::Str("#SPILL!".to_string())
        );
        assert_eq!(
            sheets["sheet1"][&(0, 2)].value,
            Variant::Str("kept".to_string())
        );
    }

    #[test]
    fn sequence_spill_materializes_vertical_and_rectangular_shapes() {
        let mut sheets = HashMap::from([(
            "sheet1".to_string(),
            HashMap::from([(
                (0, 0),
                CellContent {
                    formula: Some("=SEQUENCE(2,2)".to_string()),
                    value: Variant::Array(vec![
                        Variant::Integer(1),
                        Variant::Integer(2),
                        Variant::Integer(3),
                        Variant::Integer(4),
                    ]),
                },
            )]),
        )]);
        materialize_formula_array_spills(&mut sheets).unwrap();
        assert_eq!(sheets["sheet1"][&(0, 1)].value, Variant::Integer(2));
        assert_eq!(sheets["sheet1"][&(1, 0)].value, Variant::Integer(3));
        assert_eq!(sheets["sheet1"][&(1, 1)].value, Variant::Integer(4));
    }

    #[test]
    fn formula_array_shape_recovers_wrapped_vector_layouts() {
        assert_eq!(
            formula_array_shape(Some("=RANDARRAY(3,2)"), 6),
            ArrayShape::new(3, 2)
        );
        assert_eq!(
            formula_array_shape(Some("=WRAPROWS(SEQUENCE(5),2)"), 5),
            ArrayShape::new(3, 2)
        );
        assert_eq!(
            formula_array_shape(Some("=WRAPCOLS(SEQUENCE(5),2)"), 5),
            ArrayShape::new(2, 3)
        );
        assert_eq!(
            formula_array_shape(Some("=UNIQUE(A1:A5)"), 5),
            ArrayShape::new(1, 5)
        );
    }
}
