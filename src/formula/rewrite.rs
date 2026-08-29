//! Reference rewriting for row/column insert-delete, workbook-wide (0.14.0-A,
//! 0.14.0-A2).
//!
//! Patches only the text span of each affected cell/range reference --
//! everything else in the formula (operators, function names, literals,
//! whitespace, unaffected references) is left untouched. See
//! `super::parser::parse_with_refs` for how spans are captured; this module
//! never builds or walks a full expression tree, and there is deliberately
//! no general AST-to-text serializer -- see `internal_docs/ROADMAP.md`'s
//! 0.14.0-A note for why a span-patching design replaced that idea.
//!
//! A formula lives on one sheet (`host_sheet_key`) and may reference cells on
//! any sheet, unqualified (bare `A1`, implicitly the host sheet) or qualified
//! (`Sheet2!A1`, explicitly named). A structural edit happens on exactly one
//! sheet (`edited_sheet_key`). A reference is only rewritten when it actually
//! points at the edited sheet:
//!
//! - unqualified: only if `host_sheet_key == edited_sheet_key`
//! - qualified: only if the qualifier's name resolves (case-insensitively) to
//!   `edited_sheet_key`, regardless of which sheet hosts the formula
//!
//! Cross-sheet formula *evaluation* is a separate, still-unsupported concern
//! (see `eval::references_another_sheet`) -- this module only rewrites text,
//! it never reads a cell's value.
//!
//! Also covers sheet rename (`rename_sheet_references`, qualifier-text-only
//! rewrite) and same-sheet range move (`translate_references_for_move`,
//! 0.14.0-A4 -- reference-identity tracking, see its own doc comment and
//! `internal_docs/range-move-0.14.0-a4-design.md` for the semantics
//! research behind its design). Range move is scoped to same-sheet moves
//! only this round; workbook-wide qualified-reference following for a moved
//! range is a follow-up, matching how 0.14.0-A2 extended insert/delete from
//! same-sheet to workbook-wide after the same-sheet case shipped first.

use std::collections::HashMap;

use super::eval::col_to_letter;
use super::parser::{RefOccurrence, parse_with_refs};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RefAxis {
    Row,
    Col,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StructuralEdit {
    Insert { at: u32, count: u32 },
    Delete { at: u32, count: u32 },
}

/// Generic outcome of shifting a single coordinate by a structural edit --
/// not formula-specific despite living in this module (its historical home,
/// alongside its own tests); `pub(crate)` so 0.14.0-B's cell-metadata
/// transform can reuse the exact same arithmetic instead of re-deriving
/// equivalent logic that could silently drift from this one (see
/// `internal_docs/cell-metadata-transform-0.14.0-b-design.md` §6). What
/// `Deleted` MEANS differs per consumer: a formula reference becomes
/// `#REF!` text; a merge/style/hidden-interval entry is simply dropped.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum CellShift {
    Unchanged,
    Moved(u32),
    Deleted,
}

/// A single cell reference's coordinate on the edited axis. Deletion of a
/// row/column the reference points at is unconditionally `#REF!` -- unlike a
/// range corner, there's no surviving neighbor for a single cell to clamp to.
pub(crate) fn shift_cell_coord(idx: u32, edit: StructuralEdit) -> CellShift {
    match edit {
        StructuralEdit::Insert { at, count } if count > 0 && idx >= at => {
            CellShift::Moved(idx + count)
        }
        StructuralEdit::Delete { at, count } if count > 0 && idx >= at => {
            if idx < at + count {
                CellShift::Deleted
            } else {
                CellShift::Moved(idx - count)
            }
        }
        _ => CellShift::Unchanged,
    }
}

/// Lower bound of a range corner: on deletion, an index inside the deleted
/// band clamps to `at` (the surviving row/col that slides into its place).
/// `pub(crate)`, same reuse reasoning as `CellShift`/`shift_cell_coord` above.
pub(crate) fn shift_bound_low(idx: u32, edit: StructuralEdit) -> u32 {
    match edit {
        StructuralEdit::Insert { at, count } => {
            if idx >= at {
                idx + count
            } else {
                idx
            }
        }
        StructuralEdit::Delete { at, count } => {
            if idx < at {
                idx
            } else if idx < at + count {
                at
            } else {
                idx - count
            }
        }
    }
}

/// Upper bound of a range corner: on deletion, an index inside the deleted
/// band clamps to `at - 1` (the surviving row/col just before the deletion).
/// Returned as `i64` so a range whose top clamps to `at` and whose bottom
/// clamps to `at - 1` correctly reads as collapsed (`low > high`) even when
/// `at` is 1 -- 0 is not a valid Excel row/col but is a valid sentinel here.
/// `pub(crate)`, same reuse reasoning as `CellShift`/`shift_cell_coord` above.
pub(crate) fn shift_bound_high(idx: u32, edit: StructuralEdit) -> i64 {
    match edit {
        StructuralEdit::Insert { at, count } => {
            if idx >= at {
                (idx + count) as i64
            } else {
                idx as i64
            }
        }
        StructuralEdit::Delete { at, count } => {
            if idx < at {
                idx as i64
            } else if idx < at + count {
                at as i64 - 1
            } else {
                (idx - count) as i64
            }
        }
    }
}

fn format_cell_ref(col: u32, row: u32, abs_col: bool, abs_row: bool) -> String {
    let mut s = String::new();
    if abs_col {
        s.push('$');
    }
    s.push_str(&col_to_letter(col));
    if abs_row {
        s.push('$');
    }
    s.push_str(&row.to_string());
    s
}

/// Does this reference (qualified or not) point at the sheet being edited?
/// `host_sheet_key`/`edited_sheet_key` are already-lowercased sheet keys
/// (this codebase's own identity convention, see `ensure_sheet_at`); a
/// qualifier's `normalized_name` is lowercased here at compare-time since it
/// preserves the formula's original casing (e.g. `Sheet2!`, not `sheet2!`).
fn ref_targets_edited_sheet(
    sheet: &Option<super::ast::SheetQualifier>,
    host_sheet_key: &str,
    edited_sheet_key: &str,
) -> bool {
    match sheet {
        None => host_sheet_key == edited_sheet_key,
        Some(q) => q.normalized_name.to_lowercase() == edited_sheet_key,
    }
}

/// Splice non-overlapping `(start, end, replacement)` patches (char-offset
/// spans, half-open) into `input`, sorted by start position.
fn apply_patches(input: &str, mut patches: Vec<(usize, usize, String)>) -> String {
    patches.sort_by_key(|p| p.0);
    let chars: Vec<char> = input.chars().collect();
    let mut out = String::new();
    let mut cursor = 0usize;
    for (start, end, replacement) in patches {
        out.extend(&chars[cursor..start]);
        out.push_str(&replacement);
        cursor = end;
    }
    out.extend(&chars[cursor..]);
    out
}

/// Rewrite `formula`'s cell/range references for a row or column insert/delete
/// happening on `edited_sheet_key`. `formula` is hosted on `host_sheet_key`
/// (the sheet its own cell lives on) -- an unqualified reference is relative
/// to the host sheet, so it only shifts when the host sheet IS the edited
/// sheet; a qualified reference (`Sheet2!A1`) shifts whenever it names the
/// edited sheet, regardless of which sheet hosts the formula. See this
/// module's doc comment for the full targeting rule.
///
/// Returns `Ok(None)` when no reference in the formula was affected (caller
/// can skip rewriting the stored formula string). A reference that falls
/// entirely inside a deleted band becomes `#REF!` (the qualifier, if any, is
/// preserved -- only the coordinate collapses, matching real Excel).
///
/// `$` (absolute) flags never change whether a reference shifts here --
/// unlike copy/fill, real Excel shifts every reference on row/column
/// insert-delete regardless of `$` -- they're only preserved as-is in the
/// rewritten text. A formula this parser can't parse (external workbook
/// refs, 3D refs, and anything else 0.14.0-A2 doesn't cover) is reported as
/// `Err` by `parse_with_refs` and must be left completely untouched by the
/// caller -- see `Vm::rewrite_formulas_for_structural_edit`.
///
/// `formula` is normalized internally the same way `parse`/`parse_with_refs`
/// normalize it (leading `=` and surrounding whitespace stripped); the
/// returned string carries no leading `=`, matching how a `<f>` element's
/// text is stored. A caller whose own storage convention expects a leading
/// `=` (see `CellContent::formula`'s doc comment in `src/vm/mod.rs`) must
/// reapply it itself, same as `xlsx_cell_xml` already does defensively when
/// writing formulas back out.
pub fn shift_references(
    formula: &str,
    host_sheet_key: &str,
    edited_sheet_key: &str,
    axis: RefAxis,
    edit: StructuralEdit,
) -> Result<Option<String>, String> {
    let is_noop_edit = matches!(
        edit,
        StructuralEdit::Insert { count: 0, .. } | StructuralEdit::Delete { count: 0, .. }
    );
    if is_noop_edit {
        return Ok(None);
    }

    let input = formula.trim().trim_start_matches('=').to_string();
    let (_, refs) = parse_with_refs(&input)?;
    if refs.is_empty() {
        return Ok(None);
    }

    let mut patches: Vec<(usize, usize, String)> = Vec::new();
    for r in &refs {
        match r {
            RefOccurrence::Cell {
                span,
                col,
                row,
                abs_col,
                abs_row,
                sheet,
            } => {
                if !ref_targets_edited_sheet(sheet, host_sheet_key, edited_sheet_key) {
                    continue;
                }
                let idx = match axis {
                    RefAxis::Row => *row,
                    RefAxis::Col => *col,
                };
                match shift_cell_coord(idx, edit) {
                    CellShift::Unchanged => {}
                    CellShift::Deleted => {
                        patches.push((span.0, span.1, "#REF!".to_string()));
                    }
                    CellShift::Moved(new_idx) => {
                        let (new_col, new_row) = match axis {
                            RefAxis::Row => (*col, new_idx),
                            RefAxis::Col => (new_idx, *row),
                        };
                        patches.push((
                            span.0,
                            span.1,
                            format_cell_ref(new_col, new_row, *abs_col, *abs_row),
                        ));
                    }
                }
            }
            RefOccurrence::Range {
                span,
                c1,
                r1,
                abs_c1,
                abs_r1,
                c1_span,
                c2,
                r2,
                abs_c2,
                abs_r2,
                c2_span,
                sheet,
            } => {
                if !ref_targets_edited_sheet(sheet, host_sheet_key, edited_sheet_key) {
                    continue;
                }
                let (v1, v2) = match axis {
                    RefAxis::Row => (*r1, *r2),
                    RefAxis::Col => (*c1, *c2),
                };
                // Ranges are order-independent here (matches this codebase's
                // own evaluator, which always does `r1.min(r2)`/`.max(r2)`).
                let (v1_is_low, low, high) = if v1 <= v2 {
                    (true, v1, v2)
                } else {
                    (false, v2, v1)
                };
                let new_low = shift_bound_low(low, edit);
                let new_high = shift_bound_high(high, edit);
                if new_low as i64 > new_high {
                    // The whole range fell inside a deleted band -- the
                    // entire reference collapses, not just one corner.
                    patches.push((span.0, span.1, "#REF!".to_string()));
                    continue;
                }
                let new_high = new_high as u32;
                let (new_v1, new_v2) = if v1_is_low {
                    (new_low, new_high)
                } else {
                    (new_high, new_low)
                };
                if new_v1 != v1 {
                    let (new_c1, new_r1) = match axis {
                        RefAxis::Row => (*c1, new_v1),
                        RefAxis::Col => (new_v1, *r1),
                    };
                    patches.push((
                        c1_span.0,
                        c1_span.1,
                        format_cell_ref(new_c1, new_r1, *abs_c1, *abs_r1),
                    ));
                }
                if new_v2 != v2 {
                    let (new_c2, new_r2) = match axis {
                        RefAxis::Row => (*c2, new_v2),
                        RefAxis::Col => (new_v2, *r2),
                    };
                    patches.push((
                        c2_span.0,
                        c2_span.1,
                        format_cell_ref(new_c2, new_r2, *abs_c2, *abs_r2),
                    ));
                }
            }
        }
    }

    if patches.is_empty() {
        return Ok(None);
    }
    Ok(Some(apply_patches(&input, patches)))
}

/// Would an unquoted sheet name written into a formula (`Name!A1`, no
/// surrounding `'...'`) parse back out to exactly this `name`? Mirrors
/// `FormulaParser::try_parse_sheet_qualifier`'s unquoted-branch acceptance
/// grammar exactly (Unicode-aware first-letter/alphanumeric/`_`, same as
/// there) -- this function only exists so that grammar has a single
/// semantic inverse, not a second, independently-drifting copy of the rule.
fn sheet_name_needs_quoting(name: &str) -> bool {
    let mut chars = name.chars();
    match chars.next() {
        Some(c) if c.is_alphabetic() || c == '_' => {}
        _ => return true,
    }
    !chars.all(|c| c.is_alphanumeric() || c == '_')
}

/// Format `name` as a formula sheet-qualifier (no trailing `!`), quoting and
/// `'`-escaping it only when required -- never preserves whatever quoting
/// style the reference being replaced used, since that's the OLD name's
/// business, not the new one's.
fn format_sheet_qualifier(name: &str) -> String {
    if sheet_name_needs_quoting(name) {
        format!("'{}'", name.replace('\'', "''"))
    } else {
        name.to_string()
    }
}

/// A rectangular range on one sheet, 1-indexed, inclusive on both corners.
/// `r1 <= r2` and `c1 <= c2` is a precondition the caller must normalize
/// before constructing this -- unlike a formula's own reference (which may
/// legitimately be written reversed, e.g. `B10:A1`, and is handled as such
/// below), a move's source rectangle has no such ambiguity to preserve.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct MoveRect {
    pub r1: u32,
    pub c1: u32,
    pub r2: u32,
    pub c2: u32,
}

impl MoveRect {
    /// `pub(crate)`, same reuse reasoning as `CellShift`/`shift_cell_coord`
    /// above -- 0.14.0-B's metadata transform needs the identical
    /// corner-containment check for range-move, not a re-derived copy.
    pub(crate) fn contains(&self, col: u32, row: u32) -> bool {
        row >= self.r1 && row <= self.r2 && col >= self.c1 && col <= self.c2
    }
}

/// Outcome of translating one formula's references for a range move.
/// Deliberately not `Result<Option<String>, String>` like
/// `shift_references`/`rename_sheet_references` above -- range move has a
/// third, genuinely different "leave this formula alone" case that the
/// other two rewrites don't: a range reference with exactly one corner
/// inside the moved rectangle. Real Excel's behavior for that shape is only
/// confirmed for one narrow sub-case (destination still inside the same
/// range, where the range shrinks rather than follows -- see
/// `internal_docs/range-move-0.14.0-a4-design.md` §3) and left unconfirmed
/// for the general case (§4-A), so this rewrite refuses to guess. Unlike a
/// genuine parse failure (`Err`, external/3D refs, same as the other two
/// rewrites -- non-fatal, caller skips just this formula), `Ambiguous` must
/// reject the WHOLE move, not just this one formula -- collapsing it into
/// `Ok(None)` would let the caller silently treat "don't know" the same as
/// "nothing to do."
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MoveRewrite {
    Unchanged,
    Rewritten(String),
    Ambiguous,
}

/// Rewrite `formula`'s references for a same-sheet range move: `source` is
/// moving by `(d_row, d_col)` on `move_sheet_key`. `formula` must itself be
/// hosted on `move_sheet_key` -- range move is scoped to same-sheet moves
/// for this round (`internal_docs/range-move-0.14.0-a4-design.md` §5), so
/// unlike `shift_references` this is never called workbook-wide; a caller
/// iterating other sheets' formulas has nothing to rewrite here by
/// construction (an unqualified reference on another sheet can't possibly
/// name this sheet's cells, and a qualified reference naming THIS sheet from
/// elsewhere is explicitly out of scope -- see the design doc's cross-sheet
/// open question).
///
/// Every reference (unqualified, or qualified naming `move_sheet_key`
/// itself) whose target cell falls inside `source` is translated by the move
/// offset -- this applies uniformly whether the referencing formula's own
/// cell is inside or outside `source`, matching real Excel's single
/// "reference tracks cell identity" mechanism (design doc §1), not two
/// separate "internal" and "external" rules. A reference targeting outside
/// `source`, or qualified naming a different sheet, is left untouched.
///
/// A range reference with both corners inside `source` translates as a
/// whole (every cell in it moved together). Both corners outside is a
/// no-op. Exactly one corner inside is the unresolved case above --
/// `Ok(MoveRewrite::Ambiguous)`.
///
/// `d_row`/`d_col` of `(0, 0)` is always `Unchanged` without parsing.
pub fn translate_references_for_move(
    formula: &str,
    move_sheet_key: &str,
    source: MoveRect,
    d_row: i64,
    d_col: i64,
) -> Result<MoveRewrite, String> {
    if d_row == 0 && d_col == 0 {
        return Ok(MoveRewrite::Unchanged);
    }

    let input = formula.trim().trim_start_matches('=').to_string();
    let (_, refs) = parse_with_refs(&input)?;
    if refs.is_empty() {
        return Ok(MoveRewrite::Unchanged);
    }

    // Reused with host==edited==move_sheet_key: `ref_targets_edited_sheet`
    // already encodes exactly the "unqualified or self-qualified" rule this
    // needs (an unqualified ref always matches when host==edited; a
    // qualifier only matches when it names that same sheet) -- see its own
    // doc comment above.
    let targets_move_sheet = |sheet: &Option<super::ast::SheetQualifier>| -> bool {
        ref_targets_edited_sheet(sheet, move_sheet_key, move_sheet_key)
    };

    let mut patches: Vec<(usize, usize, String)> = Vec::new();
    for r in &refs {
        match r {
            RefOccurrence::Cell {
                span,
                col,
                row,
                abs_col,
                abs_row,
                sheet,
            } => {
                if !targets_move_sheet(sheet) || !source.contains(*col, *row) {
                    continue;
                }
                let new_col = (*col as i64 + d_col) as u32;
                let new_row = (*row as i64 + d_row) as u32;
                patches.push((
                    span.0,
                    span.1,
                    format_cell_ref(new_col, new_row, *abs_col, *abs_row),
                ));
            }
            RefOccurrence::Range {
                c1,
                r1,
                abs_c1,
                abs_r1,
                c1_span,
                c2,
                r2,
                abs_c2,
                abs_r2,
                c2_span,
                sheet,
                ..
            } => {
                if !targets_move_sheet(sheet) {
                    continue;
                }
                let c1_inside = source.contains(*c1, *r1);
                let c2_inside = source.contains(*c2, *r2);
                match (c1_inside, c2_inside) {
                    (false, false) => {}
                    (true, true) => {
                        let new_c1 = (*c1 as i64 + d_col) as u32;
                        let new_r1 = (*r1 as i64 + d_row) as u32;
                        patches.push((
                            c1_span.0,
                            c1_span.1,
                            format_cell_ref(new_c1, new_r1, *abs_c1, *abs_r1),
                        ));
                        let new_c2 = (*c2 as i64 + d_col) as u32;
                        let new_r2 = (*r2 as i64 + d_row) as u32;
                        patches.push((
                            c2_span.0,
                            c2_span.1,
                            format_cell_ref(new_c2, new_r2, *abs_c2, *abs_r2),
                        ));
                    }
                    (true, false) | (false, true) => return Ok(MoveRewrite::Ambiguous),
                }
            }
        }
    }

    if patches.is_empty() {
        return Ok(MoveRewrite::Unchanged);
    }
    Ok(MoveRewrite::Rewritten(apply_patches(&input, patches)))
}

/// Rewrite every reference in `formula` qualified with `old_sheet_key`
/// (case-insensitively, this codebase's own sheet-key convention) to name
/// `new_sheet_name` instead -- only the qualifier text changes, coordinates
/// are never touched. The replacement is quoted/escaped per Excel's own
/// rules for `new_sheet_name` itself, regardless of how the old reference
/// was written (an unquoted `Sheet1!A1` becomes `'New Name'!A1` if
/// `new_sheet_name` needs quoting, and vice versa).
///
/// Unqualified references are never touched here: renaming a sheet doesn't
/// change what a bare `A1` means to a formula already living ON it (still
/// "this same sheet, whatever it's now called") -- only an explicit `!`
/// qualifier can name a sheet by a text that might now be stale.
///
/// Returns `Ok(None)` when nothing in `formula` was qualified with
/// `old_sheet_key`. Same normalization / leading-`=` / parse-error contract
/// as `shift_references` -- see its doc comment; a formula this parser can't
/// parse at all (external/3D references) must be left completely untouched
/// by the caller on `Err`, same as there.
pub fn rename_sheet_references(
    formula: &str,
    old_sheet_key: &str,
    new_sheet_name: &str,
) -> Result<Option<String>, String> {
    let input = formula.trim().trim_start_matches('=').to_string();
    let (_, refs) = parse_with_refs(&input)?;
    if refs.is_empty() {
        return Ok(None);
    }

    let new_qualifier = format_sheet_qualifier(new_sheet_name);
    let mut patches: Vec<(usize, usize, String)> = Vec::new();
    for r in &refs {
        let sheet = match r {
            RefOccurrence::Cell { sheet, .. } => sheet,
            RefOccurrence::Range { sheet, .. } => sheet,
        };
        if let Some(q) = sheet
            && q.normalized_name.to_lowercase() == old_sheet_key
        {
            patches.push((q.raw_span.0, q.raw_span.1, new_qualifier.clone()));
        }
    }

    if patches.is_empty() {
        return Ok(None);
    }
    Ok(Some(apply_patches(&input, patches)))
}

/// Outcome of `rewrite_defined_name_for_renames` for one `<definedName>`
/// value.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum DefinedNameRewrite {
    /// No tracked rename affects this value -- pass its original text
    /// through unmodified.
    Unchanged,
    /// At least one tracked rename's old sheet-qualifier was found and
    /// rewritten; carries the new value.
    Rewritten(String),
    /// This value could not be confirmed either way (parses under neither
    /// `parse_with_refs` nor the narrower reference-list grammar below) and
    /// its text plausibly names one of the renamed sheets -- drop this ONE
    /// `<definedName>` element rather than risk leaving it stale, matching
    /// the wholesale-drop precedent this narrows from a whole-file decision
    /// down to a single name.
    Drop,
}

/// Rewrites a `<definedName>` element's raw text value for every sheet rename
/// tracked in `renames` (`Vm::sheet_renames_since_load`: original lowercased
/// name -> current display name). Named ranges, and the `_xlnm.Print_Area`/
/// `_xlnm.Print_Titles` builtins, share this exact same text shape (see
/// `internal_docs/defined-names-rename-preservation-scoping.md`).
///
/// Two paths, tried in order:
/// 1. The general formula-reference rewriter (`rename_sheet_references`,
///    itself just `parse_with_refs` plus a qualifier patch) -- handles a
///    plain single reference/range and a genuine formula-valued name (e.g. a
///    dynamic named range using `OFFSET`/`INDEX`), the same machinery already
///    used for formula cells on rename. Can't parse a full-row/full-column
///    reference or a top-level comma-separated multi-area list.
/// 2. If that fails to parse the value at all, `rewrite_reference_list_for_renames`:
///    a narrower grammar covering exactly what path 1 is missing --
///    comma-separated, optionally sheet-qualified cell/range/full-row/
///    full-column references, nothing else. This is what real Print_Area
///    (often multi-area) and Print_Titles (almost always full-row and/or
///    full-column) values actually look like.
///
/// If NEITHER path can make sense of the value, it's left alone only when
/// none of the tracked renames' old names plausibly appear in it at all;
/// otherwise it's dropped individually (`Drop`) -- e.g. a dynamic named
/// range's formula that embeds an unsupported full-column reference inside a
/// function call, which path 2's grammar correctly refuses (a function call
/// is not a bare reference list, and text-splicing a qualifier inside syntax
/// this function doesn't understand is unsafe).
pub(crate) fn rewrite_defined_name_for_renames(
    value: &str,
    renames: &HashMap<String, String>,
) -> DefinedNameRewrite {
    if renames.is_empty() {
        return DefinedNameRewrite::Unchanged;
    }

    let mut current = value.to_string();
    let mut changed = false;
    let mut general_path_failed = false;
    for (old_key, new_name) in renames {
        match rename_sheet_references(&current, old_key, new_name) {
            Ok(Some(rewritten)) => {
                current = rewritten;
                changed = true;
            }
            Ok(None) => {}
            Err(_) => {
                general_path_failed = true;
                break;
            }
        }
    }
    if !general_path_failed {
        return if changed {
            DefinedNameRewrite::Rewritten(current)
        } else {
            DefinedNameRewrite::Unchanged
        };
    }

    match rewrite_reference_list_for_renames(value, renames) {
        Some(rewritten) if rewritten != value => DefinedNameRewrite::Rewritten(rewritten),
        Some(_) => DefinedNameRewrite::Unchanged,
        None => {
            let lower = value.to_lowercase();
            if renames.keys().any(|old_key| lower.contains(old_key)) {
                DefinedNameRewrite::Drop
            } else {
                DefinedNameRewrite::Unchanged
            }
        }
    }
}

/// A comma-separated reference-list rewrite: every item is optionally
/// sheet-qualified, followed by nothing but column letters / row digits /
/// `$` / `:` -- covers a plain multi-area reference and Print_Titles' full-row
/// (`$1:$3`) / full-column (`$A:$A`) shapes, none of which `parse_with_refs`
/// accepts. Returns `None` -- not "unchanged" -- the instant any item falls
/// outside this narrow grammar, since that means the value is something this
/// function doesn't understand (already tried and failed as a general formula
/// via the caller's first path) and text-splicing inside it would be a guess.
fn rewrite_reference_list_for_renames(
    value: &str,
    renames: &HashMap<String, String>,
) -> Option<String> {
    let mut out_items = Vec::with_capacity(value.matches(',').count() + 1);
    for item in value.split(',') {
        let trimmed = item.trim();
        match scan_leading_qualifier(trimmed) {
            Some((raw_qualifier, normalized, rest)) => {
                if !is_bare_reference_body(&rest) {
                    return None;
                }
                match renames.get(&normalized) {
                    Some(new_name) => {
                        out_items.push(format!("{}!{}", format_sheet_qualifier(new_name), rest));
                    }
                    None => out_items.push(format!("{}{}", raw_qualifier, rest)),
                }
            }
            None => {
                if !is_bare_reference_body(trimmed) {
                    return None;
                }
                out_items.push(trimmed.to_string());
            }
        }
    }
    Some(out_items.join(","))
}

/// `true` iff every char is a bare reference component (column letters, row
/// digits, `$`, `:`) -- deliberately excludes `(`, operators, and quotes, so
/// a genuine function call or expression is rejected rather than
/// misinterpreted as a reference. Empty is rejected too (a comma-list with an
/// empty item, e.g. a trailing comma, isn't valid reference-list syntax).
fn is_bare_reference_body(body: &str) -> bool {
    !body.is_empty()
        && body
            .chars()
            .all(|c| c.is_ascii_alphanumeric() || c == '$' || c == ':')
}

/// Scans `item`'s leading `Sheet!` / `'Sheet Name'!` qualifier, duplicating
/// (in miniature) the exact same quoting rule as
/// `FormulaParser::try_parse_sheet_qualifier` (`''` inside a quoted name is a
/// literal `'`; an unquoted name is Unicode-alphanumeric-or-`_`) -- not reused
/// directly since that method is private and tied to full formula-expression
/// parsing, whereas this runs on bare `<definedName>` text that is never a
/// full formula. Returns `(raw_qualifier_including_bang, normalized_lowercase_name,
/// rest_of_item)`, or `None` if `item` has no qualifier at all (not an error --
/// it simply isn't rewritten).
fn scan_leading_qualifier(item: &str) -> Option<(String, String, String)> {
    let chars: Vec<char> = item.chars().collect();
    let (qualifier_len, normalized) = if chars.first() == Some(&'\'') {
        let mut i = 1;
        let mut normalized = String::new();
        loop {
            match chars.get(i) {
                Some('\'') if chars.get(i + 1) == Some(&'\'') => {
                    normalized.push('\'');
                    i += 2;
                }
                Some('\'') => {
                    i += 1;
                    break;
                }
                Some(c) => {
                    normalized.push(*c);
                    i += 1;
                }
                None => return None,
            }
        }
        if chars.get(i) != Some(&'!') {
            return None;
        }
        (i + 1, normalized)
    } else if chars
        .first()
        .is_some_and(|c| c.is_alphabetic() || *c == '_')
    {
        let mut i = 0;
        while chars
            .get(i)
            .is_some_and(|c| c.is_alphanumeric() || *c == '_')
        {
            i += 1;
        }
        if chars.get(i) != Some(&'!') {
            return None;
        }
        let normalized = chars[0..i].iter().collect::<String>();
        (i + 1, normalized)
    } else {
        return None;
    };
    let raw_qualifier: String = chars[0..qualifier_len].iter().collect();
    let rest: String = chars[qualifier_len..].iter().collect();
    Some((raw_qualifier, normalized.to_lowercase(), rest))
}

#[cfg(test)]
mod tests {
    use super::*;

    const S1: &str = "sheet1";
    const S2: &str = "sheet2";

    fn insert(at: u32, count: u32) -> StructuralEdit {
        StructuralEdit::Insert { at, count }
    }
    fn delete(at: u32, count: u32) -> StructuralEdit {
        StructuralEdit::Delete { at, count }
    }
    // Convenience for the common same-sheet case most existing tests exercise.
    fn shift_same_sheet(
        formula: &str,
        axis: RefAxis,
        edit: StructuralEdit,
    ) -> Result<Option<String>, String> {
        shift_references(formula, S1, S1, axis, edit)
    }

    #[test]
    fn unaffected_reference_is_a_noop() {
        assert_eq!(
            shift_same_sheet("=A1+1", RefAxis::Row, insert(5, 1)).unwrap(),
            None
        );
        assert_eq!(
            shift_same_sheet("=A1+1", RefAxis::Row, delete(5, 1)).unwrap(),
            None
        );
        assert_eq!(
            shift_same_sheet("=A1+1", RefAxis::Row, insert(1, 0)).unwrap(),
            None
        );
    }

    #[test]
    fn insert_shifts_cell_ref_at_or_after_the_insertion_point() {
        // Insert 2 rows before row 5: a ref at row 5 or later moves down by 2.
        assert_eq!(
            shift_same_sheet("=A5+1", RefAxis::Row, insert(5, 2)).unwrap(),
            Some("A7+1".to_string())
        );
        assert_eq!(
            shift_same_sheet("=A10+1", RefAxis::Row, insert(5, 2)).unwrap(),
            Some("A12+1".to_string())
        );
        // Strictly before the insertion point: unaffected.
        assert_eq!(
            shift_same_sheet("=A4+1", RefAxis::Row, insert(5, 2)).unwrap(),
            None
        );
    }

    #[test]
    fn insert_columns_shifts_only_the_column() {
        assert_eq!(
            shift_same_sheet("=C10", RefAxis::Col, insert(2, 1)).unwrap(),
            Some("D10".to_string())
        );
        // A row-axis edit that DOES affect the row must still leave the
        // column letter untouched -- only the row digits change.
        assert_eq!(
            shift_same_sheet("=C10", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("C11".to_string())
        );
        // And a row-axis edit that doesn't reach this row must be a no-op,
        // even though the column axis has its own affecting edit elsewhere.
        assert_eq!(
            shift_same_sheet("=C10", RefAxis::Row, insert(20, 1)).unwrap(),
            None
        );
    }

    #[test]
    fn delete_shifts_cell_ref_after_the_deleted_band() {
        assert_eq!(
            shift_same_sheet("=A10+1", RefAxis::Row, delete(5, 2)).unwrap(),
            Some("A8+1".to_string())
        );
    }

    #[test]
    fn delete_turns_a_reference_into_the_deleted_band_into_ref_error() {
        assert_eq!(
            shift_same_sheet("=A5+1", RefAxis::Row, delete(5, 2)).unwrap(),
            Some("#REF!+1".to_string())
        );
        assert_eq!(
            shift_same_sheet("=A6+1", RefAxis::Row, delete(5, 2)).unwrap(),
            Some("#REF!+1".to_string())
        );
    }

    #[test]
    fn absolute_flags_are_preserved_through_a_shift() {
        assert_eq!(
            shift_same_sheet("=$A$10", RefAxis::Row, insert(5, 2)).unwrap(),
            Some("$A$12".to_string())
        );
        assert_eq!(
            shift_same_sheet("=A$10", RefAxis::Row, delete(5, 2)).unwrap(),
            Some("A$8".to_string())
        );
    }

    #[test]
    fn range_shrinks_when_deletion_falls_inside_it_without_collapsing() {
        // A1:A10, delete rows 3-4 -> A1:A8 (both corners survive).
        assert_eq!(
            shift_same_sheet("=SUM(A1:A10)", RefAxis::Row, delete(3, 2)).unwrap(),
            Some("SUM(A1:A8)".to_string())
        );
        // A3:A10, delete rows 3-4 -> top corner clamps to the surviving row
        // that slides into position 3; bottom shifts down by 2.
        assert_eq!(
            shift_same_sheet("=SUM(A3:A10)", RefAxis::Row, delete(3, 2)).unwrap(),
            Some("SUM(A3:A8)".to_string())
        );
        // A1:A4, delete rows 3-4 -> bottom corner clamps to just above the
        // deletion; top is untouched (both survive as A1:A2).
        assert_eq!(
            shift_same_sheet("=SUM(A1:A4)", RefAxis::Row, delete(3, 2)).unwrap(),
            Some("SUM(A1:A2)".to_string())
        );
    }

    #[test]
    fn range_becomes_ref_error_when_fully_deleted() {
        // A3:A4, delete rows 3-4: nothing survives.
        assert_eq!(
            shift_same_sheet("=SUM(A3:A4)", RefAxis::Row, delete(3, 2)).unwrap(),
            Some("SUM(#REF!)".to_string())
        );
        // A single-row range exactly matching a 1-row deletion at row 1.
        assert_eq!(
            shift_same_sheet("=SUM(A1:A1)", RefAxis::Row, delete(1, 1)).unwrap(),
            Some("SUM(#REF!)".to_string())
        );
    }

    #[test]
    fn range_insert_never_collapses_and_grows_the_range() {
        // Inserting a row inside a range grows it, never splits/collapses it.
        assert_eq!(
            shift_same_sheet("=SUM(A1:A10)", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("SUM(A1:A11)".to_string())
        );
    }

    #[test]
    fn reversed_range_order_is_handled_like_the_evaluator_treats_it() {
        // B10:A1 (reversed): min/max logic still applies per-corner-slot.
        assert_eq!(
            shift_same_sheet("=SUM(A10:A1)", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("SUM(A11:A1)".to_string())
        );
    }

    #[test]
    fn mixed_corner_absolute_flags_survive_a_range_shift() {
        assert_eq!(
            shift_same_sheet("=SUM($A1:B$10)", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("SUM($A1:B$11)".to_string())
        );
    }

    #[test]
    fn leading_equals_and_whitespace_are_normalized_away() {
        assert_eq!(
            shift_same_sheet("  =A10  ", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("A11".to_string())
        );
        // A formula with no leading '=' at all (matches how <f> XML content
        // and this function's own return value are stored) works the same.
        assert_eq!(
            shift_same_sheet("A10", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("A11".to_string())
        );
    }

    #[test]
    fn multiple_references_in_one_formula_each_patch_independently() {
        assert_eq!(
            shift_same_sheet("=A10+B2*SUM(C10:C20)", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("A11+B2*SUM(C11:C21)".to_string())
        );
    }

    #[test]
    fn propagates_a_parse_error_instead_of_silently_ignoring_it() {
        assert!(shift_same_sheet("=A1+", RefAxis::Row, insert(1, 1)).is_err());
    }

    // ── 0.14.0-A2: sheet-qualified references, workbook-wide targeting ──────

    #[test]
    fn qualified_reference_targeting_the_edited_sheet_shifts_regardless_of_host() {
        // Formula hosted on sheet2, referencing sheet1 -- sheet1 is being edited.
        assert_eq!(
            shift_references("=Sheet1!A10+1", S2, S1, RefAxis::Row, insert(5, 1)).unwrap(),
            Some("Sheet1!A11+1".to_string())
        );
    }

    #[test]
    fn unqualified_reference_on_a_different_host_sheet_is_untouched() {
        // Formula hosted on sheet2; sheet1 is being edited. A bare A10 on
        // sheet2 means sheet2!A10, not sheet1!A10 -- must not shift.
        assert_eq!(
            shift_references("=A10+1", S2, S1, RefAxis::Row, insert(5, 1)).unwrap(),
            None
        );
    }

    #[test]
    fn unqualified_reference_on_the_edited_sheet_still_shifts() {
        assert_eq!(
            shift_references("=A10+1", S1, S1, RefAxis::Row, insert(5, 1)).unwrap(),
            Some("A11+1".to_string())
        );
    }

    #[test]
    fn qualified_reference_to_a_different_sheet_is_untouched_even_when_hosted_on_the_edited_sheet()
    {
        // Formula lives ON sheet1 (the sheet being edited), but this
        // particular reference explicitly names sheet2 -- must not shift.
        assert_eq!(
            shift_references("=Sheet2!A10+1", S1, S1, RefAxis::Row, insert(5, 1)).unwrap(),
            None
        );
    }

    #[test]
    fn quoted_sheet_name_reference_shifts_and_keeps_its_quoting() {
        assert_eq!(
            shift_references(
                "='Sales 2026'!A10",
                S2,
                "sales 2026",
                RefAxis::Row,
                insert(5, 1)
            )
            .unwrap(),
            Some("'Sales 2026'!A11".to_string())
        );
    }

    #[test]
    fn escaped_apostrophe_in_a_quoted_sheet_name_resolves_correctly() {
        // 'Bob''s Data' normalizes to "Bob's Data"; the sheet key convention
        // lowercases the unescaped name, so the edited key is "bob's data".
        assert_eq!(
            shift_references(
                "='Bob''s Data'!A10",
                S2,
                "bob's data",
                RefAxis::Row,
                insert(5, 1)
            )
            .unwrap(),
            Some("'Bob''s Data'!A11".to_string())
        );
    }

    #[test]
    fn qualifier_case_does_not_matter_for_targeting() {
        assert_eq!(
            shift_references("=SHEET1!A10", S2, S1, RefAxis::Row, insert(5, 1)).unwrap(),
            Some("SHEET1!A11".to_string())
        );
    }

    #[test]
    fn only_the_reference_naming_the_edited_sheet_is_rewritten_among_several() {
        assert_eq!(
            shift_references(
                "=Sheet1!A10+A10+Sheet3!A10",
                S1,
                S1,
                RefAxis::Row,
                insert(5, 1)
            )
            .unwrap(),
            // Sheet1!A10 (qualified, targets edited sheet1) and the bare A10
            // (unqualified, host IS sheet1) both shift; Sheet3!A10 does not.
            Some("Sheet1!A11+A11+Sheet3!A10".to_string())
        );
    }

    #[test]
    fn qualified_reference_into_a_deleted_band_becomes_ref_error_keeping_the_qualifier() {
        assert_eq!(
            shift_references("=Sheet1!A5+1", S2, S1, RefAxis::Row, delete(5, 2)).unwrap(),
            Some("Sheet1!#REF!+1".to_string())
        );
    }

    #[test]
    fn qualified_range_partial_delete_shrinks_keeping_the_qualifier() {
        assert_eq!(
            shift_references("=SUM(Sheet1!A1:A10)", S2, S1, RefAxis::Row, delete(3, 2)).unwrap(),
            Some("SUM(Sheet1!A1:A8)".to_string())
        );
    }

    #[test]
    fn external_and_3d_references_are_left_completely_untouched() {
        // External workbook references ([Book2.xlsx]Sheet1!A1) and 3D
        // references (Sheet1:Sheet3!A1) aren't supported syntax -- the whole
        // formula must fail to parse and be reported as Err, not partially
        // rewritten. The caller (Vm::rewrite_formulas_for_structural_edit)
        // is what actually leaves such formulas untouched on an Err.
        assert!(
            shift_references("=[Book2.xlsx]Sheet1!A1", S1, S1, RefAxis::Row, insert(1, 1)).is_err()
        );
        assert!(shift_references("=Sheet1:Sheet3!A1", S1, S1, RefAxis::Row, insert(1, 1)).is_err());
    }

    #[test]
    fn a_sheet_qualifier_inside_a_string_literal_is_never_treated_as_a_reference() {
        assert_eq!(
            shift_references("=\"Sheet1!A1\"&A10", S1, S1, RefAxis::Row, insert(5, 1)).unwrap(),
            Some("\"Sheet1!A1\"&A11".to_string())
        );
    }

    #[test]
    fn a_ref_error_literal_in_the_formula_does_not_confuse_the_qualifier_lookahead() {
        // #REF! contains '!' -- must not be misread as a sheet qualifier.
        // Elixcee's parser doesn't parse error literals at all (pre-existing,
        // unrelated to this round), so the whole formula is Err either way.
        assert!(shift_references("=#REF!+A10", S1, S1, RefAxis::Row, insert(5, 1)).is_err());
    }

    // ── sheet rename: qualifier rewrite ──────────────────────────────────

    #[test]
    fn rename_rewrites_a_qualifier_naming_the_renamed_sheet() {
        assert_eq!(
            rename_sheet_references("=Sheet1!A10+1", S1, "Data").unwrap(),
            Some("Data!A10+1".to_string())
        );
    }

    #[test]
    fn rename_never_touches_unqualified_references() {
        // Formula lives ON the sheet being renamed -- a bare A10 still means
        // "this same sheet", whatever it's now called.
        assert_eq!(rename_sheet_references("=A10+1", S1, "Data").unwrap(), None);
    }

    #[test]
    fn rename_never_touches_a_qualifier_naming_a_different_sheet() {
        assert_eq!(
            rename_sheet_references("=Sheet2!A10+1", S1, "Data").unwrap(),
            None
        );
    }

    #[test]
    fn rename_quotes_the_new_name_when_it_needs_quoting() {
        assert_eq!(
            rename_sheet_references("=Sheet1!A10", S1, "Sales 2026").unwrap(),
            Some("'Sales 2026'!A10".to_string())
        );
    }

    #[test]
    fn rename_unquotes_when_the_new_name_no_longer_needs_it() {
        assert_eq!(
            rename_sheet_references("='Sales 2026'!A10", "sales 2026", "Sheet1").unwrap(),
            Some("Sheet1!A10".to_string())
        );
    }

    #[test]
    fn rename_escapes_an_apostrophe_in_the_new_name() {
        assert_eq!(
            rename_sheet_references("=Sheet1!A10", S1, "Bob's Data").unwrap(),
            Some("'Bob''s Data'!A10".to_string())
        );
    }

    #[test]
    fn rename_matches_the_old_qualifier_case_insensitively() {
        assert_eq!(
            rename_sheet_references("=SHEET1!A10", S1, "Data").unwrap(),
            Some("Data!A10".to_string())
        );
    }

    #[test]
    fn rename_rewrites_a_qualified_range_keeping_the_coordinates_untouched() {
        assert_eq!(
            rename_sheet_references("=SUM(Sheet1!A1:B10)", S1, "Data").unwrap(),
            Some("SUM(Data!A1:B10)".to_string())
        );
    }

    #[test]
    fn rename_only_rewrites_references_naming_the_renamed_sheet_among_several() {
        assert_eq!(
            rename_sheet_references("=Sheet1!A1+Sheet2!A1+A1", S1, "Data").unwrap(),
            Some("Data!A1+Sheet2!A1+A1".to_string())
        );
    }

    #[test]
    fn rename_case_only_still_rewrites_to_match_the_new_display_casing() {
        // rename_sheet allows renaming "Sheet1" -> "SHEET1" (a pure casing
        // change, same key) -- existing formula references should still
        // pick up the new display casing, matching real Excel.
        assert_eq!(
            rename_sheet_references("=Sheet1!A1", S1, "SHEET1").unwrap(),
            Some("SHEET1!A1".to_string())
        );
    }

    #[test]
    fn rename_is_a_noop_when_nothing_references_the_renamed_sheet() {
        assert_eq!(rename_sheet_references("=1+2", S1, "Data").unwrap(), None);
    }

    #[test]
    fn rename_propagates_a_parse_error_instead_of_silently_ignoring_it() {
        assert!(rename_sheet_references("=[Book2.xlsx]Sheet1!A1", S1, "Data").is_err());
    }

    // ── range move: reference-identity translation ──────────────────────

    fn rect(r1: u32, c1: u32, r2: u32, c2: u32) -> MoveRect {
        MoveRect { r1, c1, r2, c2 }
    }
    fn move_same_sheet(
        formula: &str,
        source: MoveRect,
        d_row: i64,
        d_col: i64,
    ) -> Result<MoveRewrite, String> {
        translate_references_for_move(formula, S1, source, d_row, d_col)
    }

    #[test]
    fn zero_offset_is_unchanged_without_parsing() {
        // "=A1+" would fail to parse if it were parsed -- confirms the
        // (0,0) short-circuit happens before parse_with_refs runs.
        assert_eq!(
            move_same_sheet("=A1+", rect(1, 1, 5, 5), 0, 0).unwrap(),
            MoveRewrite::Unchanged
        );
    }

    #[test]
    fn cell_ref_inside_the_source_rect_follows_the_move() {
        // Move A1:E5 down 2, right 1: a reference to B2 (inside) follows to C4.
        assert_eq!(
            move_same_sheet("=B2+1", rect(1, 1, 5, 5), 2, 1).unwrap(),
            MoveRewrite::Rewritten("C4+1".to_string())
        );
    }

    #[test]
    fn cell_ref_outside_the_source_rect_is_unchanged() {
        assert_eq!(
            move_same_sheet("=Z9+1", rect(1, 1, 5, 5), 2, 1).unwrap(),
            MoveRewrite::Unchanged
        );
    }

    #[test]
    fn absolute_flags_are_preserved_through_a_move() {
        assert_eq!(
            move_same_sheet("=$B$2", rect(1, 1, 5, 5), 2, 1).unwrap(),
            MoveRewrite::Rewritten("$C$4".to_string())
        );
    }

    #[test]
    fn a_formula_hosted_inside_the_moved_block_still_follows_its_own_internal_ref() {
        // Design doc §1: a formula physically inside the moved block that
        // references another cell also inside it uses the SAME mechanism as
        // an external formula following a moved cell -- not a separate
        // relative-offset translation.
        assert_eq!(
            move_same_sheet("=B3+1", rect(1, 1, 5, 5), 2, 1).unwrap(),
            MoveRewrite::Rewritten("C5+1".to_string())
        );
    }

    #[test]
    fn range_reference_fully_inside_the_source_rect_translates_as_a_whole() {
        assert_eq!(
            move_same_sheet("=SUM(B2:C3)", rect(1, 1, 5, 5), 2, 1).unwrap(),
            MoveRewrite::Rewritten("SUM(C4:D5)".to_string())
        );
    }

    #[test]
    fn range_reference_fully_outside_the_source_rect_is_unchanged() {
        assert_eq!(
            move_same_sheet("=SUM(Y1:Z2)", rect(1, 1, 5, 5), 2, 1).unwrap(),
            MoveRewrite::Unchanged
        );
    }

    #[test]
    fn range_reference_with_exactly_one_corner_inside_is_ambiguous() {
        // A2:D2, only A2 falls inside the moved rectangle -- real Excel's
        // behavior here is confirmed only for the narrow "destination still
        // inside the same range" sub-case (design doc §3) and unconfirmed
        // in general (§4-A); refuse rather than guess.
        assert_eq!(
            move_same_sheet("=SUM(A2:D2)", rect(2, 1, 2, 1), 0, 1).unwrap(),
            MoveRewrite::Ambiguous
        );
        // Same shape with the inside corner on the other side.
        assert_eq!(
            move_same_sheet("=SUM(A2:D2)", rect(2, 4, 2, 4), 0, 1).unwrap(),
            MoveRewrite::Ambiguous
        );
    }

    #[test]
    fn multiple_references_only_the_ones_inside_the_rect_translate() {
        assert_eq!(
            move_same_sheet("=B2+Z9+SUM(B2:C3)", rect(1, 1, 5, 5), 2, 1).unwrap(),
            MoveRewrite::Rewritten("C4+Z9+SUM(C4:D5)".to_string())
        );
    }

    #[test]
    fn no_reference_touches_the_source_rect_is_unchanged() {
        assert_eq!(
            move_same_sheet("=1+2", rect(1, 1, 5, 5), 2, 1).unwrap(),
            MoveRewrite::Unchanged
        );
    }

    #[test]
    fn a_negative_offset_moves_references_up_and_left() {
        assert_eq!(
            move_same_sheet("=C4", rect(1, 1, 5, 5), -2, -1).unwrap(),
            MoveRewrite::Rewritten("B2".to_string())
        );
    }

    #[test]
    fn self_qualified_reference_naming_the_move_sheet_still_translates() {
        assert_eq!(
            translate_references_for_move("=Sheet1!B2", S1, rect(1, 1, 5, 5), 2, 1).unwrap(),
            MoveRewrite::Rewritten("Sheet1!C4".to_string())
        );
    }

    #[test]
    fn qualified_reference_to_a_different_sheet_is_never_touched_by_a_move() {
        assert_eq!(
            move_same_sheet("=Sheet2!B2", rect(1, 1, 5, 5), 2, 1).unwrap(),
            MoveRewrite::Unchanged
        );
    }

    #[test]
    fn move_propagates_a_parse_error_instead_of_silently_ignoring_it() {
        assert!(move_same_sheet("=[Book2.xlsx]Sheet1!A1", rect(1, 1, 5, 5), 2, 1).is_err());
        assert!(move_same_sheet("=A1+", rect(1, 1, 5, 5), 2, 1).is_err());
    }

    #[test]
    fn a_reversed_range_reference_is_handled_by_corner_not_by_sort_order() {
        // B10:A1 written reversed -- corner containment is checked exactly
        // as parsed (c1=B,r1=10 / c2=A,r2=1), not normalized min/max first.
        // Both corners inside a rect covering both -> whole thing translates.
        assert_eq!(
            move_same_sheet("=SUM(B10:A1)", rect(1, 1, 10, 2), 0, 2).unwrap(),
            MoveRewrite::Rewritten("SUM(D10:C1)".to_string())
        );
    }

    fn renames(pairs: &[(&str, &str)]) -> HashMap<String, String> {
        pairs
            .iter()
            .map(|(k, v)| (k.to_string(), v.to_string()))
            .collect()
    }

    #[test]
    fn defined_name_plain_reference_rewritten_via_the_general_path() {
        assert_eq!(
            rewrite_defined_name_for_renames("Sheet1!$F$5", &renames(&[("sheet1", "NewName")])),
            DefinedNameRewrite::Rewritten("NewName!$F$5".to_string())
        );
    }

    #[test]
    fn defined_name_referencing_an_unrelated_sheet_is_unchanged() {
        assert_eq!(
            rewrite_defined_name_for_renames("Sheet2!$F$5", &renames(&[("sheet1", "NewName")])),
            DefinedNameRewrite::Unchanged
        );
    }

    #[test]
    fn defined_name_with_no_tracked_renames_is_unchanged() {
        assert_eq!(
            rewrite_defined_name_for_renames("Sheet1!$F$5", &HashMap::new()),
            DefinedNameRewrite::Unchanged
        );
    }

    #[test]
    fn defined_name_print_titles_style_full_row_and_column_rewritten() {
        // Print_Titles' real-world shape: a full-row spec and a full-column spec,
        // comma-joined, both qualified to the same sheet -- exactly what
        // `parse_with_refs` cannot parse (no top-level comma, no full-row/column refs).
        assert_eq!(
            rewrite_defined_name_for_renames(
                "Sheet1!$1:$3,Sheet1!$A:$A",
                &renames(&[("sheet1", "Data")])
            ),
            DefinedNameRewrite::Rewritten("Data!$1:$3,Data!$A:$A".to_string())
        );
    }

    #[test]
    fn defined_name_multi_area_print_area_rewritten() {
        assert_eq!(
            rewrite_defined_name_for_renames(
                "Sheet1!$A$1:$B$2,Sheet1!$D$1:$E$2",
                &renames(&[("sheet1", "Renamed")])
            ),
            DefinedNameRewrite::Rewritten("Renamed!$A$1:$B$2,Renamed!$D$1:$E$2".to_string())
        );
    }

    #[test]
    fn defined_name_multi_area_only_rewrites_the_matching_sheet() {
        assert_eq!(
            rewrite_defined_name_for_renames(
                "Sheet1!$A$1,Sheet2!$B$2",
                &renames(&[("sheet1", "Renamed")])
            ),
            DefinedNameRewrite::Rewritten("Renamed!$A$1,Sheet2!$B$2".to_string())
        );
    }

    #[test]
    fn defined_name_quoted_sheet_name_with_escaped_quote_rewritten() {
        assert_eq!(
            rewrite_defined_name_for_renames(
                "'My ''Old'' Sheet'!$A$1:$B$2",
                &renames(&[("my 'old' sheet", "New Sheet")])
            ),
            DefinedNameRewrite::Rewritten("'New Sheet'!$A$1:$B$2".to_string())
        );
    }

    #[test]
    fn defined_name_multi_area_with_no_matching_rename_is_unchanged() {
        assert_eq!(
            rewrite_defined_name_for_renames(
                "Sheet2!$1:$3,Sheet2!$A:$A",
                &renames(&[("sheet1", "Data")])
            ),
            DefinedNameRewrite::Unchanged
        );
    }

    #[test]
    fn defined_name_unparseable_value_mentioning_the_renamed_sheet_is_dropped() {
        // A dynamic named range embedding a full-column reference inside a function
        // call -- not a bare reference list (rejected by the `(` inside it), and not
        // parseable by `parse_with_refs` either (full-column ref). Mentions the
        // renamed sheet, so it can't be confirmed safe to leave untouched.
        assert_eq!(
            rewrite_defined_name_for_renames(
                "OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)",
                &renames(&[("sheet1", "Data")])
            ),
            DefinedNameRewrite::Drop
        );
    }

    #[test]
    fn defined_name_unparseable_value_not_mentioning_any_renamed_sheet_is_unchanged() {
        assert_eq!(
            rewrite_defined_name_for_renames(
                "OFFSET(Sheet3!$A$1,0,0,COUNTA(Sheet3!$A:$A),1)",
                &renames(&[("sheet1", "Data")])
            ),
            DefinedNameRewrite::Unchanged
        );
    }

    #[test]
    fn defined_name_dynamic_range_with_simple_refs_rewritten_via_general_path() {
        // No full-column/row refs here, so `parse_with_refs` handles the whole
        // formula -- the general path, not the reference-list fallback.
        assert_eq!(
            rewrite_defined_name_for_renames(
                "OFFSET(Sheet1!$A$1,0,0,10,1)",
                &renames(&[("sheet1", "Data")])
            ),
            DefinedNameRewrite::Rewritten("OFFSET(Data!$A$1,0,0,10,1)".to_string())
        );
    }

    #[test]
    fn defined_name_multiple_renames_chained_in_one_pass() {
        assert_eq!(
            rewrite_defined_name_for_renames(
                "Sheet1!$A$1,Sheet2!$B$2",
                &renames(&[("sheet1", "First"), ("sheet2", "Second")])
            ),
            DefinedNameRewrite::Rewritten("First!$A$1,Second!$B$2".to_string())
        );
    }
}
