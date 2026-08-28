//! Reference rewriting for same-sheet row/column insert-delete (0.14.0-A).
//!
//! Patches only the text span of each affected cell/range reference --
//! everything else in the formula (operators, function names, literals,
//! whitespace, unaffected references) is left untouched. See
//! `super::parser::parse_with_refs` for how spans are captured; this module
//! never builds or walks a full expression tree, and there is deliberately
//! no general AST-to-text serializer -- see `internal_docs/ROADMAP.md`'s
//! 0.14.0-A note for why a span-patching design replaced that idea.
//!
//! Scope, matching what was explicitly approved for this round: same-sheet
//! row/column insert/delete only. Cross-sheet syntax, sheet rename, and
//! range move are out of scope here, as is wiring this into the VM's
//! `insert_rows_on_sheet`/etc (that's the next round).

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

enum CellShift {
    Unchanged,
    Moved(u32),
    Deleted,
}

/// A single cell reference's coordinate on the edited axis. Deletion of a
/// row/column the reference points at is unconditionally `#REF!` -- unlike a
/// range corner, there's no surviving neighbor for a single cell to clamp to.
fn shift_cell_coord(idx: u32, edit: StructuralEdit) -> CellShift {
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
fn shift_bound_low(idx: u32, edit: StructuralEdit) -> u32 {
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
fn shift_bound_high(idx: u32, edit: StructuralEdit) -> i64 {
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

/// Rewrite `formula`'s cell/range references for a same-sheet row or column
/// insert/delete. Returns `Ok(None)` when no reference in the formula was
/// affected (caller can skip rewriting the stored formula string). A
/// reference that falls entirely inside a deleted band becomes `#REF!`,
/// matching real Excel -- this is deliberate, tested behavior, not an
/// omission.
///
/// `$` (absolute) flags never change whether a reference shifts here --
/// unlike copy/fill, real Excel shifts every reference on row/column
/// insert-delete regardless of `$` -- they're only preserved as-is in the
/// rewritten text.
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
            } => {
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
            } => {
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

#[cfg(test)]
mod tests {
    use super::*;

    fn insert(at: u32, count: u32) -> StructuralEdit {
        StructuralEdit::Insert { at, count }
    }
    fn delete(at: u32, count: u32) -> StructuralEdit {
        StructuralEdit::Delete { at, count }
    }

    #[test]
    fn unaffected_reference_is_a_noop() {
        assert_eq!(
            shift_references("=A1+1", RefAxis::Row, insert(5, 1)).unwrap(),
            None
        );
        assert_eq!(
            shift_references("=A1+1", RefAxis::Row, delete(5, 1)).unwrap(),
            None
        );
        assert_eq!(
            shift_references("=A1+1", RefAxis::Row, insert(1, 0)).unwrap(),
            None
        );
    }

    #[test]
    fn insert_shifts_cell_ref_at_or_after_the_insertion_point() {
        // Insert 2 rows before row 5: a ref at row 5 or later moves down by 2.
        assert_eq!(
            shift_references("=A5+1", RefAxis::Row, insert(5, 2)).unwrap(),
            Some("A7+1".to_string())
        );
        assert_eq!(
            shift_references("=A10+1", RefAxis::Row, insert(5, 2)).unwrap(),
            Some("A12+1".to_string())
        );
        // Strictly before the insertion point: unaffected.
        assert_eq!(
            shift_references("=A4+1", RefAxis::Row, insert(5, 2)).unwrap(),
            None
        );
    }

    #[test]
    fn insert_columns_shifts_only_the_column() {
        assert_eq!(
            shift_references("=C10", RefAxis::Col, insert(2, 1)).unwrap(),
            Some("D10".to_string())
        );
        // A row-axis edit that DOES affect the row must still leave the
        // column letter untouched -- only the row digits change.
        assert_eq!(
            shift_references("=C10", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("C11".to_string())
        );
        // And a row-axis edit that doesn't reach this row must be a no-op,
        // even though the column axis has its own affecting edit elsewhere.
        assert_eq!(
            shift_references("=C10", RefAxis::Row, insert(20, 1)).unwrap(),
            None
        );
    }

    #[test]
    fn delete_shifts_cell_ref_after_the_deleted_band() {
        assert_eq!(
            shift_references("=A10+1", RefAxis::Row, delete(5, 2)).unwrap(),
            Some("A8+1".to_string())
        );
    }

    #[test]
    fn delete_turns_a_reference_into_the_deleted_band_into_ref_error() {
        assert_eq!(
            shift_references("=A5+1", RefAxis::Row, delete(5, 2)).unwrap(),
            Some("#REF!+1".to_string())
        );
        assert_eq!(
            shift_references("=A6+1", RefAxis::Row, delete(5, 2)).unwrap(),
            Some("#REF!+1".to_string())
        );
    }

    #[test]
    fn absolute_flags_are_preserved_through_a_shift() {
        assert_eq!(
            shift_references("=$A$10", RefAxis::Row, insert(5, 2)).unwrap(),
            Some("$A$12".to_string())
        );
        assert_eq!(
            shift_references("=A$10", RefAxis::Row, delete(5, 2)).unwrap(),
            Some("A$8".to_string())
        );
    }

    #[test]
    fn range_shrinks_when_deletion_falls_inside_it_without_collapsing() {
        // A1:A10, delete rows 3-4 -> A1:A8 (both corners survive).
        assert_eq!(
            shift_references("=SUM(A1:A10)", RefAxis::Row, delete(3, 2)).unwrap(),
            Some("SUM(A1:A8)".to_string())
        );
        // A3:A10, delete rows 3-4 -> top corner clamps to the surviving row
        // that slides into position 3; bottom shifts down by 2.
        assert_eq!(
            shift_references("=SUM(A3:A10)", RefAxis::Row, delete(3, 2)).unwrap(),
            Some("SUM(A3:A8)".to_string())
        );
        // A1:A4, delete rows 3-4 -> bottom corner clamps to just above the
        // deletion; top is untouched (both survive as A1:A2).
        assert_eq!(
            shift_references("=SUM(A1:A4)", RefAxis::Row, delete(3, 2)).unwrap(),
            Some("SUM(A1:A2)".to_string())
        );
    }

    #[test]
    fn range_becomes_ref_error_when_fully_deleted() {
        // A3:A4, delete rows 3-4: nothing survives.
        assert_eq!(
            shift_references("=SUM(A3:A4)", RefAxis::Row, delete(3, 2)).unwrap(),
            Some("SUM(#REF!)".to_string())
        );
        // A single-row range exactly matching a 1-row deletion at row 1.
        assert_eq!(
            shift_references("=SUM(A1:A1)", RefAxis::Row, delete(1, 1)).unwrap(),
            Some("SUM(#REF!)".to_string())
        );
    }

    #[test]
    fn range_insert_never_collapses_and_grows_the_range() {
        // Inserting a row inside a range grows it, never splits/collapses it.
        assert_eq!(
            shift_references("=SUM(A1:A10)", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("SUM(A1:A11)".to_string())
        );
    }

    #[test]
    fn reversed_range_order_is_handled_like_the_evaluator_treats_it() {
        // B10:A1 (reversed): min/max logic still applies per-corner-slot.
        assert_eq!(
            shift_references("=SUM(A10:A1)", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("SUM(A11:A1)".to_string())
        );
    }

    #[test]
    fn mixed_corner_absolute_flags_survive_a_range_shift() {
        assert_eq!(
            shift_references("=SUM($A1:B$10)", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("SUM($A1:B$11)".to_string())
        );
    }

    #[test]
    fn leading_equals_and_whitespace_are_normalized_away() {
        assert_eq!(
            shift_references("  =A10  ", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("A11".to_string())
        );
        // A formula with no leading '=' at all (matches how <f> XML content
        // and this function's own return value are stored) works the same.
        assert_eq!(
            shift_references("A10", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("A11".to_string())
        );
    }

    #[test]
    fn multiple_references_in_one_formula_each_patch_independently() {
        assert_eq!(
            shift_references("=A10+B2*SUM(C10:C20)", RefAxis::Row, insert(5, 1)).unwrap(),
            Some("A11+B2*SUM(C11:C21)".to_string())
        );
    }

    #[test]
    fn propagates_a_parse_error_instead_of_silently_ignoring_it() {
        assert!(shift_references("=A1+", RefAxis::Row, insert(1, 1)).is_err());
    }
}
