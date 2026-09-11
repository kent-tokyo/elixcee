pub mod ast;
pub mod eval;
pub mod parser;
pub mod rewrite;
mod workbook;

pub use ast::{FormulaExpr, SheetQualifier};
pub use eval::evaluate;
pub(crate) use eval::references_another_sheet;
pub(crate) use eval::with_sheet_context;
pub use parser::{RefOccurrence, parse, parse_with_refs};
pub use rewrite::{
    MoveRect, MoveRewrite, RefAxis, StructuralEdit, rename_sheet_references, shift_references,
    translate_references_for_move,
};
pub(crate) use workbook::recalculate as recalculate_workbook;
#[cfg(feature = "python")]
pub(crate) use workbook::{
    formula_dependencies, formula_dependency_diagnostics, formula_io_candidates,
};

/// Return whether a workbook's parsed formula dependency graph is cyclic.
/// This is a diagnostic signal; it does not change best-effort calculation.
pub fn workbook_has_formula_cycle(
    sheets: &std::collections::HashMap<
        String,
        std::collections::HashMap<(u32, u32), crate::types::CellContent>,
    >,
) -> bool {
    workbook::has_formula_cycle(sheets)
}

/// Recalculate a complete workbook through the shared Rust formula runtime.
///
/// This is the public, runtime-neutral entry point used by the WASM bridge. The
/// VM keeps the lower-level dirty-set variant private because its cache
/// bookkeeping is VM-specific; callers that own a plain workbook map get a
/// deterministic full evaluation and can then serialize the updated values in
/// their own host format.
pub fn calculate_workbook(
    sheets: &mut std::collections::HashMap<
        String,
        std::collections::HashMap<(u32, u32), crate::types::CellContent>,
    >,
    named_ranges: &std::collections::HashMap<String, String>,
) -> Result<bool, String> {
    if workbook::recalculate(
        sheets,
        named_ranges,
        &std::collections::HashMap::new(),
        &std::collections::HashMap::new(),
        None,
        true,
    )? {
        return Ok(true);
    }
    let mut calculated = false;
    for cells in sheets.values_mut() {
        let formulas = cells
            .iter()
            .filter_map(|(&position, cell)| {
                cell.formula.as_deref().map(|formula| (position, formula))
            })
            .map(|(position, formula)| (position, formula.to_string()))
            .collect::<Vec<_>>();
        for (position, formula) in formulas {
            let expr = parse(&formula)?;
            let value = evaluate(&expr, cells)?;
            if let Some(cell) = cells.get_mut(&position) {
                cell.value = value;
                calculated = true;
            }
        }
    }
    Ok(calculated)
}
// `pub(crate)`, not `pub` -- these are 0.14.0-A/A4's own internal arithmetic,
// reused (not reimplemented) by 0.14.0-B's cell-metadata transform; see
// internal_docs/cell-metadata-transform-0.14.0-b-design.md §6.
pub(crate) use rewrite::{CellShift, shift_bound_high, shift_bound_low, shift_cell_coord};
// `pub(crate)`, not `pub` -- internal to `save_xlsx_impl`'s `<definedNames>`
// rename-preservation pass (`src/lib.rs`), see
// internal_docs/defined-names-rename-preservation-scoping.md.
pub(crate) use rewrite::{DefinedNameRewrite, rewrite_defined_name_for_renames};
