pub mod ast;
pub mod eval;
pub mod parser;
pub mod rewrite;

pub use ast::{FormulaExpr, SheetQualifier};
pub use eval::evaluate;
pub(crate) use eval::references_another_sheet;
pub use parser::{RefOccurrence, parse, parse_with_refs};
pub use rewrite::{
    MoveRect, MoveRewrite, RefAxis, StructuralEdit, rename_sheet_references, shift_references,
    translate_references_for_move,
};
// `pub(crate)`, not `pub` -- these are 0.14.0-A/A4's own internal arithmetic,
// reused (not reimplemented) by 0.14.0-B's cell-metadata transform; see
// internal_docs/cell-metadata-transform-0.14.0-b-design.md §6.
pub(crate) use rewrite::{CellShift, shift_bound_high, shift_bound_low, shift_cell_coord};
// `pub(crate)`, not `pub` -- internal to `save_xlsx_impl`'s `<definedNames>`
// rename-preservation pass (`src/lib.rs`), see
// internal_docs/defined-names-rename-preservation-scoping.md.
pub(crate) use rewrite::{DefinedNameRewrite, rewrite_defined_name_for_renames};
