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
// `rewrite::{CellShift, shift_cell_coord, shift_bound_low, shift_bound_high}` are
// `pub(crate)` (not `pub`) and deliberately not re-exported here yet -- no
// consumer exists until 0.14.0-B's metadata transform lands, and an unused
// re-export would just be clippy noise until then. Reachable in the meantime
// via `formula::rewrite::shift_cell_coord` etc.; add the re-export alongside
// whichever round first calls one.
