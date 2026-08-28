/// A `Sheet2!`/`'Sales 2026'!`/`'Bob''s Data'!` prefix on a cell/range reference
/// (0.14.0-A2). `raw_span` covers only the qualifier text itself (not the `!`
/// and not the coordinate that follows) — the rewriter never touches it.
#[derive(Debug, Clone, PartialEq)]
pub struct SheetQualifier {
    pub raw_span: (usize, usize),
    /// Exact original text, quotes and `''`-escapes included (e.g.
    /// `Sheet2`, `'Sales 2026'`, `'Bob''s Data'`) — for diagnostics/tests,
    /// never used to reconstruct formula text (the span is copied verbatim).
    pub raw_text: String,
    /// Unescaped, case-preserved sheet name (quotes stripped, `''` -> `'`).
    /// Compare case-insensitively (`.to_lowercase()`) against a sheet key —
    /// this codebase's own sheet-identity convention (see `ensure_sheet_at`).
    pub normalized_name: String,
}

#[derive(Debug, Clone, PartialEq)]
pub enum BinOpKind {
    Add,
    Sub,
    Mul,
    Div,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    Concat,
}

#[derive(Debug, Clone, PartialEq)]
pub enum FormulaExpr {
    Number(f64),
    Str(String),
    Bool(bool),
    CellRef {
        col: u32,
        row: u32,
        /// Whether the column/row carried a `$` (e.g. `$A1`, `A$1`, `$A$1`).
        /// Does not affect evaluation — only whether a reference-shift rewrite
        /// (0.14.0-A) treats it as anchored, and how it round-trips to text.
        abs_col: bool,
        abs_row: bool,
        /// `Some` for `Sheet2!A1` (0.14.0-A2). `evaluate` explicitly refuses to
        /// evaluate any expression containing a qualified reference — see
        /// `eval::references_another_sheet` — rather than ever silently reading
        /// the *active* sheet's cell as if it were the qualified one.
        sheet: Option<SheetQualifier>,
    },
    Range {
        c1: u32,
        r1: u32,
        c2: u32,
        r2: u32,
        /// Per-corner `$` flags — Excel allows mixed anchoring, e.g. `$A1:B$10`.
        abs_c1: bool,
        abs_r1: bool,
        abs_c2: bool,
        abs_r2: bool,
        /// `Some` for `Sheet2!A1:B10` — applies to the whole range, never
        /// per-corner (Excel has no such syntax; a per-sheet-range qualifier
        /// always precedes the first corner only). See `CellRef::sheet`.
        sheet: Option<SheetQualifier>,
    },
    BinOp {
        op: BinOpKind,
        lhs: Box<FormulaExpr>,
        rhs: Box<FormulaExpr>,
    },
    UnaryMinus(Box<FormulaExpr>),
    FuncCall {
        name: String,
        args: Vec<FormulaExpr>,
    },
}
