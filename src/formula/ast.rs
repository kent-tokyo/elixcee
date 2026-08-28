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
