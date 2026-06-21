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
    CellRef { col: u32, row: u32 },
    Range { c1: u32, r1: u32, c2: u32, r2: u32 },
    BinOp { op: BinOpKind, lhs: Box<FormulaExpr>, rhs: Box<FormulaExpr> },
    UnaryMinus(Box<FormulaExpr>),
    FuncCall { name: String, args: Vec<FormulaExpr> },
}
