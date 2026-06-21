#[derive(Debug, Clone, PartialEq)]
pub enum XlDir { Up, Down, Left, Right }

#[derive(Debug, Clone, PartialEq)]
pub enum XlEndProp { Row, Column }

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Integer(i64),
    Float(f64),
    Str(String),
    Bool(bool),
    Var(String),
    BinOp { op: VbaBinOp, lhs: Box<Expr>, rhs: Box<Expr> },
    UnaryMinus(Box<Expr>),
    UnaryNot(Box<Expr>),
    CellRead { row: Box<Expr>, col: Box<Expr> },
    FuncCall { name: String, args: Vec<Expr> },
    RangeRead { addr: String },           // Range("A1").Value in expressions
    RowsCount,                            // Rows.Count → 1048576
    ColsCount,                            // Columns.Count → 16384
    CellsEndProp {                        // Cells(r,c).End(dir).Row/Column
        row:  Box<Expr>,
        col:  Box<Expr>,
        dir:  XlDir,
        prop: XlEndProp,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub enum VbaBinOp {
    Add, Sub, Mul, Div,
    Eq, Ne, Lt, Le, Gt, Ge,
    Concat,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CalcModeValue {
    Automatic,
    Manual,
}

/// A single value matcher inside a Case clause.
#[derive(Debug, Clone, PartialEq)]
pub enum CaseMatch {
    Value(Expr),
    Range(Expr, Expr),    // Case 1 To 5
    IsOp(VbaBinOp, Expr), // Case Is > 5
}

#[derive(Debug, Clone, PartialEq)]
pub enum Stmt {
    Assignment { var: String, value: Expr },
    CellWrite { row: Expr, col: Expr, value: Expr },
    SetCalcMode(CalcModeValue),
    SetAppProp { prop: String, value: Expr }, // Application.ScreenUpdating etc. (no-op)
    RangeWrite { addr: String, is_formula: bool, value: Expr },
    RangeCopy  { src: String, dst: String },
    For {
        var: String,
        from: Expr,
        to: Expr,
        step: Option<Expr>,
        body: Vec<Stmt>,
    },
    If {
        condition: Expr,
        then_body: Vec<Stmt>,
        else_body: Vec<Stmt>,
    },
    DoLoop {
        pre_cond: Option<(bool, Expr)>,
        post_cond: Option<(bool, Expr)>,
        body: Vec<Stmt>,
    },
    SelectCase {
        expr: Expr,
        cases: Vec<(Vec<CaseMatch>, Vec<Stmt>)>,
        else_body: Vec<Stmt>,
    },
    Dim,
    With { body: Vec<Stmt> },
    MsgBox { message: Expr },
}

#[derive(Debug, Clone)]
pub struct SubDef {
    pub name: String,
    pub body: Vec<Stmt>,
}

#[derive(Debug, Clone)]
pub struct Program {
    pub subs: Vec<SubDef>,
}
