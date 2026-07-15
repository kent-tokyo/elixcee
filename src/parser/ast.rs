/// A character-offset range into the original source text (not bytes — the
/// hand-written tokenizer already indexes into `Vec<char>`, and column-by-
/// character is what matters for CJK text). Statement-level granularity:
/// see `SpannedStmt`.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct SourceSpan {
    pub start: u32,
    pub end: u32,
}

/// A statement paired with the span of source text it was parsed from.
/// `PartialEq` ignores nothing — two `SpannedStmt`s are equal only if both
/// the statement and its span match — but existing tests never compare
/// spans directly (see `parse_body` in `src/parser/mod.rs`, which strips
/// them before comparing plain `Stmt`s).
#[derive(Debug, Clone, PartialEq)]
pub struct SpannedStmt {
    pub stmt: Stmt,
    pub span: SourceSpan,
}

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
    RangeRead { addr: String },
    RangeOffsetRead { addr: String, row_off: Box<Expr>, col_off: Box<Expr> },
    CellsFind { what: Box<Expr>, find_row: bool },
    SheetCellRead { sheet: Box<Expr>, row: Box<Expr>, col: Box<Expr> },
    SheetRangeRead { sheet: Box<Expr>, addr: String },
    /// `Workbooks(workbook).Worksheets(sheet)` / `Workbooks(workbook).Sheets(sheet)`
    /// — wraps a plain sheet key (name or 1-based index) with a workbook
    /// identity to check first. elixcee only ever loads one workbook at a
    /// time (see `Vm::loaded_workbook_name`), so this does not model real
    /// multi-workbook switching — it only lets a mismatched workbook name
    /// be diagnosed (Milestone B6a). Valid wherever a plain sheet `Expr` is
    /// (`SheetCellRead`/`SheetRangeRead`/`SheetCellWrite`/`SheetRangeWrite`/
    /// `SheetsDelete`'s `sheet` field).
    WorkbookQualifiedSheet {
        workbook: Box<Expr>,
        sheet: Box<Expr>,
    },
    RowsCount,
    ColsCount,
    CellsEndProp { row: Box<Expr>, col: Box<Expr>, dir: XlDir, prop: XlEndProp },
    RecordGet       { var: String, field: String },           // p.x
    RecordGetNested { var: String, fields: Vec<String> },    // p.a.b.c
    ArrayRecordGet  { name: String, indices: Vec<Expr>, field: String }, // arr(i).f
}

#[derive(Debug, Clone, PartialEq)]
pub enum VbaBinOp {
    Add, Sub, Mul, Div,
    Eq, Ne, Lt, Le, Gt, Ge,
    Concat,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CalcModeValue { Automatic, Manual }

#[derive(Debug, Clone, PartialEq)]
pub enum CaseMatch {
    Value(Expr),
    Range(Expr, Expr),
    IsOp(VbaBinOp, Expr),
}

#[derive(Debug, Clone, PartialEq)]
pub enum Stmt {
    Assignment { var: String, value: Expr },
    CellWrite { row: Expr, col: Expr, value: Expr },
    SetCalcMode(CalcModeValue),
    SetAppProp { prop: String, value: Expr },
    RangeWrite { addr: String, is_formula: bool, value: Expr },
    /// `dst` is `None` for a bare `Range(src).Copy` (populates the VM's
    /// clipboard only); `Some(addr)` for `Range(src).Copy
    /// Destination:=Range(addr)` (also writes `addr` immediately).
    RangeCopy  { src: String, dst: Option<String> },
    /// `Range(dest_addr).Paste` / `Range(dest_addr).PasteSpecial
    /// [Transpose:=<expr>]` (Milestone B6b) — pastes the VM's clipboard
    /// contents into `dest_addr`. Real VBA only exposes `Transpose:=` on
    /// `.PasteSpecial`, not plain `.Paste`, so the parser only ever
    /// produces `Some(_)` for a `.PasteSpecial` statement.
    RangePaste { dest_addr: String, transpose: Option<Expr> },
    /// `Worksheets(sheet).Paste Destination:=Range(dest_addr)` (Milestone
    /// B6b). No `Transpose:=` here, matching real VBA's `Worksheet.Paste`.
    SheetRangePaste { sheet: Expr, dest_addr: String },
    RangeClear { addr: String, contents_only: bool },
    RangeOffsetWrite { addr: String, row_off: Expr, col_off: Expr, value: Expr },
    RangeDelete { addr: String },
    RangeInsert { addr: String },
    RangeSort { addr: String, key_col: u32, descending: bool },
    RangeName { addr: String, name: String },  // Range("A1:B3").Name = "MyRange"
    SheetCellWrite { sheet: Expr, row: Expr, col: Expr, value: Expr },
    SheetRangeWrite { sheet: Expr, addr: String, is_formula: bool, value: Expr },
    WithSheet { sheet_name: String, body: Vec<SpannedStmt> },
    SheetsAdd,
    SheetsDelete { sheet: Expr },
    For {
        var: String, from: Expr, to: Expr, step: Option<Expr>, body: Vec<SpannedStmt>,
    },
    ForEach {
        var: String,
        range_addr: String, // Range("A1:B10") address; variable iterables TBD
        body: Vec<SpannedStmt>,
    },
    If {
        condition: Expr, then_body: Vec<SpannedStmt>, else_body: Vec<SpannedStmt>,
    },
    DoLoop {
        pre_cond: Option<(bool, Expr)>, post_cond: Option<(bool, Expr)>, body: Vec<SpannedStmt>,
    },
    SelectCase {
        expr: Expr,
        cases: Vec<(Vec<CaseMatch>, Vec<SpannedStmt>)>,
        else_body: Vec<SpannedStmt>,
    },
    ExitFor,
    ExitDo,
    ExitSub,
    ExitFunction,
    OnError { resume_next: bool },     // On Error Resume Next (true) / GoTo 0 (false)
    OnErrorGoTo(String),               // On Error GoTo <label>
    Label(String),                     // <name>:  — marks a jump target
    GoTo(String),                      // GoTo <label>
    Resume { next: bool },             // Resume (false) / Resume Next (true)
    CallSub { name: String, args: Vec<Expr> },
    Dim,
    DimArray { name: String, sizes: Vec<Expr> },
    ReDim { name: String, sizes: Vec<Expr>, preserve: bool },
    ArrayWrite { name: String, indices: Vec<Expr>, value: Expr },
    With { body: Vec<SpannedStmt> },
    MsgBox { message: Expr },
    RecordSet { var: String, field: String, value: Expr }, // p.x = val
    DimRecord      { var: String, type_name: String },      // Dim p As PersonType
    DimArrayRecord { name: String, sizes: Vec<Expr>, type_name: String }, // Dim arr(10) As MyType
    RecordSetNested { var: String, fields: Vec<String>, value: Expr },    // p.a.b = val
    ArrayRecordSet  { name: String, indices: Vec<Expr>, field: String, value: Expr }, // arr(i).f=v
    WithRecord      { var: String, body: Vec<SpannedStmt> },      // With p ... End With
    /// A no-op the parser inserted because the construct on this line isn't
    /// recognized/implemented (as opposed to `Dim`, which is intentionally
    /// a no-op by design). Executes as a true no-op in the VM, same as
    /// `Dim` — this variant only exists so `check` can surface *why* a line
    /// silently did nothing.
    Unsupported { reason: String },
}

#[derive(Debug, Clone)]
pub struct SubDef {
    pub name: String,
    pub params: Vec<String>,
    pub body: Vec<SpannedStmt>,
}

#[derive(Debug, Clone)]
pub struct FuncDef {
    pub name: String,
    pub params: Vec<String>,
    pub body: Vec<SpannedStmt>,
}

/// A user-defined type field: (field_name_lowercase, vba_type_name_lowercase).
pub type TypeField = (String, String);

/// A `Type ... End Type` definition.
#[derive(Debug, Clone)]
pub struct TypeDef {
    pub name:   String,          // lowercase type name
    pub fields: Vec<TypeField>,  // (field_name, vba_type) in declaration order
}

#[derive(Debug, Clone)]
pub struct Program {
    pub subs:      Vec<SubDef>,
    pub funcs:     Vec<FuncDef>,
    pub type_defs: Vec<TypeDef>,
    /// Module-level lines that are unsupported/unevaluated (e.g. a
    /// module-level `Const`, which never actually sets its value —
    /// see `check::run_check`). Each entry is `(reason, span)`.
    pub module_diagnostics: Vec<(String, SourceSpan)>,
    /// The module's declared name, captured from `Attribute VB_Name =
    /// "..."` if present (as real VBA does). `None` if the module has no
    /// such line — callers fall back to a file-stem-derived name.
    pub module_name: Option<String>,
}
