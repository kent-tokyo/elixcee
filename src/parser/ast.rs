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
    /// `ActiveSheet` used as a sheet qualifier (Milestone B7c item 6) —
    /// e.g. `ActiveSheet.Range("A1").Value`. Valid wherever a plain sheet
    /// `Expr` is (see `WorkbookQualifiedSheet`'s doc for the list); resolved
    /// to `Vm.active_sheet` by `resolve_sheet_expr`. `ThisWorkbook`/
    /// `ActiveWorkbook` don't need their own node — elixcee only ever loads
    /// one workbook (see `WorkbookQualifiedSheet`'s doc), so
    /// `ThisWorkbook.Worksheets(x)`/`ActiveWorkbook.Worksheets(x)` parse as
    /// a plain `Worksheets(x)`, the qualifier prefix simply skipped.
    ActiveSheetRef,
    /// A `Set`-assigned object variable used as a sheet qualifier (Phase 2C
    /// items 7/8) — e.g. `ws.Range("A1")` after `Set ws = ActiveSheet`, or
    /// `wb.Worksheets("X")` after `Set wb = ThisWorkbook`. Valid wherever a
    /// plain sheet `Expr` is, same as `ActiveSheetRef`. Holds the lowercase
    /// variable name; resolved against `Vm::object_variables` at *runtime*
    /// by `resolve_sheet_expr` — the parser has no variable-type tracking,
    /// so it can't tell `ws.Range(...)` apart from an ordinary UDT field
    /// access at parse time (see `parse_ident_stmt`'s dispatch order, which
    /// only special-cases `.Range(`/`.Cells(`/`.Worksheets(`/`.Sheets(`
    /// specifically to avoid misreading a genuine `p.range = ...` UDT
    /// field). A name that turns out not to hold an `ObjectRef::Worksheet`
    /// (unset, wrong type, or — for the `.Worksheets(...)`/`.Sheets(...)`
    /// qualifier-skip form — anything at all, since elixcee never checks
    /// that path holds a Workbook either) surfaces as a runtime error, not
    /// a parse error, same as `ObjectRef::Range`'s existing "'x' is
    /// Nothing" precedent.
    ObjectVarSheet(String),
    /// `<var> Is Nothing` — VBA's object-identity test against the null
    /// object reference. Holds the lowercase variable name; resolved against
    /// `Vm::object_variables` at *runtime* (the parser has no variable-type
    /// tracking — same situation `Expr::ObjectVarSheet` already accepts).
    /// Only the `Is Nothing` shape is modeled: a general `a Is b` object-
    /// identity comparison is still unparsed, exactly as before.
    IsNothing(String),
    /// A bare `.member` *read* inside a `With` body, e.g. the right-hand
    /// side of `.Value = .Value + 1`. Same runtime resolution as
    /// `Stmt::WithDot`, and likewise valid at any nesting depth.
    WithDot(Vec<String>),
    RecordGet       { var: String, field: String },           // p.x
    RecordGetNested { var: String, fields: Vec<String> },    // p.a.b.c
    ArrayRecordGet  { name: String, indices: Vec<Expr>, field: String }, // arr(i).f
}

#[derive(Debug, Clone, PartialEq)]
pub enum VbaBinOp {
    Add, Sub, Mul, Div,
    /// `Mod` — modulus. Result sign follows the dividend (left operand),
    /// same convention as Rust's `%`.
    Mod,
    /// `\` — integer division, truncating toward zero (same convention as
    /// Rust's integer `/`).
    IntDiv,
    /// `^` — exponentiation.
    Pow,
    Eq, Ne, Lt, Le, Gt, Ge,
    Concat,
    And, Or, Xor,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CalcModeValue { Automatic, Manual }

/// A reference-typed expression (Milestone B7c) — evaluates to an object
/// reference (currently only `Range`, possibly multi-area), never a plain
/// `Variant`. Kept as its own small AST, separate from `Expr`, matching
/// VBA's own value-vs-object distinction (`Set` vs plain `=`): `Expr`
/// always evaluates to a `Variant` in `Vm::eval_expr`, `ObjectExpr` always
/// evaluates to an `ObjectRef` in `Vm::eval_object_expr`. Only ever appears
/// as `Stmt::Set`'s RHS, or nested inside another `ObjectExpr` (`Union`'s
/// args, `Area`'s target, `SpecialCellsVisible`'s target) — the parser only
/// ever constructs one when it can fully recognize the shape; an
/// unrecognized `Set <var> = <rhs>` becomes `Stmt::Unsupported` instead of
/// a hard parse error (see `parse_set`), same no-op-on-unknown-construct
/// precedent as `Stmt::Dim`/`Stmt::Unsupported` elsewhere.
#[derive(Debug, Clone, PartialEq)]
pub enum ObjectExpr {
    /// `Range("A1:B2")` / `Range("A1:A3,C1:C3")` — a literal address,
    /// resolved against the active sheet at `Set`-evaluation time (real
    /// VBA fixes a Range object's parent worksheet at creation, not at
    /// each later `.Value` access — see `Vm::eval_object_expr`).
    RangeLit(String),
    /// An existing object variable, e.g. `Set b = a`. Copying the
    /// underlying `RangeRef` by value already gives real `Set` reference
    /// semantics here — see the doc on `Vm::object_variables`.
    Var(String),
    /// `Union(range1, range2, ...)` (Milestone B7c item 2).
    Union(Vec<ObjectExpr>),
    /// `<range>.Areas(n)` (Milestone B7c item 3) — 1-based, real VBA
    /// indexing. `n` is a plain VBA expression (evaluated as an Integer).
    Area(Box<ObjectExpr>, Box<Expr>),
    /// `<range>.SpecialCells(xlCellTypeVisible)` (Milestone B7c item 4) —
    /// only the `xlCellTypeVisible` constant (or its literal value, `12`)
    /// is recognized; every other `SpecialCells` type is unmodeled and
    /// falls back to `Stmt::Unsupported` at parse time.
    SpecialCellsVisible(Box<ObjectExpr>),
}

/// What a `With` block targets, kept unevaluated so the VM can resolve it
/// **once, at runtime, when the block is entered** — not as the parse-time
/// literal-string rewrite this used to be. That rewrite is why a computed
/// target (`With Cells(r, c)`) couldn't be expressed at all, and why a bare
/// `.member` nested inside another block construct in the body didn't parse.
#[derive(Debug, Clone, PartialEq)]
pub enum WithTarget {
    /// `With Range("A1")`, `With Union(...)`, … — anything `ObjectExpr`
    /// already models. Evaluated by `Vm::eval_object_expr` on block entry.
    Object(ObjectExpr),
    /// `With Cells(r, c)` — a computed single-cell target. Both index
    /// expressions are evaluated exactly once, on block entry.
    Cells(Expr, Expr),
    /// `With <identifier>` — resolved at runtime, since the parser has no
    /// variable-type tracking: the name may hold a `Set`-assigned Range or
    /// Worksheet reference (`Vm::object_variables`) or a UDT record
    /// (`Vm::variables`). Same runtime-resolution precedent as
    /// `Expr::ObjectVarSheet`.
    Var(String),
    /// `With Application` and any other target elixcee doesn't model — the
    /// body still runs, and a bare `.member` inside it is a no-op, matching
    /// the previous behavior for an unrecognized `With` header.
    Unmodeled,
}

/// The member path of a statement that begins with a bare `.` — resolved
/// against the innermost active `With` target at runtime, wherever in the
/// AST it appears (including inside an `If`/`For`/`Do`/`Select Case` nested
/// in the With body).
#[derive(Debug, Clone, PartialEq)]
pub enum WithMember {
    /// `.Value` / `.Formula` / `.a.b.c`
    Fields(Vec<String>),
    /// `.Cells(r, c).Value`
    Cells { row: Box<Expr>, col: Box<Expr>, fields: Vec<String> },
    /// `.Range("A1").Value`
    Range { addr: String, fields: Vec<String> },
}

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
    /// `<var>.Copy` / `<var>.Copy Destination:=Range(addr)` (Milestone
    /// B7c) — the object-variable sibling of `RangeCopy`, for a source
    /// produced by `Set` (a `Range("...")` reference, or a `Union`/
    /// `.Areas(n)`/`.SpecialCells(...)` result). `dst` stays a literal
    /// address, same convention as `RangeCopy.dst` — pasting *into* an
    /// object-variable destination isn't modeled.
    RangeObjectCopy { var: String, dst: Option<String> },
    /// `Set <var> = <value>` (Milestone B7c item 1) — VBA's object-
    /// assignment statement. Distinct from `Assignment` (plain `=`)
    /// because object variables live in `Vm::object_variables`, a
    /// namespace separate from `Vm::variables`, matching real VBA's own
    /// requirement that object references be assigned via `Set`.
    Set { var: String, value: ObjectExpr },
    /// `Range(dest_addr).Paste` / `Range(dest_addr).PasteSpecial
    /// [Transpose:=<expr>]` (Milestone B6b) — pastes the VM's clipboard
    /// contents into `dest_addr`. Real VBA only exposes `Transpose:=` on
    /// `.PasteSpecial`, not plain `.Paste`, so the parser only ever
    /// produces `Some(_)` for a `.PasteSpecial` statement.
    RangePaste { dest_addr: String, transpose: Option<Expr> },
    /// `Worksheets(sheet).Paste Destination:=Range(dest_addr)` (Milestone
    /// B6b). No `Transpose:=` here, matching real VBA's `Worksheet.Paste`.
    SheetRangePaste { sheet: Expr, dest_addr: String },
    /// `Sheets(sheet).Protect` (`protect: true`) / `.Unprotect`
    /// (`protect: false`) (Milestone B6c) — one variant with a bool flag,
    /// same convention as `Stmt::OnError { resume_next: bool }`.
    /// `ui_only` is `.Protect UserInterfaceOnly:=<expr>` — when truthy,
    /// real Excel blocks manual UI edits but *not* macro writes, so the VM
    /// must not add the sheet to `protected_sheets` in that case.
    SheetProtection {
        sheet: Expr,
        protect: bool,
        ui_only: Option<Expr>,
    },
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
    /// A truly no-op `Dim`-shaped statement: reached only when no variable
    /// name was actually available to record (a modifier — `Static`/
    /// `Friend` — followed by neither `Dim` nor `Const`, or a malformed
    /// `Dim` line with no identifier at all). Never introduces a name. A
    /// genuine `Dim x [As BuiltinType]` is `DimBare` instead, below.
    Dim,
    /// `Dim x` or `Dim x As <builtin type>` — a real declaration: `x` now
    /// exists as `Empty` until assigned, matching real VBA (`IsEmpty(x)`
    /// is `True` right after this runs, not an "undefined variable" error
    /// — this variant is what makes that possible; the old bare `Dim`
    /// never recorded the name at all).
    DimBare { var: String },
    DimArray { name: String, sizes: Vec<Expr> },
    ReDim { name: String, sizes: Vec<Expr>, preserve: bool },
    ArrayWrite { name: String, indices: Vec<Expr>, value: Expr },
    /// `With <target> ... End With`. `target` is resolved once at runtime on
    /// block entry and pushed onto the VM's With stack; every bare-`.member`
    /// statement/expression in `body` (at any nesting depth) resolves
    /// against the innermost entry. Replaces the old `WithRecord` variant
    /// too — a bare-identifier target is `WithTarget::Var`, disambiguated at
    /// runtime rather than at parse time.
    With { target: WithTarget, body: Vec<SpannedStmt> },
    /// A statement beginning with a bare `.` inside a `With` body, e.g.
    /// `.Value = 1` / `.Cells(i, 1).Value = i` / `.a.b = 2`. Valid wherever
    /// a statement is; resolving it against the innermost active With target
    /// is the VM's job, not the parser's.
    WithDot { member: WithMember, value: Expr },
    MsgBox { message: Expr },
    RecordSet { var: String, field: String, value: Expr }, // p.x = val
    DimRecord      { var: String, type_name: String },      // Dim p As PersonType
    DimArrayRecord { name: String, sizes: Vec<Expr>, type_name: String }, // Dim arr(10) As MyType
    /// `Dim a As Integer, b As Range, c(3) As MyType` — a comma-separated
    /// multi-declarator `Dim`. Each element is exactly what a single-
    /// declarator `Dim` would have produced on its own (`Dim`/`DimRecord`/
    /// `DimArray`/`DimArrayRecord`); the VM just runs them in order.
    DimMulti(Vec<Stmt>),
    RecordSetNested { var: String, fields: Vec<String>, value: Expr },    // p.a.b = val
    ArrayRecordSet  { name: String, indices: Vec<Expr>, field: String, value: Expr }, // arr(i).f=v
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
