use std::collections::HashMap;

use crate::formula;
use crate::parser::ast::{CalcModeValue, CaseMatch, Expr, Program, Stmt, VbaBinOp, XlDir, XlEndProp};

/// Excel worksheet error values (#DIV/0!, #N/A, etc.)
#[derive(Debug, Clone, PartialEq)]
pub enum ExcelError {
    DivZero, // #DIV/0!
    NA,      // #N/A
    Value,   // #VALUE!
    Ref,     // #REF!
    Name,    // #NAME?
    Num,     // #NUM!
    Null,    // #NULL!
}

impl ExcelError {
    pub fn as_str(&self) -> &'static str {
        match self {
            ExcelError::DivZero => "#DIV/0!",
            ExcelError::NA      => "#N/A",
            ExcelError::Value   => "#VALUE!",
            ExcelError::Ref     => "#REF!",
            ExcelError::Name    => "#NAME?",
            ExcelError::Num     => "#NUM!",
            ExcelError::Null    => "#NULL!",
        }
    }
}

impl std::fmt::Display for ExcelError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.as_str())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Variant {
    Integer(i64),
    Float(f64),
    Str(String),
    Boolean(bool),
    Date(i64),           // Excel serial date — displays as "YYYY-MM-DD"
    Error(ExcelError),   // Excel error value (#DIV/0!, #N/A, …)
    Empty,
}

pub fn serial_to_display(s: i64) -> String {
    // Reuse formula engine's serial_to_ymd
    let (y, m, d) = crate::formula::eval::serial_to_ymd_pub(s);
    format!("{:04}-{:02}-{:02}", y, m, d)
}

impl std::fmt::Display for Variant {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Variant::Integer(n) => write!(f, "{}", n),
            Variant::Float(v)   => write!(f, "{}", v),
            Variant::Str(s)     => write!(f, "{}", s),
            Variant::Boolean(b) => write!(f, "{}", if *b { "True" } else { "False" }),
            Variant::Date(s)    => write!(f, "{}", serial_to_display(*s)),
            Variant::Error(e)   => write!(f, "{}", e),
            Variant::Empty      => write!(f, ""),
        }
    }
}

#[derive(Debug, Clone)]
pub struct CellContent {
    pub formula: Option<String>,
    pub value: Variant,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CalculationMode {
    Automatic,
    Manual,
}

pub struct Vm {
    pub cells: HashMap<(u32, u32), CellContent>,
    pub variables: HashMap<String, Variant>,
    pub calc_mode: CalculationMode,
    pub error_on_msgbox: bool,
}

impl Vm {
    pub fn new() -> Self {
        Vm {
            cells: HashMap::new(),
            variables: HashMap::new(),
            calc_mode: CalculationMode::Automatic,
            error_on_msgbox: false,
        }
    }

    pub fn run_sub(&mut self, program: &Program, sub_name: &str) -> Result<(), String> {
        let name = sub_name.to_lowercase();
        let body = program.subs.iter()
            .find(|s| s.name == name)
            .ok_or_else(|| format!("Sub '{}' not found", sub_name))?
            .body.clone();
        for stmt in &body { self.exec_stmt(stmt)?; }
        Ok(())
    }

    fn exec_stmt(&mut self, stmt: &Stmt) -> Result<(), String> {
        match stmt {
            Stmt::Assignment { var, value } => {
                let v = self.eval_expr(value)?;
                self.variables.insert(var.clone(), v);
            }
            Stmt::CellWrite { row, col, value } => {
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                let v = self.eval_expr(value)?;
                self.cells.insert((r, c), CellContent { formula: None, value: v });
            }
            Stmt::SetCalcMode(mode) => {
                let m = match mode {
                    CalcModeValue::Automatic => CalculationMode::Automatic,
                    CalcModeValue::Manual    => CalculationMode::Manual,
                };
                self.set_calc_mode(m)?;
            }
            Stmt::For { var, from, to, step, body } => {
                let mut i  = to_f64(&self.eval_expr(from)?)?;
                let to_f   = to_f64(&self.eval_expr(to)?)?;
                let step_f = match step {
                    Some(s) => to_f64(&self.eval_expr(s)?)?,
                    None    => 1.0,
                };
                if step_f == 0.0 { return Err("For loop: step cannot be zero".into()); }
                while (step_f > 0.0 && i <= to_f) || (step_f < 0.0 && i >= to_f) {
                    self.variables.insert(var.clone(), as_int_if_whole(i));
                    for s in body { self.exec_stmt(s)?; }
                    i += step_f;
                }
            }
            Stmt::If { condition, then_body, else_body } => {
                let branch = if is_truthy(&self.eval_expr(condition)?) { then_body } else { else_body };
                for s in branch { self.exec_stmt(s)?; }
            }
            Stmt::DoLoop { pre_cond, post_cond, body } => {
                let check = |vm: &Vm, cond: &Option<(bool, Expr)>| -> Result<bool, String> {
                    match cond {
                        None => Ok(true),
                        Some((is_until, expr)) => {
                            let v = vm.eval_expr(expr)?;
                            Ok(if *is_until { !is_truthy(&v) } else { is_truthy(&v) })
                        }
                    }
                };
                while check(self, pre_cond)? {
                    for s in body { self.exec_stmt(s)?; }
                    if !check(self, post_cond)? { break; }
                }
            }
            Stmt::SelectCase { expr, cases, else_body } => {
                let val = self.eval_expr(expr)?;
                let mut matched = false;
                'outer: for (matchers, body) in cases {
                    for m in matchers {
                        let hit = match m {
                            CaseMatch::Value(v) => vba_eq(&val, &self.eval_expr(v)?),
                            CaseMatch::Range(lo, hi) => {
                                let l = self.eval_expr(lo)?;
                                let h = self.eval_expr(hi)?;
                                vba_cmp(&val, &l)? != std::cmp::Ordering::Less
                                    && vba_cmp(&val, &h)? != std::cmp::Ordering::Greater
                            }
                            CaseMatch::IsOp(op, rhs) => {
                                let r = self.eval_expr(rhs)?;
                                match op {
                                    VbaBinOp::Eq => vba_eq(&val, &r),
                                    VbaBinOp::Ne => !vba_eq(&val, &r),
                                    VbaBinOp::Lt => vba_cmp(&val, &r)? == std::cmp::Ordering::Less,
                                    VbaBinOp::Le => vba_cmp(&val, &r)? != std::cmp::Ordering::Greater,
                                    VbaBinOp::Gt => vba_cmp(&val, &r)? == std::cmp::Ordering::Greater,
                                    VbaBinOp::Ge => vba_cmp(&val, &r)? != std::cmp::Ordering::Less,
                                    _ => false,
                                }
                            }
                        };
                        if hit {
                            for s in body { self.exec_stmt(s)?; }
                            matched = true;
                            break 'outer;
                        }
                    }
                }
                if !matched {
                    for s in else_body { self.exec_stmt(s)?; }
                }
            }
            Stmt::SetAppProp { prop: _, value } => {
                let _ = self.eval_expr(value); // evaluate for side-effect checking; result ignored
            }
            Stmt::RangeWrite { addr, is_formula, value } => {
                let v = self.eval_expr(value)?;
                let (row, col) = parse_cell_addr(addr)
                    .ok_or_else(|| format!("RangeWrite: invalid address '{}'", addr))?;
                if *is_formula {
                    let s = vba_to_str(&v);
                    self.set_cell_formula(row, col, &s)?;
                } else {
                    self.cells.insert((row, col), CellContent { formula: None, value: v });
                }
            }
            Stmt::RangeCopy { src, dst } => {
                let ((r1,c1),(r2,c2)) = parse_range_addr(src)
                    .ok_or_else(|| format!("RangeCopy: invalid source range '{}'", src))?;
                let (dr, dc) = parse_cell_addr(dst)
                    .unwrap_or((r1, c1));
                let vals: Vec<(u32, u32, Variant)> = (r1..=r2)
                    .flat_map(|r| (c1..=c2).map(move |c| (r, c)))
                    .map(|(r, c)| (r, c, self.get_cell(r, c)))
                    .collect();
                for (r, c, v) in vals {
                    let nr = dr + r - r1;
                    let nc = dc + c - c1;
                    self.cells.insert((nr, nc), CellContent { formula: None, value: v });
                }
            }
            Stmt::Dim => {} // no-op
            Stmt::With { body } => {
                for s in body { self.exec_stmt(s)?; }
            }
            Stmt::MsgBox { message } => {
                let msg = self.eval_expr(message)?;
                if self.error_on_msgbox {
                    return Err(format!("MsgBox: {}", msg));
                }
            }
        }
        Ok(())
    }

    pub fn eval_expr(&self, expr: &Expr) -> Result<Variant, String> {
        match expr {
            Expr::Integer(n)    => Ok(Variant::Integer(*n)),
            Expr::Float(f)      => Ok(Variant::Float(*f)),
            Expr::Str(s)        => Ok(Variant::Str(s.clone())),
            Expr::Bool(b)       => Ok(Variant::Boolean(*b)),
            Expr::Var(name) => {
                if let Some(v) = self.variables.get(name) { return Ok(v.clone()); }
                // Excel built-in constants
                Ok(match name.as_str() {
                    "xlcalculationmanual"        => Variant::Integer(-4135),
                    "xlcalculationautomatic"     => Variant::Integer(-4105),
                    "xlcalculationsemiautomatic" => Variant::Integer(2),
                    "xlup"                       => Variant::Integer(-4162),
                    "xldown"                     => Variant::Integer(-4121),
                    "xltoleft"                   => Variant::Integer(-4159),
                    "xltoright"                  => Variant::Integer(-4161),
                    "xlwait"                     => Variant::Integer(2),
                    "xldefault"                  => Variant::Integer(1),
                    "xlibeam"                    => Variant::Integer(3),
                    "xlnorthwestarrow"           => Variant::Integer(4),
                    _ => return Err(format!("Undefined variable: '{}'", name)),
                })
            }
            Expr::UnaryMinus(inner) => match self.eval_expr(inner)? {
                Variant::Integer(n) => Ok(Variant::Integer(-n)),
                Variant::Float(f)   => Ok(Variant::Float(-f)),
                other => Err(format!("Unary minus on non-numeric: {}", other)),
            },
            Expr::UnaryNot(inner) => {
                Ok(Variant::Boolean(!is_truthy(&self.eval_expr(inner)?)))
            }
            Expr::BinOp { op, lhs, rhs } => {
                let l = self.eval_expr(lhs)?;
                let r = self.eval_expr(rhs)?;
                eval_binop(op, l, r)
            }
            Expr::CellRead { row, col } => {
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                Ok(self.get_cell(r, c))
            }
            Expr::FuncCall { name, args } => self.eval_vba_func(name, args),
            Expr::RangeRead { addr } => {
                let (row, col) = parse_cell_addr(addr)
                    .ok_or_else(|| format!("RangeRead: invalid address '{}'", addr))?;
                Ok(self.get_cell(row, col))
            }
            Expr::RowsCount => Ok(Variant::Integer(1_048_576)),
            Expr::ColsCount => Ok(Variant::Integer(16_384)),
            Expr::CellsEndProp { row, col, dir, prop } => {
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                let result = match (dir, prop) {
                    (XlDir::Up,    XlEndProp::Row)    => self.last_nonempty_row(c, r),
                    (XlDir::Down,  XlEndProp::Row)    => self.first_empty_row(c, r).saturating_sub(1),
                    (XlDir::Left,  XlEndProp::Column) => self.last_nonempty_col(r, c),
                    (XlDir::Right, XlEndProp::Column) => self.first_empty_col(r, c).saturating_sub(1),
                    (XlDir::Up,    XlEndProp::Column) |
                    (XlDir::Down,  XlEndProp::Column) => c,
                    (XlDir::Left,  XlEndProp::Row)    |
                    (XlDir::Right, XlEndProp::Row)    => r,
                };
                Ok(Variant::Integer(result as i64))
            }
        }
    }

    fn eval_vba_func(&self, name: &str, args: &[Expr]) -> Result<Variant, String> {
        let vals: Vec<Variant> = args.iter().map(|a| self.eval_expr(a)).collect::<Result<_, _>>()?;
        match name {
            "int" => {
                let f = to_f64(vals.first().ok_or("INT requires 1 argument")?)?;
                Ok(as_int_if_whole(f.floor()))
            }
            "clng" | "cint" | "cbool" => {
                let f = to_f64(vals.first().ok_or("CInt/CLng requires 1 argument")?)?;
                Ok(Variant::Integer(f.round() as i64))
            }
            "cdbl" | "csng" => {
                let f = to_f64(vals.first().ok_or("CDbl requires 1 argument")?)?;
                Ok(Variant::Float(f))
            }
            "cstr" | "str" => {
                let s = vals.first().ok_or("CStr requires 1 argument")?.to_string();
                Ok(Variant::Str(s))
            }
            "val" => {
                let s = match vals.first().ok_or("Val requires 1 argument")? {
                    Variant::Str(s) => s.trim().to_string(),
                    v => v.to_string(),
                };
                let f = s.parse::<f64>().unwrap_or(0.0);
                Ok(as_int_if_whole(f))
            }
            "len" => {
                let s = match vals.first().ok_or("Len requires 1 argument")? {
                    Variant::Str(s) => s.chars().count() as i64,
                    Variant::Empty  => 0,
                    v               => v.to_string().chars().count() as i64,
                };
                Ok(Variant::Integer(s))
            }
            "left" => {
                let s = vba_to_str(vals.get(0).ok_or("Left requires 2 arguments")?);
                let n = to_f64(vals.get(1).ok_or("Left requires 2 arguments")?)? as usize;
                Ok(Variant::Str(s.chars().take(n).collect()))
            }
            "right" => {
                let s = vba_to_str(vals.get(0).ok_or("Right requires 2 arguments")?);
                let n = to_f64(vals.get(1).ok_or("Right requires 2 arguments")?)? as usize;
                let chars: Vec<char> = s.chars().collect();
                Ok(Variant::Str(chars[chars.len().saturating_sub(n)..].iter().collect()))
            }
            "mid" => {
                if vals.len() < 2 { return Err("Mid requires at least 2 arguments".into()); }
                let s = vba_to_str(&vals[0]);
                let start = (to_f64(&vals[1])? as usize).saturating_sub(1);
                let len = if vals.len() >= 3 { to_f64(&vals[2])? as usize } else { usize::MAX };
                Ok(Variant::Str(s.chars().skip(start).take(len).collect()))
            }
            "ucase" => {
                Ok(Variant::Str(vba_to_str(vals.first().ok_or("UCase requires 1 argument")?).to_uppercase()))
            }
            "lcase" => {
                Ok(Variant::Str(vba_to_str(vals.first().ok_or("LCase requires 1 argument")?).to_lowercase()))
            }
            "trim" => {
                let s = vba_to_str(vals.first().ok_or("Trim requires 1 argument")?);
                Ok(Variant::Str(s.trim().to_string()))
            }
            "ltrim" => {
                let s = vba_to_str(vals.first().ok_or("LTrim requires 1 argument")?);
                Ok(Variant::Str(s.trim_start().to_string()))
            }
            "rtrim" => {
                let s = vba_to_str(vals.first().ok_or("RTrim requires 1 argument")?);
                Ok(Variant::Str(s.trim_end().to_string()))
            }
            "abs" => {
                let f = to_f64(vals.first().ok_or("Abs requires 1 argument")?)?;
                Ok(as_int_if_whole(f.abs()))
            }
            "sqr" => {
                let f = to_f64(vals.first().ok_or("Sqr requires 1 argument")?)?;
                Ok(Variant::Float(f.sqrt()))
            }
            "isnull" | "isempty" => {
                Ok(Variant::Boolean(matches!(vals.first(), Some(Variant::Empty) | None)))
            }
            "isnumeric" => {
                Ok(Variant::Boolean(matches!(vals.first(), Some(Variant::Integer(_)) | Some(Variant::Float(_)))))
            }
            "chr" => {
                let n = to_f64(vals.first().ok_or("Chr requires 1 argument")?)? as u32;
                char::from_u32(n)
                    .map(|c| Variant::Str(c.to_string()))
                    .ok_or_else(|| format!("Chr: invalid code {}", n))
            }
            "asc" => {
                let s = vba_to_str(vals.first().ok_or("Asc requires 1 argument")?);
                s.chars().next()
                    .map(|c| Variant::Integer(c as i64))
                    .ok_or_else(|| "Asc: empty string".into())
            }
            "instr" => {
                // InStr([start,] string1, string2 [, compare])
                let (start, s1, s2) = if vals.len() >= 3 {
                    (to_f64(&vals[0])? as usize, vba_to_str(&vals[1]), vba_to_str(&vals[2]))
                } else {
                    (1, vba_to_str(vals.get(0).ok_or("InStr requires at least 2 arguments")?),
                        vba_to_str(vals.get(1).ok_or("InStr requires at least 2 arguments")?))
                };
                let h: Vec<char> = s1.chars().collect();
                let n: Vec<char> = s2.chars().collect();
                let pos = h[(start.saturating_sub(1))..].windows(n.len())
                    .position(|w| w.iter().map(|c| c.to_uppercase().next().unwrap_or(*c))
                        .eq(n.iter().map(|c| c.to_uppercase().next().unwrap_or(*c))))
                    .map(|p| p + start)
                    .unwrap_or(0);
                Ok(Variant::Integer(pos as i64))
            }
            "replace" => {
                if vals.len() < 3 { return Err("Replace requires at least 3 arguments".into()); }
                let s = vba_to_str(&vals[0]);
                let old = vba_to_str(&vals[1]);
                let new = vba_to_str(&vals[2]);
                Ok(Variant::Str(if old.is_empty() { s } else { s.replace(&old as &str, &new as &str) }))
            }
            "now" | "date" | "time" => {
                // Return as string representation (simplified)
                Ok(Variant::Str(format!("{:?}", std::time::SystemTime::now())))
            }
            _ => Err(format!("Unknown VBA function: '{}'", name)),
        }
    }

    pub fn get_cell(&self, row: u32, col: u32) -> Variant {
        self.cells.get(&(row, col)).map(|c| c.value.clone()).unwrap_or(Variant::Empty)
    }

    pub fn set_cell_formula(&mut self, row: u32, col: u32, formula: &str) -> Result<(), String> {
        let expr  = formula::parse(formula)?;
        let value = formula::evaluate(&expr, &self.cells)?;
        self.cells.insert((row, col), CellContent { formula: Some(formula.to_string()), value });
        Ok(())
    }

    pub fn recalculate_all(&mut self) -> Result<(), String> {
        let formulas: Vec<(u32, u32, String)> = self.cells.iter()
            .filter_map(|((r, c), cell)| cell.formula.as_ref().map(|f| (*r, *c, f.clone())))
            .collect();
        for (row, col, formula) in formulas {
            let expr  = formula::parse(&formula)?;
            let value = formula::evaluate(&expr, &self.cells)?;
            if let Some(cell) = self.cells.get_mut(&(row, col)) { cell.value = value; }
        }
        Ok(())
    }

    pub fn set_calc_mode(&mut self, mode: CalculationMode) -> Result<(), String> {
        let was_manual = self.calc_mode == CalculationMode::Manual;
        self.calc_mode = mode;
        if was_manual && self.calc_mode == CalculationMode::Automatic {
            self.recalculate_all()?;
        }
        Ok(())
    }

    /// Find the last non-empty row in `col` at or above `max_row` (xlUp).
    pub fn last_nonempty_row(&self, col: u32, max_row: u32) -> u32 {
        self.cells.iter()
            .filter(|((r, c), v)| *c == col && *r <= max_row && !matches!(v.value, Variant::Empty))
            .map(|((r, _), _)| *r)
            .max()
            .unwrap_or(1)
    }

    /// Find the first empty row in `col` at or below `start_row` (xlDown helper).
    pub fn first_empty_row(&self, col: u32, start_row: u32) -> u32 {
        let max = self.cells.keys().filter(|(_, c)| *c == col).map(|(r, _)| *r).max().unwrap_or(0);
        for r in start_row..=max + 1 {
            if !self.cells.contains_key(&(r, col)) || matches!(self.get_cell(r, col), Variant::Empty) {
                return r;
            }
        }
        max + 1
    }

    /// Find the last non-empty column in `row` at or left of `max_col` (xlToLeft).
    pub fn last_nonempty_col(&self, row: u32, max_col: u32) -> u32 {
        self.cells.iter()
            .filter(|((r, c), v)| *r == row && *c <= max_col && !matches!(v.value, Variant::Empty))
            .map(|((_, c), _)| *c)
            .max()
            .unwrap_or(1)
    }

    /// Find the first empty column in `row` at or right of `start_col` (xlToRight helper).
    pub fn first_empty_col(&self, row: u32, start_col: u32) -> u32 {
        let max = self.cells.keys().filter(|(r, _)| *r == row).map(|(_, c)| *c).max().unwrap_or(0);
        for c in start_col..=max + 1 {
            if !self.cells.contains_key(&(row, c)) || matches!(self.get_cell(row, c), Variant::Empty) {
                return c;
            }
        }
        max + 1
    }
}

impl Default for Vm {
    fn default() -> Self { Self::new() }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

// ── Range address helpers ─────────────────────────────────────────────────────

fn col_letters_to_num_vm(s: &str) -> u32 {
    s.chars().fold(0u32, |acc, c| acc * 26 + (c.to_ascii_uppercase() as u32 - b'A' as u32 + 1))
}

pub fn parse_cell_addr(addr: &str) -> Option<(u32, u32)> {
    let addr = addr.trim().to_uppercase();
    let alpha_end = addr.find(|c: char| c.is_ascii_digit())?;
    if alpha_end == 0 { return None; }
    let col = col_letters_to_num_vm(&addr[..alpha_end]);
    let row: u32 = addr[alpha_end..].parse().ok()?;
    Some((row, col))
}

pub fn parse_range_addr(addr: &str) -> Option<((u32, u32), (u32, u32))> {
    let addr = addr.trim();
    if let Some(i) = addr.find(':') {
        Some((parse_cell_addr(&addr[..i])?, parse_cell_addr(&addr[i+1..])?))
    } else {
        let c = parse_cell_addr(addr)?;
        Some((c, c))
    }
}

fn vba_to_str(v: &Variant) -> String {
    match v {
        Variant::Str(s)     => s.clone(),
        Variant::Integer(n) => n.to_string(),
        Variant::Float(f)   => f.to_string(),
        Variant::Boolean(b) => if *b { "True".into() } else { "False".into() },
        Variant::Date(s)    => serial_to_display(*s),
        Variant::Error(e)   => e.as_str().to_string(),
        Variant::Empty      => String::new(),
    }
}

fn to_f64(v: &Variant) -> Result<f64, String> {
    match v {
        Variant::Integer(n) => Ok(*n as f64),
        Variant::Float(f)   => Ok(*f),
        Variant::Boolean(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Variant::Date(s)    => Ok(*s as f64),
        Variant::Error(e)   => Err(e.to_string()),
        Variant::Empty      => Ok(0.0),
        Variant::Str(s)     => s.parse::<f64>().map_err(|_| format!("Cannot convert '{}' to number", s)),
    }
}

fn is_truthy(v: &Variant) -> bool {
    match v {
        Variant::Boolean(b) => *b,
        Variant::Integer(n) => *n != 0,
        Variant::Float(f)   => *f != 0.0,
        Variant::Str(s)     => !s.is_empty(),
        Variant::Date(_)    => true,
        Variant::Error(_)   => false,
        Variant::Empty      => false,
    }
}

fn vba_eq(a: &Variant, b: &Variant) -> bool {
    match (a, b) {
        (Variant::Integer(x), Variant::Integer(y)) => x == y,
        (Variant::Float(x),   Variant::Float(y))   => x == y,
        (Variant::Integer(x), Variant::Float(y))   => (*x as f64) == *y,
        (Variant::Float(x),   Variant::Integer(y)) => *x == (*y as f64),
        (Variant::Date(x),    Variant::Date(y))    => x == y,
        (Variant::Date(x),    Variant::Integer(y)) => x == y,
        (Variant::Integer(x), Variant::Date(y))    => x == y,
        (Variant::Str(x),     Variant::Str(y))     => x.to_uppercase() == y.to_uppercase(),
        (Variant::Boolean(x), Variant::Boolean(y)) => x == y,
        (Variant::Empty,      Variant::Empty)       => true,
        (Variant::Error(_),   _) | (_, Variant::Error(_)) => false,
        _ => false,
    }
}

fn vba_cmp(a: &Variant, b: &Variant) -> Result<std::cmp::Ordering, String> {
    to_f64(a)?.partial_cmp(&to_f64(b)?)
        .ok_or_else(|| "Cannot compare NaN values".into())
}

fn as_int_if_whole(f: f64) -> Variant {
    if f.fract() == 0.0 && f >= i64::MIN as f64 && f <= i64::MAX as f64 {
        Variant::Integer(f as i64)
    } else {
        Variant::Float(f)
    }
}

fn to_cell_index(v: Variant, label: &str) -> Result<u32, String> {
    let f = to_f64(&v)?;
    if f < 1.0 || f.fract() != 0.0 {
        return Err(format!("Cell {} must be a positive integer, got {}", label, f));
    }
    Ok(f as u32)
}

fn eval_binop(op: &VbaBinOp, l: Variant, r: Variant) -> Result<Variant, String> {
    match op {
        VbaBinOp::Add | VbaBinOp::Sub | VbaBinOp::Mul | VbaBinOp::Div => {
            let lf = to_f64(&l)?;
            let rf = to_f64(&r)?;
            let result = match op {
                VbaBinOp::Add => lf + rf,
                VbaBinOp::Sub => lf - rf,
                VbaBinOp::Mul => lf * rf,
                VbaBinOp::Div => {
                    if rf == 0.0 { return Err("Division by zero".into()); }
                    lf / rf
                }
                _ => unreachable!(),
            };
            Ok(as_int_if_whole(result))
        }
        VbaBinOp::Concat => Ok(Variant::Str(format!("{}{}", l, r))),
        VbaBinOp::Eq  => Ok(Variant::Boolean(vba_eq(&l, &r))),
        VbaBinOp::Ne  => Ok(Variant::Boolean(!vba_eq(&l, &r))),
        VbaBinOp::Lt  => Ok(Variant::Boolean(vba_cmp(&l, &r)? == std::cmp::Ordering::Less)),
        VbaBinOp::Le  => Ok(Variant::Boolean(vba_cmp(&l, &r)? != std::cmp::Ordering::Greater)),
        VbaBinOp::Gt  => Ok(Variant::Boolean(vba_cmp(&l, &r)? == std::cmp::Ordering::Greater)),
        VbaBinOp::Ge  => Ok(Variant::Boolean(vba_cmp(&l, &r)? != std::cmp::Ordering::Less)),
    }
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser;

    fn run(code: &str) -> Vm {
        let prog = parser::parse(code).unwrap();
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
        vm
    }

    #[test]
    fn test_variable_assignment_integer() {
        assert_eq!(run("Sub MySub()\n    a = 42\nEnd Sub\n").variables["a"], Variant::Integer(42));
    }

    #[test]
    fn test_variable_assignment_float() {
        assert_eq!(run("Sub MySub()\n    x = 1.5\nEnd Sub\n").variables["x"], Variant::Float(1.5));
    }

    #[test]
    fn test_variable_assignment_string() {
        assert_eq!(run("Sub MySub()\n    s = \"hello\"\nEnd Sub\n").variables["s"], Variant::Str("hello".into()));
    }

    #[test]
    fn test_cell_write_literal() {
        assert_eq!(run("Sub MySub()\n    Cells(1, 1).Value = 100\nEnd Sub\n").get_cell(1, 1), Variant::Integer(100));
    }

    #[test]
    fn test_cell_write_from_variable() {
        assert_eq!(run("Sub MySub()\n    x = 99\n    Cells(2, 3).Value = x\nEnd Sub\n").get_cell(2, 3), Variant::Integer(99));
    }

    #[test]
    fn test_cell_write_string() {
        assert_eq!(run("Sub MySub()\n    Cells(1, 2).Value = \"world\"\nEnd Sub\n").get_cell(1, 2), Variant::Str("world".into()));
    }

    #[test]
    fn test_cell_empty_by_default() {
        assert_eq!(Vm::new().get_cell(1, 1), Variant::Empty);
    }

    #[test]
    fn test_multiple_cells() {
        let vm = run("Sub MySub()\n    Cells(1, 1).Value = 1\n    Cells(1, 2).Value = 2\n    Cells(2, 1).Value = 3\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(1));
        assert_eq!(vm.get_cell(1, 2), Variant::Integer(2));
        assert_eq!(vm.get_cell(2, 1), Variant::Integer(3));
    }

    #[test]
    fn test_sub_not_found() {
        let prog = parser::parse("Sub MySub()\nEnd Sub\n").unwrap();
        assert!(Vm::new().run_sub(&prog, "nonexistent").is_err());
    }

    #[test]
    fn test_undefined_variable_error() {
        let prog = parser::parse("Sub MySub()\n    Cells(1, 1).Value = x\nEnd Sub\n").unwrap();
        assert!(Vm::new().run_sub(&prog, "mysub").is_err());
    }

    #[test]
    fn test_calculation_mode_default() {
        assert_eq!(Vm::new().calc_mode, CalculationMode::Automatic);
    }

    #[test]
    fn test_variant_display() {
        assert_eq!(Variant::Integer(42).to_string(), "42");
        assert_eq!(Variant::Float(3.14).to_string(), "3.14");
        assert_eq!(Variant::Boolean(true).to_string(), "True");
        assert_eq!(Variant::Empty.to_string(), "");
    }

    // ── Arithmetic ────────────────────────────────────────────────────────────

    #[test]
    fn test_arithmetic_assignment() {
        let vm = run("Sub MySub()\n    a = 3 + 4\n    b = 10 - 3\n    c = 2 * 5\n    d = 10 / 4\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(7));
        assert_eq!(vm.variables["b"], Variant::Integer(7));
        assert_eq!(vm.variables["c"], Variant::Integer(10));
        assert_eq!(vm.variables["d"], Variant::Float(2.5));
    }

    #[test]
    fn test_precedence_mul_over_add() {
        let vm = run("Sub MySub()\n    a = 1 + 2 * 3\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(7));
    }

    #[test]
    fn test_string_concat() {
        let vm = run("Sub MySub()\n    a = \"Hello\" & \" World\"\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Str("Hello World".into()));
    }

    #[test]
    fn test_comparison_result() {
        let vm = run("Sub MySub()\n    a = 5 > 3\n    b = 5 < 3\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Boolean(false));
    }

    // ── For loop ──────────────────────────────────────────────────────────────

    #[test]
    fn test_for_loop_sum() {
        // sum = 1 + 2 + 3 + 4 + 5 = 15
        let vm = run(
            "Sub MySub()\n    sum = 0\n    For i = 1 To 5\n        sum = sum + i\n    Next i\nEnd Sub\n",
        );
        assert_eq!(vm.variables["sum"], Variant::Integer(15));
    }

    #[test]
    fn test_for_loop_writes_cells() {
        let vm = run(
            "Sub MySub()\n    For i = 1 To 3\n        Cells(i, 1).Value = i\n    Next i\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(1));
        assert_eq!(vm.get_cell(2, 1), Variant::Integer(2));
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(3));
    }

    #[test]
    fn test_for_loop_step() {
        let vm = run(
            "Sub MySub()\n    s = 0\n    For i = 0 To 10 Step 2\n        s = s + i\n    Next i\nEnd Sub\n",
        );
        // 0 + 2 + 4 + 6 + 8 + 10 = 30
        assert_eq!(vm.variables["s"], Variant::Integer(30));
    }

    #[test]
    fn test_for_loop_negative_step() {
        let vm = run(
            "Sub MySub()\n    s = 0\n    For i = 5 To 1 Step -1\n        s = s + i\n    Next i\nEnd Sub\n",
        );
        // 5 + 4 + 3 + 2 + 1 = 15
        assert_eq!(vm.variables["s"], Variant::Integer(15));
    }

    // ── If / Else ─────────────────────────────────────────────────────────────

    #[test]
    fn test_if_true_branch() {
        let vm = run(
            "Sub MySub()\n    x = 10\n    If x > 5 Then\n        result = 1\n    End If\nEnd Sub\n",
        );
        assert_eq!(vm.variables["result"], Variant::Integer(1));
    }

    #[test]
    fn test_if_false_branch_not_taken() {
        let prog = parser::parse(
            "Sub MySub()\n    x = 1\n    If x > 5 Then\n        result = 1\n    End If\nEnd Sub\n",
        ).unwrap();
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
        assert!(!vm.variables.contains_key("result"));
    }

    #[test]
    fn test_if_else() {
        let vm = run(
            "Sub MySub()\n    x = 3\n    If x > 5 Then\n        result = 1\n    Else\n        result = 0\n    End If\nEnd Sub\n",
        );
        assert_eq!(vm.variables["result"], Variant::Integer(0));
    }

    // ── Do While / Until ──────────────────────────────────────────────────────

    #[test]
    fn test_do_while_loop() {
        let vm = run("Sub MySub()\n    x = 0\n    Do While x < 5\n        x = x + 1\n    Loop\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test]
    fn test_do_until_loop() {
        let vm = run("Sub MySub()\n    x = 0\n    Do Until x >= 5\n        x = x + 1\n    Loop\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test]
    fn test_do_loop_while_post() {
        // Post-check: body runs at least once even if condition is already false
        let vm = run("Sub MySub()\n    x = 99\n    Do\n        x = x + 1\n    Loop While x < 5\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(100));
    }

    // ── Select Case ───────────────────────────────────────────────────────────

    #[test]
    fn test_select_case_value() {
        let vm = run("Sub MySub()\n    x = 2\n    Select Case x\n        Case 1\n            r = \"one\"\n        Case 2\n            r = \"two\"\n        Case Else\n            r = \"other\"\n    End Select\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Str("two".into()));
    }

    #[test]
    fn test_select_case_else() {
        let vm = run("Sub MySub()\n    x = 99\n    Select Case x\n        Case 1\n            r = 1\n        Case Else\n            r = 0\n    End Select\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(0));
    }

    #[test]
    fn test_select_case_multi_value() {
        let vm = run("Sub MySub()\n    x = 3\n    Select Case x\n        Case 1, 2\n            r = 12\n        Case 3, 4\n            r = 34\n    End Select\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(34));
    }

    #[test]
    fn test_select_case_is_op() {
        let vm = run("Sub MySub()\n    x = 10\n    Select Case x\n        Case Is > 5\n            r = 1\n        Case Else\n            r = 0\n    End Select\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(1));
    }

    #[test]
    fn test_select_case_range() {
        let vm = run("Sub MySub()\n    x = 3\n    Select Case x\n        Case 1 To 5\n            r = 1\n        Case Else\n            r = 0\n    End Select\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(1));
    }

    // ── Dim ───────────────────────────────────────────────────────────────────

    #[test]
    fn test_dim_noop() {
        let vm = run("Sub MySub()\n    Dim x As Integer\n    x = 42\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    // ── With ... End With ─────────────────────────────────────────────────────

    #[test]
    fn test_with_block() {
        let vm = run("Sub MySub()\n    With Sheet1\n        .Cells(1, 1).Value = 100\n        .Cells(2, 1).Value = 200\n    End With\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(100));
        assert_eq!(vm.get_cell(2, 1), Variant::Integer(200));
    }

    // ── VBA built-in functions ────────────────────────────────────────────────

    #[test]
    fn test_vba_int() {
        let vm = run("Sub MySub()\n    a = Int(3.9)\n    b = Int(-3.1)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(3));
        assert_eq!(vm.variables["b"], Variant::Integer(-4));
    }

    #[test]
    fn test_vba_clng() {
        let vm = run("Sub MySub()\n    a = CLng(3.7)\n    b = CLng(-2.5)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(4));
    }

    #[test]
    fn test_vba_len() {
        let vm = run("Sub MySub()\n    a = Len(\"Hello\")\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(5));
    }

    #[test]
    fn test_vba_mid_left_right() {
        let vm = run("Sub MySub()\n    a = Mid(\"Hello\", 2, 3)\n    b = Left(\"Hello\", 3)\n    c = Right(\"Hello\", 2)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Str("ell".into()));
        assert_eq!(vm.variables["b"], Variant::Str("Hel".into()));
        assert_eq!(vm.variables["c"], Variant::Str("lo".into()));
    }

    #[test]
    fn test_vba_ucase_lcase() {
        let vm = run("Sub MySub()\n    a = UCase(\"hello\")\n    b = LCase(\"WORLD\")\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Str("HELLO".into()));
        assert_eq!(vm.variables["b"], Variant::Str("world".into()));
    }

    #[test]
    fn test_vba_not_and_bool() {
        let vm = run("Sub MySub()\n    a = Not True\n    b = Not False\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Boolean(false));
        assert_eq!(vm.variables["b"], Variant::Boolean(true));
    }

    // ── Application properties ────────────────────────────────────────────────

    #[test]
    fn test_app_prop_noop() {
        let vm = run("Sub MySub()\n    Application.ScreenUpdating = False\n    Application.EnableEvents = False\n    x = 1\n    Application.ScreenUpdating = True\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1));
    }

    #[test]
    fn test_xl_constants() {
        let vm = run("Sub MySub()\n    a = xlUp\n    b = xlDown\n    c = xlCalculationManual\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(-4162));
        assert_eq!(vm.variables["b"], Variant::Integer(-4121));
        assert_eq!(vm.variables["c"], Variant::Integer(-4135));
    }

    // ── Range write / copy / read ─────────────────────────────────────────────

    #[test]
    fn test_range_write_value() {
        let vm = run("Sub MySub()\n    Range(\"A1\").Value = 42\n    Range(\"B2\").Value = \"hello\"\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
        assert_eq!(vm.get_cell(2, 2), Variant::Str("hello".into()));
    }

    #[test]
    fn test_range_write_formula() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(1,2).Value = 20\n    Range(\"C1\").Formula = \"=SUM(A1:B1)\"\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 3), Variant::Integer(30));
    }

    #[test]
    fn test_range_copy() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    Range(\"A1:A3\").Copy Destination:=Range(\"B1\")\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 2), Variant::Integer(10));
        assert_eq!(vm.get_cell(2, 2), Variant::Integer(20));
        assert_eq!(vm.get_cell(3, 2), Variant::Integer(30));
    }

    #[test]
    fn test_range_read_expr() {
        let vm = run("Sub MySub()\n    Cells(5,1).Value = 99\n    x = Range(\"A5\").Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(99));
    }

    // ── Rows.Count / Columns.Count ────────────────────────────────────────────

    #[test]
    fn test_rows_count() {
        let vm = run("Sub MySub()\n    x = Rows.Count\n    y = Columns.Count\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1_048_576));
        assert_eq!(vm.variables["y"], Variant::Integer(16_384));
    }

    // ── Cells.End ─────────────────────────────────────────────────────────────

    #[test]
    fn test_cells_end_row_up() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    lastRow = Cells(Rows.Count,1).End(xlUp).Row\nEnd Sub\n");
        assert_eq!(vm.variables["lastrow"], Variant::Integer(3));
    }

    #[test]
    fn test_cells_end_col_left() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = \"a\"\n    Cells(1,2).Value = \"b\"\n    Cells(1,3).Value = \"c\"\n    lastCol = Cells(1,Columns.Count).End(xlToLeft).Column\nEnd Sub\n");
        assert_eq!(vm.variables["lastcol"], Variant::Integer(3));
    }

    // ── Nested For / If ───────────────────────────────────────────────────────

    #[test]
    fn test_nested_for_if() {
        // Sum values > 3 in range 1..5: 4 + 5 = 9
        let vm = run(
            "Sub MySub()\n    s = 0\n    For i = 1 To 5\n        If i > 3 Then\n            s = s + i\n        End If\n    Next i\nEnd Sub\n",
        );
        assert_eq!(vm.variables["s"], Variant::Integer(9));
    }
}
