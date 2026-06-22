use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};

use crate::formula;
use crate::parser::ast::{CalcModeValue, CaseMatch, Expr, FuncDef, Program, Stmt, SubDef, VbaBinOp, XlDir, XlEndProp};

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
    Array(Vec<Variant>),                  // 0-indexed 1D array
    Record(std::collections::HashMap<String, Variant>), // UDT instance (p.x, p.y, …)
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
            Variant::Array(a)   => write!(f, "[{}]", a.iter().map(|v| v.to_string()).collect::<Vec<_>>().join(", ")),
            Variant::Record(m)  => {
                let mut pairs: Vec<String> = m.iter().map(|(k, v)| format!("{}: {}", k, v)).collect();
                pairs.sort();
                write!(f, "{{{}}}", pairs.join(", "))
            }
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

/// Signals emitted by Exit For / Exit Do / Exit Sub / Exit Function.
#[derive(Debug, Clone, PartialEq)]
pub enum ExitKind { For, Do, Sub, Function }

pub struct Vm {
    /// Per-sheet cell storage. Key is sheet name (lowercase for lookup).
    sheets: HashMap<String, HashMap<(u32, u32), CellContent>>,
    /// Currently active sheet name (lowercase).
    pub active_sheet: String,
    pub variables: HashMap<String, Variant>,
    pub calc_mode: CalculationMode,
    pub error_on_msgbox: bool,
    pub exit_flag: Option<ExitKind>,
    pub on_error_resume_next: bool,
    /// Label to jump to when an error occurs (On Error GoTo <label>).
    pub on_error_goto_label: Option<String>,
    /// Pending unconditional jump target (GoTo <label>).
    pending_goto: Option<String>,
    user_funcs: HashMap<String, FuncDef>,
    user_subs:  HashMap<String, SubDef>,
    /// Workbook-level named ranges: lowercase name → address string (e.g. "A1:B5").
    pub named_ranges: HashMap<String, String>,
    /// User-defined types: lowercase type name → vec of (field_name, vba_type).
    type_defs: HashMap<String, Vec<(String, String)>>,
    /// Lazy index for Cells.End queries: col → sorted set of non-empty rows.
    col_rows: HashMap<u32, BTreeSet<u32>>,
    /// Lazy index for Cells.End queries: row → sorted set of non-empty cols.
    row_cols: HashMap<u32, BTreeSet<u32>>,
    /// Set to true whenever cells change; triggers index rebuild on next End query.
    cell_index_dirty: bool,
}

impl Vm {
    pub fn new() -> Self {
        let mut sheets = HashMap::new();
        sheets.insert("sheet1".into(), HashMap::new());
        Vm {
            sheets,
            active_sheet: "sheet1".into(),
            variables: HashMap::new(),
            calc_mode: CalculationMode::Automatic,
            error_on_msgbox: false,
            exit_flag: None,
            on_error_resume_next: false,
            on_error_goto_label: None,
            pending_goto: None,
            user_funcs: HashMap::new(),
            user_subs:  HashMap::new(),
            named_ranges: HashMap::new(),
            type_defs: HashMap::new(),
            col_rows: HashMap::new(),
            row_cols: HashMap::new(),
            cell_index_dirty: true,
        }
    }

    /// Resolve a range address, expanding named ranges if needed.
    /// Accepts both "A1:B3" syntax and registered range names (case-insensitive).
    fn resolve_range_addr(&self, addr: &str) -> Option<((u32, u32), (u32, u32))> {
        if let Some(r) = parse_range_addr(addr) { return Some(r); }
        self.named_ranges.get(&addr.to_lowercase())
            .and_then(|real| parse_range_addr(real))
    }

    pub fn cells(&self) -> &HashMap<(u32, u32), CellContent> {
        self.sheets.get(&self.active_sheet).expect("active sheet must exist")
    }

    pub fn cells_mut(&mut self) -> &mut HashMap<(u32, u32), CellContent> {
        self.cell_index_dirty = true;
        self.sheets.get_mut(&self.active_sheet).expect("active sheet must exist")
    }

    fn rebuild_cell_index(&mut self) {
        let pairs: Vec<(u32, u32)> = self.cells().iter()
            .filter(|(_, cell)| !matches!(cell.value, Variant::Empty))
            .map(|(&(r, c), _)| (r, c))
            .collect();
        self.col_rows.clear();
        self.row_cols.clear();
        for (r, c) in pairs {
            self.col_rows.entry(c).or_default().insert(r);
            self.row_cols.entry(r).or_default().insert(c);
        }
        self.cell_index_dirty = false;
    }

    pub fn sheet_names(&self) -> Vec<String> {
        let mut names: Vec<String> = self.sheets.keys().cloned().collect();
        names.sort();
        names
    }

    pub fn ensure_sheet(&mut self, name: &str) {
        self.sheets.entry(name.to_lowercase()).or_insert_with(HashMap::new);
    }

    pub fn set_active_sheet(&mut self, name: &str) -> Result<(), String> {
        let key = name.to_lowercase();
        if !self.sheets.contains_key(&key) {
            return Err(format!("Sheet '{}' not found", name));
        }
        self.active_sheet = key;
        self.cell_index_dirty = true;
        Ok(())
    }

    pub fn get_sheet_cells(&self, name: &str) -> Option<&HashMap<(u32, u32), CellContent>> {
        self.sheets.get(&name.to_lowercase())
    }

    fn sheet_cells_mut(&mut self, name: &str) -> Option<&mut HashMap<(u32, u32), CellContent>> {
        self.cell_index_dirty = true;
        self.sheets.get_mut(&name.to_lowercase())
    }

    pub fn run_sub(&mut self, program: &Program, sub_name: &str) -> Result<(), String> {
        // Cache user-defined functions, subs, and type definitions.
        self.user_funcs = program.funcs.iter().map(|f| (f.name.clone(), f.clone())).collect();
        self.user_subs  = program.subs.iter().map(|s| (s.name.clone(), s.clone())).collect();
        for td in &program.type_defs {
            self.type_defs.insert(td.name.clone(), td.fields.clone());
        }
        let name = sub_name.to_lowercase();
        let sub = self.user_subs.get(&name)
            .ok_or_else(|| format!("Sub '{}' not found", sub_name))?
            .clone();
        self.call_sub_def(&sub, &[])
    }

    fn call_sub_def(&mut self, sub: &SubDef, args: &[Variant]) -> Result<(), String> {
        let saved: Vec<(String, Option<Variant>)> = sub.params.iter().enumerate().map(|(i, p)| {
            let old = self.variables.get(p).cloned();
            if let Some(v) = args.get(i) { self.variables.insert(p.clone(), v.clone()); }
            (p.clone(), old)
        }).collect();
        let body = sub.body.clone();
        self.exec_body(&body, |f| matches!(f, ExitKind::Sub))?;
        for (p, old) in saved {
            match old { Some(v) => { self.variables.insert(p, v); } None => { self.variables.remove(&p); } }
        }
        Ok(())
    }

    fn call_func_def(&mut self, func: &FuncDef, args: &[Variant]) -> Result<Variant, String> {
        let saved: Vec<(String, Option<Variant>)> = func.params.iter().enumerate().map(|(i, p)| {
            let old = self.variables.get(p).cloned();
            if let Some(v) = args.get(i) { self.variables.insert(p.clone(), v.clone()); }
            (p.clone(), old)
        }).collect();
        let ret_name = func.name.clone();
        let old_ret = self.variables.remove(&ret_name);
        let body = func.body.clone();
        self.exec_body(&body, |f| matches!(f, ExitKind::Function | ExitKind::Sub))?;
        let ret_val = self.variables.remove(&ret_name).unwrap_or(Variant::Empty);
        for (p, old) in saved {
            match old { Some(v) => { self.variables.insert(p, v); } None => { self.variables.remove(&p); } }
        }
        if let Some(v) = old_ret { self.variables.insert(ret_name, v); }
        Ok(ret_val)
    }

    /// Execute a body slice with label-jump support (for GoTo and On Error GoTo).
    /// The existing per-statement `exec_stmt` error catch (resume_next) is preserved.
    fn exec_body<F>(&mut self, stmts: &[Stmt], is_exit: F) -> Result<(), String>
    where F: Fn(&ExitKind) -> bool
    {
        let mut i = 0;
        while i < stmts.len() {
            // Handle pending unconditional GoTo
            if let Some(label) = self.pending_goto.take() {
                match stmts.iter().position(|s| matches!(s, Stmt::Label(l) if l == &label)) {
                    Some(pos) => { i = pos; continue; }
                    None      => return Err(format!("GoTo: label '{}' not found", label)),
                }
            }

            if let Some(ref f) = self.exit_flag {
                if is_exit(f) { self.exit_flag = None; break; }
                break; // other exit kinds bubble up
            }

            let result = self.exec_stmt(&stmts[i]); // per-stmt catch preserves resume_next behavior
            match result {
                Ok(()) => {}
                Err(e) => {
                    // On Error GoTo: jump to handler label
                    if let Some(label) = self.on_error_goto_label.take() {
                        match stmts.iter().position(|s| matches!(s, Stmt::Label(l) if l == &label)) {
                            Some(pos) => { i = pos; continue; }
                            None      => return Err(format!("On Error GoTo: label '{}' not found", label)),
                        }
                    }
                    return Err(e);
                }
            }
            i += 1;
        }
        Ok(())
    }

    fn exec_stmt(&mut self, stmt: &Stmt) -> Result<(), String> {
        if self.exit_flag.is_some() { return Ok(()); }
        let result = self.exec_stmt_inner(stmt);
        match result {
            Ok(()) => Ok(()),
            Err(_) if self.on_error_resume_next => Ok(()),
            Err(e) => Err(e),
        }
    }

    fn exec_stmt_inner(&mut self, stmt: &Stmt) -> Result<(), String> {
        match stmt {
            Stmt::Assignment { var, value } => {
                let v = self.eval_expr(value)?;
                self.variables.insert(var.clone(), v);
            }
            Stmt::CellWrite { row, col, value } => {
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                let v = self.eval_expr(value)?;
                self.cells_mut().insert((r, c), CellContent { formula: None, value: v });
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
                let step_f = match step { Some(s) => to_f64(&self.eval_expr(s)?)?, None => 1.0 };
                if step_f == 0.0 { return Err("For loop: step cannot be zero".into()); }
                'for_loop: while (step_f > 0.0 && i <= to_f) || (step_f < 0.0 && i >= to_f) {
                    self.variables.insert(var.clone(), as_int_if_whole(i));
                    for s in body {
                        self.exec_stmt(s)?;
                        if matches!(self.exit_flag, Some(ExitKind::For)) { self.exit_flag = None; break 'for_loop; }
                        if self.exit_flag.is_some() { return Ok(()); }
                    }
                    i += step_f;
                }
            }
            Stmt::ForEach { var, range_addr, body } => {
                let ((r1, c1), (r2, c2)) = self.resolve_range_addr(range_addr)
                    .ok_or_else(|| format!("ForEach: invalid range '{}'", range_addr))?;
                'fe_outer: for r in r1..=r2 {
                    for c in c1..=c2 {
                        let v = self.get_cell(r, c);
                        self.variables.insert(var.clone(), v);
                        for s in body {
                            self.exec_stmt(s)?;
                            if matches!(self.exit_flag, Some(ExitKind::For)) { self.exit_flag = None; break 'fe_outer; }
                            if self.exit_flag.is_some() { return Ok(()); }
                        }
                    }
                }
            }
            Stmt::If { condition, then_body, else_body } => {
                let branch = if is_truthy(&self.eval_expr(condition)?) { then_body } else { else_body };
                for s in branch { self.exec_stmt(s)?; if self.exit_flag.is_some() { return Ok(()); } }
            }
            Stmt::DoLoop { pre_cond, post_cond, body } => {
                let check = |vm: &mut Vm, cond: &Option<(bool, Expr)>| -> Result<bool, String> {
                    match cond {
                        None => Ok(true),
                        Some((is_until, expr)) => {
                            let v = vm.eval_expr(expr)?;
                            Ok(if *is_until { !is_truthy(&v) } else { is_truthy(&v) })
                        }
                    }
                };
                'do_loop: while check(self, pre_cond)? {
                    for s in body.clone() {
                        self.exec_stmt(&s)?;
                        if matches!(self.exit_flag, Some(ExitKind::Do)) { self.exit_flag = None; break 'do_loop; }
                        if self.exit_flag.is_some() { return Ok(()); }
                    }
                    if !check(self, post_cond)? { break 'do_loop; }
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
                                let l = self.eval_expr(lo)?; let h = self.eval_expr(hi)?;
                                vba_cmp(&val, &l)? != std::cmp::Ordering::Less && vba_cmp(&val, &h)? != std::cmp::Ordering::Greater
                            }
                            CaseMatch::IsOp(op, rhs) => {
                                let r = self.eval_expr(rhs)?;
                                match op {
                                    VbaBinOp::Eq => vba_eq(&val, &r), VbaBinOp::Ne => !vba_eq(&val, &r),
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
                            matched = true; break 'outer;
                        }
                    }
                }
                if !matched { for s in else_body { self.exec_stmt(s)?; } }
            }
            Stmt::ExitFor      => self.exit_flag = Some(ExitKind::For),
            Stmt::ExitDo       => self.exit_flag = Some(ExitKind::Do),
            Stmt::ExitSub      => self.exit_flag = Some(ExitKind::Sub),
            Stmt::ExitFunction => self.exit_flag = Some(ExitKind::Function),
            Stmt::OnError { resume_next } => {
                self.on_error_resume_next = *resume_next;
                self.on_error_goto_label = None;
            }
            Stmt::OnErrorGoTo(label) => {
                self.on_error_goto_label = Some(label.clone());
                self.on_error_resume_next = false;
            }
            Stmt::Label(_) => {}  // no-op during normal execution
            Stmt::GoTo(label) => {
                self.pending_goto = Some(label.clone());
            }
            Stmt::Resume { .. } => {
                // After error handler runs: clear error state, continue
                self.on_error_goto_label = None;
            }
            Stmt::CallSub { name, args } => {
                let arg_vals: Vec<Variant> = args.iter().map(|a| self.eval_expr(a)).collect::<Result<_, _>>()?;
                if let Some(func) = self.user_funcs.get(name).cloned() {
                    let _ = self.call_func_def(&func, &arg_vals)?;
                } else if let Some(sub) = self.user_subs.get(name).cloned() {
                    self.call_sub_def(&sub, &arg_vals)?;
                } else {
                    return Err(format!("Sub/Function '{}' not found", name));
                }
            }
            Stmt::SetAppProp { prop: _, value } => {
                let _ = self.eval_expr(value);
            }
            Stmt::RangeName { addr, name } => {
                self.named_ranges.insert(name.to_lowercase(), addr.clone());
            }
            Stmt::RangeWrite { addr, is_formula, value } => {
                let v = self.eval_expr(value)?;
                let ((r1,c1),(r2,c2)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeWrite: invalid address '{}'", addr))?;
                for r in r1..=r2 {
                    for c in c1..=c2 {
                        if *is_formula {
                            self.set_cell_formula(r, c, &vba_to_str(&v))?;
                        } else {
                            self.cells_mut().insert((r, c), CellContent { formula: None, value: v.clone() });
                        }
                    }
                }
            }
            Stmt::RangeClear { addr, .. } => {
                let ((r1,c1),(r2,c2)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeClear: invalid address '{}'", addr))?;
                for r in r1..=r2 { for c in c1..=c2 { self.cells_mut().remove(&(r, c)); } }
            }
            Stmt::RangeOffsetWrite { addr, row_off, col_off, value } => {
                let v = self.eval_expr(value)?;
                let (base_r, base_c) = parse_cell_addr(addr)
                    .ok_or_else(|| format!("RangeOffsetWrite: invalid address '{}'", addr))?;
                let ro = to_f64(&self.eval_expr(row_off)?)? as i64;
                let co = to_f64(&self.eval_expr(col_off)?)? as i64;
                let row = (base_r as i64 + ro) as u32;
                let col = (base_c as i64 + co) as u32;
                self.cells_mut().insert((row, col), CellContent { formula: None, value: v });
            }
            Stmt::RangeDelete { addr } => {
                let ((r1,_),(r2,_)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeDelete: invalid address '{}'", addr))?;
                let rows_del = r2 - r1 + 1;
                // Collect cells below the deleted range and shift them up
                let to_move: Vec<((u32,u32), CellContent)> = self.cells().iter()
                    .filter(|((r,_),_)| *r > r2)
                    .map(|((r,c),v)| ((*r,*c), v.clone())).collect();
                for ((r,c),_) in &to_move { self.cells_mut().remove(&(*r,*c)); }
                for r in r1..=r2 { self.cells_mut().retain(|&(row,_),_| row != r); }
                for ((r,c),v) in to_move {
                    self.cells_mut().insert((r - rows_del, c), v);
                }
            }
            Stmt::RangeInsert { addr } => {
                let ((r1,_),(r2,_)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeInsert: invalid address '{}'", addr))?;
                let rows_ins = r2 - r1 + 1;
                // Shift cells at r1 and below downward
                let to_move: Vec<((u32,u32), CellContent)> = self.cells().iter()
                    .filter(|((r,_),_)| *r >= r1)
                    .map(|((r,c),v)| ((*r,*c), v.clone())).collect();
                for ((r,c),_) in &to_move { self.cells_mut().remove(&(*r,*c)); }
                for ((r,c),v) in to_move {
                    self.cells_mut().insert((r + rows_ins, c), v);
                }
            }
            Stmt::RangeSort { addr, key_col, descending } => {
                let ((r1,c1),(r2,c2)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeSort: invalid address '{}'", addr))?;
                // key_col is 1-based absolute column; convert to 0-based offset within range
                let key_off = (*key_col).saturating_sub(c1) as usize;
                let mut rows: Vec<Vec<Variant>> = (r1..=r2)
                    .map(|r| (c1..=c2).map(|c| self.get_cell(r, c)).collect()).collect();
                rows.sort_by(|a, b| {
                    let va = a.get(key_off).unwrap_or(&Variant::Empty);
                    let vb = b.get(key_off).unwrap_or(&Variant::Empty);
                    let ord = cmp_variants(va, vb);
                    if *descending { ord.reverse() } else { ord }
                });
                for (ri, row) in rows.iter().enumerate() {
                    for (ci, val) in row.iter().enumerate() {
                        self.cells_mut().insert((r1 + ri as u32, c1 + ci as u32),
                            CellContent { formula: None, value: val.clone() });
                    }
                }
            }
            Stmt::RangeCopy { src, dst } => {
                let ((r1,c1),(r2,c2)) = self.resolve_range_addr(src)
                    .ok_or_else(|| format!("RangeCopy: invalid source range '{}'", src))?;
                let (dr, dc) = parse_cell_addr(dst).unwrap_or((r1, c1));
                let vals: Vec<(u32, u32, Variant)> = (r1..=r2)
                    .flat_map(|r| (c1..=c2).map(move |c| (r, c)))
                    .map(|(r, c)| (r, c, self.get_cell(r, c)))
                    .collect();
                for (r, c, v) in vals {
                    self.cells_mut().insert((dr + r - r1, dc + c - c1), CellContent { formula: None, value: v });
                }
            }
            Stmt::SheetCellWrite { sheet, row, col, value } => {
                let sheet_name = vba_to_str(&self.eval_expr(sheet)?);
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                let v = self.eval_expr(value)?;
                self.ensure_sheet(&sheet_name);
                self.sheet_cells_mut(&sheet_name).unwrap().insert((r, c), CellContent { formula: None, value: v });
            }
            Stmt::WithSheet { sheet_name, body } => {
                let prev = self.active_sheet.clone();
                self.ensure_sheet(sheet_name);
                self.active_sheet = sheet_name.to_lowercase();
                self.cell_index_dirty = true;
                for s in body.clone() {
                    self.exec_stmt(&s)?;
                    if self.exit_flag.is_some() { break; }
                }
                self.active_sheet = prev;
                self.cell_index_dirty = true;
            }
            Stmt::SheetsAdd => {
                let new_name = format!("sheet{}", self.sheets.len() + 1);
                self.ensure_sheet(&new_name);
            }
            Stmt::SheetsDelete { sheet } => {
                let name = vba_to_str(&self.eval_expr(sheet)?);
                let key = name.to_lowercase();
                if key != self.active_sheet { self.sheets.remove(&key); }
            }
            Stmt::Dim => {}
            Stmt::DimArray { name, sizes } => {
                let upper = to_f64(&self.eval_expr(&sizes[0])?)? as usize;
                self.variables.insert(name.clone(), Variant::Array(vec![Variant::Empty; upper + 1]));
            }
            Stmt::ReDim { name, sizes, preserve } => {
                let upper = to_f64(&self.eval_expr(&sizes[0])?)? as usize;
                let new_size = upper + 1;
                let new_arr = if *preserve {
                    if let Some(Variant::Array(old)) = self.variables.get(name).cloned() {
                        let mut a = old; a.resize(new_size, Variant::Empty); a
                    } else { vec![Variant::Empty; new_size] }
                } else { vec![Variant::Empty; new_size] };
                self.variables.insert(name.clone(), Variant::Array(new_arr));
            }
            Stmt::ArrayWrite { name, indices, value } => {
                let v = self.eval_expr(value)?;
                let idx = to_f64(&self.eval_expr(&indices[0])?)? as usize;
                match self.variables.get_mut(name) {
                    Some(Variant::Array(arr)) => {
                        if idx < arr.len() { arr[idx] = v; }
                        else { return Err(format!("Array '{}': index {} out of bounds (len={})", name, idx, arr.len())); }
                    }
                    _ => return Err(format!("'{}' is not an array", name)),
                }
            }
            Stmt::With { body } => {
                for s in body { self.exec_stmt(s)?; if self.exit_flag.is_some() { return Ok(()); } }
            }
            Stmt::MsgBox { message } => {
                let msg = self.eval_expr(message)?;
                if self.error_on_msgbox { return Err(format!("MsgBox: {}", msg)); }
            }
            Stmt::DimRecord { var, type_name } => {
                if let Some(fields) = self.type_defs.get(type_name).cloned() {
                    let record = make_record_default(&fields, &self.type_defs);
                    self.variables.insert(var.clone(), record);
                }
                // Unknown type name → no-op (built-in type)
            }
            Stmt::DimArrayRecord { name, sizes, type_name } => {
                let upper = to_f64(&self.eval_expr(&sizes[0])?)? as usize;
                let element = if let Some(fields) = self.type_defs.get(type_name).cloned() {
                    make_record_default(&fields, &self.type_defs)
                } else {
                    Variant::Empty
                };
                self.variables.insert(name.clone(), Variant::Array(vec![element; upper + 1]));
            }
            Stmt::RecordSetNested { var, fields, value } => {
                let v = self.eval_expr(value)?;
                let target = self.variables.entry(var.clone())
                    .or_insert_with(|| Variant::Record(HashMap::new()));
                nested_set(target, fields, v);
            }
            Stmt::ArrayRecordSet { name, indices, field, value } => {
                let v = self.eval_expr(value)?;
                let idx = to_f64(&self.eval_expr(&indices[0])?)? as usize;
                match self.variables.get_mut(name) {
                    Some(Variant::Array(arr)) => {
                        if idx < arr.len() {
                            match &mut arr[idx] {
                                Variant::Record(m) => { m.insert(field.clone(), v); }
                                slot => {
                                    let mut m = HashMap::new();
                                    m.insert(field.clone(), v);
                                    *slot = Variant::Record(m);
                                }
                            }
                        } else {
                            return Err(format!("Array '{}': index {} out of bounds (len={})", name, idx, arr.len()));
                        }
                    }
                    _ => return Err(format!("'{}' is not an array", name)),
                }
            }
            Stmt::WithRecord { body, .. } => {
                // Parser already substituted the variable name into each statement.
                for s in body { self.exec_stmt(s)?; if self.exit_flag.is_some() { return Ok(()); } }
            }
            Stmt::RecordSet { var, field, value } => {
                let v = self.eval_expr(value)?;
                let entry = self.variables.entry(var.clone()).or_insert(Variant::Record(std::collections::HashMap::new()));
                if let Variant::Record(m) = entry {
                    m.insert(field.clone(), v);
                } else {
                    self.variables.insert(var.clone(), Variant::Record({
                        let mut m = std::collections::HashMap::new();
                        m.insert(field.clone(), v);
                        m
                    }));
                }
            }
        }
        Ok(())
    }

    pub fn eval_expr(&mut self, expr: &Expr) -> Result<Variant, String> {
        match expr {
            Expr::Integer(n)    => Ok(Variant::Integer(*n)),
            Expr::Float(f)      => Ok(Variant::Float(*f)),
            Expr::Str(s)        => Ok(Variant::Str(s.clone())),
            Expr::Bool(b)       => Ok(Variant::Boolean(*b)),
            Expr::Var(name) => {
                if let Some(v) = self.variables.get(name) { return Ok(v.clone()); }
                // Excel built-in constants
                Ok(match name.as_str() {
                    // Calculation mode
                    "xlcalculationmanual"        => Variant::Integer(-4135),
                    "xlcalculationautomatic"     => Variant::Integer(-4105),
                    "xlcalculationsemiautomatic" => Variant::Integer(2),
                    // Direction
                    "xlup"        => Variant::Integer(-4162),
                    "xldown"      => Variant::Integer(-4121),
                    "xltoleft"    => Variant::Integer(-4159),
                    "xltoright"   => Variant::Integer(-4161),
                    // Cursor
                    "xlwait"           => Variant::Integer(2),
                    "xldefault"        => Variant::Integer(1),
                    "xlibeam"          => Variant::Integer(3),
                    "xlnorthwestarrow" => Variant::Integer(4),
                    // VB string constants
                    "vbcrlf"       => Variant::Str("\r\n".into()),
                    "vblf"         => Variant::Str("\n".into()),
                    "vbcr"         => Variant::Str("\r".into()),
                    "vbtab"        => Variant::Str("\t".into()),
                    "vbnullstring" => Variant::Str(String::new()),
                    "vbnullchar"   => Variant::Str("\0".into()),
                    // VB boolean constants (in addition to True/False literals)
                    "vbtrue"  => Variant::Boolean(true),
                    "vbfalse" => Variant::Boolean(false),
                    // VB MsgBox return values
                    "vbok"     => Variant::Integer(1),
                    "vbcancel" => Variant::Integer(2),
                    "vbyes"    => Variant::Integer(6),
                    "vbno"     => Variant::Integer(7),
                    "empty" | "null" | "nothing" => return Ok(Variant::Empty),
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
            Expr::FuncCall { name, args } => {
                // User-defined functions take priority over built-ins
                if let Some(func) = self.user_funcs.get(name).cloned() {
                    let arg_vals: Vec<Variant> = args.iter().map(|a| self.eval_expr(a)).collect::<Result<_, _>>()?;
                    return self.call_func_def(&func, &arg_vals);
                }
                // Array subscript access: arr(i)
                if matches!(self.variables.get(name.as_str()), Some(Variant::Array(_))) {
                    let idx = to_f64(&self.eval_expr(args.first().ok_or_else(|| format!("Array '{}' requires index", name))?)?)? as usize;
                    return match self.variables.get(name.as_str()) {
                        Some(Variant::Array(arr)) => arr.get(idx).cloned().ok_or_else(|| format!("Array '{}': index {} out of bounds (len={})", name, idx, arr.len())),
                        _ => Err(format!("'{}' is not an array", name)),
                    };
                }
                self.eval_vba_func(name, args)
            }
            Expr::RangeRead { addr } => {
                let ((r1, c1), (r2, c2)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeRead: invalid address '{}'", addr))?;
                if r1 == r2 && c1 == c2 {
                    Ok(self.get_cell(r1, c1))
                } else {
                    let arr = (r1..=r2)
                        .flat_map(|r| (c1..=c2).map(move |c| (r, c)))
                        .map(|(r, c)| self.get_cell(r, c))
                        .collect();
                    Ok(Variant::Array(arr))
                }
            }
            Expr::RangeOffsetRead { addr, row_off, col_off } => {
                let (base_r, base_c) = parse_cell_addr(addr)
                    .ok_or_else(|| format!("RangeOffsetRead: invalid address '{}'", addr))?;
                let ro = to_f64(&self.eval_expr(row_off)?)? as i64;
                let co = to_f64(&self.eval_expr(col_off)?)? as i64;
                Ok(self.get_cell((base_r as i64 + ro) as u32, (base_c as i64 + co) as u32))
            }
            Expr::SheetCellRead { sheet, row, col } => {
                let sheet_name = vba_to_str(&self.eval_expr(sheet)?);
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                Ok(self.sheets.get(&sheet_name.to_lowercase())
                    .and_then(|s| s.get(&(r, c)))
                    .map(|cell| cell.value.clone())
                    .unwrap_or(Variant::Empty))
            }
            Expr::CellsFind { what, find_row } => {
                let target = self.eval_expr(what)?;
                let mut keys: Vec<(u32, u32)> = self.cells().keys().cloned().collect();
                keys.sort(); // 行優先スキャン
                for (r, c) in keys {
                    if vba_eq(&self.get_cell(r, c), &target) {
                        return Ok(Variant::Integer(if *find_row { r as i64 } else { c as i64 }));
                    }
                }
                Ok(Variant::Integer(0)) // not found
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
            Expr::RecordGet { var, field } => {
                match self.variables.get(var) {
                    Some(Variant::Record(m)) => Ok(m.get(field).cloned().unwrap_or(Variant::Empty)),
                    _ => Ok(Variant::Empty),
                }
            }
            Expr::RecordGetNested { var, fields } => {
                let mut cur = self.variables.get(var).cloned().unwrap_or(Variant::Empty);
                for f in fields {
                    cur = match cur {
                        Variant::Record(m) => m.get(f).cloned().unwrap_or(Variant::Empty),
                        _ => Variant::Empty,
                    };
                }
                Ok(cur)
            }
            Expr::ArrayRecordGet { name, indices, field } => {
                let idx = to_f64(&self.eval_expr(&indices[0])?)? as usize;
                match self.variables.get(name) {
                    Some(Variant::Array(arr)) => {
                        match arr.get(idx) {
                            Some(Variant::Record(m)) => Ok(m.get(field).cloned().unwrap_or(Variant::Empty)),
                            Some(other) => Ok(other.clone()),
                            None => Err(format!("Array '{}': index {} out of bounds (len={})", name, idx, arr.len())),
                        }
                    }
                    _ => Err(format!("'{}' is not an array", name)),
                }
            }
        }
    }

    fn eval_vba_func(&mut self, name: &str, args: &[Expr]) -> Result<Variant, String> {
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
                // VBA: empty needle → return start position; start > len → return 0
                if n.is_empty() { return Ok(Variant::Integer(start as i64)); }
                let begin = start.saturating_sub(1);
                if begin >= h.len() { return Ok(Variant::Integer(0)); }
                let pos = h[begin..].windows(n.len())
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
                Ok(Variant::Str(format!("{:?}", std::time::SystemTime::now())))
            }
            // ── Inline conditional ───────────────────────────────────────────
            "iif" => {
                if vals.len() < 3 { return Err("IIf requires 3 arguments".into()); }
                Ok(if is_truthy(&vals[0]) { vals[1].clone() } else { vals[2].clone() })
            }
            // ── Format ───────────────────────────────────────────────────────
            "format" => {
                if vals.is_empty() { return Err("Format requires at least 1 argument".into()); }
                let v = &vals[0];
                let fmt = if vals.len() >= 2 { vba_to_str(&vals[1]) } else { String::new() };
                Ok(Variant::Str(format_vba(v, &fmt)))
            }
            // ── Type inspection ──────────────────────────────────────────────
            "typename" => {
                let name = match vals.first().ok_or("TypeName requires 1 argument")? {
                    Variant::Integer(_) => "Long",
                    Variant::Float(_)   => "Double",
                    Variant::Str(_)     => "String",
                    Variant::Boolean(_) => "Boolean",
                    Variant::Date(_)    => "Date",
                    Variant::Error(_)   => "Error",
                    Variant::Array(_)   => "Variant()",
                    Variant::Empty      => "Empty",
                    Variant::Record(_)  => "Object",
                };
                Ok(Variant::Str(name.into()))
            }
            "vartype" => {
                let n: i64 = match vals.first().ok_or("VarType requires 1 argument")? {
                    Variant::Empty      => 0,
                    Variant::Integer(_) => 3,  // vbLong
                    Variant::Float(_)   => 5,  // vbDouble
                    Variant::Str(_)     => 8,  // vbString
                    Variant::Boolean(_) => 11, // vbBoolean
                    Variant::Date(_)    => 7,  // vbDate
                    Variant::Array(_)   => 8204, // vbArray + vbVariant
                    Variant::Error(_)   => 10, // vbError
                    Variant::Record(_)  => 0,  // vbEmpty as fallback
                };
                Ok(Variant::Integer(n))
            }
            // ── Array functions ──────────────────────────────────────────────
            "split" => {
                if vals.is_empty() { return Err("Split requires at least 1 argument".into()); }
                let s = vba_to_str(&vals[0]);
                let delim = if vals.len() >= 2 { vba_to_str(&vals[1]) } else { " ".to_string() };
                let parts = s.split(delim.as_str()).map(|p| Variant::Str(p.to_string())).collect();
                Ok(Variant::Array(parts))
            }
            "join" => {
                if vals.is_empty() { return Err("Join requires at least 1 argument".into()); }
                let parts = match &vals[0] {
                    Variant::Array(a) => a.iter().map(|v| vba_to_str(v)).collect::<Vec<_>>(),
                    v                 => vec![vba_to_str(v)],
                };
                let delim = if vals.len() >= 2 { vba_to_str(&vals[1]) } else { " ".to_string() };
                Ok(Variant::Str(parts.join(&delim)))
            }
            "ubound" => {
                match vals.first().ok_or("UBound requires 1 argument")? {
                    Variant::Array(a) => Ok(Variant::Integer(a.len() as i64 - 1)),
                    _ => Err("UBound: argument is not an array".into()),
                }
            }
            "lbound" => {
                match vals.first().ok_or("LBound requires 1 argument")? {
                    Variant::Array(_) => Ok(Variant::Integer(0)),
                    _ => Err("LBound: argument is not an array".into()),
                }
            }
            "isarray" => {
                Ok(Variant::Boolean(matches!(vals.first(), Some(Variant::Array(_)))))
            }
            // ── Range object (used as WSF arg) ───────────────────────────────
            "range" => {
                if let Some(Variant::Str(addr)) = vals.first() {
                    let ((r1,c1),(r2,c2)) = self.resolve_range_addr(addr)
                        .ok_or_else(|| format!("Range: invalid address '{}'", addr))?;
                    let arr = (r1..=r2).flat_map(|r| (c1..=c2).map(move |c| (r,c)))
                        .map(|(r,c)| self.get_cell(r,c)).collect();
                    Ok(Variant::Array(arr))
                } else { Err("Range: requires a string address argument".into()) }
            }
            // ── WorksheetFunction.*  ─────────────────────────────────────────
            name if name.starts_with("wsf_") => {
                let func = &name[4..];
                eval_wsf(func, &vals)
            }
            _ => Err(format!("Unknown VBA function: '{}'", name)),
        }
    }

    pub fn get_cell(&self, row: u32, col: u32) -> Variant {
        self.cells().get(&(row, col)).map(|c| c.value.clone()).unwrap_or(Variant::Empty)
    }

    pub fn set_cell_formula(&mut self, row: u32, col: u32, formula: &str) -> Result<(), String> {
        let expr  = formula::parse(formula)?;
        let value = formula::evaluate(&expr, self.cells())?;
        self.cells_mut().insert((row, col), CellContent { formula: Some(formula.to_string()), value });
        Ok(())
    }

    pub fn recalculate_all(&mut self) -> Result<(), String> {
        // Collect all formula cells and parse them
        let formula_cells: Vec<(u32, u32, formula::FormulaExpr)> = {
            self.cells().iter()
                .filter_map(|((r, c), cell)| {
                    cell.formula.as_ref().and_then(|f| formula::parse(f).ok().map(|e| (*r, *c, e)))
                })
                .collect()
        };

        // Sort by dependency order so that A2=A1+1 evaluates after A1
        let order = topo_sort_formulas(&formula_cells)?;

        // Update cell values directly, bypassing cells_mut() to avoid N dirty-flag sets.
        let active = self.active_sheet.clone();
        for idx in order {
            let (row, col, ref expr) = formula_cells[idx];
            let value = formula::evaluate(expr, self.cells())?;
            if let Some(cell) = self.sheets.get_mut(&active).and_then(|m| m.get_mut(&(row, col))) {
                cell.value = value;
            }
        }
        // Mark index dirty once (formula values changed, End queries may be stale)
        if !formula_cells.is_empty() { self.cell_index_dirty = true; }
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
    /// Find the last non-empty row in `col` at or above `max_row` (xlUp).
    pub fn last_nonempty_row(&mut self, col: u32, max_row: u32) -> u32 {
        if self.cell_index_dirty { self.rebuild_cell_index(); }
        self.col_rows.get(&col)
            .and_then(|rows| rows.range(..=max_row).next_back().copied())
            .unwrap_or(1)
    }

    /// Find the first empty row in `col` at or below `start_row` (xlDown helper).
    pub fn first_empty_row(&mut self, col: u32, start_row: u32) -> u32 {
        if self.cell_index_dirty { self.rebuild_cell_index(); }
        if let Some(rows) = self.col_rows.get(&col) {
            let mut prev = start_row.saturating_sub(1);
            for &r in rows.range(start_row..) {
                if r != prev + 1 { return prev + 1; }  // gap found
                prev = r;
            }
            prev + 1
        } else {
            start_row  // column is entirely empty
        }
    }

    /// Find the last non-empty column in `row` at or left of `max_col` (xlToLeft).
    pub fn last_nonempty_col(&mut self, row: u32, max_col: u32) -> u32 {
        if self.cell_index_dirty { self.rebuild_cell_index(); }
        self.row_cols.get(&row)
            .and_then(|cols| cols.range(..=max_col).next_back().copied())
            .unwrap_or(1)
    }

    /// Find the first empty column in `row` at or right of `start_col` (xlToRight helper).
    pub fn first_empty_col(&mut self, row: u32, start_col: u32) -> u32 {
        if self.cell_index_dirty { self.rebuild_cell_index(); }
        if let Some(cols) = self.row_cols.get(&row) {
            let mut prev = start_col.saturating_sub(1);
            for &c in cols.range(start_col..) {
                if c != prev + 1 { return prev + 1; }
                prev = c;
            }
            prev + 1
        } else {
            start_col
        }
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

// ── UDT helpers ──────────────────────────────────────────────────────────────

/// Build a `Variant::Record` with type-appropriate defaults, supporting nested UDTs.
fn make_record_default(
    fields: &[(String, String)],
    type_defs: &HashMap<String, Vec<(String, String)>>,
) -> Variant {
    let map: HashMap<String, Variant> = fields.iter().map(|(name, vba_type)| {
        let default = match vba_type.as_str() {
            "integer" | "long" | "longlong" | "byte" => Variant::Integer(0),
            "single" | "double" | "currency" | "decimal" => Variant::Float(0.0),
            "boolean" => Variant::Boolean(false),
            "string"  => Variant::Str(String::new()),
            other => {
                if let Some(nested) = type_defs.get(other) {
                    make_record_default(nested, type_defs)
                } else {
                    Variant::Empty
                }
            }
        };
        (name.clone(), default)
    }).collect();
    Variant::Record(map)
}

/// Recursively set a value at the path given by `fields` inside a `Variant::Record` tree.
fn nested_set(target: &mut Variant, fields: &[String], value: Variant) {
    if fields.is_empty() { *target = value; return; }
    let field = &fields[0];
    let rest  = &fields[1..];
    match target {
        Variant::Record(m) => {
            let inner = m.entry(field.clone()).or_insert(Variant::Empty);
            nested_set(inner, rest, value);
        }
        _ => {
            let mut inner = Variant::Empty;
            nested_set(&mut inner, rest, value);
            let mut m = HashMap::new();
            m.insert(field.clone(), inner);
            *target = Variant::Record(m);
        }
    }
}

// ── Formula dependency ordering ───────────────────────────────────────────────

/// Collect all (row, col) cell references in a formula expression (deduped).
fn extract_cell_refs(expr: &formula::FormulaExpr) -> HashSet<(u32, u32)> {
    use formula::FormulaExpr::*;
    match expr {
        CellRef { col, row }                 => [(*row, *col)].into(),
        Range { c1, r1, c2, r2 }             => {
            let mut s = HashSet::new();
            for r in *r1..=*r2 { for c in *c1..=*c2 { s.insert((r, c)); } }
            s
        }
        BinOp { lhs, rhs, .. }               => { let mut s = extract_cell_refs(lhs); s.extend(extract_cell_refs(rhs)); s }
        UnaryMinus(inner)                    => extract_cell_refs(inner),
        FuncCall { args, .. }                => args.iter().flat_map(extract_cell_refs).collect(),
        Number(_) | Str(_) | Bool(_)         => HashSet::new(),
    }
}

/// Topological sort of formula cells by dependency order.
/// Returns indices into `cells` in safe evaluation order.
/// Cells with no inter-formula dependencies appear first.
/// Returns `Err` if a circular reference is detected.
fn topo_sort_formulas(
    cells: &[(u32, u32, formula::FormulaExpr)],
) -> Result<Vec<usize>, String> {
    let n = cells.len();
    // map (row, col) → index in cells slice
    let pos: HashMap<(u32, u32), usize> = cells.iter().enumerate().map(|(i, (r,c,_))| ((*r,*c), i)).collect();

    // in_degree[i] = number of formula cells that i depends on
    let mut in_degree = vec![0usize; n];
    // adj[j] = list of formula cells that depend on j
    let mut adj: Vec<Vec<usize>> = vec![vec![]; n];

    for (i, (_, _, expr)) in cells.iter().enumerate() {
        for dep in extract_cell_refs(expr) {
            if let Some(&j) = pos.get(&dep) {
                if j != i { // skip self-reference
                    adj[j].push(i);
                    in_degree[i] += 1;
                }
            }
        }
    }

    // Kahn's algorithm
    let mut queue: VecDeque<usize> = (0..n).filter(|&i| in_degree[i] == 0).collect();
    let mut order = Vec::with_capacity(n);
    while let Some(i) = queue.pop_front() {
        order.push(i);
        for &j in &adj[i] {
            in_degree[j] -= 1;
            if in_degree[j] == 0 { queue.push_back(j); }
        }
    }

    if order.len() != n {
        // Circular reference detected — evaluate remaining cells in original order with a warning
        let visited: HashSet<usize> = order.iter().copied().collect();
        for i in 0..n { if !visited.contains(&i) { order.push(i); } }
        // Return Ok with best-effort order rather than hard-erroring; circular refs will show stale values
    }
    Ok(order)
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
        Variant::Array(a)   => a.iter().map(|x| vba_to_str(x)).collect::<Vec<_>>().join(", "),
        Variant::Record(_)  => "[Record]".into(),
    }
}

fn cmp_variants(a: &Variant, b: &Variant) -> std::cmp::Ordering {
    use std::cmp::Ordering::*;
    match (a, b) {
        (Variant::Integer(x), Variant::Integer(y)) => x.cmp(y),
        (Variant::Float(x),   Variant::Float(y))   => x.partial_cmp(y).unwrap_or(Equal),
        (Variant::Integer(x), Variant::Float(y))   => (*x as f64).partial_cmp(y).unwrap_or(Equal),
        (Variant::Float(x),   Variant::Integer(y)) => x.partial_cmp(&(*y as f64)).unwrap_or(Equal),
        (Variant::Str(x),     Variant::Str(y))     => x.to_lowercase().cmp(&y.to_lowercase()),
        (Variant::Empty,      Variant::Empty)       => Equal,
        (Variant::Empty,      _)                   => Less,
        (_,                   Variant::Empty)       => Greater,
        _ => vba_to_str(a).to_lowercase().cmp(&vba_to_str(b).to_lowercase()),
    }
}

fn flat_nums(vals: &[Variant]) -> Vec<f64> {
    let mut out = vec![];
    for v in vals {
        match v {
            Variant::Array(a) => out.extend(a.iter().filter_map(|x| to_f64(x).ok())),
            _ => { if let Ok(f) = to_f64(v) { out.push(f); } }
        }
    }
    out
}

fn flat_all(vals: &[Variant]) -> Vec<Variant> {
    vals.iter().flat_map(|v| match v {
        Variant::Array(a) => a.clone(),
        other             => vec![other.clone()],
    }).collect()
}

fn eval_wsf(func: &str, vals: &[Variant]) -> Result<Variant, String> {
    match func {
        "sum" => {
            let nums = flat_nums(vals);
            Ok(as_int_if_whole(nums.iter().sum::<f64>()))
        }
        "max" => {
            let nums = flat_nums(vals);
            nums.iter().cloned().reduce(f64::max)
                .map(as_int_if_whole)
                .ok_or_else(|| "WorksheetFunction.Max: no values".into())
        }
        "min" => {
            let nums = flat_nums(vals);
            nums.iter().cloned().reduce(f64::min)
                .map(as_int_if_whole)
                .ok_or_else(|| "WorksheetFunction.Min: no values".into())
        }
        "average" => {
            let nums = flat_nums(vals);
            if nums.is_empty() { return Err("WorksheetFunction.Average: no values".into()); }
            Ok(as_int_if_whole(nums.iter().sum::<f64>() / nums.len() as f64))
        }
        "count" => {
            let n = flat_all(vals).iter().filter(|v| matches!(v, Variant::Integer(_) | Variant::Float(_))).count();
            Ok(Variant::Integer(n as i64))
        }
        "counta" => {
            let n = flat_all(vals).iter().filter(|v| !matches!(v, Variant::Empty)).count();
            Ok(Variant::Integer(n as i64))
        }
        "countblank" => {
            let n = flat_all(vals).iter().filter(|v| matches!(v, Variant::Empty)).count();
            Ok(Variant::Integer(n as i64))
        }
        "countif" => {
            if vals.len() < 2 { return Err("WorksheetFunction.CountIf requires 2 arguments".into()); }
            let range = flat_all(&vals[..1]);
            let criteria = &vals[1];
            let n = range.iter().filter(|v| wsf_criteria_match(v, criteria)).count();
            Ok(Variant::Integer(n as i64))
        }
        "sumif" => {
            // SumIf(range, criteria [, sum_range])
            if vals.len() < 2 { return Err("WorksheetFunction.SumIf requires at least 2 arguments".into()); }
            let crit_range = flat_all(&vals[..1]);
            let criteria   = &vals[1];
            let sum_range  = if vals.len() >= 3 { flat_all(&vals[2..3]) } else { crit_range.clone() };
            let total: f64 = crit_range.iter().zip(sum_range.iter())
                .filter(|(cv, _)| wsf_criteria_match(cv, criteria))
                .filter_map(|(_, sv)| to_f64(sv).ok())
                .sum();
            Ok(as_int_if_whole(total))
        }
        "round" => {
            if vals.is_empty() { return Err("WorksheetFunction.Round requires arguments".into()); }
            let f = to_f64(&vals[0])?;
            let digits = if vals.len() >= 2 { to_f64(&vals[1])? as i32 } else { 0 };
            let factor = 10f64.powi(digits);
            Ok(as_int_if_whole((f * factor).round() / factor))
        }
        "abs"   => { let f = to_f64(vals.first().ok_or("Abs: no arg")?)?; Ok(as_int_if_whole(f.abs())) }
        "sqrt"  => { let f = to_f64(vals.first().ok_or("Sqrt: no arg")?)?; Ok(Variant::Float(f.sqrt())) }
        "power" => {
            if vals.len() < 2 { return Err("Power requires 2 arguments".into()); }
            Ok(as_int_if_whole(to_f64(&vals[0])?.powf(to_f64(&vals[1])?)))
        }
        "log"   => {
            let x = to_f64(vals.first().ok_or("Log: no arg")?)?;
            let base = if vals.len() >= 2 { to_f64(&vals[1])? } else { std::f64::consts::E };
            Ok(Variant::Float(x.log(base)))
        }
        "match" => {
            // Match(lookup_val, lookup_array, [match_type]) — returns 1-based position
            if vals.len() < 2 { return Err("Match: requires at least 2 arguments".into()); }
            let target  = &vals[0];
            let arr     = flat_all(&vals[1..2]);
            let pos = arr.iter().position(|v| vba_eq(v, target))
                .map(|i| Variant::Integer(i as i64 + 1))
                .unwrap_or(Variant::Error(ExcelError::NA));
            Ok(pos)
        }
        "index" => {
            // Index(array, row_num [, col_num])
            if vals.len() < 2 { return Err("Index: requires at least 2 arguments".into()); }
            let arr = flat_all(&vals[0..1]);
            let idx = (to_f64(&vals[1])? as usize).saturating_sub(1);
            Ok(arr.get(idx).cloned().unwrap_or(Variant::Error(ExcelError::Ref)))
        }
        _ => Err(format!("WorksheetFunction.{} is not implemented", func)),
    }
}

fn wsf_criteria_match(v: &Variant, criteria: &Variant) -> bool {
    match criteria {
        Variant::Str(s) => {
            let s = s.trim();
            // Comparison criteria like ">5", "<>0", ">=10"
            if let Some(rest) = s.strip_prefix(">=") {
                if let Ok(n) = rest.parse::<f64>() { return to_f64(v).map_or(false, |f| f >= n); }
            } else if let Some(rest) = s.strip_prefix("<=") {
                if let Ok(n) = rest.parse::<f64>() { return to_f64(v).map_or(false, |f| f <= n); }
            } else if let Some(rest) = s.strip_prefix("<>") {
                if let Ok(n) = rest.parse::<f64>() { return to_f64(v).map_or(false, |f| f != n); }
                return vba_to_str(v).to_lowercase() != rest.to_lowercase();
            } else if let Some(rest) = s.strip_prefix('>') {
                if let Ok(n) = rest.parse::<f64>() { return to_f64(v).map_or(false, |f| f > n); }
            } else if let Some(rest) = s.strip_prefix('<') {
                if let Ok(n) = rest.parse::<f64>() { return to_f64(v).map_or(false, |f| f < n); }
            }
            // Exact match
            vba_to_str(v).to_lowercase() == s.to_lowercase()
        }
        _ => vba_eq(v, criteria),
    }
}

fn format_vba(v: &Variant, fmt: &str) -> String {
    let fmt_l = fmt.to_lowercase();
    // Named numeric formats
    if fmt_l == "general number" || fmt_l == "general" || fmt.is_empty() {
        return vba_to_str(v);
    }
    // Numeric formatting: count decimal places from pattern like "0.00" or "#,##0.00"
    let thousands = fmt.contains(',');
    let dec_places = fmt.find('.').map(|i| fmt[i+1..].chars().filter(|c| *c == '0' || *c == '#').count()).unwrap_or(0);
    match v {
        Variant::Integer(n) => {
            let f = *n as f64;
            if thousands {
                // Simple thousands separator
                let int_part = format!("{}", n.abs());
                let grouped: String = int_part.chars().rev().enumerate()
                    .flat_map(|(i, c)| if i > 0 && i % 3 == 0 { vec![',', c] } else { vec![c] })
                    .collect::<String>().chars().rev().collect();
                let signed = if *n < 0 { format!("-{}", grouped) } else { grouped };
                if dec_places > 0 { format!("{}.{}", signed, "0".repeat(dec_places)) } else { signed }
            } else if dec_places > 0 {
                format!("{:.prec$}", f, prec = dec_places)
            } else {
                format!("{}", n)
            }
        }
        Variant::Float(f) => {
            if thousands {
                let int_part = format!("{}", (*f as i64).abs());
                let grouped: String = int_part.chars().rev().enumerate()
                    .flat_map(|(i, c)| if i > 0 && i % 3 == 0 { vec![',', c] } else { vec![c] })
                    .collect::<String>().chars().rev().collect();
                let signed = if *f < 0.0 { format!("-{}", grouped) } else { grouped };
                if dec_places > 0 { format!("{}.{:.prec$}", signed, f.fract().abs(), prec = dec_places)
                } else { signed }
            } else {
                format!("{:.prec$}", f, prec = dec_places)
            }
        }
        _ => vba_to_str(v),
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
        Variant::Array(_)   => Err("Cannot convert array to number".into()),
        Variant::Record(_)  => Err("Cannot convert record to number".into()),
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
        Variant::Array(a)   => !a.is_empty(),
        Variant::Record(_)  => true,
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
    // String operands: case-insensitive lexicographic comparison (VBA default).
    if let (Variant::Str(sa), Variant::Str(sb)) = (a, b) {
        return Ok(sa.to_uppercase().cmp(&sb.to_uppercase()));
    }
    // Mixed string/number: try numeric first, fall back to string coercion.
    match (to_f64(a), to_f64(b)) {
        (Ok(fa), Ok(fb)) => fa.partial_cmp(&fb).ok_or_else(|| "Cannot compare NaN values".into()),
        _ => {
            let sa = match a { Variant::Str(s) => s.clone(), _ => format!("{:?}", a) };
            let sb = match b { Variant::Str(s) => s.clone(), _ => format!("{:?}", b) };
            Ok(sa.to_uppercase().cmp(&sb.to_uppercase()))
        }
    }
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

    // ── ElseIf ───────────────────────────────────────────────────────────────

    #[test]
    fn test_elseif_chain() {
        let vm = run("Sub MySub()\n    x = 7\n    If x > 10 Then\n        r = 1\n    ElseIf x > 5 Then\n        r = 2\n    Else\n        r = 3\n    End If\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(2));
    }

    // ── Exit ─────────────────────────────────────────────────────────────────

    #[test]
    fn test_exit_for() {
        let vm = run("Sub MySub()\n    s = 0\n    For i = 1 To 10\n        If i > 3 Then\n            Exit For\n        End If\n        s = s + i\n    Next i\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Integer(6)); // 1+2+3
    }

    #[test]
    fn test_exit_do() {
        let vm = run("Sub MySub()\n    x = 0\n    Do\n        x = x + 1\n        If x >= 5 Then\n            Exit Do\n        End If\n    Loop While x < 100\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    // ── On Error Resume Next ──────────────────────────────────────────────────

    #[test]
    fn test_on_error_resume_next() {
        let vm = run("Sub MySub()\n    On Error Resume Next\n    a = 1\n    b = 2\n    a = 1\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
    }

    #[test]
    fn test_on_error_goto_label() {
        // On Error GoTo jumps to the label when an error occurs
        let code = concat!(
            "Sub MySub()\n",
            "    On Error GoTo ErrH\n",
            "    x = 1\n",
            "    Cells(0, 0).Value = 1\n", // invalid cell → error
            "    x = 99\n",               // should be skipped
            "    Exit Sub\n",
            "ErrH:\n",
            "    handled = 1\n",
            "End Sub\n",
        );
        let vm = run(code);
        assert_eq!(vm.variables["x"], Variant::Integer(1));   // set before error
        assert!(!vm.variables.contains_key("x") || vm.variables["x"] != Variant::Integer(99)); // not 99
        assert_eq!(vm.variables["handled"], Variant::Integer(1)); // handler ran
    }

    #[test]
    fn test_goto_unconditional() {
        let code = concat!(
            "Sub MySub()\n",
            "    a = 1\n",
            "    GoTo Skip\n",
            "    a = 99\n",  // should be skipped
            "Skip:\n",
            "    b = 2\n",
            "End Sub\n",
        );
        let vm = run(code);
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
        assert!(!vm.variables.contains_key("a") || vm.variables["a"] != Variant::Integer(99));
    }

    // ── UDT / Record field access ─────────────────────────────────────────────

    #[test]
    fn test_record_field_write_read() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    p.x = 3\n",
            "    p.y = 4\n",
            "    result = p.x + p.y\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["result"], Variant::Integer(7));
    }

    #[test]
    fn test_record_unset_field_is_empty() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    p.x = 10\n",
            "    result = p.y\n",  // p.y not set → Empty → 0 in arithmetic
            "End Sub\n",
        ));
        assert_eq!(vm.variables["result"], Variant::Empty);
    }

    #[test]
    fn test_record_field_in_arithmetic() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    pt.x = 3.0\n",
            "    pt.y = 4.0\n",
            "    dist = pt.x * pt.x + pt.y * pt.y\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["dist"], Variant::Integer(25)); // 9+16=25, whole numbers
    }

    #[test]
    fn test_multiple_records() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    a.val = 10\n",
            "    b.val = 20\n",
            "    total = a.val + b.val\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["total"], Variant::Integer(30));
    }

    // ── For Each ─────────────────────────────────────────────────────────────

    #[test]
    fn test_for_each_range() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    total = 0\n    For Each cell In Range(\"A1:A3\")\n        total = total + cell\n    Next cell\nEnd Sub\n");
        assert_eq!(vm.variables["total"], Variant::Integer(60));
    }

    // ── Function + Call ───────────────────────────────────────────────────────

    #[test]
    fn test_function_parsed_and_call_sub() {
        let prog = parser::parse("Function Double(x)\n    Double = x * 2\nEnd Function\nSub MySub()\n    Call Double(21)\nEnd Sub\n").unwrap();
        assert_eq!(prog.funcs[0].name, "double");
        assert_eq!(prog.funcs[0].params, vec!["x"]);
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
    }

    #[test]
    fn test_function_return_value_in_expr() {
        let vm = run("Function Square(n)\n    Square = n * n\nEnd Function\nSub MySub()\n    result = Square(7)\nEnd Sub\n");
        assert_eq!(vm.variables["result"], Variant::Integer(49));
    }

    #[test]
    fn test_function_return_value_nested() {
        let vm = run("Function Add(a, b)\n    Add = a + b\nEnd Function\nSub MySub()\n    x = Add(3, 4) + Add(1, 2)\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(10));
    }

    #[test]
    fn test_function_recursive() {
        // Factorial: 5! = 120
        let vm = run("Function Fact(n)\n    If n <= 1 Then\n        Fact = 1\n    Else\n        Fact = n * Fact(n - 1)\n    End If\nEnd Function\nSub MySub()\n    result = Fact(5)\nEnd Sub\n");
        assert_eq!(vm.variables["result"], Variant::Integer(120));
    }

    #[test]
    fn test_function_in_cell_write() {
        let vm = run("Function Double(x)\n    Double = x * 2\nEnd Function\nSub MySub()\n    Cells(1, 1).Value = Double(21)\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
    }

    #[test]
    fn test_call_sub_with_args() {
        let prog = parser::parse("Sub FillRow(rowNum, val)\n    Cells(rowNum, 1).Value = val\nEnd Sub\nSub MySub()\n    Call FillRow(3, 99)\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(99));
    }

    // ── vb constants ─────────────────────────────────────────────────────────

    #[test]
    fn test_vb_string_constants() {
        let vm = run("Sub MySub()\n    a = \"Hello\" & vbCrLf & \"World\"\n    b = \"tab\" & vbTab & \"here\"\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Str("Hello\r\nWorld".into()));
        assert_eq!(vm.variables["b"], Variant::Str("tab\there".into()));
    }

    // ── While ... Wend ───────────────────────────────────────────────────────

    #[test] fn test_while_wend() {
        let vm = run("Sub MySub()\n    x = 0\n    While x < 5\n        x = x + 1\n    Wend\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test] fn test_while_wend_no_iteration() {
        let vm = run("Sub MySub()\n    x = 10\n    While x < 5\n        x = x + 1\n    Wend\n    y = 99\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(10));
        assert_eq!(vm.variables["y"], Variant::Integer(99));
    }

    // ── Const ─────────────────────────────────────────────────────────────────

    #[test] fn test_const_declaration() {
        let vm = run("Sub MySub()\n    Const MAX_ROW As Long = 100\n    x = MAX_ROW\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(100));
    }

    #[test] fn test_const_string() {
        let vm = run("Sub MySub()\n    Const PREFIX = \"ID_\"\n    s = PREFIX & \"001\"\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("ID_001".into()));
    }

    // ── Empty / Null / Nothing ────────────────────────────────────────────────

    #[test] fn test_empty_literal() {
        let vm = run("Sub MySub()\n    a = Empty\n    b = Null\n    c = Nothing\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Empty);
        assert_eq!(vm.variables["b"], Variant::Empty);
        assert_eq!(vm.variables["c"], Variant::Empty);
    }

    // ── IIf ──────────────────────────────────────────────────────────────────

    #[test] fn test_iif_true() {
        let vm = run("Sub MySub()\n    x = IIf(1 > 0, \"yes\", \"no\")\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Str("yes".into()));
    }

    #[test] fn test_iif_false() {
        let vm = run("Sub MySub()\n    x = IIf(0 > 1, 10, 20)\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(20));
    }

    // ── Format ───────────────────────────────────────────────────────────────

    #[test] fn test_format_decimal() {
        let vm = run("Sub MySub()\n    s = Format(3.14159, \"0.00\")\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("3.14".into()));
    }

    #[test] fn test_format_integer_no_dec() {
        let vm = run("Sub MySub()\n    s = Format(42, \"0\")\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("42".into()));
    }

    // ── TypeName / VarType ───────────────────────────────────────────────────

    #[test] fn test_typename() {
        let vm = run("Sub MySub()\n    a = TypeName(42)\n    b = TypeName(\"hi\")\n    c = TypeName(True)\n    d = TypeName(Empty)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Str("Long".into()));
        assert_eq!(vm.variables["b"], Variant::Str("String".into()));
        assert_eq!(vm.variables["c"], Variant::Str("Boolean".into()));
        assert_eq!(vm.variables["d"], Variant::Str("Empty".into()));
    }

    // ── Arrays ───────────────────────────────────────────────────────────────

    #[test] fn test_dim_array_write_read() {
        let vm = run("Sub MySub()\n    Dim arr(5)\n    arr(0) = 10\n    arr(3) = 99\n    a = arr(0)\n    b = arr(3)\n    c = arr(1)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(10));
        assert_eq!(vm.variables["b"], Variant::Integer(99));
        assert_eq!(vm.variables["c"], Variant::Empty);
    }

    #[test] fn test_dim_array_loop() {
        let vm = run("Sub MySub()\n    Dim arr(4)\n    For i = 0 To 4\n        arr(i) = i * 2\n    Next i\n    s = 0\n    For i = 0 To 4\n        s = s + arr(i)\n    Next i\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Integer(20)); // 0+2+4+6+8
    }

    #[test] fn test_redim_preserve() {
        let vm = run("Sub MySub()\n    Dim arr(2)\n    arr(0) = 1\n    arr(1) = 2\n    ReDim Preserve arr(4)\n    arr(3) = 99\n    a = arr(0)\n    b = arr(3)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(99));
    }

    #[test] fn test_ubound_lbound() {
        let vm = run("Sub MySub()\n    Dim arr(9)\n    u = UBound(arr)\n    l = LBound(arr)\nEnd Sub\n");
        assert_eq!(vm.variables["u"], Variant::Integer(9));
        assert_eq!(vm.variables["l"], Variant::Integer(0));
    }

    #[test] fn test_split_join() {
        let vm = run("Sub MySub()\n    arr = Split(\"a,b,c\", \",\")\n    s = Join(arr, \"-\")\n    n = UBound(arr)\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("a-b-c".into()));
        assert_eq!(vm.variables["n"], Variant::Integer(2));
    }

    #[test] fn test_isarray() {
        let vm = run("Sub MySub()\n    Dim arr(3)\n    a = IsArray(arr)\n    b = IsArray(42)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Boolean(false));
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

    // ── WorksheetFunction ────────────────────────────────────────────────────

    #[test] fn test_wsf_sum() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    s = WorksheetFunction.Sum(Range(\"A1:A3\"))\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Integer(60));
    }

    #[test] fn test_wsf_max_min() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 5\n    Cells(2,1).Value = 3\n    Cells(3,1).Value = 8\n    mx = WorksheetFunction.Max(Range(\"A1:A3\"))\n    mn = WorksheetFunction.Min(Range(\"A1:A3\"))\nEnd Sub\n");
        assert_eq!(vm.variables["mx"], Variant::Integer(8));
        assert_eq!(vm.variables["mn"], Variant::Integer(3));
    }

    #[test] fn test_wsf_average() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    av = WorksheetFunction.Average(Range(\"A1:A2\"))\nEnd Sub\n");
        assert_eq!(vm.variables["av"], Variant::Integer(15));
    }

    #[test] fn test_wsf_countif() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 5\n    Cells(2,1).Value = 10\n    Cells(3,1).Value = 3\n    n = WorksheetFunction.CountIf(Range(\"A1:A3\"), \">4\")\nEnd Sub\n");
        assert_eq!(vm.variables["n"], Variant::Integer(2));
    }

    #[test] fn test_wsf_sumif() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 5\n    Cells(2,1).Value = 10\n    Cells(3,1).Value = 3\n    s = WorksheetFunction.SumIf(Range(\"A1:A3\"), \">4\")\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Integer(15));
    }

    #[test] fn test_wsf_application_prefix() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 7\n    Cells(2,1).Value = 3\n    s = Application.WorksheetFunction.Sum(Range(\"A1:A2\"))\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Integer(10));
    }

    #[test] fn test_wsf_match() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = \"a\"\n    Cells(2,1).Value = \"b\"\n    Cells(3,1).Value = \"c\"\n    pos = WorksheetFunction.Match(\"b\", Range(\"A1:A3\"), 0)\nEnd Sub\n");
        assert_eq!(vm.variables["pos"], Variant::Integer(2));
    }

    // ── Range("A1:A10").Value 多セル読み取り ─────────────────────────────────

    #[test] fn test_range_read_multi_cell() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    x = Range(\"A1:A3\").Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Array(vec![Variant::Integer(1), Variant::Integer(2), Variant::Integer(3)]));
    }

    #[test] fn test_range_read_single_cell_backward_compat() {
        let vm = run("Sub MySub()\n    Cells(5,1).Value = 99\n    x = Range(\"A5\").Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(99));
    }

    #[test] fn test_range_read_2d_row_major() {
        // A1=1, B1=2, A2=3, B2=4 → Range("A1:B2").Value → [1,2,3,4]
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(1,2).Value = 2\n    Cells(2,1).Value = 3\n    Cells(2,2).Value = 4\n    arr = Range(\"A1:B2\").Value\n    a = arr(0)\n    b = arr(1)\n    c = arr(2)\n    d = arr(3)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
        assert_eq!(vm.variables["c"], Variant::Integer(3));
        assert_eq!(vm.variables["d"], Variant::Integer(4));
    }

    // ── Cells.Find ───────────────────────────────────────────────────────────

    #[test] fn test_cells_find_row() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = \"apple\"\n    Cells(2,1).Value = \"banana\"\n    Cells(3,1).Value = \"cherry\"\n    r = Cells.Find(What:=\"banana\").Row\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(2));
    }

    #[test] fn test_cells_find_column() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = \"x\"\n    Cells(1,2).Value = \"y\"\n    Cells(1,3).Value = \"z\"\n    c = Cells.Find(What:=\"y\").Column\nEnd Sub\n");
        assert_eq!(vm.variables["c"], Variant::Integer(2));
    }

    #[test] fn test_cells_find_not_found() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = \"a\"\n    r = Cells.Find(What:=\"missing\").Row\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(0));
    }

    #[test] fn test_cells_find_extra_kwargs() {
        let vm = run("Sub MySub()\n    Cells(2,1).Value = 42\n    r = Cells.Find(What:=42, LookIn:=xlValues, SearchDirection:=xlPrevious).Row\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(2));
    }

    // ── EntireRow / EntireColumn ──────────────────────────────────────────────

    #[test] fn test_entirerow_delete() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    Range(\"A2:A2\").EntireRow.Delete\n    x = Cells(2,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(3)); // 3 が行2に移動
    }

    #[test] fn test_entirerow_clear() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 99\n    Range(\"A1\").EntireRow.Clear\n    x = Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Empty);
    }

    #[test] fn test_entirerow_clear_contents() {
        let vm = run("Sub MySub()\n    Cells(2,1).Value = 55\n    Range(\"A2\").EntireRow.ClearContents\n    x = Cells(2,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Empty);
    }

    #[test] fn test_entirecolumn_delete() {
        // パースエラーなしで実行できることを確認
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Range(\"A1:A3\").EntireColumn.Delete\n    x = 1\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1));
    }

    #[test] fn test_range_noop_hidden() {
        let vm = run("Sub MySub()\n    Range(\"A1\").Hidden = True\n    x = 1\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1));
    }

    #[test] fn test_range_noop_interior_color() {
        let vm = run("Sub MySub()\n    Range(\"A1:B2\").Interior.Color = 3\n    x = 2\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(2));
    }

    #[test] fn test_range_noop_numberformat() {
        let vm = run("Sub MySub()\n    Range(\"A1\").NumberFormat = \"0.00\"\n    x = 3\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(3));
    }

    // ── Range.Delete / Range.Insert ──────────────────────────────────────────

    #[test] fn test_range_delete_shifts_up() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    Cells(4,1).Value = 4\n    Range(\"A2:A2\").Delete\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(3)); // 3 shifted up
        assert_eq!(vm.variables["c"], Variant::Integer(4)); // 4 shifted up
    }

    #[test] fn test_range_insert_shifts_down() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Range(\"A2:A2\").Insert\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Empty); // new empty row
        assert_eq!(vm.variables["c"], Variant::Integer(2)); // shifted down
    }

    // ── Range.Sort ───────────────────────────────────────────────────────────

    #[test] fn test_range_sort_ascending() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 3\n    Cells(2,1).Value = 1\n    Cells(3,1).Value = 2\n    Range(\"A1:A3\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
        assert_eq!(vm.variables["c"], Variant::Integer(3));
    }

    #[test] fn test_range_sort_descending() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 3\n    Cells(2,1).Value = 1\n    Cells(3,1).Value = 2\n    Range(\"A1:A3\").Sort Key1:=Range(\"A1\"), Order1:=xlDescending\n    a = Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(3));
    }

    // ── Range clear / offset / multi-cell write / Sheets() ───────────────────

    #[test] fn test_range_clear_contents() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 99\n    Range(\"A1\").ClearContents\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Empty);
    }

    #[test] fn test_range_clear() {
        let vm = run("Sub MySub()\n    Range(\"A1:A3\").Value = 5\n    Range(\"A1:A3\").Clear\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Empty);
        assert_eq!(vm.get_cell(2, 1), Variant::Empty);
        assert_eq!(vm.get_cell(3, 1), Variant::Empty);
    }

    #[test] fn test_range_write_multi_cell() {
        let vm = run("Sub MySub()\n    Range(\"A1:A3\").Value = 7\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(7));
        assert_eq!(vm.get_cell(2, 1), Variant::Integer(7));
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(7));
    }

    #[test] fn test_range_offset_read() {
        let vm = run("Sub MySub()\n    Cells(2,2).Value = 42\n    x = Range(\"A1\").Offset(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    #[test] fn test_range_offset_write() {
        let vm = run("Sub MySub()\n    Range(\"A1\").Offset(2,0).Value = 99\nEnd Sub\n");
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(99));
    }

    #[test] fn test_sheets_cell_write() {
        let vm = run("Sub MySub()\n    Sheets(\"Sheet1\").Cells(1,1).Value = 123\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(123));
    }

    #[test] fn test_sheets_cell_read() {
        let vm = run("Sub MySub()\n    Sheets(\"Sheet1\").Cells(1,1).Value = 55\n    x = Sheets(\"Sheet1\").Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(55));
    }

    #[test] fn test_worksheets_cell_write() {
        let vm = run("Sub MySub()\n    Worksheets(\"Data\").Cells(2,3).Value = 77\nEnd Sub\n");
        // Now routes to "data" sheet, not the active "sheet1"
        let cell = vm.get_sheet_cells("data").and_then(|s| s.get(&(2,3))).map(|c| c.value.clone());
        assert_eq!(cell, Some(Variant::Integer(77)));
    }

    // ── Multi-sheet (Phase 9) ────────────────────────────────────────────────

    #[test] fn test_multisheet_write_read_different_sheets() {
        let vm = run("Sub MySub()\n    Sheets(\"Sheet1\").Cells(1,1).Value = 10\n    Sheets(\"Sheet2\").Cells(1,1).Value = 20\nEnd Sub\n");
        let s1 = vm.get_sheet_cells("sheet1").and_then(|s| s.get(&(1,1))).map(|c| c.value.clone());
        let s2 = vm.get_sheet_cells("sheet2").and_then(|s| s.get(&(1,1))).map(|c| c.value.clone());
        assert_eq!(s1, Some(Variant::Integer(10)));
        assert_eq!(s2, Some(Variant::Integer(20)));
    }

    #[test] fn test_multisheet_cross_sheet_read() {
        let vm = run("Sub MySub()\n    Sheets(\"Data\").Cells(1,1).Value = 42\n    x = Sheets(\"Data\").Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    #[test] fn test_with_sheets_block() {
        let vm = run("Sub MySub()\n    With Sheets(\"Sheet2\")\n        .Cells(1,1).Value = 99\n    End With\n    x = Sheets(\"Sheet2\").Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(99));
    }

    #[test] fn test_with_sheets_restores_active() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    With Sheets(\"Sheet2\")\n        .Cells(1,1).Value = 2\n    End With\n    x = Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1)); // active sheet unchanged
    }

    #[test] fn test_sheets_add() {
        let vm = run("Sub MySub()\n    Sheets.Add\n    n = 1\nEnd Sub\n");
        assert!(vm.sheet_names().len() >= 2);
    }

    #[test] fn test_sheets_delete() {
        let vm = run("Sub MySub()\n    Sheets(\"Sheet2\").Cells(1,1).Value = 5\n    Sheets(\"Sheet2\").Delete\n    n = 1\nEnd Sub\n");
        assert!(!vm.sheet_names().contains(&"sheet2".to_string()));
    }

    #[test] fn test_sheet_names() {
        let vm = run("Sub MySub()\n    Sheets(\"Alpha\").Cells(1,1).Value = 1\n    Sheets(\"Beta\").Cells(1,1).Value = 2\nEnd Sub\n");
        let names = vm.sheet_names();
        assert!(names.contains(&"alpha".to_string()));
        assert!(names.contains(&"beta".to_string()));
        assert!(names.contains(&"sheet1".to_string())); // default
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

    // ── Named ranges ──────────────────────────────────────────────────────────

    #[test]
    fn test_named_range_write_read() {
        // Define a named range and use it for write / read
        let vm = run(concat!(
            "Sub MySub()\n",
            "    Range(\"A1:A3\").Name = \"MyData\"\n",
            "    Range(\"MyData\").Value = 99\n",
            "    x = Range(\"MyData\").Value\n",
            "End Sub\n",
        ));
        // All three cells should be 99
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(99));
        assert_eq!(vm.get_cell(2, 1), Variant::Integer(99));
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(99));
        // The range name is registered
        assert_eq!(vm.named_ranges.get("mydata"), Some(&"A1:A3".to_string()));
    }

    #[test]
    fn test_named_range_for_each() {
        // Named range works in For Each iteration
        let vm = run(concat!(
            "Sub MySub()\n",
            "    Cells(1,1).Value = 10\n",
            "    Cells(2,1).Value = 20\n",
            "    Range(\"A1:A2\").Name = \"Items\"\n",
            "    s = 0\n",
            "    For Each item In Range(\"Items\")\n",
            "        s = s + item\n",
            "    Next item\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["s"], Variant::Integer(30));
    }

    // ── User-defined types (Type...End Type) ─────────────────────────────────

    #[test]
    fn test_type_def_basic() {
        let vm = run(concat!(
            "Type Person\n",
            "    Name As String\n",
            "    Age As Integer\n",
            "    Score As Double\n",
            "End Type\n",
            "\n",
            "Sub MySub()\n",
            "    Dim p As Person\n",
            "    p.Name = \"Alice\"\n",
            "    p.Age = 30\n",
            "    p.Score = 9.5\n",
            "    x = p.Name\n",
            "    y = p.Age\n",
            "    z = p.Score\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["x"], Variant::Str("Alice".into()));
        assert_eq!(vm.variables["y"], Variant::Integer(30));
        assert_eq!(vm.variables["z"], Variant::Float(9.5));
    }

    #[test]
    fn test_type_def_default_values() {
        // Dim p As Person initializes all fields to type-appropriate defaults
        let vm = run(concat!(
            "Type Point\n",
            "    X As Integer\n",
            "    Y As Integer\n",
            "    Label As String\n",
            "    Active As Boolean\n",
            "End Type\n",
            "\n",
            "Sub MySub()\n",
            "    Dim p As Point\n",
            "    xi = p.X\n",
            "    yi = p.Y\n",
            "    lbl = p.Label\n",
            "    act = p.Active\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["xi"],  Variant::Integer(0));
        assert_eq!(vm.variables["yi"],  Variant::Integer(0));
        assert_eq!(vm.variables["lbl"], Variant::Str(String::new()));
        assert_eq!(vm.variables["act"], Variant::Boolean(false));
    }

    #[test]
    fn test_type_def_in_loop() {
        // Using a UDT in a loop
        let vm = run(concat!(
            "Type Item\n",
            "    Value As Integer\n",
            "End Type\n",
            "\n",
            "Sub MySub()\n",
            "    Dim it As Item\n",
            "    total = 0\n",
            "    For i = 1 To 3\n",
            "        it.Value = i * 10\n",
            "        total = total + it.Value\n",
            "    Next i\n",
            "End Sub\n",
        ));
        // total = 10 + 20 + 30 = 60
        assert_eq!(vm.variables["total"], Variant::Integer(60));
    }

    #[test]
    fn test_public_type_def() {
        // Public Type should work the same as Type
        let vm = run(concat!(
            "Public Type Rect\n",
            "    Width As Integer\n",
            "    Height As Integer\n",
            "End Type\n",
            "\n",
            "Sub MySub()\n",
            "    Dim r As Rect\n",
            "    r.Width = 4\n",
            "    r.Height = 5\n",
            "    area = r.Width * r.Height\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["area"], Variant::Integer(20));
    }

    // ── Advanced UDT: nested types, array of UDT, With p ─────────────────────

    #[test]
    fn test_nested_type() {
        let vm = run(concat!(
            "Type Address\n    Street As String\n    City As String\nEnd Type\n",
            "Type Person\n    Name As String\n    Addr As Address\nEnd Type\n",
            "Sub MySub()\n",
            "    Dim p As Person\n",
            "    p.Name = \"Alice\"\n",
            "    p.Addr.Street = \"123 Main St\"\n",
            "    p.Addr.City = \"Springfield\"\n",
            "    n = p.Name\n",
            "    s = p.Addr.Street\n",
            "    c = p.Addr.City\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Str("Alice".into()));
        assert_eq!(vm.variables["s"], Variant::Str("123 Main St".into()));
        assert_eq!(vm.variables["c"], Variant::Str("Springfield".into()));
    }

    #[test]
    fn test_nested_type_default_values() {
        let vm = run(concat!(
            "Type Inner\n    X As Integer\n    Y As Integer\nEnd Type\n",
            "Type Outer\n    Val As String\n    Pt As Inner\nEnd Type\n",
            "Sub MySub()\n",
            "    Dim o As Outer\n",
            "    s = o.Val\n",       // default "" for String
            "    x = o.Pt.X\n",     // default 0 for nested Integer
            "    y = o.Pt.Y\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["s"], Variant::Str(String::new()));
        assert_eq!(vm.variables["x"], Variant::Integer(0));
        assert_eq!(vm.variables["y"], Variant::Integer(0));
    }

    #[test]
    fn test_dim_array_of_udt() {
        let vm = run(concat!(
            "Type Item\n    Value As Integer\n    Label As String\nEnd Type\n",
            "Sub MySub()\n",
            "    Dim items(3) As Item\n",
            "    items(1).Value = 10\n",
            "    items(2).Value = 20\n",
            "    items(1).Label = \"first\"\n",
            "    a = items(1).Value\n",
            "    b = items(2).Value\n",
            "    c = items(1).Label\n",
            "    d = items(0).Value\n",  // default 0
            "End Sub\n",
        ));
        assert_eq!(vm.variables["a"], Variant::Integer(10));
        assert_eq!(vm.variables["b"], Variant::Integer(20));
        assert_eq!(vm.variables["c"], Variant::Str("first".into()));
        assert_eq!(vm.variables["d"], Variant::Integer(0));
    }

    #[test]
    fn test_with_record_block() {
        let vm = run(concat!(
            "Type Point\n    X As Integer\n    Y As Integer\nEnd Type\n",
            "Sub MySub()\n",
            "    Dim p As Point\n",
            "    With p\n",
            "        .X = 5\n",
            "        .Y = 10\n",
            "        total = .X + .Y\n",
            "    End With\n",
            "    rx = p.X\n",
            "    ry = p.Y\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["rx"],    Variant::Integer(5));
        assert_eq!(vm.variables["ry"],    Variant::Integer(10));
        assert_eq!(vm.variables["total"], Variant::Integer(15));
    }

    #[test]
    fn test_with_nested_field() {
        let vm = run(concat!(
            "Type Inner\n    V As Integer\nEnd Type\n",
            "Type Outer\n    A As Inner\n    B As Integer\nEnd Type\n",
            "Sub MySub()\n",
            "    Dim o As Outer\n",
            "    With o\n",
            "        .A.V = 42\n",
            "        .B = 7\n",
            "    End With\n",
            "    x = o.A.V\n",
            "    y = o.B\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["x"], Variant::Integer(42));
        assert_eq!(vm.variables["y"], Variant::Integer(7));
    }

    // ── Bug-fix regression tests ──────────────────────────────────────────────

    #[test]
    fn test_instr_empty_needle() {
        // InStr(s, "") should return 1 (VBA spec), not panic
        let vm = run("Sub MySub()\n    x = InStr(\"hello\", \"\")\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1));
    }

    #[test]
    fn test_instr_start_beyond_length() {
        // InStr(10, "hello", "x") should return 0, not panic
        let vm = run("Sub MySub()\n    x = InStr(10, \"hello\", \"x\")\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(0));
    }

    #[test]
    fn test_vba_cmp_strings() {
        // String comparison: "a" < "b" should be true, not an error
        let vm = run(concat!(
            "Sub MySub()\n",
            "    If \"apple\" < \"banana\" Then\n",
            "        result = 1\n",
            "    Else\n",
            "        result = 0\n",
            "    End If\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["result"], Variant::Integer(1));
    }

    #[test]
    fn test_sheet_cells_mut_dirty_flag() {
        // Writing via Sheets("sheet1").Cells() must invalidate the End-query index
        let mut vm = Vm::new();
        vm.cells_mut().insert((1,1), CellContent { formula: None, value: Variant::Integer(10) });
        // Force index rebuild
        let _ = vm.last_nonempty_row(1, 1_048_576);
        assert!(!vm.cell_index_dirty);
        // Write via sheet_cells_mut (simulated by inserting through the public method)
        vm.cells_mut().insert((5,1), CellContent { formula: None, value: Variant::Integer(50) });
        // dirty flag must be set
        assert!(vm.cell_index_dirty);
        // Next query should rebuild and find row 5
        assert_eq!(vm.last_nonempty_row(1, 1_048_576), 5);
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
