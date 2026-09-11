use std::cell::RefCell;
use std::cmp::Ordering;
use std::collections::HashMap;

use regex::Regex;

use super::ast::{BinOpKind, FormulaExpr};
use crate::types::{MAX_ARRAY_ELEMENTS, days_in_month, is_leap, serial_to_ymd};
use crate::vm::{CellContent, ExcelError, Variant};

// ── LET/LAMBDA name-binding stack ────────────────────────────────────────────
// A stack of binding frames; each frame is pushed by LET or a lambda call.

thread_local! {
    static BINDINGS: RefCell<Vec<HashMap<String, BindingValue>>> = const { RefCell::new(vec![]) };
    static SHEET_CONTEXT: RefCell<(usize, usize)> = const { RefCell::new((1, 1)) };
}

#[derive(Clone)]
enum BindingValue {
    Value(Variant),
    Omitted,
}

pub(crate) fn with_sheet_context<T>(
    sheet_number: usize,
    sheet_count: usize,
    f: impl FnOnce() -> T,
) -> T {
    SHEET_CONTEXT.with(|context| {
        let previous = *context.borrow();
        *context.borrow_mut() = (sheet_number.max(1), sheet_count.max(1));
        let result = f();
        *context.borrow_mut() = previous;
        result
    })
}

fn sheet_context() -> (usize, usize) {
    SHEET_CONTEXT.with(|context| *context.borrow())
}

fn push_bindings(frame: HashMap<String, BindingValue>) {
    BINDINGS.with(|b| b.borrow_mut().push(frame));
}

fn pop_bindings() {
    BINDINGS.with(|b| {
        b.borrow_mut().pop();
    });
}

fn lookup_binding(name: &str) -> Option<Variant> {
    BINDINGS.with(|b| {
        let stack = b.borrow();
        for frame in stack.iter().rev() {
            if let Some(value) = frame.get(name) {
                return Some(match value {
                    BindingValue::Value(value) => value.clone(),
                    BindingValue::Omitted => Variant::Empty,
                });
            }
        }
        None
    })
}

fn binding_is_omitted(name: &str) -> bool {
    BINDINGS.with(|b| {
        let stack = b.borrow();
        for frame in stack.iter().rev() {
            if let Some(value) = frame.get(name) {
                return matches!(value, BindingValue::Omitted);
            }
        }
        false
    })
}

/// Does `expr` contain a sheet-qualified reference (`Sheet2!A1`) anywhere in
/// its tree? `cells` is intentionally a single-sheet map. The fast evaluator
/// rejects a qualified reference before any range helper can accidentally read
/// the host sheet. Workbook callers use `formula::workbook`, which builds an
/// explicit dependency order and remaps each referenced sheet before calling
/// this evaluator.
pub(crate) fn references_another_sheet(expr: &FormulaExpr) -> bool {
    match expr {
        FormulaExpr::Number(_)
        | FormulaExpr::Str(_)
        | FormulaExpr::Bool(_)
        | FormulaExpr::Omitted => false,
        FormulaExpr::CellRef { sheet, .. } => sheet.is_some(),
        FormulaExpr::Range { sheet, .. } => sheet.is_some(),
        FormulaExpr::BinOp { lhs, rhs, .. } => {
            references_another_sheet(lhs) || references_another_sheet(rhs)
        }
        FormulaExpr::UnaryMinus(inner) => references_another_sheet(inner),
        FormulaExpr::FuncCall { args, .. } => args.iter().any(references_another_sheet),
        FormulaExpr::Call { callee, args } => {
            references_another_sheet(callee) || args.iter().any(references_another_sheet)
        }
    }
}

pub fn evaluate(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if references_another_sheet(expr) {
        return Err("cross-sheet formula evaluation requires a workbook context".into());
    }
    match expr {
        FormulaExpr::Number(n) => Ok(as_integer_if_whole(*n)),
        FormulaExpr::Str(s) => Ok(Variant::Str(s.clone())),
        FormulaExpr::Bool(b) => Ok(Variant::Boolean(*b)),
        FormulaExpr::Omitted => Ok(Variant::Empty),
        FormulaExpr::CellRef { col, row, .. } => Ok(cells
            .get(&(*row, *col))
            .map(|c| c.value.clone())
            .unwrap_or(Variant::Empty)),
        FormulaExpr::Range { .. } => Err("Range cannot be used as a scalar value".into()),
        FormulaExpr::UnaryMinus(inner) => match evaluate(inner, cells)? {
            Variant::Integer(n) => Ok(Variant::Integer(-n)),
            Variant::Float(f) => Ok(Variant::Float(-f)),
            Variant::Array(values) => Ok(Variant::Array(
                values
                    .into_iter()
                    .map(|value| match value {
                        Variant::Error(error) => Variant::Error(error),
                        value => to_float(&value)
                            .map(|number| as_integer_if_whole(-number))
                            .unwrap_or(Variant::Error(ExcelError::Value)),
                    })
                    .collect(),
            )),
            Variant::VbaArray(array) => Ok(Variant::Array(
                array
                    .elements
                    .into_iter()
                    .map(|value| match value {
                        Variant::Error(error) => Variant::Error(error),
                        value => to_float(&value)
                            .map(|number| as_integer_if_whole(-number))
                            .unwrap_or(Variant::Error(ExcelError::Value)),
                    })
                    .collect(),
            )),
            other => Err(format!("Unary minus on non-numeric value: {}", other)),
        },
        FormulaExpr::BinOp { op, lhs, rhs } => eval_binop(op, lhs, rhs, cells),
        FormulaExpr::FuncCall { name, args } => {
            // 0-arg "call" is a name reference (LET / LAMBDA parameter)
            if args.is_empty()
                && let Some(v) = lookup_binding(name)
            {
                return Ok(v);
            }
            eval_func(name, args, cells)
        }
        FormulaExpr::Call { callee, args } => evaluate_lambda_call(callee, args, cells),
    }
}

// ── Binary operators ──────────────────────────────────────────────────────────

fn eval_binop(
    op: &BinOpKind,
    lhs: &FormulaExpr,
    rhs: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let l = evaluate(lhs, cells)?;
    let r = evaluate(rhs, cells)?;
    // Propagate Excel error values through arithmetic
    if let Variant::Error(_) = &l {
        return Ok(l);
    }
    if let Variant::Error(_) = &r {
        return Ok(r);
    }
    let l_array = match &l {
        Variant::Array(values) => Some(values.clone()),
        Variant::VbaArray(array) => Some(array.elements.clone()),
        _ => None,
    };
    let r_array = match &r {
        Variant::Array(values) => Some(values.clone()),
        Variant::VbaArray(array) => Some(array.elements.clone()),
        _ => None,
    };
    if l_array.is_some() || r_array.is_some() {
        let l_is_array = l_array.is_some();
        let r_is_array = r_array.is_some();
        let l_values = l_array.unwrap_or_else(|| vec![l.clone()]);
        let r_values = r_array.unwrap_or_else(|| vec![r.clone()]);
        let l_shape = if l_is_array {
            array_shape_for_expr(lhs, cells, l_values.len())
        } else {
            (1, 1)
        };
        let r_shape = if r_is_array {
            array_shape_for_expr(rhs, cells, r_values.len())
        } else {
            (1, 1)
        };
        if l_shape.0.saturating_mul(l_shape.1) != l_values.len()
            || r_shape.0.saturating_mul(r_shape.1) != r_values.len()
        {
            return Ok(Variant::Error(ExcelError::Value));
        }
        let rows = broadcast_dimension(l_shape.0, r_shape.0);
        let cols = broadcast_dimension(l_shape.1, r_shape.1);
        let (Some(rows), Some(cols)) = (rows, cols) else {
            return Ok(Variant::Error(ExcelError::Value));
        };
        let mut result = Vec::with_capacity(rows.saturating_mul(cols));
        for row in 0..rows {
            for col in 0..cols {
                let left_row = if l_shape.0 == 1 { 0 } else { row };
                let left_col = if l_shape.1 == 1 { 0 } else { col };
                let right_row = if r_shape.0 == 1 { 0 } else { row };
                let right_col = if r_shape.1 == 1 { 0 } else { col };
                let left = &l_values[left_row * l_shape.1 + left_col];
                let right = &r_values[right_row * r_shape.1 + right_col];
                result.push(eval_scalar_binop(op, left, right)?);
            }
        }
        return Ok(Variant::Array(result));
    }
    eval_scalar_binop(op, &l, &r)
}

fn broadcast_dimension(left: usize, right: usize) -> Option<usize> {
    if left == right {
        Some(left)
    } else if left == 1 {
        Some(right)
    } else if right == 1 {
        Some(left)
    } else {
        None
    }
}

fn eval_scalar_binop(op: &BinOpKind, l: &Variant, r: &Variant) -> Result<Variant, String> {
    if let Variant::Error(error) = l {
        return Ok(Variant::Error(error.clone()));
    }
    if let Variant::Error(error) = r {
        return Ok(Variant::Error(error.clone()));
    }
    match op {
        BinOpKind::Add | BinOpKind::Sub | BinOpKind::Mul | BinOpKind::Div => {
            let lf = to_float(l)?;
            let rf = to_float(r)?;
            let result = match op {
                BinOpKind::Add => lf + rf,
                BinOpKind::Sub => lf - rf,
                BinOpKind::Mul => lf * rf,
                BinOpKind::Div => {
                    if rf == 0.0 {
                        return Ok(Variant::Error(ExcelError::DivZero));
                    }
                    lf / rf
                }
                _ => unreachable!(),
            };
            Ok(as_integer_if_whole(result))
        }
        BinOpKind::Concat => Ok(Variant::Str(format!("{}{}", l, r))),
        BinOpKind::Eq => Ok(Variant::Boolean(variant_eq(l, r))),
        BinOpKind::Ne => Ok(Variant::Boolean(!variant_eq(l, r))),
        BinOpKind::Lt => Ok(Variant::Boolean(variant_cmp(l, r)? == Ordering::Less)),
        BinOpKind::Le => Ok(Variant::Boolean(variant_cmp(l, r)? != Ordering::Greater)),
        BinOpKind::Gt => Ok(Variant::Boolean(variant_cmp(l, r)? == Ordering::Greater)),
        BinOpKind::Ge => Ok(Variant::Boolean(variant_cmp(l, r)? != Ordering::Less)),
    }
}

// ── Type helpers ──────────────────────────────────────────────────────────────

fn to_float(v: &Variant) -> Result<f64, String> {
    match v {
        Variant::Integer(n) => Ok(*n as f64),
        Variant::Float(f) => Ok(*f),
        Variant::Boolean(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Variant::Date(s) => Ok(*s as f64),
        Variant::Error(e) => Err(e.to_string()),
        Variant::Empty => Ok(0.0),
        // VBA Null never reaches the worksheet-formula engine (a Null is
        // never stored in a cell -- see the write path, which treats it as
        // blank). Grouped with Empty for the same reason a blank cell is 0.
        Variant::Null => Ok(0.0),
        Variant::Str(s) => s
            .parse::<f64>()
            .map_err(|_| format!("Cannot convert '{}' to a number", s)),
        Variant::Array(_) | Variant::VbaArray(_) => Err("Cannot convert array to number".into()),
        Variant::Record(_) => Err("Cannot convert record to number".into()),
    }
}

fn to_str(v: &Variant) -> String {
    match v {
        Variant::Str(s) => s.clone(),
        Variant::Integer(n) => n.to_string(),
        Variant::Float(f) => f.to_string(),
        Variant::Boolean(b) => {
            if *b {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
        Variant::Date(s) => serial_to_display(*s),
        Variant::Error(e) => e.as_str().to_string(),
        Variant::Empty | Variant::Null => String::new(),
        Variant::Array(a) => a.iter().map(to_str).collect::<Vec<_>>().join(", "),
        Variant::VbaArray(a) => a.elements.iter().map(to_str).collect::<Vec<_>>().join(", "),
        Variant::Record(_) => "[Record]".into(),
    }
}

fn is_truthy(v: &Variant) -> bool {
    match v {
        Variant::Boolean(b) => *b,
        Variant::Integer(n) => *n != 0,
        Variant::Float(f) => *f != 0.0,
        Variant::Str(s) => !s.is_empty(),
        Variant::Date(_) => true,
        Variant::Error(_) => false,
        Variant::Empty | Variant::Null => false,
        Variant::Array(a) => !a.is_empty(),
        Variant::VbaArray(a) => !a.elements.is_empty(),
        Variant::Record(_) => true,
    }
}

fn variant_eq(a: &Variant, b: &Variant) -> bool {
    match (a, b) {
        (Variant::Integer(x), Variant::Integer(y)) => x == y,
        (Variant::Float(x), Variant::Float(y)) => x == y,
        (Variant::Integer(x), Variant::Float(y)) => (*x as f64) == *y,
        (Variant::Float(x), Variant::Integer(y)) => *x == (*y as f64),
        (Variant::Date(x), Variant::Date(y)) => x == y,
        (Variant::Date(x), Variant::Integer(y)) => x == y,
        (Variant::Integer(x), Variant::Date(y)) => x == y,
        (Variant::Str(x), Variant::Str(y)) => x.eq_ignore_ascii_case(y),
        (Variant::Boolean(x), Variant::Boolean(y)) => x == y,
        (Variant::Empty, Variant::Empty) => true,
        (Variant::Error(_), _) | (_, Variant::Error(_)) => false,
        _ => false,
    }
}

fn variant_cmp(a: &Variant, b: &Variant) -> Result<Ordering, String> {
    let af = to_float(a)?;
    let bf = to_float(b)?;
    af.partial_cmp(&bf)
        .ok_or_else(|| "Cannot compare NaN values".into())
}

fn as_integer_if_whole(f: f64) -> Variant {
    if f.fract() == 0.0 && f >= i64::MIN as f64 && f <= i64::MAX as f64 {
        Variant::Integer(f as i64)
    } else {
        Variant::Float(f)
    }
}

/// Extract a numeric value from a Variant, ignoring non-numeric types.
fn as_f64(v: &Variant) -> Option<f64> {
    match v {
        Variant::Integer(n) => Some(*n as f64),
        Variant::Float(f) => Some(*f),
        _ => None,
    }
}

/// Expand a single formula argument into a flat list of Variant values.
/// Range → all cells row-major. Scalar → single value.
fn collect_values(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Vec<Variant>, String> {
    match expr {
        FormulaExpr::Range { c1, r1, c2, r2, .. } => {
            let (rmin, rmax) = (r1.min(r2), r1.max(r2));
            let (cmin, cmax) = (c1.min(c2), c1.max(c2));
            let rows = (rmax - rmin + 1) as u64;
            let cols = (cmax - cmin + 1) as u64;
            if rows * cols > 1_000_000 {
                return Err(format!(
                    "Range too large ({} cells); maximum is 1,000,000",
                    rows * cols
                ));
            }
            let mut vals = Vec::with_capacity((rows * cols) as usize);
            for row in *rmin..=*rmax {
                for col in *cmin..=*cmax {
                    vals.push(cell_val(cells, row, col));
                }
            }
            Ok(vals)
        }
        FormulaExpr::FuncCall { name, args }
            if name.eq_ignore_ascii_case("OFFSET") && args.len() >= 3 =>
        {
            Ok(offset_values(args, cells)?)
        }
        other => match evaluate(other, cells)? {
            Variant::Array(values) => Ok(values),
            Variant::VbaArray(array) => Ok(array.elements),
            value => Ok(vec![value]),
        },
    }
}

/// Resolve an OFFSET reference to its bounded row-major cell values.
///
/// OFFSET is a reference-producing function.  Keeping the reference expansion
/// here (rather than returning only its top-left cell) lets aggregators and
/// dynamic-array consumers share the same height/width semantics.
fn offset_values(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Vec<Variant>, String> {
    if args.len() < 3 || args.len() > 5 {
        return Err("OFFSET requires 3 to 5 arguments".into());
    }
    let (base_row, base_col, base_height, base_width) = match &args[0] {
        FormulaExpr::CellRef { row, col, .. } => (*row, *col, 1_i64, 1_i64),
        FormulaExpr::Range { r1, c1, r2, c2, .. } => (
            *r1,
            *c1,
            (i64::from(r1.max(r2) - r1.min(r2)) + 1),
            (i64::from(c1.max(c2) - c1.min(c2)) + 1),
        ),
        _ => return Err("OFFSET: first argument must be a cell or range reference".into()),
    };
    let integer_argument = |arg: &FormulaExpr, label: &str| -> Result<i64, String> {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || value < i64::MIN as f64 {
            return Err(format!("OFFSET: {label} must be an integer"));
        }
        if value > i64::MAX as f64 {
            return Err(format!("OFFSET: {label} is out of range"));
        }
        Ok(value as i64)
    };
    let row_offset = integer_argument(&args[1], "row offset")?;
    let col_offset = integer_argument(&args[2], "column offset")?;
    let height = args
        .get(3)
        .map(|arg| integer_argument(arg, "height"))
        .transpose()?
        .unwrap_or(base_height);
    let width = args
        .get(4)
        .map(|arg| integer_argument(arg, "width"))
        .transpose()?
        .unwrap_or(base_width);
    if height <= 0 || width <= 0 {
        return Err("OFFSET: height and width must be positive".into());
    }
    let row = i64::from(base_row)
        .checked_add(row_offset)
        .ok_or_else(|| "OFFSET: row overflow".to_string())?;
    let col = i64::from(base_col)
        .checked_add(col_offset)
        .ok_or_else(|| "OFFSET: column overflow".to_string())?;
    if row < 1 || col < 1 {
        return Err("OFFSET: reference is outside the worksheet".into());
    }
    let total = (height as u64)
        .checked_mul(width as u64)
        .ok_or_else(|| "OFFSET: range is too large".to_string())?;
    if total > 1_000_000 {
        return Err("OFFSET: range too large (maximum is 1,000,000 cells)".into());
    }
    let mut values = Vec::with_capacity(total as usize);
    for row_delta in 0..height {
        for col_delta in 0..width {
            let target_row = row
                .checked_add(row_delta)
                .ok_or_else(|| "OFFSET: row overflow".to_string())?;
            let target_col = col
                .checked_add(col_delta)
                .ok_or_else(|| "OFFSET: column overflow".to_string())?;
            if target_row > u32::MAX as i64 || target_col > u32::MAX as i64 {
                return Err("OFFSET: reference is outside the worksheet".into());
            }
            values.push(cell_val(cells, target_row as u32, target_col as u32));
        }
    }
    Ok(values)
}

fn collect_all(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Vec<Variant>, String> {
    let mut out = vec![];
    for a in args {
        out.extend(collect_values(a, cells)?);
    }
    Ok(out)
}

/// Flatten the engine's bounded one-dimensional array representation for
/// functions whose worksheet contract consumes a value range.
fn flatten_values(values: Vec<Variant>) -> Vec<Variant> {
    values.into_iter().flat_map(variant_values).collect()
}

/// Collect numeric values for scalar statistical aggregators while retaining
/// Excel worksheet errors.  Range text and logical values remain excluded;
/// array-produced values are flattened one level, matching the existing
/// bounded-array representation.
fn collect_numeric_values_with_errors(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Vec<f64>, ExcelError> {
    let values = collect_all(args, cells).map_err(|_| ExcelError::Value)?;
    let mut numbers = Vec::new();
    for value in values {
        match value {
            Variant::Error(error) => return Err(error),
            Variant::Array(values) => {
                for value in values {
                    if let Variant::Error(error) = value {
                        return Err(error);
                    }
                    if let Some(number) = as_f64(&value) {
                        numbers.push(number);
                    }
                }
            }
            Variant::VbaArray(array) => {
                for value in array.elements {
                    if let Variant::Error(error) = value {
                        return Err(error);
                    }
                    if let Some(number) = as_f64(&value) {
                        numbers.push(number);
                    }
                }
            }
            value => {
                if let Some(number) = as_f64(&value) {
                    numbers.push(number);
                }
            }
        }
    }
    Ok(numbers)
}

fn expression_shape(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<(u64, u64), String> {
    match expr {
        FormulaExpr::Range { c1, r1, c2, r2, .. } => Ok((
            (r1.max(r2) - r1.min(r2) + 1) as u64,
            (c1.max(c2) - c1.min(c2) + 1) as u64,
        )),
        _ => Ok((collect_values(expr, cells)?.len() as u64, 1)),
    }
}

fn ensure_same_shape(
    expected: (u64, u64),
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
    function_name: &str,
) -> Result<(), String> {
    if expression_shape(expr, cells)? != expected {
        return Err(format!("{function_name}: ranges must have the same shape"));
    }
    Ok(())
}

fn cell_val(cells: &HashMap<(u32, u32), CellContent>, row: u32, col: u32) -> Variant {
    cells
        .get(&(row, col))
        .map(|c| c.value.clone())
        .unwrap_or(Variant::Empty)
}

/// Borrow a cell value without cloning. Returns `None` for missing/empty cells.
fn cell_ref(cells: &HashMap<(u32, u32), CellContent>, row: u32, col: u32) -> Option<&Variant> {
    cells.get(&(row, col)).map(|c| &c.value)
}

// ── Function dispatch ─────────────────────────────────────────────────────────

fn eval_func(
    name: &str,
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    match name {
        "SUM" => func_sum(args, cells),
        "AVERAGE" => func_average(args, cells),
        "AVERAGEA" => func_averagea(args, cells),
        "MIN" => func_min(args, cells),
        "MINA" => func_mina(args, cells),
        "MAX" => func_max(args, cells),
        "MAXA" => func_maxa(args, cells),
        "COUNT" => func_count(args, cells),
        "COUNTA" => func_counta(args, cells),
        "IF" => func_if(args, cells),
        "AND" => func_and(args, cells),
        "OR" => func_or(args, cells),
        "NOT" => func_not(args, cells),
        "TRUE" => func_boolean_constant(args, true),
        "FALSE" => func_boolean_constant(args, false),
        "IFERROR" => func_iferror(args, cells),
        "IFNA" => func_ifna(args, cells),
        "LEFT" => func_left(args, cells),
        "RIGHT" => func_right(args, cells),
        "MID" => func_mid(args, cells),
        "LEN" => func_len(args, cells),
        "LEFTB" => func_leftb(args, cells),
        "RIGHTB" => func_rightb(args, cells),
        "MIDB" => func_midb(args, cells),
        "LENB" => func_lenb(args, cells),
        "ROUND" => func_round(args, cells),
        "ROUNDUP" => func_roundup(args, cells),
        "ROUNDDOWN" => func_rounddown(args, cells),
        "CONCATENATE" => func_concatenate(args, cells),
        "CONCAT" => func_concatenate(args, cells),
        "TEXT" => func_text(args, cells),
        "COUNTIF" => func_countif(args, cells),
        "SUMIF" => func_sumif(args, cells),
        "SUMIFS" => func_sumifs(args, cells),
        "COUNTIFS" => func_countifs(args, cells),
        "MEDIAN" => func_median(args, cells),
        "MODE.MULT" => func_mode_mult(args, cells, true),
        "MODE" | "MODE.SNGL" => func_mode_mult(args, cells, false),
        "FREQUENCY" => func_frequency(args, cells),
        "PROB" => func_prob(args, cells),
        "PRODUCT" => func_product(args, cells),
        "ROW" => func_row(args, cells),
        "ROWS" => func_rows(args, cells),
        "DATE" => func_date(args, cells),
        "TODAY" => func_today(args, cells),
        "NETWORKDAYS" => func_networkdays(args, cells),
        "WORKDAY" => func_workday(args, cells),
        "RANK" | "RANK.EQ" => func_rank(args, cells),
        "RANK.AVG" => func_rank_avg(args, cells),
        "IFS" => func_ifs(args, cells),
        "XLOOKUP" => func_xlookup(args, cells),
        "EOMONTH" => func_eomonth(args, cells),
        "SUBTOTAL" => func_subtotal(args, cells),
        "AGGREGATE" => func_aggregate(args, cells),
        // -- Numerical --
        "AVERAGEIF" => func_averageif(args, cells),
        "AVERAGEIFS" => func_averageifs(args, cells),
        "INT" => func_int(args, cells),
        "ISEVEN" => func_parity(args, cells, true),
        "ISODD" => func_parity(args, cells, false),
        "LARGE" => func_large(args, cells),
        "MAXIFS" => func_maxifs(args, cells),
        "MINIFS" => func_minifs(args, cells),
        "MOD" => func_mod(args, cells),
        "PERCENTILE" | "PERCENTILE.INC" => func_percentile(args, cells),
        "PERCENTILE.EXC" => func_percentile_exc(args, cells),
        "PERCENTOF" => func_percentof(args, cells),
        "PERCENTRANK" | "PERCENTRANK.INC" => func_percentrank(args, cells),
        "PERCENTRANK.EXC" => func_percentrank_exc(args, cells),
        "QUARTILE" | "QUARTILE.INC" => func_quartile(args, cells, false),
        "QUARTILE.EXC" => func_quartile(args, cells, true),
        "RAND" => func_rand(args, cells),
        "RANDBETWEEN" => func_randbetween(args, cells),
        "SMALL" => func_small(args, cells),
        "SUMPRODUCT" => func_sumproduct(args, cells),
        "SUMSQ" => func_sumsq(args, cells),
        "GEOMEAN" => func_geomean(args, cells),
        "HARMEAN" => func_harmean(args, cells),
        "DEVSQ" => func_devsq(args, cells),
        "AVEDEV" => func_avedev(args, cells),
        "TRIMMEAN" => func_trimmean(args, cells),
        "SKEW" => func_skew(args, cells),
        "SKEW.P" => func_skew_p(args, cells),
        "KURT" => func_kurt(args, cells),
        "PHI" => func_phi(args, cells),
        "GAUSS" => func_gauss(args, cells),
        "TRUNC" => func_trunc(args, cells),
        // -- String --
        "ASC" => func_asc(args, cells),
        "CHAR" => func_char(args, cells),
        "CODE" => func_code(args, cells),
        "EXACT" => func_exact(args, cells),
        "FIND" => func_find(args, cells),
        "JIS" | "DBCS" => func_jis(args, cells),
        "LOWER" => func_lower(args, cells),
        "PHONETIC" => func_phonetic(args, cells),
        "PROPER" => func_proper(args, cells),
        "REPLACE" => func_replace(args, cells),
        "REPLACEB" => func_replaceb(args, cells),
        "SEARCH" => func_search(args, cells),
        "SEARCHB" => func_searchb(args, cells),
        "FINDB" => func_findb(args, cells),
        "SUBSTITUTE" => func_substitute(args, cells),
        "TEXTJOIN" => func_textjoin(args, cells),
        "TEXTSPLIT" => func_textsplit(args, cells),
        "TEXTBEFORE" => func_textbefore(args, cells),
        "TEXTAFTER" => func_textafter(args, cells),
        "ENCODEURL" => func_encodeurl(args, cells),
        "VALUETOTEXT" => func_valuetotext(args, cells),
        "ARRAYTOTEXT" => func_arraytotext(args, cells),
        "HYPERLINK" => func_hyperlink(args, cells),
        "TRIM" => func_trim(args, cells),
        "CLEAN" => func_clean(args, cells),
        "T" => func_t(args, cells),
        "FIXED" => func_fixed(args, cells),
        "DOLLAR" => func_dollar(args, cells),
        "BAHTTEXT" => func_bahttext(args, cells),
        "BASE" => func_base(args, cells),
        "DECIMAL" => func_decimal(args, cells),
        "UNICHAR" => func_unichar(args, cells),
        "UNICODE" => func_unicode(args, cells),
        "UPPER" => func_upper(args, cells),
        "VALUE" => func_value(args, cells),
        "REPT" => func_rept(args, cells),
        "NUMBERVALUE" => func_numbervalue(args, cells),
        "REGEXTEST" => func_regextest(args, cells),
        "REGEXEXTRACT" => func_regexextract(args, cells),
        "REGEXREPLACE" => func_regexreplace(args, cells),
        "DETECTLANGUAGE" => func_detectlanguage(args, cells),
        // -- Date/Time --
        "YEAR" => func_year(args, cells),
        "MONTH" => func_month(args, cells),
        "DAY" => func_day(args, cells),
        "WEEKDAY" => func_weekday(args, cells),
        "WEEKNUM" => func_weeknum(args, cells),
        "ISOWEEKNUM" => func_isoweeknum(args, cells),
        "DAYS" => func_days(args, cells),
        "DAYS360" => func_days360(args, cells),
        "YEARFRAC" => func_yearfrac(args, cells),
        "EDATE" => func_edate(args, cells),
        "DATEDIF" => func_datedif(args, cells),
        "DATEVALUE" => func_datevalue(args, cells),
        "NOW" => func_now(args, cells),
        "TIME" => func_time_fn(args, cells),
        "TIMEVALUE" => func_timevalue(args, cells),
        "HOUR" => func_hour(args, cells),
        "MINUTE" => func_minute(args, cells),
        "SECOND" => func_second(args, cells),
        "NETWORKDAYS.INTL" => func_networkdays_intl(args, cells),
        "WORKDAY.INTL" => func_workday_intl(args, cells),
        // -- Logic --
        "SWITCH" => func_switch(args, cells),
        "XOR" => func_xor(args, cells),
        // -- Lookup --
        "CHOOSE" => func_choose(args, cells),
        "COLUMN" => func_column(args, cells),
        "COLUMNS" => func_columns(args, cells),
        "LOOKUP" => func_lookup(args, cells),
        "XMATCH" => func_xmatch(args, cells),
        // -- Info --
        "ISBLANK" => func_isblank(args, cells),
        "ISFORMULA" => func_isformula(args, cells),
        "ISOMITTED" => func_isomitted(args, cells),
        "SHEET" => func_sheet(args, cells),
        "SHEETS" => func_sheets(args, cells),
        "ISREF" => func_isref(args, cells),
        "ISERROR" => func_iserror(args, cells),
        "ISERR" => func_iserror(args, cells),
        "ISNA" => func_isna(args, cells),
        "ISNUMBER" => func_isnumber(args, cells),
        "ISTEXT" => func_istext(args, cells),
        "ISLOGICAL" => func_islogical(args, cells),
        "ISNONTEXT" => func_isnontext(args, cells),
        "N" => func_n(args, cells),
        "NA" => func_na(args, cells),
        "TYPE" => func_type_fn(args, cells),
        "ERROR.TYPE" => func_error_type(args, cells),
        "FORMULATEXT" => func_formulatext(args, cells),
        "CELL" => func_cell(args, cells),
        "AREAS" => func_areas(args, cells),
        "INFO" => func_info(args, cells),
        "VLOOKUP" => func_vlookup(args, cells),
        "HLOOKUP" => func_hlookup(args, cells),
        "INDEX" => func_index(args, cells),
        "MATCH" => func_match_fn(args, cells),
        "GETPIVOTDATA" => func_getpivotdata(args, cells),
        // ── Statistics ───────────────────────────────────────────────────────
        "STDEV" | "STDEV.S" => func_stdev_s(args, cells),
        "STDEVA" => func_stdev_a(args, cells),
        "STDEVP" | "STDEV.P" => func_stdev_p(args, cells),
        "STDEVPA" => func_stdev_pa(args, cells),
        "VAR" | "VAR.S" => func_var_s(args, cells),
        "VARA" => func_var_a(args, cells),
        "VARP" | "VAR.P" => func_var_p(args, cells),
        "VARPA" => func_var_pa(args, cells),
        "CORREL" | "PEARSON" => func_correl(args, cells),
        "SLOPE" => func_slope(args, cells),
        "INTERCEPT" => func_intercept(args, cells),
        "RSQ" => func_rsq(args, cells),
        "FORECAST.LINEAR" | "FORECAST" => func_forecast_linear(args, cells),
        "FORECAST.ETS" => func_forecast_ets(args, cells),
        "FORECAST.ETS.SEASONALITY" => func_forecast_ets_seasonality(args, cells),
        "FORECAST.ETS.CONFINT" => func_forecast_ets_confint(args, cells),
        "FORECAST.ETS.STAT" => func_forecast_ets_stat(args, cells),
        "TREND" => func_trend(args, cells),
        "GROWTH" => func_growth(args, cells),
        "LINEST" => func_linest(args, cells),
        "LOGEST" => func_logest(args, cells),
        "STEYX" => func_steyx(args, cells),
        "FISHER" => func_fisher(args, cells),
        "FISHERINV" => func_fisherinv(args, cells),
        "STANDARDIZE" => func_standardize(args, cells),
        "Z.TEST" | "ZTEST" => func_z_test(args, cells),
        "CONFIDENCE.NORM" | "CONFIDENCE" => func_confidence_norm(args, cells),
        "CONFIDENCE.T" => func_confidence_t(args, cells),
        "COVARIANCE.S" | "COVAR" => func_covariance_s(args, cells),
        "COVARIANCE.P" => func_covariance_p(args, cells),
        "FTEST" | "F.TEST" => func_ftest(args, cells),
        "CHITEST" | "CHISQ.TEST" => func_chitest(args, cells),
        "NORM.DIST" | "NORMDIST" => func_norm_dist(args, cells),
        "NORM.INV" | "NORMINV" => func_norm_inv(args, cells),
        "NORM.S.DIST" | "NORMSDIST" => func_norm_s_dist(args, cells),
        "NORM.S.INV" | "NORMSINV" => func_norm_s_inv(args, cells),
        "BINOM.DIST" | "BINOMDIST" => func_binom_dist(args, cells),
        "BINOM.DIST.RANGE" => func_binom_dist_range(args, cells),
        "BINOM.INV" | "CRITBINOM" => func_binom_inv(args, cells),
        "NEGBINOM.DIST" | "NEGBINOMDIST" => func_negbinom_dist(args, cells),
        "HYPGEOM.DIST" | "HYPGEOMDIST" => func_hypgeom_dist(args, cells),
        "POISSON.DIST" | "POISSON" => func_poisson_dist(args, cells),
        "GAMMA" => func_gamma(args, cells),
        "GAMMALN" | "GAMMALN.PRECISE" => func_gammaln(args, cells),
        "GAMMA.DIST" | "GAMMADIST" => func_gamma_dist(args, cells),
        "GAMMA.INV" | "GAMMAINV" => func_gamma_inv(args, cells),
        "BETA.DIST" | "BETADIST" => func_beta_dist(args, cells),
        "BETA.INV" | "BETAINV" => func_beta_inv(args, cells),
        "CHISQ.DIST" | "CHIDIST" => func_chisq_dist(args, cells),
        "CHISQ.DIST.RT" => func_chisq_dist_rt(args, cells),
        "CHISQ.INV" => func_chisq_inv(args, cells),
        "CHISQ.INV.RT" | "CHIINV" => func_chisq_inv_rt(args, cells),
        "F.DIST" | "FDIST" => func_f_dist(args, cells),
        "F.DIST.RT" => func_f_dist_rt(args, cells),
        "F.DIST.2T" => func_f_dist_2t(args, cells),
        "F.INV" => func_f_inv(args, cells),
        "F.INV.RT" | "FINV" => func_f_inv_rt(args, cells),
        "WEIBULL.DIST" | "WEIBULL" => func_weibull_dist(args, cells),
        "EXPON.DIST" | "EXPONDIST" => func_expon_dist(args, cells),
        "LOGNORM.DIST" | "LOGNORMDIST" => func_lognorm_dist(args, cells),
        "LOGNORM.INV" | "LOGINV" => func_lognorm_inv(args, cells),
        "T.DIST" => func_t_dist(args, cells),
        "T.DIST.2T" => func_t_dist_2t(args, cells),
        "T.DIST.RT" => func_t_dist_rt(args, cells),
        "T.INV" => func_t_inv(args, cells),
        "T.INV.2T" => func_t_inv_2t(args, cells),
        "TDIST" => func_tdist_legacy(args, cells),
        "TINV" => func_t_inv_2t(args, cells),
        "TTEST" | "T.TEST" => func_ttest(args, cells),
        // ── Rounding ─────────────────────────────────────────────────────────
        "FLOOR" | "FLOOR.MATH" => func_floor(args, cells),
        "CEILING" | "CEILING.MATH" => func_ceiling(args, cells),
        "FLOOR.PRECISE" | "ISO.FLOOR" => func_precise_round(args, cells, false),
        "CEILING.PRECISE" | "ISO.CEILING" | "ECMA.CEILING" => func_precise_round(args, cells, true),
        "EVEN" => func_even_odd(args, cells, true),
        "ODD" => func_even_odd(args, cells, false),
        "MROUND" => func_mround(args, cells),
        // ── Math ─────────────────────────────────────────────────────────────
        "ABS" => func_abs(args, cells),
        "SQRT" => func_sqrt(args, cells),
        "SQRTPI" => func_sqrtpi(args, cells),
        "POWER" => func_power(args, cells),
        "EXP" => func_exp(args, cells),
        "LOG" => func_log(args, cells),
        "LOG10" => func_log10(args, cells),
        "LN" => func_ln(args, cells),
        // ── Engineering ─────────────────────────────────────────────────────
        "BITAND" => func_bitwise(args, cells, "BITAND"),
        "BITOR" => func_bitwise(args, cells, "BITOR"),
        "BITXOR" => func_bitwise(args, cells, "BITXOR"),
        "BITLSHIFT" => func_bitshift(args, cells, true),
        "BITRSHIFT" => func_bitshift(args, cells, false),
        "DELTA" => func_delta(args, cells),
        "GESTEP" => func_gestep(args, cells),
        "ERF" | "ERF.PRECISE" => func_erf(args, cells),
        "ERFC" | "ERFC.PRECISE" => func_erfc(args, cells),
        "BESSELJ" => func_bessel(args, cells, false),
        "BESSELI" => func_bessel(args, cells, true),
        "BESSELY" => func_bessel_second_kind(args, cells, false),
        "BESSELK" => func_bessel_second_kind(args, cells, true),
        "DEC2BIN" => func_dec_to_radix(args, cells, 2),
        "DEC2HEX" => func_dec_to_radix(args, cells, 16),
        "DEC2OCT" => func_dec_to_radix(args, cells, 8),
        "BIN2DEC" => func_radix_to_dec(args, cells, 2),
        "HEX2DEC" => func_radix_to_dec(args, cells, 16),
        "OCT2DEC" => func_radix_to_dec(args, cells, 8),
        "BIN2HEX" => func_radix_to_radix(args, cells, 2, 16),
        "BIN2OCT" => func_radix_to_radix(args, cells, 2, 8),
        "HEX2BIN" => func_radix_to_radix(args, cells, 16, 2),
        "HEX2OCT" => func_radix_to_radix(args, cells, 16, 8),
        "OCT2BIN" => func_radix_to_radix(args, cells, 8, 2),
        "OCT2HEX" => func_radix_to_radix(args, cells, 8, 16),
        "CONVERT" => func_convert(args, cells),
        "ROMAN" => func_roman(args, cells),
        "ARABIC" => func_arabic(args, cells),
        "COMPLEX" => func_complex(args, cells),
        "IMREAL" => func_imreal(args, cells),
        "IMAGINARY" => func_imaginary(args, cells),
        "IMARGUMENT" => func_imargument(args, cells),
        "IMABS" => func_imabs(args, cells),
        "IMSUM" => func_imsum(args, cells),
        "IMSUB" => func_imsub(args, cells),
        "IMPRODUCT" => func_improduct(args, cells),
        "IMDIV" => func_imdiv(args, cells),
        "IMCONJUGATE" => func_imconjugate(args, cells),
        "IMEXP" => func_imexp(args, cells),
        "IMLN" => func_imln(args, cells),
        "IMLOG10" => func_imlog10(args, cells),
        "IMLOG2" => func_imlog2(args, cells),
        "IMSQRT" => func_imsqrt(args, cells),
        "IMPOWER" => func_impower(args, cells),
        "IMSIN" => func_imsin(args, cells),
        "IMCOS" => func_imcos(args, cells),
        "IMTAN" => func_imtan(args, cells),
        "IMSINH" => func_imsinh(args, cells),
        "IMCOSH" => func_imcosh(args, cells),
        "IMTANH" => func_imtanh(args, cells),
        "IMSEC" => func_imsec(args, cells),
        "IMCSC" => func_imcsc(args, cells),
        "IMCOT" => func_imcot(args, cells),
        "IMASIN" => func_imasin(args, cells),
        "IMACOS" => func_imacos(args, cells),
        "IMATAN" => func_imatan(args, cells),
        "IMACOT" => func_imacot(args, cells),
        "IMASINH" => func_imasinh(args, cells),
        "IMACOSH" => func_imacosh(args, cells),
        "IMATANH" => func_imatanh(args, cells),
        "IMSECH" => func_imsech(args, cells),
        "IMCSCH" => func_imcsch(args, cells),
        "IMCOTH" => func_imcoth(args, cells),
        // ── Trigonometry ──────────────────────────────────────────────────────
        "PI" => func_pi(args, cells),
        "SIN" => func_trig1(args, cells, f64::sin),
        "COS" => func_trig1(args, cells, f64::cos),
        "TAN" => func_trig1(args, cells, f64::tan),
        "ASIN" => func_trig1(args, cells, f64::asin),
        "ACOS" => func_trig1(args, cells, f64::acos),
        "ATAN" => func_trig1(args, cells, f64::atan),
        "ATAN2" => func_atan2(args, cells),
        "DEGREES" => func_trig1(args, cells, f64::to_degrees),
        "RADIANS" => func_trig1(args, cells, f64::to_radians),
        "SINH" | "COSH" | "TANH" | "ASINH" | "ACOSH" | "ATANH" | "SEC" | "SECH" | "CSC"
        | "CSCH" | "COT" | "COTH" | "ACOT" | "ACOTH" => func_extended_trig(args, cells, name),
        // ── Info ─────────────────────────────────────────────────────────────
        "COUNTBLANK" => func_countblank(args, cells),
        "ADDRESS" => func_address(args, cells),
        "INDIRECT" => func_indirect(args, cells),
        "OFFSET" => func_offset(args, cells),
        // ── Array / spill functions ───────────────────────────────────────────
        "FILTER" => func_filter(args, cells),
        "FILTERXML" => func_filterxml(args, cells),
        "GROUPBY" => func_groupby(args, cells),
        "PIVOTBY" => func_pivotby(args, cells),
        "UNIQUE" => func_unique(args, cells),
        "SORT" => func_sort(args, cells),
        "SORTBY" => func_sortby(args, cells),
        "SEQUENCE" => func_sequence(args, cells),
        "TRANSPOSE" => func_transpose(args, cells),
        "MUNIT" => func_munit(args, cells),
        "MMULT" => func_mmult(args, cells),
        "MDETERM" => func_mdeterm(args, cells),
        "MINVERSE" => func_minverse(args, cells),
        "TOCOL" => func_tocol(args, cells),
        "TOROW" => func_torow(args, cells),
        "WRAPCOLS" => func_wrapcols(args, cells),
        "WRAPROWS" => func_wraprows(args, cells),
        "RANDARRAY" => func_randarray(args, cells),
        "TAKE" => func_take(args, cells),
        "DROP" => func_drop(args, cells),
        "VSTACK" => func_vstack(args, cells),
        "HSTACK" => func_hstack(args, cells),
        "CHOOSECOLS" => func_choosecols(args, cells),
        "CHOOSEROWS" => func_chooserows(args, cells),
        "EXPAND" => func_expand(args, cells),
        "TRIMRANGE" => func_trimrange(args, cells),
        "SINGLE" => func_single(args, cells),
        "MAKEARRAY" => func_makearray(args, cells),
        // ── Math / Financial ─────────────────────────────────────────────────
        "COMBIN" => func_combin(args, cells),
        "COMBINA" => func_combina(args, cells),
        "FACT" => func_fact(args, cells),
        "FACTDOUBLE" => func_factdouble(args, cells),
        "PERMUT" => func_permut(args, cells),
        "PERMUTATIONA" => func_permutationa(args, cells),
        "MULTINOMIAL" => func_multinomial(args, cells),
        "GCD" => func_gcd(args, cells),
        "LCM" => func_lcm(args, cells),
        "QUOTIENT" => func_quotient(args, cells),
        "SIGN" => func_sign(args, cells),
        "SUMX2MY2" => func_sumx2my2(args, cells),
        "SUMX2PY2" => func_sumx2py2(args, cells),
        "SUMXMY2" => func_sumxmy2(args, cells),
        "SERIESSUM" => func_seriessum(args, cells),
        "PMT" => func_pmt(args, cells),
        "FV" => func_fv(args, cells),
        "PV" => func_pv(args, cells),
        "NPER" => func_nper(args, cells),
        "RATE" => func_rate(args, cells),
        "IPMT" => func_ipmt(args, cells),
        "PPMT" => func_ppmt(args, cells),
        "ISPMT" => func_ispmt(args, cells),
        "NPV" => func_npv(args, cells),
        "IRR" => func_irr(args, cells),
        "MIRR" => func_mirr(args, cells),
        "XNPV" => func_xnpv(args, cells),
        "XIRR" => func_xirr(args, cells),
        "SLN" => func_sln(args, cells),
        "SYD" => func_syd(args, cells),
        "DB" => func_db(args, cells),
        "DDB" => func_ddb(args, cells),
        "VDB" => func_vdb(args, cells),
        "AMORLINC" => func_amorlinc(args, cells),
        "AMORDEGRC" => func_amordegrc(args, cells),
        "EFFECT" => func_effect(args, cells),
        "NOMINAL" => func_nominal(args, cells),
        "RRI" => func_rri(args, cells),
        "CUMIPMT" => func_cumipmt(args, cells),
        "CUMPRINC" => func_cumprinc(args, cells),
        "FVSCHEDULE" => func_fvschedule(args, cells),
        "DOLLARDE" => func_dollarde(args, cells),
        "DOLLARFR" => func_dollarfr(args, cells),
        "EUROCONVERT" => func_euroconvert(args, cells),
        "PDURATION" => func_pduration(args, cells),
        "PRICEDISC" => func_pricedisc(args, cells),
        "DISC" => func_disc(args, cells),
        "RECEIVED" => func_received(args, cells),
        "YIELDDISC" => func_yielddisc(args, cells),
        "ACCRINTM" => func_accrintm(args, cells),
        "ACCRINT" => func_accrint(args, cells),
        "COUPDAYBS" => func_coupdaybs(args, cells),
        "COUPDAYS" => func_coupdays(args, cells),
        "COUPDAYSNC" => func_coupdaysnc(args, cells),
        "COUPNCD" => func_coupncd(args, cells),
        "COUPNUM" => func_coupnum(args, cells),
        "COUPPCD" => func_couppcd(args, cells),
        "INTRATE" => func_intrate(args, cells),
        "PRICEMAT" => func_pricemat(args, cells),
        "YIELDMAT" => func_yieldmat(args, cells),
        "DURATION" => func_duration(args, cells, false),
        "MDURATION" => func_duration(args, cells, true),
        "PRICE" => func_price(args, cells),
        "YIELD" => func_yield(args, cells),
        "ODDFPRICE" => func_oddfprice(args, cells),
        "ODDFYIELD" => func_oddfyield(args, cells),
        "ODDLPRICE" => func_oddlprice(args, cells),
        "ODDLYIELD" => func_oddlyield(args, cells),
        "TBILLPRICE" => func_tbillprice(args, cells),
        "TBILLYIELD" => func_tbillyield(args, cells),
        "TBILLEQ" => func_tbilleq(args, cells),
        // ── Database ─────────────────────────────────────────────────────────
        "DGET" => func_dget(args, cells),
        "DSUM" => func_dsum(args, cells),
        "DAVERAGE" => func_daverage(args, cells),
        "DCOUNT" => func_dcount(args, cells),
        "DCOUNTA" => func_dcounta(args, cells),
        "DMAX" => func_dmax(args, cells),
        "DMIN" => func_dmin(args, cells),
        "DPRODUCT" => func_dproduct(args, cells),
        "DSTDEV" => func_dstdev(args, cells),
        "DSTDEVP" => func_dstdevp(args, cells),
        "DVAR" => func_dvar(args, cells),
        "DVARP" => func_dvarp(args, cells),
        // ── LET / higher-order ───────────────────────────────────────────────
        "LET" => func_let(args, cells),
        "LAMBDA" => func_lambda(args, cells),
        "MAP" => func_map(args, cells),
        "REDUCE" => func_reduce(args, cells),
        "SCAN" => func_scan(args, cells),
        "BYROW" => func_byrow(args, cells),
        "BYCOL" => func_bycol(args, cells),
        // These worksheet names require code loading, a remote service, or a
        // live cube/data connection. Recognize them explicitly but fail
        // closed without evaluating arguments or performing external I/O.
        "CALL" | "CUBEKPIMEMBER" | "CUBEMEMBER" | "CUBEMEMBERPROPERTY" | "CUBERANKEDMEMBER"
        | "CUBESET" | "CUBESETCOUNT" | "CUBEVALUE" | "IMAGE" | "REGISTER.ID" | "RTD"
        | "STOCKHISTORY" | "WEBSERVICE" => func_external_boundary(name),
        "TRANSLATE" => func_translate(args, cells),
        _ => Ok(Variant::Error(ExcelError::Name)),
    }
}

fn func_external_boundary(name: &str) -> Result<Variant, String> {
    let _ = name;
    Ok(Variant::Error(ExcelError::NA))
}

fn func_translate(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Ok(Variant::Error(ExcelError::NA));
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let source = to_str(&evaluate(&args[1], cells)?);
    let target = to_str(&evaluate(&args[2], cells)?);
    let normalize_language = |language: &str| language.trim().to_ascii_lowercase();
    if !source.trim().is_empty()
        && !target.trim().is_empty()
        && normalize_language(&source) == normalize_language(&target)
    {
        return Ok(Variant::Str(text));
    }
    // Actual translation requires Microsoft's remote service. Keep the
    // offline runtime deterministic and fail closed for non-identity pairs.
    Ok(Variant::Error(ExcelError::NA))
}

// ── Arithmetic ────────────────────────────────────────────────────────────────

fn func_sum(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    // Fast path: single Range — iterate cell refs without cloning into Vec<Variant>
    if args.len() == 1
        && let FormulaExpr::Range { c1, r1, c2, r2, .. } = &args[0]
    {
        let mut sum = 0f64;
        for row in *r1..=*r2 {
            for col in *c1..=*c2 {
                if let Some(value) = cell_ref(cells, row, col) {
                    if let Variant::Error(error) = value {
                        return Ok(Variant::Error(error.clone()));
                    }
                    if let Some(f) = as_f64(value) {
                        sum += f;
                    }
                }
            }
        }
        return Ok(as_integer_if_whole(sum));
    }
    let sum: f64 = match collect_direct_numeric_args(args, cells, true) {
        Ok(values) => values.into_iter().sum(),
        Err(error) => return Ok(Variant::Error(error)),
    };
    Ok(as_integer_if_whole(sum))
}

fn func_average(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    // Excel includes numeric text and logical values supplied as direct
    // arguments to AVERAGE, while values of those types inside a reference
    // remain excluded. `collect_direct_numeric_args` applies that distinction
    // per argument rather than flattening everything through one coercion rule.
    let nums = match collect_direct_numeric_args(args, cells, true) {
        Ok(values) => values,
        Err(error) => return Ok(Variant::Error(error)),
    };
    if nums.is_empty() {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
}

/// Collect values for SUM/AVERAGE using Excel's reference-vs-direct-argument
/// coercion rule: numbers in a range are numeric, while scalar text and
/// logical arguments are coerced as direct arguments. Array-producing
/// expressions contribute their numeric elements without treating text as a
/// scalar argument.
fn collect_direct_numeric_args(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    coerce_direct: bool,
) -> Result<Vec<f64>, ExcelError> {
    let mut values = Vec::new();
    for arg in args {
        let is_reference = matches!(arg, FormulaExpr::Range { .. })
            || matches!(arg, FormulaExpr::FuncCall { name, .. } if name.eq_ignore_ascii_case("OFFSET"));
        if is_reference {
            let referenced = collect_values(arg, cells).map_err(|_| ExcelError::Value)?;
            for value in &referenced {
                if let Variant::Error(error) = value {
                    return Err(error.clone());
                }
                if let Some(number) = as_f64(value) {
                    values.push(number);
                }
            }
            continue;
        }
        match evaluate(arg, cells).map_err(|_| ExcelError::Value)? {
            Variant::Array(array) => {
                for value in array {
                    if let Variant::Error(error) = value {
                        return Err(error);
                    }
                    if let Some(number) = as_f64(&value) {
                        values.push(number);
                    }
                }
            }
            Variant::VbaArray(array) => {
                for value in array.elements {
                    if let Variant::Error(error) = value {
                        return Err(error);
                    }
                    if let Some(number) = as_f64(&value) {
                        values.push(number);
                    }
                }
            }
            Variant::Error(error) => return Err(error),
            value if coerce_direct => {
                values.push(to_float(&value).map_err(|_| ExcelError::Value)?);
            }
            Variant::Integer(value) => values.push(value as f64),
            Variant::Float(value) => values.push(value),
            _ => {}
        }
    }
    Ok(values)
}

fn func_percentof(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("PERCENTOF requires 2 arguments".into());
    }
    let subset = to_float(&func_sum(&args[0..1], cells)?)?;
    let all = to_float(&func_sum(&args[1..2], cells)?)?;
    if !subset.is_finite() || !all.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    if all == 0.0 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    Ok(as_integer_if_whole(subset / all))
}

fn append_a_value(value: Variant, values: &mut Vec<f64>) -> Result<(), ExcelError> {
    match value {
        Variant::Empty => {}
        Variant::Integer(n) => values.push(n as f64),
        Variant::Float(value) => values.push(value),
        Variant::Boolean(value) => values.push(if value { 1.0 } else { 0.0 }),
        Variant::Str(text) => values.push(text.parse::<f64>().unwrap_or(0.0)),
        Variant::Error(error) => return Err(error),
        Variant::Array(values_array) => {
            for value in values_array {
                append_a_value(value, values)?;
            }
        }
        Variant::VbaArray(array) => {
            for value in array.elements {
                append_a_value(value, values)?;
            }
        }
        Variant::Null | Variant::Record(_) => {}
        Variant::Date(value) => values.push(value as f64),
    }
    Ok(())
}

fn collect_a_values(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Vec<f64>, ExcelError> {
    let mut values = Vec::new();
    for value in collect_all(args, cells).map_err(|_| ExcelError::Value)? {
        append_a_value(value, &mut values)?;
    }
    Ok(values)
}

fn func_averagea(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let values = match collect_a_values(args, cells) {
        Ok(values) => values,
        Err(error) => return Ok(Variant::Error(error)),
    };
    if values.is_empty() {
        return Err("AVERAGEA: no values".into());
    }
    Ok(Variant::Float(
        values.iter().sum::<f64>() / values.len() as f64,
    ))
}

fn func_min(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    // Excel coerces numeric text and logical values supplied directly to MIN;
    // the same values inside a referenced range remain excluded.
    let values = match collect_direct_numeric_args(args, cells, true) {
        Ok(values) => values,
        Err(error) => return Ok(Variant::Error(error)),
    };
    let min = values.into_iter().reduce(f64::min);
    min.map(as_integer_if_whole)
        .ok_or_else(|| "MIN: no numeric values".into())
}

fn func_max(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    // Keep direct-argument coercion aligned with Excel while preserving
    // reference semantics for text and logical cells.
    let values = match collect_direct_numeric_args(args, cells, true) {
        Ok(values) => values,
        Err(error) => return Ok(Variant::Error(error)),
    };
    let max = values.into_iter().reduce(f64::max);
    max.map(as_integer_if_whole)
        .ok_or_else(|| "MAX: no numeric values".into())
}

fn func_mina(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let values = match collect_a_values(args, cells) {
        Ok(values) => values,
        Err(error) => return Ok(Variant::Error(error)),
    };
    Ok(as_integer_if_whole(
        values.into_iter().reduce(f64::min).unwrap_or(0.0),
    ))
}

fn func_maxa(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let values = match collect_a_values(args, cells) {
        Ok(values) => values,
        Err(error) => return Ok(Variant::Error(error)),
    };
    Ok(as_integer_if_whole(
        values.into_iter().reduce(f64::max).unwrap_or(0.0),
    ))
}

fn func_count(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let mut count = 0usize;
    for arg in args {
        let is_reference = matches!(arg, FormulaExpr::Range { .. })
            || matches!(arg, FormulaExpr::FuncCall { name, .. } if name.eq_ignore_ascii_case("OFFSET"));
        for value in if is_reference {
            collect_values(arg, cells)?
        } else {
            vec![evaluate(arg, cells)?]
        } {
            count += match value {
                Variant::Integer(_) | Variant::Float(_) | Variant::Date(_) => 1,
                // Excel counts logical values and numeric text supplied as
                // direct scalar arguments, but not those values in a range.
                Variant::Boolean(_) if !is_reference => 1,
                Variant::Str(text) if !is_reference && text.parse::<f64>().is_ok() => 1,
                Variant::Array(values) => values
                    .into_iter()
                    .filter(|value| {
                        matches!(
                            value,
                            Variant::Integer(_) | Variant::Float(_) | Variant::Date(_)
                        )
                    })
                    .count(),
                Variant::VbaArray(array) => array
                    .elements
                    .into_iter()
                    .filter(|value| {
                        matches!(
                            value,
                            Variant::Integer(_) | Variant::Float(_) | Variant::Date(_)
                        )
                    })
                    .count(),
                _ => 0,
            };
        }
    }
    Ok(Variant::Integer(count as i64))
}

fn func_counta(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let mut count = 0usize;
    for value in collect_all(args, cells)? {
        count += match value {
            Variant::Array(values) => values
                .into_iter()
                .filter(|value| !matches!(value, Variant::Empty))
                .count(),
            Variant::VbaArray(array) => array
                .elements
                .into_iter()
                .filter(|value| !matches!(value, Variant::Empty))
                .count(),
            value if !matches!(value, Variant::Empty) => 1,
            _ => 0,
        };
    }
    Ok(Variant::Integer(count as i64))
}

fn func_parity(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    even: bool,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err(format!(
            "{} requires 1 argument",
            if even { "ISEVEN" } else { "ISODD" }
        ));
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    if !value.is_finite() || value < i64::MIN as f64 || value > i64::MAX as f64 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let integer = value.trunc() as i64;
    Ok(Variant::Boolean((integer % 2 == 0) == even))
}

// ── Logical ───────────────────────────────────────────────────────────────────

fn func_if(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("IF requires 2 or 3 arguments".into());
    }
    let condition = evaluate(&args[0], cells)?;
    if let Variant::Error(error) = condition {
        return Ok(Variant::Error(error));
    }
    let condition_values = match condition {
        Variant::Array(ref values) => Some(values.clone()),
        Variant::VbaArray(ref array) => Some(array.elements.clone()),
        _ => None,
    };
    if let Some(condition_values) = condition_values {
        let true_value = evaluate(&args[1], cells)?;
        let false_value = if args.len() == 3 {
            evaluate(&args[2], cells)?
        } else {
            Variant::Boolean(false)
        };
        let values = |value: Variant| match value {
            Variant::Array(values) => values,
            Variant::VbaArray(array) => array.elements,
            scalar => vec![scalar],
        };
        let true_values = values(true_value);
        let false_values = values(false_value);
        let broadcast = |branch: &[Variant], index: usize| -> Result<Variant, ExcelError> {
            if branch.len() == 1 {
                Ok(branch[0].clone())
            } else if branch.len() == condition_values.len() {
                Ok(branch[index].clone())
            } else {
                Err(ExcelError::Value)
            }
        };
        let mut result = Vec::with_capacity(condition_values.len());
        for (index, condition) in condition_values.iter().enumerate() {
            if let Variant::Error(error) = condition {
                result.push(Variant::Error(error.clone()));
                continue;
            }
            let branch = if is_truthy(condition) {
                &true_values
            } else {
                &false_values
            };
            match broadcast(branch, index) {
                Ok(value) => result.push(value),
                Err(error) => return Ok(Variant::Error(error)),
            }
        }
        return Ok(Variant::Array(result));
    }
    if is_truthy(&condition) {
        evaluate(&args[1], cells)
    } else if args.len() == 3 {
        evaluate(&args[2], cells)
    } else {
        Ok(Variant::Boolean(false))
    }
}

fn func_and(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let mut first_error = None;
    let mut result = true;
    for a in args {
        match evaluate(a, cells)? {
            Variant::Error(error) => {
                first_error.get_or_insert(error);
            }
            Variant::Array(values) => {
                for value in values {
                    match value {
                        Variant::Error(error) => {
                            first_error.get_or_insert(error);
                        }
                        value => result &= is_truthy(&value),
                    }
                }
            }
            Variant::VbaArray(array) => {
                for value in array.elements {
                    match value {
                        Variant::Error(error) => {
                            first_error.get_or_insert(error);
                        }
                        value => result &= is_truthy(&value),
                    }
                }
            }
            value => result &= is_truthy(&value),
        };
    }
    if let Some(error) = first_error {
        return Ok(Variant::Error(error));
    }
    Ok(Variant::Boolean(result))
}

fn func_or(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let mut first_error = None;
    let mut result = false;
    for a in args {
        match evaluate(a, cells)? {
            Variant::Error(error) => {
                first_error.get_or_insert(error);
            }
            Variant::Array(values) => {
                for value in values {
                    match value {
                        Variant::Error(error) => {
                            first_error.get_or_insert(error);
                        }
                        value => result |= is_truthy(&value),
                    }
                }
            }
            Variant::VbaArray(array) => {
                for value in array.elements {
                    match value {
                        Variant::Error(error) => {
                            first_error.get_or_insert(error);
                        }
                        value => result |= is_truthy(&value),
                    }
                }
            }
            value => result |= is_truthy(&value),
        }
    }
    if let Some(error) = first_error {
        return Ok(Variant::Error(error));
    }
    Ok(Variant::Boolean(result))
}

fn func_not(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("NOT requires 1 argument".into());
    }
    let value = evaluate(&args[0], cells)?;
    match value {
        Variant::Error(error) => Ok(Variant::Error(error)),
        Variant::Array(values) => Ok(Variant::Array(
            values
                .into_iter()
                .map(|value| match value {
                    Variant::Error(error) => Variant::Error(error),
                    value => Variant::Boolean(!is_truthy(&value)),
                })
                .collect(),
        )),
        Variant::VbaArray(array) => Ok(Variant::Array(
            array
                .elements
                .into_iter()
                .map(|value| match value {
                    Variant::Error(error) => Variant::Error(error),
                    value => Variant::Boolean(!is_truthy(&value)),
                })
                .collect(),
        )),
        value => Ok(Variant::Boolean(!is_truthy(&value))),
    }
}

fn func_boolean_constant(args: &[FormulaExpr], value: bool) -> Result<Variant, String> {
    if !args.is_empty() {
        return Err(format!(
            "{} takes no arguments",
            if value { "TRUE" } else { "FALSE" }
        ));
    }
    Ok(Variant::Boolean(value))
}

fn func_iferror(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("IFERROR requires 2 arguments".into());
    }
    match evaluate(&args[0], cells) {
        Ok(Variant::Array(values)) => {
            let fallback = variant_values(evaluate(&args[1], cells)?);
            map_error_fallback(values, fallback, |_| true)
        }
        Ok(Variant::VbaArray(array)) => {
            let fallback = variant_values(evaluate(&args[1], cells)?);
            map_error_fallback(array.elements, fallback, |_| true)
        }
        Ok(Variant::Error(_)) | Err(_) => evaluate(&args[1], cells),
        Ok(v) => Ok(v),
    }
}

/// Return the fallback only for Excel's `#N/A` error.
///
/// Unlike IFERROR, IFNA must preserve other worksheet errors such as
/// `#DIV/0!` and `#VALUE!`. The fallback remains lazy so lookup formulas do
/// not evaluate an unused branch.
fn func_ifna(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("IFNA requires 2 arguments".into());
    }
    match evaluate(&args[0], cells) {
        Ok(Variant::Array(values)) => {
            let fallback = variant_values(evaluate(&args[1], cells)?);
            map_error_fallback(values, fallback, |error| error == &ExcelError::NA)
        }
        Ok(Variant::VbaArray(array)) => {
            let fallback = variant_values(evaluate(&args[1], cells)?);
            map_error_fallback(array.elements, fallback, |error| error == &ExcelError::NA)
        }
        Ok(Variant::Error(ExcelError::NA)) => evaluate(&args[1], cells),
        Ok(value) => Ok(value),
        Err(error) => Err(error),
    }
}

fn variant_values(value: Variant) -> Vec<Variant> {
    match value {
        Variant::Array(values) => values,
        Variant::VbaArray(array) => array.elements,
        scalar => vec![scalar],
    }
}

fn map_error_fallback(
    values: Vec<Variant>,
    fallback: Vec<Variant>,
    replace: impl Fn(&ExcelError) -> bool,
) -> Result<Variant, String> {
    if fallback.len() != 1 && fallback.len() != values.len() {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let mut result = Vec::with_capacity(values.len());
    for (index, value) in values.into_iter().enumerate() {
        let replacement = if fallback.len() == 1 {
            fallback[0].clone()
        } else {
            fallback[index].clone()
        };
        match value {
            Variant::Error(error) if replace(&error) => result.push(replacement),
            value => result.push(value),
        }
    }
    Ok(Variant::Array(result))
}

// ── Text ──────────────────────────────────────────────────────────────────────

fn func_left(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("LEFT requires 1 or 2 arguments".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let n = if args.len() == 2 {
        to_float(&evaluate(&args[1], cells)?)? as usize
    } else {
        1
    };
    Ok(Variant::Str(s.chars().take(n).collect()))
}

fn func_right(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("RIGHT requires 1 or 2 arguments".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let n = if args.len() == 2 {
        to_float(&evaluate(&args[1], cells)?)? as usize
    } else {
        1
    };
    let chars: Vec<char> = s.chars().collect();
    Ok(Variant::Str(
        chars[chars.len().saturating_sub(n)..].iter().collect(),
    ))
}

fn func_mid(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("MID requires 3 arguments".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let start = (to_float(&evaluate(&args[1], cells)?)? as usize).saturating_sub(1);
    let len = to_float(&evaluate(&args[2], cells)?)? as usize;
    Ok(Variant::Str(s.chars().skip(start).take(len).collect()))
}

fn func_len(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("LEN requires 1 argument".into());
    }
    Ok(Variant::Integer(
        to_str(&evaluate(&args[0], cells)?).chars().count() as i64,
    ))
}

// DBCS byte width: ASCII = 1, everything else = 2 (matches Excel's LENB/LEFTB/etc.)
fn char_byte_width(c: char) -> usize {
    if (c as u32) <= 0x7F { 1 } else { 2 }
}

fn str_byte_len(s: &str) -> usize {
    s.chars().map(char_byte_width).sum()
}

fn func_lenb(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("LENB requires 1 argument".into());
    }
    Ok(Variant::Integer(
        str_byte_len(&to_str(&evaluate(&args[0], cells)?)) as i64,
    ))
}

fn func_leftb(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("LEFTB requires 1 or 2 arguments".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let n = if args.len() == 2 {
        to_float(&evaluate(&args[1], cells)?)? as usize
    } else {
        1
    };
    let mut result = String::new();
    let mut bytes = 0;
    for c in s.chars() {
        let w = char_byte_width(c);
        if bytes + w > n {
            break;
        }
        bytes += w;
        result.push(c);
    }
    Ok(Variant::Str(result))
}

fn func_rightb(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("RIGHTB requires 1 or 2 arguments".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let n = if args.len() == 2 {
        to_float(&evaluate(&args[1], cells)?)? as usize
    } else {
        1
    };
    let chars: Vec<char> = s.chars().collect();
    let mut result = Vec::new();
    let mut bytes = 0;
    for &c in chars.iter().rev() {
        let w = char_byte_width(c);
        if bytes + w > n {
            break;
        }
        bytes += w;
        result.push(c);
    }
    result.reverse();
    Ok(Variant::Str(result.into_iter().collect()))
}

fn func_midb(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("MIDB requires 3 arguments".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let start_b = (to_float(&evaluate(&args[1], cells)?)? as usize).saturating_sub(1);
    let num_bytes = to_float(&evaluate(&args[2], cells)?)? as usize;
    let mut result = String::new();
    let mut pos = 0usize;
    for c in s.chars() {
        let w = char_byte_width(c);
        if pos >= start_b {
            let taken = result.chars().map(char_byte_width).sum::<usize>();
            if taken + w > num_bytes {
                break;
            }
            result.push(c);
        }
        pos += w;
    }
    Ok(Variant::Str(result))
}

// ── Rounding ──────────────────────────────────────────────────────────────────

fn round_multiplier(num_digits: i32) -> f64 {
    10_f64.powi(num_digits)
}

fn func_round(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("ROUND requires 2 arguments".into());
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let digits = to_float(&evaluate(&args[1], cells)?)? as i32;
    let mult = round_multiplier(digits);
    Ok(as_integer_if_whole((num * mult).round() / mult))
}

fn func_roundup(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("ROUNDUP requires 2 arguments".into());
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let digits = to_float(&evaluate(&args[1], cells)?)? as i32;
    let mult = round_multiplier(digits);
    let result = if num >= 0.0 {
        (num * mult).ceil()
    } else {
        (num * mult).floor()
    };
    Ok(as_integer_if_whole(result / mult))
}

fn func_rounddown(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("ROUNDDOWN requires 2 arguments".into());
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let digits = to_float(&evaluate(&args[1], cells)?)? as i32;
    let mult = round_multiplier(digits);
    let result = if num >= 0.0 {
        (num * mult).floor()
    } else {
        (num * mult).ceil()
    };
    Ok(as_integer_if_whole(result / mult))
}

fn func_concatenate(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let mut s = String::new();
    for a in args {
        s.push_str(&to_str(&evaluate(a, cells)?));
    }
    Ok(Variant::Str(s))
}

fn func_text(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("TEXT requires 2 arguments".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    let fmt = to_str(&evaluate(&args[1], cells)?);
    Ok(Variant::Str(apply_text_format(n, &fmt)))
}

fn apply_text_format(n: f64, fmt: &str) -> String {
    let fmt_upper = fmt.to_uppercase();
    // Date format detection: contains YYYY, YY, or DD (not percentage)
    if !fmt.ends_with('%')
        && (fmt_upper.contains("YYYY") || fmt_upper.contains("YY") || fmt_upper.contains("DD"))
    {
        let (y, m, d) = serial_to_ymd(n as i64);
        return fmt
            .replace("YYYY", &format!("{:04}", y))
            .replace("yyyy", &format!("{:04}", y))
            .replace("YY", &format!("{:02}", y % 100))
            .replace("yy", &format!("{:02}", y % 100))
            .replace("MM", &format!("{:02}", m))
            .replace("mm", &format!("{:02}", m))
            .replace("DD", &format!("{:02}", d))
            .replace("dd", &format!("{:02}", d));
    }
    if fmt.ends_with('%') {
        let dec = fmt
            .trim_end_matches('%')
            .split('.')
            .nth(1)
            .map(|s| s.len())
            .unwrap_or(0);
        format!("{:.prec$}%", n * 100.0, prec = dec)
    } else if fmt.contains('.') {
        let dec = fmt.split('.').nth(1).map(|s| s.len()).unwrap_or(2);
        format!("{:.prec$}", n, prec = dec)
    } else {
        format!("{}", n as i64)
    }
}

// ── Lookup ────────────────────────────────────────────────────────────────────

fn require_range(expr: &FormulaExpr, fname: &str) -> Result<(u32, u32, u32, u32), String> {
    match expr {
        FormulaExpr::Range { c1, r1, c2, r2, .. } => Ok((*c1, *r1, *c2, *r2)),
        _ => Err(format!("{}: table argument must be a range", fname)),
    }
}

fn lookup_mode(value: &Variant) -> Result<i32, String> {
    let number = to_float(value)?;
    if !number.is_finite()
        || number.fract() != 0.0
        || number < i32::MIN as f64
        || number > i32::MAX as f64
    {
        return Err("lookup mode must be an integer".into());
    }
    Ok(number as i32)
}

fn func_vlookup(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 4 {
        return Err("VLOOKUP requires 3 or 4 arguments".into());
    }
    let key = evaluate(&args[0], cells)?;
    let (c1, r1, c2, r2) = require_range(&args[1], "VLOOKUP")?;
    let col_n = match lookup_mode(&evaluate(&args[2], cells)?) {
        Ok(value) => i64::from(value),
        Err(_) => return Ok(Variant::Error(ExcelError::Value)),
    };
    if col_n <= 0 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let exact = if args.len() == 4 {
        !is_truthy(&evaluate(&args[3], cells)?)
    } else {
        false
    };
    let ret_col = i64::from(c1) + col_n - 1;
    if ret_col > i64::from(c2) {
        return Ok(Variant::Error(ExcelError::Ref));
    }

    if exact {
        for row in r1..=r2 {
            if variant_eq(&cell_val(cells, row, c1), &key) {
                return Ok(cell_val(cells, row, ret_col as u32));
            }
        }
        Ok(Variant::Error(ExcelError::NA))
    } else {
        let mut sorted = true;
        for row in r1..r2 {
            let left = cell_val(cells, row, c1);
            let right = cell_val(cells, row + 1, c1);
            let order = match variant_cmp(&left, &right) {
                Ok(order) => order,
                Err(_) => return Ok(Variant::Error(ExcelError::Value)),
            };
            if order == Ordering::Greater {
                sorted = false;
                break;
            }
        }
        if !sorted {
            return Ok(Variant::Error(ExcelError::NA));
        }
        // Binary search for largest row where cell <= key (data assumed sorted ascending)
        let (mut lo, mut hi) = (r1 as i64, r2 as i64);
        let mut best: Option<u32> = None;
        while lo <= hi {
            let mid = lo + (hi - lo) / 2;
            match variant_cmp(&cell_val(cells, mid as u32, c1), &key) {
                Ok(Ordering::Less) | Ok(Ordering::Equal) => {
                    best = Some(mid as u32);
                    lo = mid + 1;
                }
                _ => {
                    hi = mid - 1;
                }
            }
        }
        Ok(best
            .map(|row| cell_val(cells, row, ret_col as u32))
            .unwrap_or(Variant::Error(ExcelError::NA)))
    }
}

fn func_hlookup(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 4 {
        return Err("HLOOKUP requires 3 or 4 arguments".into());
    }
    let key = evaluate(&args[0], cells)?;
    let (c1, r1, c2, r2) = require_range(&args[1], "HLOOKUP")?;
    let row_n = match lookup_mode(&evaluate(&args[2], cells)?) {
        Ok(value) => i64::from(value),
        Err(_) => return Ok(Variant::Error(ExcelError::Value)),
    };
    if row_n <= 0 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let height = i64::from(r2 - r1 + 1);
    if row_n > height {
        return Ok(Variant::Error(ExcelError::Ref));
    }
    let exact = if args.len() == 4 {
        !is_truthy(&evaluate(&args[3], cells)?)
    } else {
        false
    };
    let ret_row = r1 + row_n as u32 - 1;

    if exact {
        for col in c1..=c2 {
            if variant_eq(&cell_val(cells, r1, col), &key) {
                return Ok(cell_val(cells, ret_row, col));
            }
        }
        Ok(Variant::Error(ExcelError::NA))
    } else {
        let mut sorted = true;
        for col in c1..c2 {
            let left = cell_val(cells, r1, col);
            let right = cell_val(cells, r1, col + 1);
            let order = match variant_cmp(&left, &right) {
                Ok(order) => order,
                Err(_) => return Ok(Variant::Error(ExcelError::Value)),
            };
            if order == Ordering::Greater {
                sorted = false;
                break;
            }
        }
        if !sorted {
            return Ok(Variant::Error(ExcelError::NA));
        }
        // Binary search for largest col where cell <= key (data assumed sorted ascending)
        let (mut lo, mut hi) = (c1 as i64, c2 as i64);
        let mut best: Option<u32> = None;
        while lo <= hi {
            let mid = lo + (hi - lo) / 2;
            match variant_cmp(&cell_val(cells, r1, mid as u32), &key) {
                Ok(Ordering::Less) | Ok(Ordering::Equal) => {
                    best = Some(mid as u32);
                    lo = mid + 1;
                }
                _ => {
                    hi = mid - 1;
                }
            }
        }
        Ok(best
            .map(|col| cell_val(cells, ret_row, col))
            .unwrap_or(Variant::Error(ExcelError::NA)))
    }
}

fn func_index(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 4 {
        return Err("INDEX requires 2 to 4 arguments".into());
    }
    let range = require_range(&args[0], "INDEX").ok();
    let array_values = if range.is_none() {
        Some(flatten_array_vals(collect_values(&args[0], cells)?))
    } else {
        None
    };
    let (height, width) = if let Some((c1, r1, c2, r2)) = range.as_ref() {
        (i64::from(r2 - r1 + 1), i64::from(c2 - c1 + 1))
    } else {
        let values = array_values.as_ref().expect("array values are present");
        let (rows, cols) = array_shape_for_expr(&args[0], cells, values.len());
        (rows as i64, cols as i64)
    };
    let integer_index = |arg: &FormulaExpr| -> Result<i64, String> {
        match evaluate(arg, cells)? {
            Variant::Integer(value) => Ok(value),
            Variant::Float(value)
                if value.is_finite()
                    && value.fract() == 0.0
                    && value >= i64::MIN as f64
                    && value <= i64::MAX as f64 =>
            {
                Ok(value as i64)
            }
            _ => Err("INDEX row and column numbers must be integers".into()),
        }
    };
    let row_off = match integer_index(&args[1]) {
        Ok(value) => value,
        Err(_) => return Ok(Variant::Error(ExcelError::Value)),
    };
    let col_off = if args.len() >= 3 {
        match integer_index(&args[2]) {
            Ok(value) => value,
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        }
    } else {
        1i64
    };
    if args.len() == 4 {
        let area_num = match integer_index(&args[3]) {
            Ok(value) => value,
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        };
        // The parser has no union/reference-area AST yet.  Accept the
        // single-area reference form explicitly, but never silently select a
        // different area when a caller asks for one.
        if area_num != 1 {
            return Ok(Variant::Error(ExcelError::Ref));
        }
    }
    if row_off < 0 || col_off < 0 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    if row_off > height || col_off > width {
        return Ok(Variant::Error(ExcelError::Ref));
    }
    if row_off == 0 || col_off == 0 {
        let row_start = if row_off == 0 {
            0
        } else {
            row_off as usize - 1
        };
        let row_end = if row_off == 0 {
            height as usize
        } else {
            row_start + 1
        };
        let col_start = if col_off == 0 {
            0
        } else {
            col_off as usize - 1
        };
        let col_end = if col_off == 0 {
            width as usize
        } else {
            col_start + 1
        };
        let values = if let Some(source) = array_values {
            (row_start..row_end)
                .flat_map(|row| {
                    source[row * width as usize + col_start..row * width as usize + col_end]
                        .iter()
                        .cloned()
                })
                .collect()
        } else {
            let (c1, r1, _c2, _r2) = range.expect("range is present");
            (row_start..row_end)
                .flat_map(|row| {
                    (col_start..col_end)
                        .map(move |col| cell_val(cells, r1 + row as u32, c1 + col as u32))
                })
                .collect()
        };
        return Ok(Variant::Array(values));
    }
    if let Some(source) = array_values {
        Ok(source[(row_off as usize - 1) * width as usize + col_off as usize - 1].clone())
    } else {
        let (c1, r1, _c2, _r2) = range.expect("range is present");
        Ok(cell_val(
            cells,
            r1 + row_off as u32 - 1,
            c1 + col_off as u32 - 1,
        ))
    }
}

fn func_match_fn(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("MATCH requires 2 or 3 arguments".into());
    }
    let key = evaluate(&args[0], cells)?;
    let lookup_shape = expression_shape(&args[1], cells)?;
    if lookup_shape.0 != 1 && lookup_shape.1 != 1 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let vals = collect_values(&args[1], cells)?;
    let mtype = if args.len() == 3 {
        match evaluate(&args[2], cells)? {
            Variant::Integer(value) if matches!(value, -1..=1) => value as i32,
            Variant::Float(value)
                if value.is_finite() && value.fract() == 0.0 && matches!(value as i64, -1..=1) =>
            {
                value as i32
            }
            _ => return Ok(Variant::Error(ExcelError::Value)),
        }
    } else {
        1
    };

    if matches!(mtype, 1 | -1) {
        let ascending = mtype == 1;
        let mut sorted = true;
        for pair in vals.windows(2) {
            let order = match variant_cmp(&pair[0], &pair[1]) {
                Ok(order) => order,
                Err(_) => return Ok(Variant::Error(ExcelError::Value)),
            };
            if (ascending && order == Ordering::Greater) || (!ascending && order == Ordering::Less)
            {
                sorted = false;
                break;
            }
        }
        if !sorted {
            return Ok(Variant::Error(ExcelError::NA));
        }
    }

    match mtype {
        0 => Ok(vals
            .iter()
            .position(|v| match (&key, v) {
                (Variant::Str(pattern), Variant::Str(text))
                    if pattern.contains('*') || pattern.contains('?') || pattern.contains('~') =>
                {
                    wildcard_match(&text.to_uppercase(), &pattern.to_uppercase())
                }
                _ => variant_eq(v, &key),
            })
            .map(|i| Variant::Integer((i + 1) as i64))
            .unwrap_or(Variant::Error(ExcelError::NA))),
        1 => {
            let mut best = None;
            for (i, v) in vals.iter().enumerate() {
                if matches!(
                    variant_cmp(v, &key),
                    Ok(Ordering::Less) | Ok(Ordering::Equal)
                ) {
                    best = Some(i);
                }
            }
            Ok(best
                .map(|i| Variant::Integer((i + 1) as i64))
                .unwrap_or(Variant::Error(ExcelError::NA)))
        }
        -1 => {
            let mut best = None;
            for (i, v) in vals.iter().enumerate() {
                if matches!(
                    variant_cmp(v, &key),
                    Ok(Ordering::Greater) | Ok(Ordering::Equal)
                ) {
                    best = Some(i);
                }
            }
            Ok(best
                .map(|i| Variant::Integer((i + 1) as i64))
                .unwrap_or(Variant::Error(ExcelError::NA)))
        }
        _t => Ok(Variant::Error(ExcelError::Value)),
    }
}

fn pivot_reference_bounds(expr: &FormulaExpr) -> Option<(u32, u32, u32, u32)> {
    match expr {
        FormulaExpr::CellRef {
            col,
            row,
            sheet: None,
            ..
        } => Some((*row, *col, *row, *col)),
        FormulaExpr::Range {
            c1,
            r1,
            c2,
            r2,
            sheet: None,
            ..
        } => Some((*r1, *c1, *r2, *c2)),
        _ => None,
    }
}

fn pivot_item_matches(actual: &Variant, wanted: &Variant) -> bool {
    match (actual, wanted) {
        (Variant::Str(left), Variant::Str(right)) => left.eq_ignore_ascii_case(right),
        _ => variant_eq(actual, wanted),
    }
}

fn pivot_data_header_matches(actual: &str, wanted: &str) -> bool {
    let actual = actual.trim().to_ascii_lowercase();
    let wanted = wanted.trim().to_ascii_lowercase();
    if actual == wanted {
        return true;
    }
    [
        "sum of ",
        "count of ",
        "average of ",
        "min of ",
        "max of ",
        "product of ",
    ]
    .iter()
    .any(|prefix| {
        actual
            .strip_prefix(prefix)
            .is_some_and(|name| name == wanted)
    })
}

/// Resolve a bounded worksheet-rendered PivotTable without reading external
/// connections or inventing cache records. The supported shape has a header
/// row containing the value field, field labels with row/column items, and an
/// optional `Grand Total` row. Unrecognized PivotTable layouts fail closed.
fn func_getpivotdata(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 254 || !(args.len() - 2).is_multiple_of(2) {
        return Err("GETPIVOTDATA requires 2 to 254 arguments with field/item pairs".into());
    }
    let data_field = match evaluate(&args[0], cells)? {
        Variant::Str(value) if !value.trim().is_empty() => value,
        _ => return Ok(Variant::Error(ExcelError::Value)),
    };
    let Some((anchor_row, anchor_col, _, _)) = pivot_reference_bounds(&args[1]) else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    let max_row = anchor_row.saturating_add(255);
    let max_col = anchor_col.saturating_add(63);
    let header_row = anchor_row;
    let data_col = (anchor_col..=max_col).find(|col| {
        matches!(
            cells.get(&(header_row, *col)).map(|cell| &cell.value),
            Some(Variant::Str(value)) if pivot_data_header_matches(value, &data_field)
        )
    });
    let Some(data_col) = data_col else {
        return Ok(Variant::Error(ExcelError::Ref));
    };

    let mut target_row = None;
    let mut target_col = Some(data_col);
    for pair in args[2..].chunks_exact(2) {
        let field = match evaluate(&pair[0], cells)? {
            Variant::Str(value) if !value.trim().is_empty() => value,
            _ => return Ok(Variant::Error(ExcelError::Value)),
        };
        let item = evaluate(&pair[1], cells)?;
        let field_position = (anchor_row..=max_row)
            .flat_map(|row| (anchor_col..=max_col).map(move |col| (row, col)))
            .find(|(row, col)| {
                matches!(
                    cells.get(&(*row, *col)).map(|cell| &cell.value),
                    Some(Variant::Str(value)) if value.eq_ignore_ascii_case(&field)
                )
            });
        let Some((field_row, field_col)) = field_position else {
            return Ok(Variant::Error(ExcelError::Ref));
        };

        let row_match = (field_row.saturating_add(1)..=max_row).find(|row| {
            cells
                .get(&(*row, field_col))
                .is_some_and(|cell| pivot_item_matches(&cell.value, &item))
        });
        let col_match = (field_col.saturating_add(1)..=max_col).find(|col| {
            cells
                .get(&(field_row, *col))
                .is_some_and(|cell| pivot_item_matches(&cell.value, &item))
        });
        if row_match.is_none() && col_match.is_none() {
            return Ok(Variant::Error(ExcelError::Ref));
        }
        if let Some(row) = row_match {
            if target_row.is_some_and(|existing| existing != row) {
                return Ok(Variant::Error(ExcelError::Ref));
            }
            target_row = Some(row);
        }
        if let Some(col) = col_match {
            if target_col.is_some_and(|existing| existing != col && existing != data_col) {
                return Ok(Variant::Error(ExcelError::Ref));
            }
            target_col = Some(col);
        }
    }

    let grand_total_row = (header_row.saturating_add(1)..=max_row).find(|row| {
        cells.get(&(*row, anchor_col)).is_some_and(|cell| {
            matches!(&cell.value, Variant::Str(value) if value.eq_ignore_ascii_case("grand total"))
        })
    });
    let grand_total_col = (anchor_col..=max_col).find(|col| {
        matches!(
            cells.get(&(header_row, *col)).map(|cell| &cell.value),
            Some(Variant::Str(value)) if value.eq_ignore_ascii_case("grand total")
        )
    });
    let row = target_row
        .or(grand_total_row)
        .ok_or_else(|| "GETPIVOTDATA: Grand Total is not visible".to_string())?;
    let col = if target_col == Some(data_col) && target_row.is_none() {
        grand_total_col.unwrap_or(data_col)
    } else {
        target_col.unwrap_or(data_col)
    };
    Ok(cell_val(cells, row, col))
}

// ── Criteria matching (for COUNTIF / SUMIF / COUNTIFS / SUMIFS) ──────────────

fn matches_criteria(val: &Variant, criteria: &Variant) -> bool {
    let crit_str = match criteria {
        Variant::Str(s) => s.clone(),
        other => return variant_eq(val, other),
    };
    // Comparison operator prefix
    let (op, rest) = if let Some(r) = crit_str.strip_prefix(">=") {
        (">=", r)
    } else if let Some(r) = crit_str.strip_prefix("<=") {
        ("<=", r)
    } else if let Some(r) = crit_str.strip_prefix("<>") {
        ("<>", r)
    } else if let Some(r) = crit_str.strip_prefix('>') {
        (">", r)
    } else if let Some(r) = crit_str.strip_prefix('<') {
        ("<", r)
    } else {
        ("=", crit_str.as_str())
    };

    if op == "=" {
        // Wildcard pattern match (case-insensitive)
        if rest.contains('*') || rest.contains('?') {
            return wildcard_match(&to_str(val).to_uppercase(), &rest.to_uppercase());
        }
        // Numeric or string equality
        if let Ok(n) = rest.parse::<f64>() {
            return variant_eq(val, &as_integer_if_whole(n));
        }
        return to_str(val).to_uppercase() == rest.to_uppercase();
    }

    // Comparison: both sides must be numeric
    let val_f = match to_float(val) {
        Ok(f) => f,
        Err(_) => return false,
    };
    let crit_f = match rest.parse::<f64>() {
        Ok(f) => f,
        Err(_) => return false,
    };
    match op {
        ">" => val_f > crit_f,
        ">=" => val_f >= crit_f,
        "<" => val_f < crit_f,
        "<=" => val_f <= crit_f,
        "<>" => (val_f - crit_f).abs() > f64::EPSILON,
        _ => false,
    }
}

pub(crate) fn wildcard_match(text: &str, pattern: &str) -> bool {
    if pattern.contains('~') {
        return wildcard_match_escaped(text, pattern);
    }
    // Fast path: patterns without '?' cover ~90% of real-world Excel wildcards.
    // Split on '*' and match segments without any Vec allocation.
    if !pattern.contains('?') {
        let parts: Vec<&str> = pattern.split('*').collect();
        // All '*' — match anything
        if parts.iter().all(|p| p.is_empty()) {
            return true;
        }
        let first = parts[0];
        let last = *parts.last().unwrap();
        // Prefix check
        if !first.is_empty() && (text.len() < first.len() || &text[..first.len()] != first) {
            return false;
        }
        // Suffix check
        if !last.is_empty() {
            if text.len() < last.len() || &text[text.len() - last.len()..] != last {
                return false;
            }
            // Prefix and suffix must not overlap
            if !first.is_empty() && first.len() + last.len() > text.len() {
                return false;
            }
        }
        // Middle segments: scan left-to-right (O(n·m) worst case but no allocation)
        let start = first.len();
        let end = text.len() - if last.is_empty() { 0 } else { last.len() };
        let mut pos = start;
        let inner = &parts[1..parts.len() - 1];
        for &mid in inner {
            if mid.is_empty() {
                continue;
            }
            match text[pos..end].find(mid) {
                Some(i) => pos += i + mid.len(),
                None => return false,
            }
        }
        return true;
    }
    // Full DP for patterns containing '?' — O(|text|×|pattern|), no exponential recursion.
    let t: Vec<char> = text.chars().collect();
    let p: Vec<char> = pattern.chars().collect();
    let (n, m) = (t.len(), p.len());
    let mut dp = vec![vec![false; m + 1]; n + 1];
    dp[0][0] = true;
    for j in 1..=m {
        if p[j - 1] == '*' {
            dp[0][j] = dp[0][j - 1];
        }
    }
    for i in 1..=n {
        for j in 1..=m {
            dp[i][j] = match p[j - 1] {
                '*' => dp[i - 1][j] || dp[i][j - 1],
                '?' => dp[i - 1][j - 1],
                c => dp[i - 1][j - 1] && t[i - 1] == c,
            };
        }
    }
    dp[n][m]
}

/// Excel wildcard matching with `~` escapes (`~*`, `~?`, and `~~`).
/// Patterns containing an escape use this bounded DP path; the common
/// unescaped path above keeps its allocation-light fast path.
fn wildcard_match_escaped(text: &str, pattern: &str) -> bool {
    let tokens: Vec<(char, bool)> = {
        let mut tokens = Vec::new();
        let mut chars = pattern.chars();
        while let Some(c) = chars.next() {
            if c == '~' {
                tokens.push((chars.next().unwrap_or('~'), true));
            } else {
                tokens.push((c, false));
            }
        }
        tokens
    };
    let text: Vec<char> = text.chars().collect();
    let (n, m) = (text.len(), tokens.len());
    let mut dp = vec![vec![false; m + 1]; n + 1];
    dp[0][0] = true;
    for j in 1..=m {
        if tokens[j - 1] == ('*', false) {
            dp[0][j] = dp[0][j - 1];
        }
    }
    for i in 1..=n {
        for j in 1..=m {
            let (token, escaped) = tokens[j - 1];
            dp[i][j] = if !escaped && token == '*' {
                dp[i - 1][j] || dp[i][j - 1]
            } else if !escaped && token == '?' {
                dp[i - 1][j - 1]
            } else {
                dp[i - 1][j - 1] && text[i - 1] == token
            };
        }
    }
    dp[n][m]
}

/// Like wildcard_match but pattern only needs to match a prefix of text (for SEARCH positioning).
fn wildcard_match_prefix(text: &[char], pattern: &[char]) -> bool {
    // Fast path: no '?', pattern is a simple literal → check if text starts with it
    if !pattern.contains(&'?') && !pattern.contains(&'*') {
        return text.len() >= pattern.len() && text[..pattern.len()] == *pattern;
    }
    if !pattern.contains(&'?') && pattern == [b'*' as char].as_slice() {
        return true; // single '*' matches any prefix
    }
    // Full DP — O(|text|×|pattern|), avoids exponential recursion.
    let (n, m) = (text.len(), pattern.len());
    let mut dp = vec![vec![false; m + 1]; n + 1];
    dp[0][0] = true;
    for j in 1..=m {
        if pattern[j - 1] == '*' {
            dp[0][j] = dp[0][j - 1];
        }
    }
    for i in 1..=n {
        for j in 1..=m {
            dp[i][j] = match pattern[j - 1] {
                '*' => dp[i - 1][j] || dp[i][j - 1],
                '?' => dp[i - 1][j - 1],
                c => dp[i - 1][j - 1] && text[i - 1] == c,
            };
        }
        if dp[i][m] {
            return true;
        }
    }
    dp[0][m]
}

// ── ParsedCriteria — criteria parsed once, matched N times ───────────────────

#[derive(Clone, Copy)]
enum CompOp {
    Gt,
    Ge,
    Lt,
    Le,
    Ne,
}

enum ParsedCriteria {
    Direct(Variant),      // non-string Variant: exact match
    EqNum(Variant),       // "5" or "5.0": numeric equality
    EqStr(String),        // "abc": case-insensitive string equality
    NeStr(String),        // "<>abc": case-insensitive string inequality
    Wildcard(String),     // "a*b": wildcard (pre-uppercased)
    CompNum(CompOp, f64), // ">5", ">=3": numeric comparison
}

fn parse_criteria(crit: &Variant) -> ParsedCriteria {
    let s = match crit {
        Variant::Str(s) => s,
        other => return ParsedCriteria::Direct(other.clone()),
    };
    let (op, rest) = if let Some(r) = s.strip_prefix(">=") {
        (Some(CompOp::Ge), r)
    } else if let Some(r) = s.strip_prefix("<=") {
        (Some(CompOp::Le), r)
    } else if let Some(r) = s.strip_prefix("<>") {
        (Some(CompOp::Ne), r)
    } else if let Some(r) = s.strip_prefix('>') {
        (Some(CompOp::Gt), r)
    } else if let Some(r) = s.strip_prefix('<') {
        (Some(CompOp::Lt), r)
    } else {
        (None, s.as_str())
    };
    match op {
        Some(CompOp::Ne) => {
            // "<>abc" → string inequality; "<>5" → numeric inequality
            if let Ok(n) = rest.parse::<f64>() {
                ParsedCriteria::CompNum(CompOp::Ne, n)
            } else {
                ParsedCriteria::NeStr(rest.to_string())
            }
        }
        Some(op) => {
            // Numeric comparisons only (">abc" → always false via NaN)
            ParsedCriteria::CompNum(op, rest.parse::<f64>().unwrap_or(f64::NAN))
        }
        None => {
            if rest.contains('*') || rest.contains('?') {
                ParsedCriteria::Wildcard(rest.to_uppercase())
            } else if let Ok(n) = rest.parse::<f64>() {
                ParsedCriteria::EqNum(as_integer_if_whole(n))
            } else {
                ParsedCriteria::EqStr(rest.to_string())
            }
        }
    }
}

fn matches_parsed(val: &Variant, crit: &ParsedCriteria) -> bool {
    match crit {
        ParsedCriteria::Direct(v) => variant_eq(val, v),
        ParsedCriteria::EqNum(v) => variant_eq(val, v),
        ParsedCriteria::EqStr(s) => match val {
            Variant::Str(vs) => vs.eq_ignore_ascii_case(s),
            Variant::Empty => s.is_empty(),
            _ => to_str(val).eq_ignore_ascii_case(s),
        },
        ParsedCriteria::NeStr(s) => match val {
            Variant::Str(vs) => !vs.eq_ignore_ascii_case(s),
            Variant::Empty => !s.is_empty(),
            _ => !to_str(val).eq_ignore_ascii_case(s),
        },
        ParsedCriteria::Wildcard(p) => wildcard_match(&to_str(val).to_uppercase(), p),
        ParsedCriteria::CompNum(op, n) => match to_float(val) {
            Ok(v) => match op {
                CompOp::Gt => v > *n,
                CompOp::Ge => v >= *n,
                CompOp::Lt => v < *n,
                CompOp::Le => v <= *n,
                CompOp::Ne => (v - n).abs() > f64::EPSILON,
            },
            Err(_) => false,
        },
    }
}

// ── Statistical ───────────────────────────────────────────────────────────────

fn func_countif(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("COUNTIF requires 2 arguments".into());
    }
    let vals = flatten_values(collect_values(&args[0], cells)?);
    let pcrit = parse_criteria(&evaluate(&args[1], cells)?);
    Ok(Variant::Integer(
        vals.iter().filter(|v| matches_parsed(v, &pcrit)).count() as i64,
    ))
}

fn func_sumif(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("SUMIF requires 2 or 3 arguments".into());
    }
    let range_vals = flatten_values(collect_values(&args[0], cells)?);
    let pcrit = parse_criteria(&evaluate(&args[1], cells)?);
    let total: f64 = if args.len() == 3 {
        let sum_vals = flatten_values(collect_values(&args[2], cells)?);
        let mut total = 0.0;
        for (rv, sv) in range_vals.iter().zip(sum_vals.iter()) {
            if matches_parsed(rv, &pcrit) {
                if let Variant::Error(error) = sv {
                    return Ok(Variant::Error(error.clone()));
                }
                if let Ok(value) = to_float(sv) {
                    total += value;
                }
            }
        }
        total
    } else {
        // 2-arg: criteria range = sum range — avoid cloning range_vals
        let mut total = 0.0;
        for value in &range_vals {
            if matches_parsed(value, &pcrit) {
                if let Variant::Error(error) = value {
                    return Ok(Variant::Error(error.clone()));
                }
                if let Ok(number) = to_float(value) {
                    total += number;
                }
            }
        }
        total
    };
    Ok(as_integer_if_whole(total))
}

fn func_sumifs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return Err("SUMIFS requires sum_range then pairs of (range,criteria)".into());
    }
    let sum_vals = flatten_values(collect_values(&args[0], cells)?);
    let sum_shape = expression_shape(&args[0], cells)?;
    let n = sum_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        if ensure_same_shape(sum_shape, &args[i], cells, "SUMIFS").is_err() {
            return Ok(Variant::Error(ExcelError::Value));
        }
        let range_vals = flatten_values(collect_values(&args[i], cells)?);
        let pcrit = parse_criteria(&evaluate(&args[i + 1], cells)?);
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_parsed(rv, &pcrit) {
                mask[j] = false;
            }
        }
        i += 2;
    }
    let mut total = 0.0;
    for (index, value) in sum_vals.iter().enumerate() {
        if mask[index] {
            if let Variant::Error(error) = value {
                return Ok(Variant::Error(error.clone()));
            }
            if let Ok(number) = to_float(value) {
                total += number;
            }
        }
    }
    Ok(as_integer_if_whole(total))
}

fn func_countifs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || !args.len().is_multiple_of(2) {
        return Err("COUNTIFS requires pairs of (range,criteria)".into());
    }
    let first_vals = flatten_values(collect_values(&args[0], cells)?);
    let first_shape = expression_shape(&args[0], cells)?;
    let n = first_vals.len();
    let mut mask = vec![true; n];
    let pcrit0 = parse_criteria(&evaluate(&args[1], cells)?);
    for (j, rv) in first_vals.iter().enumerate() {
        if !matches_parsed(rv, &pcrit0) {
            mask[j] = false;
        }
    }
    let mut i = 2;
    while i + 1 < args.len() {
        if ensure_same_shape(first_shape, &args[i], cells, "COUNTIFS").is_err() {
            return Ok(Variant::Error(ExcelError::Value));
        }
        let range_vals = flatten_values(collect_values(&args[i], cells)?);
        let pcrit = parse_criteria(&evaluate(&args[i + 1], cells)?);
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_parsed(rv, &pcrit) {
                mask[j] = false;
            }
        }
        i += 2;
    }
    Ok(Variant::Integer(mask.iter().filter(|&&b| b).count() as i64))
}

fn func_median(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("MEDIAN requires at least 1 argument".into());
    }
    let mut nums = match collect_numeric_values_with_errors(args, cells) {
        Ok(values) => values,
        Err(error) => return Ok(Variant::Error(error)),
    };
    if nums.is_empty() {
        return Err("MEDIAN: no numeric values".into());
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    let mid = nums.len() / 2;
    let result = if nums.len().is_multiple_of(2) {
        (nums[mid - 1] + nums[mid]) / 2.0
    } else {
        nums[mid]
    };
    Ok(as_integer_if_whole(result))
}

fn func_mode_mult(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    spill: bool,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("MODE.MULT requires at least 1 argument".into());
    }
    let vals: Vec<i64> = collect_all(args, cells)?
        .into_iter()
        .filter_map(|v| match v {
            Variant::Integer(n) => Some(n),
            Variant::Float(f) if f.fract() == 0.0 => Some(f as i64),
            _ => None,
        })
        .collect();
    if vals.is_empty() {
        return Err("MODE.MULT: no numeric values".into());
    }
    let mut freq: std::collections::HashMap<i64, usize> = std::collections::HashMap::new();
    for &v in &vals {
        *freq.entry(v).or_insert(0) += 1;
    }
    let max_freq = *freq.values().max().unwrap();
    let mut modes: Vec<i64> = freq
        .into_iter()
        .filter_map(|(value, count)| (count == max_freq).then_some(value))
        .collect();
    modes.sort_unstable();
    if spill {
        Ok(Variant::Array(
            modes.into_iter().map(Variant::Integer).collect(),
        ))
    } else {
        Ok(Variant::Integer(modes[0]))
    }
}

fn func_frequency(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("FREQUENCY requires 2 arguments".into());
    }
    let data = collect_values(&args[0], cells)?
        .iter()
        .filter_map(as_f64)
        .collect::<Vec<_>>();
    let bins = collect_values(&args[1], cells)?
        .iter()
        .filter_map(as_f64)
        .collect::<Vec<_>>();
    if bins.is_empty() {
        return Ok(Variant::Array(vec![Variant::Integer(data.len() as i64)]));
    }
    let mut counts = vec![0_i64; bins.len() + 1];
    for value in data {
        let bucket = bins
            .iter()
            .position(|bin| value <= *bin)
            .unwrap_or(bins.len());
        counts[bucket] += 1;
    }
    Ok(Variant::Array(
        counts.into_iter().map(Variant::Integer).collect(),
    ))
}

fn func_prob(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(3..=4).contains(&args.len()) {
        return Err("PROB requires 3 or 4 arguments".into());
    }
    let values = collect_values(&args[0], cells)?
        .into_iter()
        .filter_map(|value| as_f64(&value))
        .collect::<Vec<_>>();
    let probabilities = collect_values(&args[1], cells)?
        .into_iter()
        .filter_map(|value| as_f64(&value))
        .collect::<Vec<_>>();
    if values.is_empty() || values.len() != probabilities.len() {
        return Err("PROB: value and probability arrays must have equal non-zero length".into());
    }
    if probabilities
        .iter()
        .any(|probability| !probability.is_finite() || *probability < 0.0)
        || probabilities.iter().sum::<f64>() > 1.0 + 1e-12
    {
        return Err("PROB: probabilities must be non-negative and sum to at most 1".into());
    }
    let lower = to_float(&evaluate(&args[2], cells)?)?;
    let upper = if args.len() == 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        lower
    };
    if !lower.is_finite() || !upper.is_finite() || lower > upper {
        return Err("PROB: invalid lower or upper limit".into());
    }
    Ok(as_integer_if_whole(
        values
            .iter()
            .zip(probabilities)
            .filter(|(value, _)| **value >= lower && **value <= upper)
            .map(|(_, probability)| probability)
            .sum(),
    ))
}

fn func_product(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("PRODUCT requires at least 1 argument".into());
    }
    // Excel counts numeric text and logical values supplied directly to
    // PRODUCT, but ignores those same non-numeric types in a reference.
    let values = match collect_direct_numeric_args(args, cells, true) {
        Ok(values) => values,
        Err(error) => return Ok(Variant::Error(error)),
    };
    let product: f64 = values.into_iter().fold(1.0, |acc, x| acc * x);
    Ok(as_integer_if_whole(product))
}

fn func_rank(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("RANK requires 2 or 3 arguments".into());
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let vals: Vec<f64> = collect_values(&args[1], cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    let asc = if args.len() == 3 {
        to_float(&evaluate(&args[2], cells)?)? != 0.0
    } else {
        false
    };
    let rank = if asc {
        vals.iter().filter(|&&v| v < num).count() + 1
    } else {
        vals.iter().filter(|&&v| v > num).count() + 1
    };
    Ok(Variant::Integer(rank as i64))
}

fn func_rank_avg(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("RANK.AVG requires 2 or 3 arguments".into());
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let vals: Vec<f64> = collect_values(&args[1], cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    if vals.is_empty() {
        return Err("RANK.AVG: no numeric values".into());
    }
    let asc = if args.len() == 3 {
        to_float(&evaluate(&args[2], cells)?)? != 0.0
    } else {
        false
    };
    let before = if asc {
        vals.iter().filter(|&&value| value < num).count()
    } else {
        vals.iter().filter(|&&value| value > num).count()
    };
    let ties = vals.iter().filter(|&&value| value == num).count();
    if ties == 0 {
        return Err("RANK.AVG: number is not present in array".into());
    }
    Ok(as_integer_if_whole(
        before as f64 + (ties as f64 + 1.0) / 2.0,
    ))
}

// ── Conditional ───────────────────────────────────────────────────────────────

fn func_ifs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || !args.len().is_multiple_of(2) {
        return Err("IFS requires an even number of arguments".into());
    }
    // Keep the scalar path lazy: Excel callers commonly use IFS to guard a
    // branch that would otherwise produce an error.  Only materialize all
    // branches when at least one condition is a dynamic array.
    let mut conditions = Vec::with_capacity(args.len() / 2);
    let mut array_len = None;
    for condition in args.iter().step_by(2) {
        let value = evaluate(condition, cells)?;
        let length = match &value {
            Variant::Array(values) => Some(values.len()),
            Variant::VbaArray(array) => Some(array.elements.len()),
            _ => None,
        };
        if let Some(length) = length {
            array_len = Some(array_len.unwrap_or(0).max(length));
        }
        conditions.push(value);
    }
    let Some(array_len) = array_len else {
        let mut i = 0;
        while i + 1 < args.len() {
            if is_truthy(&conditions[i / 2]) {
                return evaluate(&args[i + 1], cells);
            }
            i += 2;
        }
        return Err("IFS: no condition matched".into());
    };

    let as_values = |value: Variant| match value {
        Variant::Array(values) => values,
        Variant::VbaArray(array) => array.elements,
        scalar => vec![scalar],
    };
    let branches: Vec<Vec<Variant>> = args
        .iter()
        .skip(1)
        .step_by(2)
        .map(|branch| evaluate(branch, cells).map(as_values))
        .collect::<Result<_, _>>()?;
    let value_at = |values: &[Variant], index: usize| -> Result<Variant, ExcelError> {
        if values.len() == 1 {
            Ok(values[0].clone())
        } else if values.len() == array_len {
            Ok(values[index].clone())
        } else {
            Err(ExcelError::Value)
        }
    };
    let mut result = Vec::with_capacity(array_len);
    for index in 0..array_len {
        let mut matched = false;
        let mut condition_error = None;
        for (condition, branch) in conditions.iter().zip(branches.iter()) {
            let condition_value = match condition {
                Variant::Array(values) => value_at(values, index),
                Variant::VbaArray(array) => value_at(&array.elements, index),
                scalar => Ok(scalar.clone()),
            };
            let condition_value = match condition_value {
                Ok(value) => value,
                Err(error) => return Ok(Variant::Error(error)),
            };
            if let Variant::Error(error) = condition_value {
                condition_error = Some(error);
                continue;
            }
            if is_truthy(&condition_value) {
                match value_at(branch, index) {
                    Ok(value) => result.push(value),
                    Err(error) => return Ok(Variant::Error(error)),
                }
                matched = true;
                break;
            }
        }
        if !matched {
            result.push(Variant::Error(condition_error.unwrap_or(ExcelError::NA)));
        }
    }
    Ok(Variant::Array(result))
}

// ── Date helpers ──────────────────────────────────────────────────────────────
// is_leap/days_in_month/serial_to_ymd moved to elixcee-types (Phase 2A),
// imported near the top of this file; serial_to_ymd_pub (a pure wrapper
// around serial_to_ymd) is dropped now that serial_to_ymd is itself pub.
// date_to_serial and serial_to_display below are eval.rs-local and stay —
// serial_to_display in particular is a separate, pre-existing private
// duplicate of elixcee-types::serial_to_display (same computation, kept
// as-is rather than deleted, since removing it is a behavior-affecting
// edit beyond what a pure extraction covers).

/// Excel serial date: Jan 1 1900 = 1. Includes the Excel leap-year bug (Feb 29 1900 = 60).
fn date_to_serial(y: i32, m: u32, d: u32) -> i64 {
    // Count days from Jan 1 1900 (serial 1)
    let mut serial: i64 = 1;
    for yr in 1900..y {
        serial += if is_leap(yr) { 366 } else { 365 };
    }
    for mo in 1..m {
        serial += days_in_month(y, mo) as i64;
    }
    serial += d as i64 - 1;
    // Excel pretends Feb 29 1900 exists (serial 60), so dates after Feb 28 1900 are +1
    if y > 1900 || (y == 1900 && (m > 2 || (m == 2 && d == 29))) {
        serial += 1;
    }
    serial
}

fn serial_to_display(s: i64) -> String {
    let (y, m, d) = serial_to_ymd(s);
    format!("{:04}-{:02}-{:02}", y, m, d)
}

fn today_serial() -> i64 {
    let unix_days = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs()
        / 86400;
    unix_days as i64 + 25569 // 25569 = Excel serial of 1970-01-01
}

/// Day of week from Excel serial: 0=Sun,1=Mon,...,6=Sat (like Excel WEEKDAY default)
fn serial_weekday(serial: i64) -> u32 {
    // Serial 1 (Jan 1 1900) was a Monday. (serial - 1 + 1) % 7 gives 0=Mon..6=Sun
    // Shift to 0=Sun: ((serial - 1 + 1) % 7 + 1) % 7
    ((serial % 7 + 6) % 7) as u32 // 0=Sun,1=Mon,...,5=Fri,6=Sat
}

fn integer_mode_arg(value: f64, function_name: &str) -> Result<u32, String> {
    if !value.is_finite() || value.fract() != 0.0 || value < 0.0 || value > u32::MAX as f64 {
        return Err(format!("{function_name}: return_type must be an integer"));
    }
    Ok(value as u32)
}

fn normalize_date_components(
    year_value: f64,
    month_value: f64,
    day_value: f64,
) -> Result<Variant, String> {
    if !year_value.is_finite() || !month_value.is_finite() || !day_value.is_finite() {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let mut year = year_value.trunc() as i64;
    // Excel interprets 0..99 as 1900..1999 for DATE arguments.
    if (0..=1899).contains(&year) {
        year += 1900;
    }
    if !(1900..=9999).contains(&year) {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let month_index = month_value.trunc() as i64 - 1;
    let normalized_year = year + month_index.div_euclid(12);
    if !(1900..=9999).contains(&normalized_year) {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let normalized_month = month_index.rem_euclid(12) as u32 + 1;
    let first_day = date_to_serial(normalized_year as i32, normalized_month, 1);
    let serial = first_day
        .checked_add(day_value.trunc() as i64 - 1)
        .ok_or_else(|| "DATE: serial overflow".to_string())?;
    if serial < 1 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Date(serial))
}

fn english_month_number(token: &str) -> Option<f64> {
    match token.trim().to_ascii_lowercase().as_str() {
        "jan" | "january" => Some(1.0),
        "feb" | "february" => Some(2.0),
        "mar" | "march" => Some(3.0),
        "apr" | "april" => Some(4.0),
        "may" => Some(5.0),
        "jun" | "june" => Some(6.0),
        "jul" | "july" => Some(7.0),
        "aug" | "august" => Some(8.0),
        "sep" | "sept" | "september" => Some(9.0),
        "oct" | "october" => Some(10.0),
        "nov" | "november" => Some(11.0),
        "dec" | "december" => Some(12.0),
        _ => None,
    }
}

fn func_date(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("DATE requires 3 arguments".into());
    }
    let year_value = to_float(&evaluate(&args[0], cells)?)?;
    let month_value = to_float(&evaluate(&args[1], cells)?)?;
    let day_value = to_float(&evaluate(&args[2], cells)?)?;
    normalize_date_components(year_value, month_value, day_value)
}

fn func_today(
    args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !args.is_empty() {
        return Err("TODAY takes no arguments".into());
    }
    Ok(Variant::Date(today_serial()))
}

fn func_eomonth(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("EOMONTH requires 2 arguments".into());
    }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let months = to_float(&evaluate(&args[1], cells)?)? as i32;
    let (mut y, mut m, _) = serial_to_ymd(start);
    let total = (m as i32 - 1) + months;
    let offset_y = total.div_euclid(12);
    let nm = (total.rem_euclid(12) + 1) as u32;
    y += offset_y;
    m = nm;
    let last_day = days_in_month(y, m);
    Ok(Variant::Date(date_to_serial(y, m, last_day)))
}

fn func_networkdays(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("NETWORKDAYS requires 2 or 3 arguments".into());
    }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let end = to_float(&evaluate(&args[1], cells)?)? as i64;
    let holidays: std::collections::HashSet<i64> = if args.len() == 3 {
        collect_values(&args[2], cells)?
            .iter()
            .filter_map(|v| to_float(v).ok().map(|f| f as i64))
            .collect()
    } else {
        std::collections::HashSet::new()
    };
    let (lo, hi, sign) = if start <= end {
        (start, end, 1i64)
    } else {
        (end, start, -1)
    };
    let count: i64 = (lo..=hi)
        .filter(|&d| {
            let wd = serial_weekday(d);
            wd != 0 && wd != 6 && !holidays.contains(&d) // 0=Sun, 6=Sat
        })
        .count() as i64;
    Ok(Variant::Integer(count * sign))
}

// ── ROW ───────────────────────────────────────────────────────────────────────

fn func_row(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    match args.first() {
        None => Ok(Variant::Integer(1)),
        Some(FormulaExpr::CellRef { row, .. }) => Ok(Variant::Integer(*row as i64)),
        Some(FormulaExpr::Range { r1, .. }) => Ok(Variant::Integer(*r1 as i64)),
        Some(other) => {
            evaluate(other, cells)?;
            Ok(Variant::Integer(1))
        }
    }
}

fn func_rows(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ROWS requires 1 argument".into());
    }
    let rows = formula_array_dimensions(&args[0], cells)?.map_or(1, |(rows, _)| rows);
    Ok(Variant::Integer(rows as i64))
}

/// Return the known two-dimensional shape of a reference or a bounded array
/// constructor. Flat `Variant::Array` values intentionally remain unknown:
/// their producer may have lost shape metadata, so guessing would make
/// ROWS/COLUMNS silently report a false dimension.
fn formula_array_dimensions(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Option<(u32, u32)>, String> {
    match expr {
        FormulaExpr::Range { r1, r2, c1, c2, .. } => Ok(Some((
            r2.saturating_sub(*r1).saturating_add(1),
            c2.saturating_sub(*c1).saturating_add(1),
        ))),
        FormulaExpr::CellRef { .. } => Ok(Some((1, 1))),
        FormulaExpr::FuncCall { name, args } if name.eq_ignore_ascii_case("SEQUENCE") => {
            if args.is_empty() || args.len() > 4 {
                return Err("SEQUENCE requires 1 to 4 arguments".into());
            }
            let dimension = |arg: &FormulaExpr| -> Result<u32, String> {
                let value = to_float(&evaluate(arg, cells)?)?;
                if !value.is_finite() || value.fract() != 0.0 || value <= 0.0 {
                    return Err("SEQUENCE dimensions must be positive integers".into());
                }
                if value > u32::MAX as f64 {
                    return Err("SEQUENCE dimension is too large".into());
                }
                Ok(value as u32)
            };
            let rows = dimension(&args[0])?;
            let columns = args.get(1).map(dimension).transpose()?.unwrap_or(1);
            Ok(Some((rows, columns)))
        }
        FormulaExpr::FuncCall { name, args }
            if name.eq_ignore_ascii_case("TRANSPOSE") && args.len() == 1 =>
        {
            Ok(formula_array_dimensions(&args[0], cells)?.map(|(rows, columns)| (columns, rows)))
        }
        _ => Ok(None),
    }
}

// ── XLOOKUP ───────────────────────────────────────────────────────────────────

pub(crate) fn xlookup_binary_index(
    lookup: &[Variant],
    key: &Variant,
    ascending: bool,
    match_mode: i32,
) -> Result<Option<usize>, String> {
    if lookup.is_empty() {
        return Ok(None);
    }
    let ordering = |value: &Variant| -> Result<Ordering, String> {
        let cmp = variant_cmp(value, key)?;
        Ok(if ascending { cmp } else { cmp.reverse() })
    };

    // Lower bound in the requested sort order: the first value >= key.
    let (mut lo, mut hi) = (0usize, lookup.len());
    while lo < hi {
        let mid = lo + (hi - lo) / 2;
        if matches!(ordering(&lookup[mid])?, Ordering::Less) {
            lo = mid + 1;
        } else {
            hi = mid;
        }
    }
    let lower = lo;
    let is_exact = lower < lookup.len() && variant_eq(&lookup[lower], key);
    let index = match match_mode {
        0 => is_exact.then_some(lower),
        -1 => {
            if is_exact {
                Some(lower)
            } else if ascending {
                lower.checked_sub(1)
            } else {
                // In descending order, the first value at the lower bound
                // is the next smaller numeric value.
                Some(lower).filter(|&i| i < lookup.len())
            }
        }
        1 => {
            if is_exact {
                Some(lower)
            } else if ascending {
                Some(lower).filter(|&i| i < lookup.len())
            } else {
                lower.checked_sub(1)
            }
        }
        mode => return Err(format!("XLOOKUP: unsupported match_mode {}", mode)),
    };
    Ok(index)
}

fn validate_lookup_sorted(lookup: &[Variant], ascending: bool) -> Result<bool, String> {
    for pair in lookup.windows(2) {
        let order = variant_cmp(&pair[0], &pair[1])?;
        if (ascending && order == Ordering::Greater) || (!ascending && order == Ordering::Less) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn func_xlookup(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 6 {
        return Err("XLOOKUP requires 3 to 6 arguments".into());
    }
    let key = evaluate(&args[0], cells)?;
    let lookup = collect_values(&args[1], cells)?;
    let return_arr = collect_values(&args[2], cells)?;
    if return_arr.len() != lookup.len() {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let not_found = if args.len() >= 4 {
        Some(evaluate(&args[3], cells)?)
    } else {
        None
    };
    let match_mode = if args.len() >= 5 {
        match lookup_mode(&evaluate(&args[4], cells)?) {
            Ok(mode) => mode,
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        }
    } else {
        0
    };
    let search_mode = if args.len() >= 6 {
        match lookup_mode(&evaluate(&args[5], cells)?) {
            Ok(mode) => mode,
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        }
    } else {
        1
    };

    if matches!(search_mode, 2 | -2) {
        if match_mode == 2 {
            return Err("XLOOKUP: wildcard match_mode is incompatible with binary search".into());
        }
        match validate_lookup_sorted(&lookup, search_mode == 2) {
            Ok(true) => {}
            Ok(false) => return Ok(Variant::Error(ExcelError::NA)),
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        }
        let index = xlookup_binary_index(&lookup, &key, search_mode == 2, match_mode)?;
        return Ok(index
            .and_then(|i| return_arr.get(i).cloned())
            .unwrap_or_else(|| not_found.unwrap_or(Variant::Error(ExcelError::NA))));
    }

    let iter: Box<dyn Iterator<Item = usize>> = match search_mode {
        -1 => Box::new((0..lookup.len()).rev()),
        1 => Box::new(0..lookup.len()),
        mode => return Err(format!("XLOOKUP: unsupported search_mode {}", mode)),
    };

    match match_mode {
        0 => {
            for i in iter {
                if variant_eq(&lookup[i], &key) {
                    return Ok(return_arr.get(i).cloned().unwrap_or(Variant::Empty));
                }
            }
        }
        -1 => {
            // exact or next smaller
            let mut best: Option<(usize, f64)> = None;
            let key_f = to_float(&key)?;
            for (i, l) in lookup.iter().enumerate() {
                if let Ok(v) = to_float(l)
                    && v <= key_f
                    && best.is_none_or(|(_, bv)| v > bv)
                {
                    best = Some((i, v));
                }
            }
            if let Some((i, _)) = best {
                return Ok(return_arr.get(i).cloned().unwrap_or(Variant::Empty));
            }
        }
        1 => {
            // exact or next larger
            let mut best: Option<(usize, f64)> = None;
            let key_f = to_float(&key)?;
            for (i, l) in lookup.iter().enumerate() {
                if let Ok(v) = to_float(l)
                    && v >= key_f
                    && best.is_none_or(|(_, bv)| v < bv)
                {
                    best = Some((i, v));
                }
            }
            if let Some((i, _)) = best {
                return Ok(return_arr.get(i).cloned().unwrap_or(Variant::Empty));
            }
        }
        2 => {
            let pattern = to_str(&key);
            for i in iter {
                if wildcard_match(&to_str(&lookup[i]), &pattern) {
                    return Ok(return_arr.get(i).cloned().unwrap_or(Variant::Empty));
                }
            }
        }
        m => return Err(format!("XLOOKUP: unsupported match_mode {}", m)),
    }

    Ok(not_found.unwrap_or(Variant::Error(ExcelError::NA)))
}

// ── SUBTOTAL ──────────────────────────────────────────────────────────────────

fn is_nested_aggregate_formula(formula: &str) -> bool {
    let normalized: String = formula
        .chars()
        .filter(|character| !character.is_whitespace())
        .flat_map(char::to_uppercase)
        .collect();
    normalized.starts_with("=SUBTOTAL(") || normalized.starts_with("=AGGREGATE(")
}

/// Collect AGGREGATE inputs while preserving worksheet errors and nested
/// aggregate provenance. Ordinary collection helpers intentionally flatten
/// values for most functions; AGGREGATE needs the extra information because
/// its options select whether errors and nested SUBTOTAL/AGGREGATE results are
/// ignored.
fn collect_aggregate_values(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    ignore_nested: bool,
) -> Result<Vec<Variant>, String> {
    let mut values = Vec::new();
    for expr in args {
        match expr {
            FormulaExpr::Range { c1, r1, c2, r2, .. } => {
                let (rmin, rmax) = (r1.min(r2), r1.max(r2));
                let (cmin, cmax) = (c1.min(c2), c1.max(c2));
                let rows = (rmax - rmin + 1) as u64;
                let cols = (cmax - cmin + 1) as u64;
                if rows * cols > 1_000_000 {
                    return Err(format!(
                        "Range too large ({} cells); maximum is 1,000,000",
                        rows * cols
                    ));
                }
                for row in *rmin..=*rmax {
                    for col in *cmin..=*cmax {
                        if ignore_nested
                            && cells
                                .get(&(row, col))
                                .and_then(|cell| cell.formula.as_deref())
                                .is_some_and(is_nested_aggregate_formula)
                        {
                            continue;
                        }
                        values.push(cell_val(cells, row, col));
                    }
                }
            }
            FormulaExpr::FuncCall { name, .. }
                if ignore_nested
                    && (name.eq_ignore_ascii_case("SUBTOTAL")
                        || name.eq_ignore_ascii_case("AGGREGATE")) => {}
            other => values.extend(collect_values(other, cells)?),
        }
    }
    Ok(values)
}

fn aggregate_option(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<u32, String> {
    let value = to_float(&evaluate(expr, cells)?)?;
    if !value.is_finite() || value.fract() != 0.0 || !(0.0..=7.0).contains(&value) {
        return Err("AGGREGATE: options must be an integer from 0 to 7".into());
    }
    Ok(value as u32)
}

fn func_subtotal(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("SUBTOTAL requires at least 2 arguments".into());
    }
    let fn_num = to_float(&evaluate(&args[0], cells)?)? as u32;
    let rest = &args[1..];
    let values = collect_aggregate_values(rest, cells, true)?;
    if let Some(error) = values.iter().find_map(|value| match value {
        Variant::Error(error) => Some(error.clone()),
        _ => None,
    }) {
        return Ok(Variant::Error(error));
    }
    let nums: Vec<f64> = values.iter().filter_map(as_f64).collect();
    // 101-111 = ignore hidden rows (same behavior here since no hidden rows)
    match fn_num % 100 {
        1 => {
            if nums.is_empty() {
                return Err("SUBTOTAL: no values".into());
            }
            Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
        }
        2 => Ok(Variant::Integer(nums.len() as i64)),
        3 => Ok(Variant::Integer(
            values
                .iter()
                .filter(|value| !matches!(value, Variant::Empty))
                .count() as i64,
        )),
        4 => nums
            .iter()
            .copied()
            .reduce(f64::max)
            .map(as_integer_if_whole)
            .ok_or_else(|| "SUBTOTAL: no values".into()),
        5 => nums
            .iter()
            .copied()
            .reduce(f64::min)
            .map(as_integer_if_whole)
            .ok_or_else(|| "SUBTOTAL: no values".into()),
        6 => Ok(as_integer_if_whole(nums.iter().product::<f64>())),
        7 => {
            if nums.len() < 2 {
                return Err("SUBTOTAL: at least 2 numeric values required".into());
            }
            let mean = nums.iter().sum::<f64>() / nums.len() as f64;
            Ok(Variant::Float(
                (nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>()
                    / (nums.len() - 1) as f64)
                    .sqrt(),
            ))
        }
        8 => {
            if nums.is_empty() {
                return Err("SUBTOTAL: at least 1 numeric value required".into());
            }
            let mean = nums.iter().sum::<f64>() / nums.len() as f64;
            Ok(Variant::Float(
                (nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>() / nums.len() as f64)
                    .sqrt(),
            ))
        }
        9 => Ok(as_integer_if_whole(nums.iter().sum::<f64>())),
        10 => {
            if nums.len() < 2 {
                return Err("SUBTOTAL: at least 2 numeric values required".into());
            }
            let mean = nums.iter().sum::<f64>() / nums.len() as f64;
            Ok(Variant::Float(
                nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>()
                    / (nums.len() - 1) as f64,
            ))
        }
        11 => {
            if nums.is_empty() {
                return Err("SUBTOTAL: at least 1 numeric value required".into());
            }
            let mean = nums.iter().sum::<f64>() / nums.len() as f64;
            Ok(Variant::Float(
                nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>() / nums.len() as f64,
            ))
        }
        n => Err(format!("SUBTOTAL: unsupported function_num {}", n)),
    }
}

// ── Additional Numerical ─────────────────────────────────────────────────────

fn func_averageif(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("AVERAGEIF requires 2 or 3 arguments".into());
    }
    let range_vals = flatten_values(collect_values(&args[0], cells)?);
    let pcrit = parse_criteria(&evaluate(&args[1], cells)?);
    let avg_vals = if args.len() == 3 {
        flatten_values(collect_values(&args[2], cells)?)
    } else {
        range_vals.clone()
    };
    let mut nums = Vec::new();
    for (rv, av) in range_vals.iter().zip(avg_vals.iter()) {
        if matches_parsed(rv, &pcrit) {
            if let Variant::Error(error) = av {
                return Ok(Variant::Error(error.clone()));
            }
            if let Ok(value) = to_float(av) {
                nums.push(value);
            }
        }
    }
    if nums.is_empty() {
        return Err("AVERAGEIF: no matching values".into());
    }
    Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
}

fn func_averageifs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return Err("AVERAGEIFS requires avg_range then pairs".into());
    }
    let avg_vals = flatten_values(collect_values(&args[0], cells)?);
    let avg_shape = expression_shape(&args[0], cells)?;
    let n = avg_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        if ensure_same_shape(avg_shape, &args[i], cells, "AVERAGEIFS").is_err() {
            return Ok(Variant::Error(ExcelError::Value));
        }
        let range_vals = flatten_values(collect_values(&args[i], cells)?);
        let pcrit = parse_criteria(&evaluate(&args[i + 1], cells)?);
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_parsed(rv, &pcrit) {
                mask[j] = false;
            }
        }
        i += 2;
    }
    let mut nums = Vec::new();
    for (index, value) in avg_vals.iter().enumerate() {
        if mask[index] {
            if let Variant::Error(error) = value {
                return Ok(Variant::Error(error.clone()));
            }
            if let Ok(number) = to_float(value) {
                nums.push(number);
            }
        }
    }
    if nums.is_empty() {
        return Err("AVERAGEIFS: no matching values".into());
    }
    Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
}

fn func_int(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("INT requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    Ok(as_integer_if_whole(n.floor()))
}

fn func_large(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("LARGE requires 2 arguments".into());
    }
    let mut nums: Vec<f64> = collect_values(&args[0], cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    let k = to_float(&evaluate(&args[1], cells)?)? as usize;
    if k == 0 || k > nums.len() {
        return Err("LARGE: k out of range".into());
    }
    nums.sort_by(|a, b| b.partial_cmp(a).unwrap_or(Ordering::Equal));
    Ok(as_integer_if_whole(nums[k - 1]))
}

fn func_small(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("SMALL requires 2 arguments".into());
    }
    let mut nums: Vec<f64> = collect_values(&args[0], cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    let k = to_float(&evaluate(&args[1], cells)?)? as usize;
    if k == 0 || k > nums.len() {
        return Err("SMALL: k out of range".into());
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    Ok(as_integer_if_whole(nums[k - 1]))
}

fn func_maxifs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return Err("MAXIFS requires max_range then pairs".into());
    }
    let max_vals = flatten_values(collect_values(&args[0], cells)?);
    let max_shape = expression_shape(&args[0], cells)?;
    let n = max_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        if ensure_same_shape(max_shape, &args[i], cells, "MAXIFS").is_err() {
            return Ok(Variant::Error(ExcelError::Value));
        }
        let range_vals = flatten_values(collect_values(&args[i], cells)?);
        let pcrit = parse_criteria(&evaluate(&args[i + 1], cells)?);
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_parsed(rv, &pcrit) {
                mask[j] = false;
            }
        }
        i += 2;
    }
    let mut max = None;
    for (index, value) in max_vals.iter().enumerate() {
        if mask[index] {
            if let Variant::Error(error) = value {
                return Ok(Variant::Error(error.clone()));
            }
            if let Some(number) = as_f64(value) {
                max = Some(max.map_or(number, |current: f64| current.max(number)));
            }
        }
    }
    max.map(as_integer_if_whole)
        .ok_or_else(|| "MAXIFS: no matching values".into())
}

fn func_minifs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return Err("MINIFS requires min_range then pairs".into());
    }
    let min_vals = flatten_values(collect_values(&args[0], cells)?);
    let min_shape = expression_shape(&args[0], cells)?;
    let n = min_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        if ensure_same_shape(min_shape, &args[i], cells, "MINIFS").is_err() {
            return Ok(Variant::Error(ExcelError::Value));
        }
        let range_vals = flatten_values(collect_values(&args[i], cells)?);
        let pcrit = parse_criteria(&evaluate(&args[i + 1], cells)?);
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_parsed(rv, &pcrit) {
                mask[j] = false;
            }
        }
        i += 2;
    }
    let mut min = None;
    for (index, value) in min_vals.iter().enumerate() {
        if mask[index] {
            if let Variant::Error(error) = value {
                return Ok(Variant::Error(error.clone()));
            }
            if let Some(number) = as_f64(value) {
                min = Some(min.map_or(number, |current: f64| current.min(number)));
            }
        }
    }
    min.map(as_integer_if_whole)
        .ok_or_else(|| "MINIFS: no matching values".into())
}

fn func_mod(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("MOD requires 2 arguments".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    let d = to_float(&evaluate(&args[1], cells)?)?;
    if d == 0.0 {
        return Err("MOD: division by zero".into());
    }
    Ok(as_integer_if_whole(n - d * (n / d).floor()))
}

fn func_percentile(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("PERCENTILE requires 2 arguments".into());
    }
    let mut nums: Vec<f64> = collect_values(&args[0], cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    let k = to_float(&evaluate(&args[1], cells)?)?;
    if !(0.0..=1.0).contains(&k) {
        return Err("PERCENTILE: k must be 0 to 1".into());
    }
    if nums.is_empty() {
        return Err("PERCENTILE: no numeric values".into());
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    let pos = k * (nums.len() - 1) as f64;
    let lo = pos.floor() as usize;
    let hi = pos.ceil() as usize;
    let result = if lo == hi {
        nums[lo]
    } else {
        nums[lo] + (pos - lo as f64) * (nums[hi] - nums[lo])
    };
    Ok(as_integer_if_whole(result))
}

fn func_percentrank(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("PERCENTRANK requires 2 or 3 arguments".into());
    }
    let nums: Vec<f64> = collect_values(&args[0], cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    let x = to_float(&evaluate(&args[1], cells)?)?;
    let sig = if args.len() == 3 {
        to_float(&evaluate(&args[2], cells)?)? as usize
    } else {
        3
    };
    if nums.is_empty() {
        return Err("PERCENTRANK: no values".into());
    }
    let below = nums.iter().filter(|&&v| v < x).count();
    let equal = nums.iter().filter(|&&v| v == x).count();
    if equal == 0 {
        return Err("PERCENTRANK: value not in array".into());
    }
    let rank = below as f64 / (nums.len() - 1) as f64;
    let mult = 10_f64.powi(sig as i32);
    Ok(Variant::Float((rank * mult).floor() / mult))
}

fn percentile_value(mut nums: Vec<f64>, k: f64, exclusive: bool) -> Result<Variant, String> {
    if nums.is_empty() || !k.is_finite() {
        return Err("PERCENTILE: no numeric values or invalid k".into());
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    let n = nums.len() as f64;
    let pos = if exclusive {
        k * (n + 1.0)
    } else {
        k * (n - 1.0) + 1.0
    };
    if pos < 1.0 || pos > n {
        return Err("PERCENTILE: k is out of range".into());
    }
    let index = pos - 1.0;
    let lo = index.floor() as usize;
    let hi = index.ceil() as usize;
    let value = if lo == hi {
        nums[lo]
    } else {
        nums[lo] + (index - lo as f64) * (nums[hi] - nums[lo])
    };
    Ok(as_integer_if_whole(value))
}

fn func_percentile_exc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("PERCENTILE.EXC requires 2 arguments".into());
    }
    let nums = collect_values(&args[0], cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    let k = to_float(&evaluate(&args[1], cells)?)?;
    percentile_value(nums, k, true)
}

fn func_percentrank_exc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("PERCENTRANK.EXC requires 2 or 3 arguments".into());
    }
    let mut nums: Vec<f64> = collect_values(&args[0], cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    let x = to_float(&evaluate(&args[1], cells)?)?;
    let sig = if args.len() == 3 {
        to_float(&evaluate(&args[2], cells)?)? as i32
    } else {
        3
    };
    if nums.is_empty() || !x.is_finite() || sig < 0 {
        return Err("PERCENTRANK.EXC: invalid argument".into());
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    if x < nums[0] || x > nums[nums.len() - 1] {
        return Err("PERCENTRANK.EXC: x is out of range".into());
    }
    let upper = nums.partition_point(|v| *v <= x);
    let rank = if upper == 0 {
        1.0
    } else if upper == nums.len() {
        nums.len() as f64
    } else {
        let hi = nums[upper];
        let lo = nums[upper - 1];
        upper as f64 + if hi == lo { 0.0 } else { (x - lo) / (hi - lo) }
    };
    let value = rank / (nums.len() as f64 + 1.0);
    let mult = 10_f64.powi(sig);
    Ok(Variant::Float((value * mult).round() / mult))
}

fn func_quartile(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    exclusive: bool,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("QUARTILE requires 2 arguments".into());
    }
    let nums = collect_values(&args[0], cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    let quart = to_float(&evaluate(&args[1], cells)?)?;
    if !quart.is_finite() || quart.fract() != 0.0 || !(0.0..=4.0).contains(&quart) {
        return Err("QUARTILE: quart must be an integer from 0 to 4".into());
    }
    percentile_value(nums, quart / 4.0, exclusive)
}

fn numeric_args(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<Vec<f64>, String> {
    let values = collect_all(args, cells)?;
    let mut nums = Vec::with_capacity(values.len());
    for value in values {
        if let Variant::Error(_error) = value {
            return Ok(vec![f64::NAN]);
        }
        if let Some(number) = as_f64(&value) {
            nums.push(number);
        }
    }
    if nums.iter().any(|v| v.is_nan()) {
        return Err(format!("{name}: error argument"));
    }
    if nums.is_empty() {
        return Err(format!("{name}: no numeric values"));
    }
    Ok(nums)
}

fn func_sumsq(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    Ok(as_integer_if_whole(
        numeric_args(args, cells, "SUMSQ")?
            .iter()
            .map(|v| v * v)
            .sum(),
    ))
}

fn func_geomean(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = numeric_args(args, cells, "GEOMEAN")?;
    if nums.iter().any(|v| *v <= 0.0) {
        return Err("GEOMEAN: values must be positive".into());
    }
    Ok(Variant::Float(
        (nums.iter().map(|v| v.ln()).sum::<f64>() / nums.len() as f64).exp(),
    ))
}

fn func_harmean(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = numeric_args(args, cells, "HARMEAN")?;
    if nums.iter().any(|v| *v <= 0.0) {
        return Err("HARMEAN: values must be positive".into());
    }
    Ok(Variant::Float(
        nums.len() as f64 / nums.iter().map(|v| 1.0 / v).sum::<f64>(),
    ))
}

fn func_devsq(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = numeric_args(args, cells, "DEVSQ")?;
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    Ok(as_integer_if_whole(
        nums.iter().map(|v| (v - mean).powi(2)).sum(),
    ))
}

fn func_avedev(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = numeric_args(args, cells, "AVEDEV")?;
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    Ok(Variant::Float(
        nums.iter().map(|v| (v - mean).abs()).sum::<f64>() / nums.len() as f64,
    ))
}

fn func_trimmean(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("TRIMMEAN requires 2 arguments".into());
    }
    let mut nums = numeric_args(&args[..1], cells, "TRIMMEAN")?;
    let percent = to_float(&evaluate(&args[1], cells)?)?;
    if !percent.is_finite() || !(0.0..1.0).contains(&percent) {
        return Err("TRIMMEAN: percent must be in [0, 1)".into());
    }
    let trim_each = (nums.len() as f64 * percent / 2.0).floor() as usize;
    if trim_each * 2 >= nums.len() {
        return Err("TRIMMEAN: trimming removes all values".into());
    }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
    let kept = &nums[trim_each..nums.len() - trim_each];
    Ok(Variant::Float(kept.iter().sum::<f64>() / kept.len() as f64))
}

fn func_skew(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = numeric_args(args, cells, "SKEW")?;
    if nums.len() < 3 {
        return Err("SKEW requires at least 3 values".into());
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let sum_sq = nums.iter().map(|v| (v - mean).powi(2)).sum::<f64>();
    if sum_sq == 0.0 {
        return Err("SKEW: standard deviation is zero".into());
    }
    let s = (sum_sq / (nums.len() - 1) as f64).sqrt();
    let sum_cubed = nums.iter().map(|v| ((v - mean) / s).powi(3)).sum::<f64>();
    Ok(Variant::Float(
        nums.len() as f64 / ((nums.len() - 1) * (nums.len() - 2)) as f64 * sum_cubed,
    ))
}

fn func_skew_p(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = numeric_args(args, cells, "SKEW.P")?;
    if nums.len() < 2 {
        return Err("SKEW.P requires at least 2 values".into());
    }
    let n = nums.len() as f64;
    let mean = nums.iter().sum::<f64>() / n;
    let second = nums.iter().map(|v| (v - mean).powi(2)).sum::<f64>() / n;
    if second == 0.0 {
        return Err("SKEW.P: standard deviation is zero".into());
    }
    let third = nums.iter().map(|v| (v - mean).powi(3)).sum::<f64>() / n;
    Ok(Variant::Float(third / second.powf(1.5)))
}

fn func_kurt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = numeric_args(args, cells, "KURT")?;
    if nums.len() < 4 {
        return Err("KURT requires at least 4 values".into());
    }
    let n = nums.len() as f64;
    let mean = nums.iter().sum::<f64>() / n;
    let sum_sq = nums.iter().map(|v| (v - mean).powi(2)).sum::<f64>();
    if sum_sq == 0.0 {
        return Err("KURT: standard deviation is zero".into());
    }
    let s = (sum_sq / (n - 1.0)).sqrt();
    let sum_fourth = nums.iter().map(|v| ((v - mean) / s).powi(4)).sum::<f64>();
    let result = n * (n + 1.0) / ((n - 1.0) * (n - 2.0) * (n - 3.0)) * sum_fourth
        - 3.0 * (n - 1.0).powi(2) / ((n - 2.0) * (n - 3.0));
    Ok(Variant::Float(result))
}

fn pseudo_rand() -> f64 {
    let nanos = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .subsec_nanos() as u64;
    let x = nanos
        .wrapping_mul(6364136223846793005)
        .wrapping_add(1442695040888963407);
    (x >> 11) as f64 / ((1u64 << 53) as f64)
}

fn func_rand(
    args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !args.is_empty() {
        return Err("RAND takes no arguments".into());
    }
    Ok(Variant::Float(pseudo_rand()))
}

fn func_randbetween(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("RANDBETWEEN requires 2 arguments".into());
    }
    let lo = to_float(&evaluate(&args[0], cells)?)? as i64;
    let hi = to_float(&evaluate(&args[1], cells)?)? as i64;
    if lo > hi {
        return Err("RANDBETWEEN: bottom > top".into());
    }
    let range = (hi - lo + 1) as f64;
    Ok(Variant::Integer(lo + (pseudo_rand() * range) as i64))
}

fn func_sumproduct(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("SUMPRODUCT requires at least 1 argument".into());
    }
    let arrays: Vec<Vec<f64>> = args
        .iter()
        .map(|a| {
            collect_values(a, cells)
                .map(|vals| vals.iter().map(|v| to_float(v).unwrap_or(0.0)).collect())
        })
        .collect::<Result<_, _>>()?;
    let len = arrays[0].len();
    let sum: f64 = (0..len)
        .map(|i| {
            arrays
                .iter()
                .map(|arr| arr.get(i).copied().unwrap_or(0.0))
                .product::<f64>()
        })
        .sum();
    Ok(as_integer_if_whole(sum))
}

fn func_trunc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("TRUNC requires 1 or 2 arguments".into());
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let digits = if args.len() == 2 {
        to_float(&evaluate(&args[1], cells)?)? as i32
    } else {
        0
    };
    let mult = 10_f64.powi(digits);
    Ok(as_integer_if_whole(
        num.signum() * (num.abs() * mult).floor() / mult,
    ))
}

fn func_aggregate(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 {
        return Err("AGGREGATE requires at least 3 arguments".into());
    }
    let fn_num = to_float(&evaluate(&args[0], cells)?)? as u32;
    let options = match aggregate_option(&args[1], cells) {
        Ok(options) => options,
        Err(_) => return Ok(Variant::Error(ExcelError::Value)),
    };
    let rest = &args[2..];
    // Options 0-3 ignore nested SUBTOTAL/AGGREGATE; options 2, 3, 6, 7
    // ignore worksheet errors. Hidden-row options are accepted, but the core
    // cell model has no hidden-row metadata, so they currently behave like
    // their visible-row counterparts.
    let ignore_nested = options < 4;
    let ignore_errors = matches!(options, 2 | 3 | 6 | 7);
    let mut values = collect_aggregate_values(rest, cells, ignore_nested)?;
    if !ignore_errors
        && let Some(error) = values.iter().find_map(|value| match value {
            Variant::Error(error) => Some(error.clone()),
            _ => None,
        })
    {
        return Ok(Variant::Error(error));
    }
    if ignore_errors {
        values.retain(|value| !matches!(value, Variant::Error(_)));
    }
    let nums: Vec<f64> = values.iter().filter_map(as_f64).collect();
    let aggregate_nums = |expr: &FormulaExpr| -> Result<Vec<f64>, String> {
        Ok(
            collect_aggregate_values(std::slice::from_ref(expr), cells, ignore_nested)?
                .into_iter()
                .filter_map(|value| match value {
                    Variant::Error(_) if ignore_errors => None,
                    Variant::Integer(number) => Some(number as f64),
                    Variant::Float(number) => Some(number),
                    _ => None,
                })
                .collect(),
        )
    };
    match fn_num % 100 {
        1 => {
            if nums.is_empty() {
                return Err("AGGREGATE: no values".into());
            }
            Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
        }
        2 => Ok(Variant::Integer(nums.len() as i64)),
        3 => Ok(Variant::Integer(
            values
                .iter()
                .filter(|value| !matches!(value, Variant::Empty))
                .count() as i64,
        )),
        4 => nums
            .iter()
            .copied()
            .reduce(f64::max)
            .map(as_integer_if_whole)
            .ok_or_else(|| "AGGREGATE: no values".into()),
        5 => nums
            .iter()
            .copied()
            .reduce(f64::min)
            .map(as_integer_if_whole)
            .ok_or_else(|| "AGGREGATE: no values".into()),
        6 => Ok(as_integer_if_whole(nums.iter().fold(1.0, |a, &x| a * x))),
        7 => {
            if nums.len() < 2 {
                return Err("AGGREGATE: at least 2 numeric values required".into());
            }
            let mean = nums.iter().sum::<f64>() / nums.len() as f64;
            Ok(Variant::Float(
                (nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>()
                    / (nums.len() - 1) as f64)
                    .sqrt(),
            ))
        }
        8 => {
            if nums.is_empty() {
                return Err("AGGREGATE: at least 1 numeric value required".into());
            }
            let mean = nums.iter().sum::<f64>() / nums.len() as f64;
            Ok(Variant::Float(
                (nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>() / nums.len() as f64)
                    .sqrt(),
            ))
        }
        9 => Ok(as_integer_if_whole(nums.iter().sum::<f64>())),
        10 => {
            if nums.len() < 2 {
                return Err("AGGREGATE: at least 2 numeric values required".into());
            }
            let mean = nums.iter().sum::<f64>() / nums.len() as f64;
            Ok(Variant::Float(
                nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>()
                    / (nums.len() - 1) as f64,
            ))
        }
        11 => {
            if nums.is_empty() {
                return Err("AGGREGATE: at least 1 numeric value required".into());
            }
            let mean = nums.iter().sum::<f64>() / nums.len() as f64;
            Ok(Variant::Float(
                nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>() / nums.len() as f64,
            ))
        }
        12 => {
            // MEDIAN
            let mut s = nums.clone();
            s.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
            if s.is_empty() {
                return Err("AGGREGATE: no values".into());
            }
            let mid = s.len() / 2;
            let r = if s.len().is_multiple_of(2) {
                (s[mid - 1] + s[mid]) / 2.0
            } else {
                s[mid]
            };
            Ok(as_integer_if_whole(r))
        }
        13 => {
            if nums.is_empty() {
                return Err("AGGREGATE: no values".into());
            }
            let mut frequencies = HashMap::new();
            for value in nums {
                *frequencies.entry(value.to_bits()).or_insert(0_usize) += 1;
            }
            let maximum = frequencies.values().copied().max().unwrap_or(0);
            frequencies
                .into_iter()
                .filter_map(|(value, count)| (count == maximum).then_some(f64::from_bits(value)))
                .reduce(f64::min)
                .map(as_integer_if_whole)
                .ok_or_else(|| "AGGREGATE: no mode".into())
        }
        14 | 15 => {
            if rest.len() != 2 {
                return Err("AGGREGATE LARGE/SMALL requires ref and k".into());
            }
            let mut data = aggregate_nums(&rest[0])?;
            let k = to_float(&evaluate(&rest[1], cells)?)?;
            if !k.is_finite() || k.fract() != 0.0 || k < 1.0 || k > data.len() as f64 {
                return Err("AGGREGATE: k out of range".into());
            }
            data.sort_by(|left, right| left.partial_cmp(right).unwrap_or(Ordering::Equal));
            let index = k as usize - 1;
            Ok(as_integer_if_whole(if fn_num % 100 == 14 {
                data[data.len() - 1 - index]
            } else {
                data[index]
            }))
        }
        16..=19 => {
            if rest.len() != 2 {
                return Err("AGGREGATE percentile mode requires ref and k".into());
            }
            let data = aggregate_nums(&rest[0])?;
            let k = to_float(&evaluate(&rest[1], cells)?)?;
            match fn_num % 100 {
                16 => percentile_value(data, k, false),
                17 => percentile_value(data, k / 4.0, false),
                18 => percentile_value(data, k, true),
                19 => percentile_value(data, k / 4.0, true),
                _ => unreachable!(),
            }
        }
        20 | 21 => {
            if !(2..=3).contains(&rest.len()) {
                return Err(
                    "AGGREGATE PERCENTRANK requires ref, x, and optional significance".into(),
                );
            }
            let mut data = aggregate_nums(&rest[0])?;
            let x = to_float(&evaluate(&rest[1], cells)?)?;
            let significance = if rest.len() == 3 {
                to_float(&evaluate(&rest[2], cells)?)?
            } else {
                3.0
            };
            if !significance.is_finite() || significance < 0.0 || significance.fract() != 0.0 {
                return Err("AGGREGATE: invalid significance".into());
            }
            if data.is_empty() || !x.is_finite() {
                return Err("AGGREGATE: invalid percent rank input".into());
            }
            data.sort_by(|left, right| left.partial_cmp(right).unwrap_or(Ordering::Equal));
            if x < data[0] || x > data[data.len() - 1] {
                return Err("AGGREGATE: x is out of range".into());
            }
            let upper = data.partition_point(|value| *value <= x);
            let rank = if upper == 0 {
                1.0
            } else if upper == data.len() {
                data.len() as f64
            } else {
                let high = data[upper];
                let low = data[upper - 1];
                upper as f64
                    + if high == low {
                        0.0
                    } else {
                        (x - low) / (high - low)
                    }
            };
            let normalized = if fn_num % 100 == 20 {
                if data.len() == 1 {
                    0.0
                } else {
                    (rank - 1.0) / (data.len() - 1) as f64
                }
            } else {
                rank / (data.len() as f64 + 1.0)
            };
            let multiplier = 10_f64.powf(significance);
            Ok(Variant::Float(
                (normalized * multiplier).round() / multiplier,
            ))
        }
        n => Err(format!("AGGREGATE: unsupported function_num {}", n)),
    }
}

// ── String functions ──────────────────────────────────────────────────────────

fn func_upper(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("UPPER requires 1 argument".into());
    }
    Ok(Variant::Str(
        to_str(&evaluate(&args[0], cells)?).to_uppercase(),
    ))
}

fn func_lower(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("LOWER requires 1 argument".into());
    }
    Ok(Variant::Str(
        to_str(&evaluate(&args[0], cells)?).to_lowercase(),
    ))
}

fn func_phonetic(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("PHONETIC requires 1 argument".into());
    }
    // OOXML phonetic runs are not currently retained by the workbook model.
    // Preserve the source text rather than invoking any locale-dependent UI or
    // speech service; this is deterministic and safe for headless execution.
    match evaluate(&args[0], cells)? {
        Variant::Str(value) => Ok(Variant::Str(value)),
        Variant::Empty => Ok(Variant::Str(String::new())),
        _ => Ok(Variant::Error(ExcelError::Value)),
    }
}

fn func_encodeurl(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ENCODEURL requires 1 argument".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let mut encoded = String::with_capacity(text.len());
    for byte in text.as_bytes() {
        if byte.is_ascii_alphanumeric() || matches!(*byte, b'-' | b'.' | b'_' | b'~') {
            encoded.push(*byte as char);
        } else {
            encoded.push('%');
            encoded.push(char::from(b"0123456789ABCDEF"[(byte >> 4) as usize]));
            encoded.push(char::from(b"0123456789ABCDEF"[(byte & 0x0f) as usize]));
        }
    }
    Ok(Variant::Str(encoded))
}

/// Detect a small, deterministic language subset without contacting a
/// translation service. Excel's implementation is service-backed; returning
/// #N/A for ambiguous text is safer than claiming a language from weak
/// evidence.
fn func_detectlanguage(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("DETECTLANGUAGE requires 1 argument".into());
    }
    let text = match evaluate(&args[0], cells)? {
        Variant::Str(text) => text,
        _ => return Ok(Variant::Error(ExcelError::Value)),
    };
    if text.is_empty() || text.chars().count() > 32_768 {
        return Ok(Variant::Error(ExcelError::NA));
    }

    let mut script = None;
    for ch in text.chars() {
        let code = ch as u32;
        script = if (0x3040..=0x30ff).contains(&code) {
            Some("ja")
        } else if (0xac00..=0xd7af).contains(&code) {
            Some("ko")
        } else if (0x0600..=0x06ff).contains(&code) {
            Some("ar")
        } else if (0x0590..=0x05ff).contains(&code) {
            Some("he")
        } else if (0x0370..=0x03ff).contains(&code) {
            Some("el")
        } else if (0x0400..=0x04ff).contains(&code) {
            Some("ru")
        } else if (0x4e00..=0x9fff).contains(&code) {
            Some("zh")
        } else {
            script
        };
        if matches!(script, Some("ja" | "ko" | "ar" | "he" | "el" | "ru")) {
            break;
        }
    }
    if let Some(code) = script {
        return Ok(Variant::Str(code.into()));
    }

    let lower = text.to_lowercase();
    let words: Vec<&str> = lower
        .split(|ch: char| !ch.is_ascii_alphabetic())
        .filter(|word| !word.is_empty())
        .collect();
    let lexicons: [(&str, &[&str]); 5] = [
        (
            "en",
            &["the", "and", "of", "to", "is", "in", "hello", "world"],
        ),
        (
            "es",
            &["el", "la", "los", "las", "de", "que", "hola", "mundo"],
        ),
        (
            "fr",
            &["le", "la", "les", "des", "de", "que", "bonjour", "monde"],
        ),
        (
            "de",
            &["der", "die", "das", "und", "ist", "ich", "hallo", "welt"],
        ),
        ("it", &["il", "la", "gli", "dei", "che", "ciao", "mondo"]),
    ];
    let mut best = None;
    let mut best_score = 0usize;
    let mut tied = false;
    for (code, vocabulary) in lexicons {
        let score = words
            .iter()
            .filter(|word| vocabulary.contains(word))
            .count();
        if score > best_score {
            best = Some(code);
            best_score = score;
            tied = false;
        } else if score != 0 && score == best_score {
            tied = true;
        }
    }
    if best_score >= 2 && !tied {
        Ok(Variant::Str(best.unwrap().into()))
    } else {
        Ok(Variant::Error(ExcelError::NA))
    }
}

fn regex_from_args(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    pattern_index: usize,
    case_index: Option<usize>,
    name: &str,
) -> Result<Regex, String> {
    let pattern = to_str(&evaluate(&args[pattern_index], cells)?);
    let case_insensitive = if let Some(index) = case_index {
        if let Some(arg) = args.get(index) {
            let value = to_float(&evaluate(arg, cells)?)?;
            if !value.is_finite() || value.fract() != 0.0 || !(0.0..=1.0).contains(&value) {
                return Err(format!("{name}: case_sensitivity must be 0 or 1"));
            }
            value == 1.0
        } else {
            false
        }
    } else {
        false
    };
    let source = if case_insensitive {
        format!("(?i:{pattern})")
    } else {
        pattern
    };
    Regex::new(&source).map_err(|error| format!("{name}: invalid pattern: {error}"))
}

fn func_regextest(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(2..=3).contains(&args.len()) {
        return Err("REGEXTEST requires 2 or 3 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let regex = regex_from_args(args, cells, 1, Some(2), "REGEXTEST")?;
    Ok(Variant::Boolean(regex.is_match(&text)))
}

fn func_regexextract(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(2..=4).contains(&args.len()) {
        return Err("REGEXEXTRACT requires 2 to 4 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let mode = if let Some(arg) = args.get(2) {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=2.0).contains(&value) {
            return Ok(Variant::Error(ExcelError::Value));
        }
        value as u8
    } else {
        0
    };
    let regex = regex_from_args(args, cells, 1, Some(3), "REGEXEXTRACT")?;
    if mode == 1 {
        return Ok(Variant::Array(
            regex
                .find_iter(&text)
                .map(|matched| Variant::Str(matched.as_str().to_string()))
                .collect(),
        ));
    }
    let Some(captures) = regex.captures(&text) else {
        return Ok(Variant::Error(ExcelError::NA));
    };
    if mode == 2 {
        if regex.captures_len() <= 1 {
            return Ok(Variant::Error(ExcelError::NA));
        }
        return Ok(Variant::Array(
            captures
                .iter()
                .skip(1)
                .map(|capture| {
                    Variant::Str(capture.map_or(String::new(), |m| m.as_str().to_string()))
                })
                .collect(),
        ));
    }
    Ok(Variant::Str(
        captures
            .get(0)
            .map_or(String::new(), |m| m.as_str().to_string()),
    ))
}

fn func_regexreplace(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("REGEXREPLACE requires 3 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let regex = regex_from_args(args, cells, 1, None, "REGEXREPLACE")?;
    let replacement = to_str(&evaluate(&args[2], cells)?);
    Ok(Variant::Str(
        regex.replace_all(&text, replacement.as_str()).to_string(),
    ))
}

fn func_proper(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("PROPER requires 1 argument".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let mut cap = true;
    let result: String = s
        .chars()
        .map(|c| {
            let out = if cap {
                c.to_uppercase().next().unwrap_or(c)
            } else {
                c.to_lowercase().next().unwrap_or(c)
            };
            cap = !c.is_alphanumeric();
            out
        })
        .collect();
    Ok(Variant::Str(result))
}

fn func_trim(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("TRIM requires 1 argument".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let result = s.split_whitespace().collect::<Vec<_>>().join(" ");
    Ok(Variant::Str(result))
}

fn func_clean(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("CLEAN requires 1 argument".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    Ok(Variant::Str(
        text.chars().filter(|c| *c as u32 > 31).collect(),
    ))
}

fn func_t(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("T requires 1 argument".into());
    }
    match evaluate(&args[0], cells)? {
        Variant::Str(text) => Ok(Variant::Str(text)),
        _ => Ok(Variant::Str(String::new())),
    }
}

fn grouped_decimal(mut text: String) -> String {
    let negative = text.starts_with('-');
    if negative {
        text.remove(0);
    }
    let (integer, fraction) = text.split_once('.').unwrap_or((&text, ""));
    let mut grouped = String::new();
    for (index, ch) in integer.chars().enumerate() {
        if index > 0 && (integer.len() - index).is_multiple_of(3) {
            grouped.push(',');
        }
        grouped.push(ch);
    }
    if !fraction.is_empty() {
        grouped.push('.');
        grouped.push_str(fraction);
    }
    if negative {
        format!("-{grouped}")
    } else {
        grouped
    }
}

fn fixed_number(value: f64, decimals: i32, no_commas: bool) -> Result<String, String> {
    if !value.is_finite() || !(-100..=100).contains(&decimals) {
        return Err("FIXED: invalid number of decimals".into());
    }
    let factor = 10_f64.powi(decimals.abs());
    let rounded = if decimals >= 0 {
        (value * factor).round() / factor
    } else {
        (value / factor).round() * factor
    };
    let text = if decimals >= 0 {
        format!("{rounded:.prec$}", prec = decimals as usize)
    } else {
        format!("{rounded:.0}")
    };
    Ok(if no_commas {
        text
    } else {
        grouped_decimal(text)
    })
}

const THAI_DIGITS: [&str; 10] = [
    "ศูนย์",
    "หนึ่ง",
    "สอง",
    "สาม",
    "สี่",
    "ห้า",
    "หก",
    "เจ็ด",
    "แปด",
    "เก้า",
];

fn thai_integer_group(mut value: u64) -> String {
    if value == 0 {
        return String::new();
    }
    let positions = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน"];
    let mut parts = Vec::new();
    for (position, suffix) in positions.iter().enumerate() {
        let digit = (value % 10) as usize;
        value /= 10;
        if digit != 0 {
            let word = if position == 1 && digit == 1 {
                "สิบ"
            } else if position == 1 && digit == 2 {
                "ยี่สิบ"
            } else if position == 0 && digit == 1 && value > 0 {
                "เอ็ด"
            } else {
                THAI_DIGITS[digit]
            };
            if position == 1 && (digit == 1 || digit == 2) {
                parts.push(word.to_string());
            } else {
                parts.push(format!("{}{}", word, suffix));
            }
        }
        if value == 0 {
            break;
        }
    }
    parts.into_iter().rev().collect()
}

fn thai_integer(value: u64) -> String {
    if value == 0 {
        return THAI_DIGITS[0].to_string();
    }
    let mut groups = Vec::new();
    let mut remaining = value;
    while remaining > 0 {
        groups.push(remaining % 1_000_000);
        remaining /= 1_000_000;
    }
    groups
        .into_iter()
        .enumerate()
        .rev()
        .filter(|&(_, group)| group != 0)
        .map(|(index, group)| format!("{}{}", thai_integer_group(group), "ล้าน".repeat(index)))
        .collect()
}

fn func_bahttext(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("BAHTTEXT requires 1 argument".into());
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    if !value.is_finite() || value.abs() >= 1e15 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let scaled = (value.abs() * 100.0).round() as u64;
    let integer = scaled / 100;
    let satang = scaled % 100;
    let mut result = String::new();
    if value < 0.0 {
        result.push_str("ลบ");
    }
    result.push_str(&thai_integer(integer));
    result.push_str("บาท");
    if satang == 0 {
        result.push_str("ถ้วน");
    } else {
        result.push_str(&thai_integer(satang));
        result.push_str("สตางค์");
    }
    Ok(Variant::Str(result))
}

fn func_fixed(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 3 {
        return Err("FIXED requires 1 to 3 arguments".into());
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    let decimals = if args.len() >= 2 {
        to_float(&evaluate(&args[1], cells)?)? as i32
    } else {
        2
    };
    let no_commas = args.len() == 3 && is_truthy(&evaluate(&args[2], cells)?);
    fixed_number(value, decimals, no_commas).map(Variant::Str)
}

fn func_dollar(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("DOLLAR requires 1 or 2 arguments".into());
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    let decimals = if args.len() == 2 {
        to_float(&evaluate(&args[1], cells)?)? as i32
    } else {
        2
    };
    let text = fixed_number(value.abs(), decimals, false).map_err(|e| format!("DOLLAR: {e}"))?;
    Ok(Variant::Str(if value < 0.0 {
        format!("-${}", text)
    } else {
        format!("${}", text)
    }))
}

fn func_base(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("BASE requires 2 or 3 arguments".into());
    }
    let number = to_float(&evaluate(&args[0], cells)?)?;
    let radix = to_float(&evaluate(&args[1], cells)?)?;
    let min_length = if args.len() == 3 {
        to_float(&evaluate(&args[2], cells)?)?
    } else {
        0.0
    };
    if !number.is_finite()
        || number < 0.0
        || number.fract() != 0.0
        || !(2.0..=36.0).contains(&radix)
        || min_length < 0.0
        || min_length.fract() != 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut n = number as u64;
    let radix = radix as u64;
    let digits = b"0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    let mut result = if n == 0 {
        "0".to_string()
    } else {
        let mut out = String::new();
        while n > 0 {
            out.push(digits[(n % radix) as usize] as char);
            n /= radix;
        }
        out.chars().rev().collect()
    };
    while result.len() < min_length as usize {
        result.insert(0, '0');
    }
    Ok(Variant::Str(result))
}

fn func_decimal(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("DECIMAL requires 2 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let radix = to_float(&evaluate(&args[1], cells)?)?;
    if !(2.0..=36.0).contains(&radix) || radix.fract() != 0.0 || text.is_empty() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    match u64::from_str_radix(&text, radix as u32) {
        Ok(value) => Ok(as_integer_if_whole(value as f64)),
        Err(_) => Ok(Variant::Error(ExcelError::Num)),
    }
}

fn func_substitute(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 4 {
        return Err("SUBSTITUTE requires 3 or 4 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let old = to_str(&evaluate(&args[1], cells)?);
    let new = to_str(&evaluate(&args[2], cells)?);
    if old.is_empty() {
        return Ok(Variant::Str(text));
    }
    if args.len() == 3 {
        return Ok(Variant::Str(text.replace(&old as &str, &new as &str)));
    }
    let instance = to_float(&evaluate(&args[3], cells)?)? as usize;
    let mut result = String::new();
    let mut count = 0usize;
    let mut search_from = 0usize;
    let text_chars: Vec<char> = text.chars().collect();
    let old_chars: Vec<char> = old.chars().collect();
    let mut i = 0;
    while i <= text_chars.len().saturating_sub(old_chars.len()) {
        if text_chars[i..].starts_with(&old_chars) {
            count += 1;
            if count == instance {
                result.extend(text_chars[search_from..i].iter());
                result.push_str(&new);
                search_from = i + old_chars.len();
                i += old_chars.len();
                continue;
            }
        }
        i += 1;
    }
    result.extend(text_chars[search_from..].iter());
    Ok(Variant::Str(result))
}

fn func_replace(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("REPLACE requires 4 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let start = (to_float(&evaluate(&args[1], cells)?)? as usize).saturating_sub(1);
    let len = to_float(&evaluate(&args[2], cells)?)? as usize;
    let new = to_str(&evaluate(&args[3], cells)?);
    let chars: Vec<char> = text.chars().collect();
    let end = (start + len).min(chars.len());
    let result: String = chars[..start]
        .iter()
        .chain(new.chars().collect::<Vec<_>>().iter())
        .chain(chars[end..].iter())
        .collect();
    Ok(Variant::Str(result))
}

fn func_replaceb(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("REPLACEB requires 4 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let start = to_float(&evaluate(&args[1], cells)?)?;
    let length = to_float(&evaluate(&args[2], cells)?)?;
    if !start.is_finite()
        || !length.is_finite()
        || start < 1.0
        || length < 0.0
        || start.fract() != 0.0
        || length.fract() != 0.0
    {
        return Err("REPLACEB: invalid byte position or length".into());
    }
    let new = to_str(&evaluate(&args[3], cells)?);
    let start_byte = start as usize - 1;
    let end_byte = start_byte.saturating_add(length as usize);
    let mut prefix = String::new();
    let mut suffix = String::new();
    let mut offset = 0;
    for ch in text.chars() {
        let next = offset + char_byte_width(ch);
        if next <= start_byte {
            prefix.push(ch);
        }
        if offset >= end_byte {
            suffix.push(ch);
        }
        offset = next;
    }
    Ok(Variant::Str(format!("{prefix}{new}{suffix}")))
}

fn func_find(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("FIND requires 2 or 3 arguments".into());
    }
    let needle = to_str(&evaluate(&args[0], cells)?);
    let haystack = to_str(&evaluate(&args[1], cells)?);
    let start = if args.len() == 3 {
        (to_float(&evaluate(&args[2], cells)?)? as usize).saturating_sub(1)
    } else {
        0
    };
    let h_chars: Vec<char> = haystack.chars().collect();
    let n_chars: Vec<char> = needle.chars().collect();
    // An empty search string matches trivially at `start` (real Excel: `FIND("","abc")` is
    // 1, `FIND("","abc",3)` is 3) -- mirrors VBA InStr's identical "empty needle -> return
    // start position" convention (`src/vm/mod.rs`). Handled before the `.windows()` call
    // below, which panics on a zero window size (`slice::windows` requires non-zero) --
    // found by the fuzz corpus on `FIND("","...")`.
    if n_chars.is_empty() {
        return if start > h_chars.len() {
            Err("FIND: value not found".into())
        } else {
            Ok(Variant::Integer((start + 1) as i64))
        };
    }
    // `start` beyond the haystack has no match -- guarded explicitly rather than letting
    // `h_chars[start..]` panic on an out-of-bounds slice.
    if start >= h_chars.len() {
        return Err("FIND: value not found".into());
    }
    let pos = h_chars[start..]
        .windows(n_chars.len())
        .position(|w| w == n_chars.as_slice());
    pos.map(|p| Variant::Integer((start + p + 1) as i64))
        .ok_or_else(|| "FIND: value not found".into())
}

fn func_findb(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("FINDB requires 2 or 3 arguments".into());
    }
    let needle = to_str(&evaluate(&args[0], cells)?);
    let haystack = to_str(&evaluate(&args[1], cells)?);
    let start_byte = if args.len() == 3 {
        let start = to_float(&evaluate(&args[2], cells)?)?;
        if !start.is_finite() || start < 1.0 || start.fract() != 0.0 {
            return Err("FINDB: invalid start position".into());
        }
        start as usize - 1
    } else {
        0
    };
    let hay_chars: Vec<char> = haystack.chars().collect();
    let needle_chars: Vec<char> = needle.chars().collect();
    if needle_chars.is_empty() {
        return if start_byte <= str_byte_len(&haystack) {
            Ok(Variant::Integer((start_byte + 1) as i64))
        } else {
            Err("FINDB: value not found".into())
        };
    }
    let mut byte_offset = 0;
    for (index, _) in hay_chars.iter().enumerate() {
        if byte_offset >= start_byte && hay_chars[index..].starts_with(&needle_chars) {
            return Ok(Variant::Integer((byte_offset + 1) as i64));
        }
        byte_offset += char_byte_width(hay_chars[index]);
    }
    Err("FINDB: value not found".into())
}

fn func_search(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("SEARCH requires 2 or 3 arguments".into());
    }
    let needle = to_str(&evaluate(&args[0], cells)?).to_uppercase();
    let haystack = to_str(&evaluate(&args[1], cells)?);
    let start = if args.len() == 3 {
        (to_float(&evaluate(&args[2], cells)?)? as usize).saturating_sub(1)
    } else {
        0
    };
    let h_chars: Vec<char> = haystack.chars().collect();
    let n_chars: Vec<char> = needle.chars().collect();
    // wildcard-aware case-insensitive search
    let h_upper: Vec<char> = h_chars
        .iter()
        .map(|c| c.to_uppercase().next().unwrap_or(*c))
        .collect();
    for i in start..=h_upper.len() {
        if wildcard_match_prefix(&h_upper[i..], &n_chars) {
            return Ok(Variant::Integer((i + 1) as i64));
        }
    }
    Err("SEARCH: value not found".into())
}

fn func_searchb(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("SEARCHB requires 2 or 3 arguments".into());
    }
    let needle = to_str(&evaluate(&args[0], cells)?).to_uppercase();
    let haystack = to_str(&evaluate(&args[1], cells)?);
    let start_byte = if args.len() == 3 {
        let start = to_float(&evaluate(&args[2], cells)?)?;
        if !start.is_finite() || start < 1.0 || start.fract() != 0.0 {
            return Err("SEARCHB: invalid start position".into());
        }
        start as usize - 1
    } else {
        0
    };
    let hay_chars: Vec<char> = haystack.chars().collect();
    let hay_upper: Vec<char> = hay_chars
        .iter()
        .map(|ch| ch.to_uppercase().next().unwrap_or(*ch))
        .collect();
    let needle_chars: Vec<char> = needle.chars().collect();
    let mut byte_offset = 0;
    for (index, _) in hay_chars.iter().enumerate() {
        if byte_offset >= start_byte && wildcard_match_prefix(&hay_upper[index..], &needle_chars) {
            return Ok(Variant::Integer((byte_offset + 1) as i64));
        }
        byte_offset += char_byte_width(hay_chars[index]);
    }
    Err("SEARCHB: value not found".into())
}

fn func_exact(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("EXACT requires 2 arguments".into());
    }
    let a = to_str(&evaluate(&args[0], cells)?);
    let b = to_str(&evaluate(&args[1], cells)?);
    Ok(Variant::Boolean(a == b))
}

fn func_hyperlink(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("HYPERLINK requires 1 or 2 arguments".into());
    }
    let link_location = to_str(&evaluate(&args[0], cells)?);
    let friendly_name = if args.len() == 2 {
        to_str(&evaluate(&args[1], cells)?)
    } else {
        link_location
    };
    Ok(Variant::Str(friendly_name))
}

fn func_textjoin(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 {
        return Err("TEXTJOIN requires at least 3 arguments".into());
    }
    let delim = to_str(&evaluate(&args[0], cells)?);
    let ignore_empty = is_truthy(&evaluate(&args[1], cells)?);
    let parts: Vec<String> = collect_all(&args[2..], cells)?
        .iter()
        .map(to_str)
        .filter(|s| !ignore_empty || !s.is_empty())
        .collect();
    Ok(Variant::Str(parts.join(&delim)))
}

fn func_value(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("VALUE requires 1 argument".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let cleaned = s.replace(',', "");
    cleaned
        .parse::<f64>()
        .map(as_integer_if_whole)
        .map_err(|_| format!("VALUE: cannot convert '{}' to number", s))
}

fn func_rept(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("REPT requires 2 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let n = to_float(&evaluate(&args[1], cells)?)? as i64;
    if n < 0 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    Ok(Variant::Str(text.repeat(n as usize)))
}

fn func_numbervalue(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 3 {
        return Err("NUMBERVALUE requires 1 to 3 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let dec = if args.len() >= 2 {
        to_str(&evaluate(&args[1], cells)?)
    } else {
        ".".into()
    };
    let grp = if args.len() >= 3 {
        to_str(&evaluate(&args[2], cells)?)
    } else {
        ",".into()
    };
    if dec.is_empty() || dec == grp {
        return Ok(Variant::Error(ExcelError::Value));
    }
    if text.trim().is_empty() {
        return Ok(Variant::Integer(0));
    }
    let normalized = text.replace(grp.as_str(), "").replace(dec.as_str(), ".");
    match normalized.parse::<f64>() {
        Ok(n) => Ok(as_integer_if_whole(n)),
        Err(_) => Ok(Variant::Error(ExcelError::Value)),
    }
}

fn func_char(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("CHAR requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)? as u32;
    char::from_u32(n)
        .map(|c| Variant::Str(c.to_string()))
        .ok_or_else(|| format!("CHAR: invalid code {}", n))
}

fn func_code(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("CODE requires 1 argument".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    s.chars()
        .next()
        .map(|c| Variant::Integer(c as i64))
        .ok_or_else(|| "CODE: empty string".into())
}

/// Return the Unicode scalar value requested by UNICHAR.
///
/// UNICHAR is deliberately separate from CHAR: CHAR is the legacy character
/// function, while UNICHAR accepts the full Unicode scalar range and must not
/// manufacture UTF-16 surrogate code points.
fn func_unichar(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("UNICHAR requires 1 argument".into());
    }
    let value = match to_float(&evaluate(&args[0], cells)?) {
        Ok(value) if value.is_finite() && value.fract() == 0.0 => value,
        _ => return Ok(Variant::Error(ExcelError::Value)),
    };
    if !(1.0..=0x10_FFFF as f64).contains(&value) {
        return Ok(Variant::Error(ExcelError::Value));
    }
    match char::from_u32(value as u32) {
        Some(character) => Ok(Variant::Str(character.to_string())),
        None => Ok(Variant::Error(ExcelError::Value)),
    }
}

/// Return the Unicode scalar value of the first character in a string.
fn func_unicode(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("UNICODE requires 1 argument".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    match text.chars().next() {
        Some(character) => Ok(Variant::Integer(character as i64)),
        None => Ok(Variant::Error(ExcelError::Value)),
    }
}

fn func_asc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ASC requires 1 argument".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let result: String = s
        .chars()
        .map(|c| {
            let cp = c as u32;
            // Full-width ASCII/punctuation U+FF01-U+FF5E → half-width U+0021-U+007E
            if (0xFF01..=0xFF5E).contains(&cp) {
                char::from_u32(cp - 0xFEE0).unwrap_or(c)
            // Full-width space U+3000 → half-width space
            } else if cp == 0x3000 {
                '\u{0020}'
            // Full-width katakana (basic) U+30A1-U+30F6 → half-width katakana U+FF71-U+FF9F
            } else if (0x30A1..=0x30F6).contains(&cp) {
                let base = [
                    '\u{FF71}', '\u{FF72}', '\u{FF73}', '\u{FF74}', '\u{FF75}', // ア-オ
                    '\u{FF76}', '\u{FF77}', '\u{FF78}', '\u{FF79}', '\u{FF7A}', // カ-コ
                    '\u{FF7B}', '\u{FF7C}', '\u{FF7D}', '\u{FF7E}', '\u{FF7F}', // サ-ソ
                    '\u{FF80}', '\u{FF81}', '\u{FF82}', '\u{FF83}', '\u{FF84}', // タ-ト
                    '\u{FF85}', '\u{FF86}', '\u{FF87}', '\u{FF88}', '\u{FF89}', // ナ-ノ
                    '\u{FF8A}', '\u{FF8B}', '\u{FF8C}', '\u{FF8D}', '\u{FF8E}', // ハ-ホ
                    '\u{FF8F}', '\u{FF90}', '\u{FF91}', '\u{FF92}', '\u{FF93}', // マ-モ
                    '\u{FF94}', '\u{FF95}', '\u{FF96}', // ヤユヨ
                    '\u{FF97}', '\u{FF98}', '\u{FF99}', '\u{FF9A}', '\u{FF9B}', // ラ-ロ
                    '\u{FF9C}', '\u{FF9D}', // ワン
                ];
                // Map full-width katakana code point to half-width index
                // Only map direct (non-voiced) correspondences; voiced/semi-voiced left as-is
                let idx_map: &[(u32, usize)] = &[
                    (0x30A2, 0),
                    (0x30A4, 1),
                    (0x30A6, 2),
                    (0x30A8, 3),
                    (0x30AA, 4),
                    (0x30AB, 5),
                    (0x30AD, 6),
                    (0x30AF, 7),
                    (0x30B1, 8),
                    (0x30B3, 9),
                    (0x30B5, 10),
                    (0x30B7, 11),
                    (0x30B9, 12),
                    (0x30BB, 13),
                    (0x30BD, 14),
                    (0x30BF, 15),
                    (0x30C1, 16),
                    (0x30C4, 17),
                    (0x30C6, 18),
                    (0x30C8, 19),
                    (0x30CA, 20),
                    (0x30CB, 21),
                    (0x30CC, 22),
                    (0x30CD, 23),
                    (0x30CE, 24),
                    (0x30CF, 25),
                    (0x30D2, 26),
                    (0x30D5, 27),
                    (0x30D8, 28),
                    (0x30DB, 29),
                    (0x30DE, 30),
                    (0x30DF, 31),
                    (0x30E0, 32),
                    (0x30E1, 33),
                    (0x30E2, 34),
                    (0x30E4, 35),
                    (0x30E6, 36),
                    (0x30E8, 37),
                    (0x30E9, 38),
                    (0x30EA, 39),
                    (0x30EB, 40),
                    (0x30EC, 41),
                    (0x30ED, 42),
                    (0x30EF, 43),
                    (0x30F3, 44),
                ];
                idx_map
                    .iter()
                    .find(|&&(k, _)| k == cp)
                    .map(|&(_, i)| base[i])
                    .unwrap_or(c)
            } else {
                c
            }
        })
        .collect();
    Ok(Variant::Str(result))
}

fn func_jis(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("JIS requires 1 argument".into());
    }
    let s = to_str(&evaluate(&args[0], cells)?);
    let result: String = s
        .chars()
        .map(|c| {
            let cp = c as u32;
            // Half-width ASCII/punctuation U+0021-U+007E → full-width U+FF01-U+FF5E
            if (0x0021..=0x007E).contains(&cp) {
                char::from_u32(cp + 0xFEE0).unwrap_or(c)
            } else if cp == 0x0020 {
                '\u{3000}' // space → ideographic space
            } else {
                c
            }
        })
        .collect();
    Ok(Variant::Str(result))
}

// ── Date/Time functions ───────────────────────────────────────────────────────

fn func_year(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("YEAR requires 1 argument".into());
    }
    let s = to_float(&evaluate(&args[0], cells)?)? as i64;
    Ok(Variant::Integer(serial_to_ymd(s).0 as i64))
}

fn func_month(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("MONTH requires 1 argument".into());
    }
    let s = to_float(&evaluate(&args[0], cells)?)? as i64;
    Ok(Variant::Integer(serial_to_ymd(s).1 as i64))
}

fn func_day(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("DAY requires 1 argument".into());
    }
    let s = to_float(&evaluate(&args[0], cells)?)? as i64;
    Ok(Variant::Integer(serial_to_ymd(s).2 as i64))
}

fn func_weekday(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("WEEKDAY requires 1 or 2 arguments".into());
    }
    let serial = to_float(&evaluate(&args[0], cells)?)? as i64;
    let return_type = if args.len() == 2 {
        integer_mode_arg(to_float(&evaluate(&args[1], cells)?)?, "WEEKDAY")?
    } else {
        1
    };
    // serial_weekday: 0=Sun,1=Mon,...,5=Fri,6=Sat
    let wd = serial_weekday(serial);
    let result = match return_type {
        1 => wd + 1,            // Sun=1..Sat=7
        2 => (wd + 6) % 7 + 1,  // Mon=1..Sun=7
        3 => (wd + 6) % 7,      // Mon=0..Sun=6
        11 => (wd + 6) % 7 + 1, // Mon=1..Sun=7
        12 => (wd + 5) % 7 + 1, // Tue=1..Mon=7
        13 => (wd + 4) % 7 + 1, // Wed=1..Tue=7
        14 => (wd + 3) % 7 + 1, // Thu=1..Wed=7
        15 => (wd + 2) % 7 + 1, // Fri=1..Thu=7
        16 => (wd + 1) % 7 + 1, // Sat=1..Fri=7
        17 => wd + 1,           // Sun=1..Sat=7
        _ => return Err(format!("WEEKDAY: unsupported return_type {}", return_type)),
    };
    Ok(Variant::Integer(result as i64))
}

/// ISO 8601 week number (internal helper, used by ISOWEEKNUM and WEEKNUM type 21).
fn iso_weeknum_impl(serial: i64) -> i64 {
    let (y, _, _) = serial_to_ymd(serial);
    let iso_dow = ((serial_weekday(serial) + 6) % 7 + 1) as i64; // 1=Mon..7=Sun
    let jan1 = date_to_serial(y, 1, 1);
    let doy = serial - jan1 + 1;
    let week = (doy - iso_dow + 10) / 7;
    if week < 1 {
        iso_weeknum_impl(date_to_serial(y - 1, 12, 31))
    } else if week > 52 {
        let dec31 = date_to_serial(y, 12, 31);
        let iso_dow_dec31 = ((serial_weekday(dec31) + 6) % 7 + 1) as i64;
        if iso_dow_dec31 >= 4 { week } else { 1 }
    } else {
        week
    }
}

fn func_isoweeknum(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISOWEEKNUM requires 1 argument".into());
    }
    let serial = to_float(&evaluate(&args[0], cells)?)? as i64;
    Ok(Variant::Integer(iso_weeknum_impl(serial)))
}

fn func_weeknum(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("WEEKNUM requires 1 or 2 arguments".into());
    }
    let serial = to_float(&evaluate(&args[0], cells)?)? as i64;
    let return_type = if args.len() == 2 {
        integer_mode_arg(to_float(&evaluate(&args[1], cells)?)?, "WEEKNUM")?
    } else {
        1
    };
    if return_type == 21 {
        return Ok(Variant::Integer(iso_weeknum_impl(serial)));
    }
    let (y, _, _) = serial_to_ymd(serial);
    let jan1 = date_to_serial(y, 1, 1);
    let doy = serial - jan1 + 1;
    let dow_jan1 = serial_weekday(jan1); // 0=Sun..6=Sat
    let offset: u32 = match return_type {
        1 | 17 => dow_jan1,           // Sun-start
        2 | 11 => (dow_jan1 + 6) % 7, // Mon-start
        12 => (dow_jan1 + 5) % 7,     // Tue-start
        13 => (dow_jan1 + 4) % 7,     // Wed-start
        14 => (dow_jan1 + 3) % 7,     // Thu-start
        15 => (dow_jan1 + 2) % 7,     // Fri-start
        16 => (dow_jan1 + 1) % 7,     // Sat-start
        _ => return Ok(Variant::Error(ExcelError::Num)),
    };
    Ok(Variant::Integer((doy - 1 + offset as i64) / 7 + 1))
}

fn func_days(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("DAYS requires 2 arguments".into());
    }
    let end = to_float(&evaluate(&args[0], cells)?)? as i64;
    let start = to_float(&evaluate(&args[1], cells)?)? as i64;
    Ok(Variant::Integer(end - start))
}

fn days360_serial(start: i64, end: i64, european: bool) -> i64 {
    let (y1, m1, mut d1) = serial_to_ymd(start);
    let (y2, m2, mut d2) = serial_to_ymd(end);
    if !european {
        if d1 == 31 || (m1 == 2 && d1 == days_in_month(y1, m1)) {
            d1 = 30;
        }
        if d2 == 31 && d1 >= 30 {
            d2 = 30;
        }
    } else {
        if d1 == 31 {
            d1 = 30;
        }
        if d2 == 31 {
            d2 = 30;
        }
    }
    (y2 - y1) as i64 * 360 + (m2 as i64 - m1 as i64) * 30 + d2 as i64 - d1 as i64
}

fn func_days360(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("DAYS360 requires 2 or 3 arguments".into());
    }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let end = to_float(&evaluate(&args[1], cells)?)? as i64;
    let method = if args.len() == 3 {
        is_truthy(&evaluate(&args[2], cells)?)
    } else {
        false
    };
    Ok(Variant::Integer(days360_serial(start, end, method)))
}

fn func_yearfrac(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("YEARFRAC requires 2 or 3 arguments".into());
    }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let end = to_float(&evaluate(&args[1], cells)?)? as i64;
    let basis = if args.len() == 3 {
        to_float(&evaluate(&args[2], cells)?)? as i32
    } else {
        0
    };
    if basis == 0 || basis == 4 {
        return Ok(Variant::Float(
            days360_serial(start, end, basis == 4) as f64 / 360.0,
        ));
    }
    let days = (end - start) as f64;
    let denom = match basis {
        2 => 360.0,
        3 => 365.0,
        1 => {
            let (y1, _, _) = serial_to_ymd(start);
            let (y2, _, _) = serial_to_ymd(end);
            if y1 == y2 {
                if is_leap(y1) { 366.0 } else { 365.0 }
            } else {
                (y1..=y2)
                    .map(|y| if is_leap(y) { 366.0 } else { 365.0 })
                    .sum::<f64>()
                    / (y2 - y1 + 1) as f64
            }
        }
        _ => return Ok(Variant::Error(ExcelError::Num)),
    };
    Ok(Variant::Float(days / denom))
}

fn func_edate(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("EDATE requires 2 arguments".into());
    }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let months = to_float(&evaluate(&args[1], cells)?)? as i32;
    let (mut y, mut m, d) = serial_to_ymd(start);
    let total = (m as i32 - 1) + months;
    y += total.div_euclid(12);
    m = (total.rem_euclid(12) + 1) as u32;
    let d_clamped = d.min(days_in_month(y, m));
    Ok(Variant::Date(date_to_serial(y, m, d_clamped)))
}

fn func_datedif(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("DATEDIF requires 3 arguments".into());
    }
    let s1 = to_float(&evaluate(&args[0], cells)?)? as i64;
    let s2 = to_float(&evaluate(&args[1], cells)?)? as i64;
    let unit = to_str(&evaluate(&args[2], cells)?).to_uppercase();
    if s1 > s2 {
        return Err("DATEDIF: start_date > end_date".into());
    }
    let (y1, m1, d1) = serial_to_ymd(s1);
    let (y2, m2, d2) = serial_to_ymd(s2);
    let result = match unit.as_str() {
        "Y" => (y2 - y1 - if (m2, d2) < (m1, d1) { 1 } else { 0 }) as i64,
        "M" => {
            let mut months = (y2 - y1) * 12 + (m2 as i32 - m1 as i32);
            if d2 < d1 {
                months -= 1;
            }
            months as i64
        }
        "D" => s2 - s1,
        "MD" => {
            let d2i = d2 as i32;
            let d1i = d1 as i32;
            (if d2i >= d1i {
                d2i - d1i
            } else {
                days_in_month(y2, m2) as i32 - d1i + d2i
            }) as i64
        }
        "YM" => {
            let mut m = m2 as i32 - m1 as i32;
            if m < 0 {
                m += 12;
            }
            if d2 < d1 {
                m -= 1;
                if m < 0 {
                    m += 12;
                }
            }
            m as i64
        }
        "YD" => {
            let base = date_to_serial(y2, m1, d1);
            if base <= s2 {
                s2 - base
            } else {
                s2 - date_to_serial(y2 - 1, m1, d1)
            }
        }
        u => return Err(format!("DATEDIF: unknown unit '{}'", u)),
    };
    Ok(Variant::Integer(result))
}

fn func_datevalue(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("DATEVALUE requires 1 argument".into());
    }
    let original = to_str(&evaluate(&args[0], cells)?);
    let trimmed = original.trim();
    let date_part = trimmed
        .split_once('T')
        .filter(|(_, time)| time.contains(':'))
        .map(|(date, _)| date)
        .or_else(|| {
            trimmed
                .split_once('t')
                .filter(|(_, time)| time.contains(':'))
                .map(|(date, _)| date)
        })
        .or_else(|| {
            trimmed
                .find(' ')
                .and_then(|index| trimmed[index..].contains(':').then_some(&trimmed[..index]))
        })
        .unwrap_or(trimmed);
    let named_tokens: Vec<&str> = date_part
        .split(|c: char| c == ',' || c.is_ascii_whitespace())
        .filter(|token| !token.is_empty())
        .collect();
    if named_tokens.len() == 3 {
        let named_date = if let Some(month) = english_month_number(named_tokens[0]) {
            named_tokens[1]
                .parse::<f64>()
                .ok()
                .zip(named_tokens[2].parse::<f64>().ok())
                .map(|(day, year)| (year, month, day))
        } else if let Some(month) = english_month_number(named_tokens[1]) {
            named_tokens[0]
                .parse::<f64>()
                .ok()
                .zip(named_tokens[2].parse::<f64>().ok())
                .map(|(day, year)| (year, month, day))
        } else {
            None
        };
        if let Some((year, month, day)) = named_date {
            return normalize_date_components(year, month, day);
        }
    }
    let separator = if date_part.contains('/') {
        '/'
    } else if date_part.contains('-') {
        '-'
    } else if date_part.contains('.') {
        '.'
    } else {
        return Err(format!("DATEVALUE: cannot parse '{}'", original));
    };
    let parts: Vec<&str> = date_part.split(separator).collect();
    if parts.len() != 3 {
        return Err(format!("DATEVALUE: cannot parse '{}'", original));
    }
    let parsed = (|| {
        Some((
            parts[0].trim().parse::<f64>().ok()?,
            parts[1].trim().parse::<f64>().ok()?,
            parts[2].trim().parse::<f64>().ok()?,
        ))
    })();
    match parsed {
        Some((year, month, day)) => normalize_date_components(year, month, day),
        None => Err(format!("DATEVALUE: cannot parse '{}'", original)),
    }
}

fn func_now(
    args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !args.is_empty() {
        return Err("NOW takes no arguments".into());
    }
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    let unix_days = secs / 86400;
    let frac = (secs % 86400) as f64 / 86400.0;
    Ok(Variant::Float(unix_days as f64 + 25569.0 + frac))
}

fn func_time_fn(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("TIME requires 3 arguments".into());
    }
    let h = to_float(&evaluate(&args[0], cells)?)?;
    let m = to_float(&evaluate(&args[1], cells)?)?;
    let s = to_float(&evaluate(&args[2], cells)?)?;
    Ok(Variant::Float((h * 3600.0 + m * 60.0 + s) / 86400.0))
}

fn func_timevalue(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("TIMEVALUE requires 1 argument".into());
    }
    let original = to_str(&evaluate(&args[0], cells)?);
    let trimmed = original.trim();
    let lower = trimmed.to_ascii_lowercase();
    let (clock, meridiem) = if let Some(clock) = lower.strip_suffix("am") {
        (clock.trim(), Some(false))
    } else if let Some(clock) = lower.strip_suffix("pm") {
        (clock.trim(), Some(true))
    } else {
        (lower.as_str(), None)
    };
    let parts: Vec<&str> = clock.split(':').collect();
    if parts.len() < 2 || parts.len() > 3 {
        return Err(format!("TIMEVALUE: cannot parse '{}'", original));
    }
    let parsed = (
        parts[0].trim().parse::<f64>(),
        parts[1].trim().parse::<f64>(),
        parts
            .get(2)
            .map_or(Ok(0.0), |part| part.trim().parse::<f64>()),
    );
    let (mut hour, minute, second) = match parsed {
        (Ok(hour), Ok(minute), Ok(second))
            if hour.is_finite()
                && minute.is_finite()
                && second.is_finite()
                && (0.0..60.0).contains(&minute)
                && (0.0..60.0).contains(&second) =>
        {
            (hour, minute, second)
        }
        _ => return Err(format!("TIMEVALUE: cannot parse '{}'", original)),
    };
    if let Some(is_pm) = meridiem {
        if !(1.0..=12.0).contains(&hour) {
            return Err(format!("TIMEVALUE: cannot parse '{}'", original));
        }
        if is_pm && hour < 12.0 {
            hour += 12.0;
        } else if !is_pm && hour == 12.0 {
            hour = 0.0;
        }
    } else if !(0.0..24.0).contains(&hour) {
        return Err(format!("TIMEVALUE: cannot parse '{}'", original));
    }
    Ok(Variant::Float(
        (hour * 3600.0 + minute * 60.0 + second) / 86400.0,
    ))
}

fn serial_frac(v: f64) -> f64 {
    v.fract().abs()
}

fn func_hour(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("HOUR requires 1 argument".into());
    }
    let v = to_float(&evaluate(&args[0], cells)?)?;
    Ok(Variant::Integer((serial_frac(v) * 24.0) as i64))
}

fn func_minute(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("MINUTE requires 1 argument".into());
    }
    let v = to_float(&evaluate(&args[0], cells)?)?;
    Ok(Variant::Integer(((serial_frac(v) * 1440.0) % 60.0) as i64))
}

fn func_second(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("SECOND requires 1 argument".into());
    }
    let v = to_float(&evaluate(&args[0], cells)?)?;
    Ok(Variant::Integer(((serial_frac(v) * 86400.0) % 60.0) as i64))
}

fn parse_weekend_mask(v: &Variant) -> Result<[bool; 7], String> {
    // Returns [Mon,Tue,Wed,Thu,Fri,Sat,Sun] true = is weekend
    match v {
        Variant::Str(s)
            if s.chars().count() == 7
                && s.chars().all(|character| matches!(character, '0' | '1')) =>
        {
            let mut mask = [false; 7];
            for (i, c) in s.chars().enumerate() {
                mask[i] = c == '1';
            }
            Ok(mask)
        }
        Variant::Str(_) => {
            Err("NETWORKDAYS.INTL: weekend mask must contain seven 0/1 characters".into())
        }
        _ => {
            let number = to_float(v)?;
            if !number.is_finite() || number.fract() != 0.0 {
                return Err("NETWORKDAYS.INTL: weekend code must be an integer".into());
            }
            let n = number as u32;
            Ok(match n {
                1 => [false, false, false, false, false, true, true],
                2 => [true, false, false, false, false, false, true],
                3 => [true, true, false, false, false, false, false],
                4 => [false, true, true, false, false, false, false],
                5 => [false, false, true, true, false, false, false],
                6 => [false, false, false, true, true, false, false],
                7 => [false, false, false, false, true, true, false],
                11 => [false, false, false, false, false, false, true],
                12 => [true, false, false, false, false, false, false],
                13 => [false, true, false, false, false, false, false],
                14 => [false, false, true, false, false, false, false],
                15 => [false, false, false, true, false, false, false],
                16 => [false, false, false, false, true, false, false],
                17 => [false, false, false, false, false, true, false],
                _ => return Err(format!("NETWORKDAYS.INTL: invalid weekend {}", n)),
            })
        }
    }
}

fn is_weekend_intl(serial: i64, mask: &[bool; 7]) -> bool {
    // serial_weekday: 0=Sun,1=Mon,...,5=Fri,6=Sat
    // mask: [Mon,Tue,Wed,Thu,Fri,Sat,Sun]
    let wd = serial_weekday(serial) as usize;
    let mask_idx = (wd + 6) % 7; // convert: 0=Sun→6, 1=Mon→0, ..., 6=Sat→5
    mask[mask_idx]
}

fn func_networkdays_intl(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 4 {
        return Err("NETWORKDAYS.INTL requires 2 to 4 arguments".into());
    }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let end = to_float(&evaluate(&args[1], cells)?)? as i64;
    let mask = if args.len() >= 3 {
        parse_weekend_mask(&evaluate(&args[2], cells)?)?
    } else {
        [false, false, false, false, false, true, true] // default Sat+Sun
    };
    let holidays: std::collections::HashSet<i64> = if args.len() == 4 {
        collect_values(&args[3], cells)?
            .iter()
            .filter_map(|v| to_float(v).ok().map(|f| f as i64))
            .collect()
    } else {
        std::collections::HashSet::new()
    };
    let (lo, hi, sign) = if start <= end {
        (start, end, 1i64)
    } else {
        (end, start, -1)
    };
    let count: i64 = (lo..=hi)
        .filter(|&d| !is_weekend_intl(d, &mask) && !holidays.contains(&d))
        .count() as i64;
    Ok(Variant::Integer(count * sign))
}

fn func_workday_intl(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 4 {
        return Err("WORKDAY.INTL requires 2 to 4 arguments".into());
    }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let days = to_float(&evaluate(&args[1], cells)?)? as i64;
    let mask = if args.len() >= 3 {
        parse_weekend_mask(&evaluate(&args[2], cells)?)?
    } else {
        [false, false, false, false, false, true, true]
    };
    let holidays: std::collections::HashSet<i64> = if args.len() == 4 {
        collect_values(&args[3], cells)?
            .iter()
            .filter_map(|v| to_float(v).ok().map(|f| f as i64))
            .collect()
    } else {
        std::collections::HashSet::new()
    };
    let mut current = start;
    let mut remaining = days.abs();
    let step = if days >= 0 { 1i64 } else { -1 };
    while remaining > 0 {
        current += step;
        if !is_weekend_intl(current, &mask) && !holidays.contains(&current) {
            remaining -= 1;
        }
    }
    Ok(Variant::Date(current))
}

// ── Logic ─────────────────────────────────────────────────────────────────────

fn func_switch(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 {
        return Err("SWITCH requires at least 3 arguments".into());
    }
    let expr = evaluate(&args[0], cells)?;
    if matches!(&expr, Variant::Array(_) | Variant::VbaArray(_)) {
        let expression_values = variant_values(expr);
        let count = expression_values.len();
        let has_default = args.len().is_multiple_of(2);
        let pair_end = if has_default {
            args.len() - 1
        } else {
            args.len()
        };
        let mut selected = vec![None; count];

        let broadcast = |value: Variant| -> Result<Vec<Variant>, ExcelError> {
            let values = variant_values(value);
            if values.len() == 1 {
                Ok(vec![values[0].clone(); count])
            } else if values.len() == count {
                Ok(values)
            } else {
                Err(ExcelError::Value)
            }
        };

        let mut i = 1;
        while i + 1 < pair_end {
            let match_values = broadcast(evaluate(&args[i], cells)?);
            let match_values = match match_values {
                Ok(values) => values,
                Err(error) => return Ok(Variant::Error(error)),
            };
            let mut matched = vec![false; count];
            for index in 0..count {
                if selected[index].is_some() {
                    continue;
                }
                if let Variant::Error(error) = &expression_values[index] {
                    selected[index] = Some(Variant::Error(error.clone()));
                    continue;
                }
                if !matches!(&match_values[index], Variant::Error(_))
                    && variant_eq(&expression_values[index], &match_values[index])
                {
                    matched[index] = true;
                }
            }
            if matched.iter().any(|matched| *matched) {
                let result_values = match broadcast(evaluate(&args[i + 1], cells)?) {
                    Ok(values) => values,
                    Err(error) => return Ok(Variant::Error(error)),
                };
                for index in 0..count {
                    if matched[index] {
                        selected[index] = Some(result_values[index].clone());
                    }
                }
            }
            i += 2;
        }

        if selected.iter().any(Option::is_none) {
            if has_default {
                let default_values = match broadcast(evaluate(&args[args.len() - 1], cells)?) {
                    Ok(values) => values,
                    Err(error) => return Ok(Variant::Error(error)),
                };
                for (index, value) in selected.iter_mut().enumerate() {
                    if value.is_none() {
                        *value = Some(default_values[index].clone());
                    }
                }
            } else {
                for value in &mut selected {
                    if value.is_none() {
                        *value = Some(Variant::Error(ExcelError::NA));
                    }
                }
            }
        }
        return Ok(Variant::Array(
            selected.into_iter().map(Option::unwrap).collect(),
        ));
    }
    let mut i = 1;
    while i + 1 < args.len() {
        let v = evaluate(&args[i], cells)?;
        if variant_eq(&expr, &v) {
            return evaluate(&args[i + 1], cells);
        }
        i += 2;
    }
    // odd remaining arg = default
    if args.len().is_multiple_of(2) {
        evaluate(&args[args.len() - 1], cells)
    } else {
        Err("SWITCH: no match found".into())
    }
}

fn func_xor(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("XOR requires at least 1 argument".into());
    }
    let mut count = 0;
    let mut first_error = None;
    for arg in args {
        match evaluate(arg, cells)? {
            Variant::Error(error) => {
                first_error.get_or_insert(error);
            }
            Variant::Array(values) => {
                for value in values {
                    match value {
                        Variant::Error(error) => {
                            first_error.get_or_insert(error);
                        }
                        value if is_truthy(&value) => count += 1,
                        _ => {}
                    }
                }
            }
            Variant::VbaArray(array) => {
                for value in array.elements {
                    match value {
                        Variant::Error(error) => {
                            first_error.get_or_insert(error);
                        }
                        value if is_truthy(&value) => count += 1,
                        _ => {}
                    }
                }
            }
            value if is_truthy(&value) => count += 1,
            _ => {}
        }
    }
    if let Some(error) = first_error {
        return Ok(Variant::Error(error));
    }
    Ok(Variant::Boolean(count % 2 == 1))
}

// ── Lookup ────────────────────────────────────────────────────────────────────

fn func_choose(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("CHOOSE requires at least 2 arguments".into());
    }
    let choose_index = |index: &Variant| -> Result<Option<usize>, String> {
        let index = to_float(index)?;
        if !index.is_finite() || index.fract() != 0.0 {
            return Ok(None);
        }
        let idx = index as i64;
        if idx < 1 || idx >= args.len() as i64 {
            return Ok(None);
        }
        Ok(Some(idx as usize))
    };
    let choose_one = |index: &Variant| -> Result<Variant, String> {
        let Some(idx) = choose_index(index)? else {
            return Ok(Variant::Error(ExcelError::Value));
        };
        // A selected reference is an array result in worksheet context.  Do
        // not force it through the scalar evaluator, which would reject a
        // Range.
        match &args[idx] {
            FormulaExpr::Range { .. } => Ok(Variant::Array(collect_values(&args[idx], cells)?)),
            _ => evaluate(&args[idx], cells),
        }
    };

    match evaluate(&args[0], cells)? {
        Variant::Array(indices) => {
            let mut result = Vec::new();
            for index in indices {
                // An invalid index is a CHOOSE argument error, not an
                // elementwise worksheet error.  Validate the complete index
                // array before allowing selected expressions to contribute
                // their own errors to the result array.
                if choose_index(&index)?.is_none() {
                    return Ok(Variant::Error(ExcelError::Value));
                }
                match choose_one(&index)? {
                    Variant::Array(values) => result.extend(values),
                    Variant::VbaArray(array) => result.extend(array.elements),
                    Variant::Error(error) => result.push(Variant::Error(error)),
                    value => result.push(value),
                }
            }
            Ok(Variant::Array(result))
        }
        Variant::VbaArray(indices) => {
            let mut result = Vec::new();
            for index in indices.elements {
                if choose_index(&index)?.is_none() {
                    return Ok(Variant::Error(ExcelError::Value));
                }
                match choose_one(&index)? {
                    Variant::Array(values) => result.extend(values),
                    Variant::VbaArray(array) => result.extend(array.elements),
                    Variant::Error(error) => result.push(Variant::Error(error)),
                    value => result.push(value),
                }
            }
            Ok(Variant::Array(result))
        }
        index => choose_one(&index),
    }
}

fn func_column(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    match args.first() {
        None => Ok(Variant::Integer(1)),
        Some(FormulaExpr::CellRef { col, .. }) => Ok(Variant::Integer(*col as i64)),
        Some(FormulaExpr::Range { c1, .. }) => Ok(Variant::Integer(*c1 as i64)),
        Some(other) => {
            evaluate(other, cells)?;
            Ok(Variant::Integer(1))
        }
    }
}

fn func_columns(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("COLUMNS requires 1 argument".into());
    }
    let columns = formula_array_dimensions(&args[0], cells)?.map_or(1, |(_, columns)| columns);
    Ok(Variant::Integer(columns as i64))
}

fn func_lookup(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("LOOKUP requires 2 or 3 arguments".into());
    }
    let key = evaluate(&args[0], cells)?;
    let lookup_shape = expression_shape(&args[1], cells)?;
    if lookup_shape.0 != 1 && lookup_shape.1 != 1 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let lookup = collect_values(&args[1], cells)?;
    let result = if args.len() == 3 {
        if expression_shape(&args[2], cells)? != lookup_shape {
            return Ok(Variant::Error(ExcelError::Value));
        }
        collect_values(&args[2], cells)?
    } else {
        lookup.clone()
    };
    if result.len() != lookup.len() {
        return Ok(Variant::Error(ExcelError::Value));
    }

    let mut sorted = true;
    for pair in lookup.windows(2) {
        let order = match variant_cmp(&pair[0], &pair[1]) {
            Ok(order) => order,
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        };
        if order == std::cmp::Ordering::Greater {
            sorted = false;
            break;
        }
    }
    if !sorted {
        return Ok(Variant::Error(ExcelError::NA));
    }

    let mut best: Option<usize> = None;
    for (i, v) in lookup.iter().enumerate() {
        match variant_cmp(v, &key) {
            Ok(std::cmp::Ordering::Less) | Ok(std::cmp::Ordering::Equal) => best = Some(i),
            _ => break,
        }
    }
    best.and_then(|i| result.get(i).cloned())
        .map(Ok)
        .unwrap_or(Ok(Variant::Error(ExcelError::NA)))
}

fn func_xmatch(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 4 {
        return Err("XMATCH requires 2 to 4 arguments".into());
    }
    let key = evaluate(&args[0], cells)?;
    let lookup_shape = expression_shape(&args[1], cells)?;
    if lookup_shape.0 != 1 && lookup_shape.1 != 1 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let lookup = collect_values(&args[1], cells)?;
    let match_mode = if args.len() >= 3 {
        match lookup_mode(&evaluate(&args[2], cells)?) {
            Ok(mode) => mode,
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        }
    } else {
        0
    };
    let search_mode = if args.len() >= 4 {
        match lookup_mode(&evaluate(&args[3], cells)?) {
            Ok(mode) => mode,
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        }
    } else {
        1
    };

    if matches!(search_mode, 2 | -2) {
        if match_mode == 2 {
            return Err("XMATCH: wildcard match_mode is incompatible with binary search".into());
        }
        match validate_lookup_sorted(&lookup, search_mode == 2) {
            Ok(true) => {}
            Ok(false) => return Ok(Variant::Error(ExcelError::NA)),
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        }
        let index = xlookup_binary_index(&lookup, &key, search_mode == 2, match_mode)?;
        return Ok(index
            .map(|i| Variant::Integer((i + 1) as i64))
            .unwrap_or(Variant::Error(ExcelError::NA)));
    }

    let iter: Box<dyn Iterator<Item = usize>> = match search_mode {
        -1 => Box::new((0..lookup.len()).rev()),
        1 => Box::new(0..lookup.len()),
        mode => return Err(format!("XMATCH: unsupported search_mode {}", mode)),
    };

    match match_mode {
        0 => {
            for i in iter {
                if variant_eq(&lookup[i], &key) {
                    return Ok(Variant::Integer((i + 1) as i64));
                }
            }
        }
        -1 => {
            let key_f = to_float(&key)?;
            let mut best: Option<(usize, f64)> = None;
            for (i, l) in lookup.iter().enumerate() {
                if let Ok(v) = to_float(l)
                    && v <= key_f
                    && best.is_none_or(|(_, bv)| v > bv)
                {
                    best = Some((i, v));
                }
            }
            if let Some((i, _)) = best {
                return Ok(Variant::Integer((i + 1) as i64));
            }
        }
        1 => {
            let key_f = to_float(&key)?;
            let mut best: Option<(usize, f64)> = None;
            for (i, l) in lookup.iter().enumerate() {
                if let Ok(v) = to_float(l)
                    && v >= key_f
                    && best.is_none_or(|(_, bv)| v < bv)
                {
                    best = Some((i, v));
                }
            }
            if let Some((i, _)) = best {
                return Ok(Variant::Integer((i + 1) as i64));
            }
        }
        2 => {
            let pattern = to_str(&key);
            for i in iter {
                if wildcard_match(&to_str(&lookup[i]), &pattern) {
                    return Ok(Variant::Integer((i + 1) as i64));
                }
            }
        }
        m => return Err(format!("XMATCH: unsupported match_mode {}", m)),
    }
    Ok(Variant::Error(ExcelError::NA))
}

// ── Info functions ────────────────────────────────────────────────────────────

fn func_isblank(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISBLANK requires 1 argument".into());
    }
    Ok(Variant::Boolean(matches!(
        evaluate(&args[0], cells)?,
        Variant::Empty
    )))
}

fn func_isformula(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISFORMULA requires 1 argument".into());
    }
    match &args[0] {
        FormulaExpr::CellRef {
            col, row, sheet, ..
        } if sheet.is_none() => Ok(Variant::Boolean(
            cells
                .get(&(*row, *col))
                .is_some_and(|cell| cell.formula.is_some()),
        )),
        FormulaExpr::Range {
            c1,
            r1,
            c2,
            r2,
            sheet,
            ..
        } if sheet.is_none() => Ok(Variant::Array(
            (*r1..=*r2)
                .flat_map(|row| {
                    (*c1..=*c2).map(move |col| {
                        Variant::Boolean(
                            cells
                                .get(&(row, col))
                                .is_some_and(|cell| cell.formula.is_some()),
                        )
                    })
                })
                .collect(),
        )),
        FormulaExpr::CellRef { .. } | FormulaExpr::Range { .. } => {
            Err("ISFORMULA: cross-sheet references require workbook context".into())
        }
        _ => Ok(Variant::Boolean(false)),
    }
}

fn func_isomitted(
    args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISOMITTED requires 1 argument".into());
    }
    Ok(Variant::Boolean(
        matches!(args[0], FormulaExpr::Omitted)
            || matches!(&args[0], FormulaExpr::FuncCall { name, args } if args.is_empty() && binding_is_omitted(name)),
    ))
}

fn func_sheet(
    args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() > 1 {
        return Err("SHEET requires 0 or 1 argument".into());
    }
    let (sheet_number, _) = sheet_context();
    if let Some(reference) = args.first() {
        match reference {
            FormulaExpr::CellRef { sheet, .. } | FormulaExpr::Range { sheet, .. }
                if sheet.is_some() =>
            {
                return Ok(Variant::Error(ExcelError::NA));
            }
            FormulaExpr::CellRef { .. }
            | FormulaExpr::Range { .. }
            | FormulaExpr::FuncCall { .. } => {}
            _ => return Ok(Variant::Error(ExcelError::Value)),
        }
    }
    Ok(Variant::Integer(sheet_number as i64))
}

fn func_sheets(
    args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() > 1 {
        return Err("SHEETS requires 0 or 1 argument".into());
    }
    let (_, sheet_count) = sheet_context();
    if let Some(reference) = args.first() {
        match reference {
            FormulaExpr::CellRef { sheet, .. } | FormulaExpr::Range { sheet, .. }
                if sheet.is_some() =>
            {
                return Ok(Variant::Integer(1));
            }
            FormulaExpr::CellRef { .. }
            | FormulaExpr::Range { .. }
            | FormulaExpr::FuncCall { .. } => return Ok(Variant::Integer(1)),
            _ => return Ok(Variant::Error(ExcelError::Value)),
        }
    }
    Ok(Variant::Integer(sheet_count as i64))
}

fn func_isref(
    args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISREF requires 1 argument".into());
    }
    let is_reference = match &args[0] {
        FormulaExpr::CellRef { .. } | FormulaExpr::Range { .. } => true,
        FormulaExpr::FuncCall { name, .. } => {
            name.eq_ignore_ascii_case("INDIRECT") || name.eq_ignore_ascii_case("OFFSET")
        }
        _ => false,
    };
    Ok(Variant::Boolean(is_reference))
}

fn func_areas(
    args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("AREAS requires 1 argument".into());
    }
    let is_reference = match &args[0] {
        FormulaExpr::CellRef { .. } | FormulaExpr::Range { .. } => true,
        FormulaExpr::FuncCall { name, .. } => {
            name.eq_ignore_ascii_case("INDIRECT") || name.eq_ignore_ascii_case("OFFSET")
        }
        _ => false,
    };
    if is_reference {
        Ok(Variant::Integer(1))
    } else {
        Ok(Variant::Error(ExcelError::Value))
    }
}

fn func_info(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("INFO requires 1 argument".into());
    }
    let info_type = to_str(&evaluate(&args[0], cells)?).to_ascii_lowercase();
    match info_type.as_str() {
        "system" => Ok(Variant::Str(
            if cfg!(target_os = "macos") {
                "mac"
            } else {
                "pcdos"
            }
            .into(),
        )),
        "osversion" => Ok(Variant::Str(std::env::consts::OS.into())),
        "release" => Ok(Variant::Str(env!("CARGO_PKG_VERSION").into())),
        "recalc" => Ok(Variant::Str("Automatic".into())),
        // Directory, memory, and file-count values are host-sensitive and can
        // disclose information in a headless service, so they remain explicit
        // unsupported results rather than guessed placeholders.
        _ => Ok(Variant::Error(ExcelError::NA)),
    }
}

fn func_iserror(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISERROR requires 1 argument".into());
    }
    Ok(Variant::Boolean(matches!(
        evaluate(&args[0], cells),
        Ok(Variant::Error(_)) | Err(_)
    )))
}

fn func_isna(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISNA requires 1 argument".into());
    }
    Ok(Variant::Boolean(matches!(
        evaluate(&args[0], cells),
        Ok(Variant::Error(ExcelError::NA))
    )))
}

fn func_isnumber(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISNUMBER requires 1 argument".into());
    }
    Ok(Variant::Boolean(matches!(
        evaluate(&args[0], cells)?,
        Variant::Integer(_) | Variant::Float(_)
    )))
}

fn func_istext(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISTEXT requires 1 argument".into());
    }
    Ok(Variant::Boolean(matches!(
        evaluate(&args[0], cells)?,
        Variant::Str(_)
    )))
}

fn func_islogical(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISLOGICAL requires 1 argument".into());
    }
    Ok(Variant::Boolean(matches!(
        evaluate(&args[0], cells)?,
        Variant::Boolean(_)
    )))
}

fn func_isnontext(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ISNONTEXT requires 1 argument".into());
    }
    Ok(Variant::Boolean(!matches!(
        evaluate(&args[0], cells).unwrap_or(Variant::Empty),
        Variant::Str(_)
    )))
}

// ── N / NA / TYPE / ERROR.TYPE / FORMULATEXT / CELL ──────────────────────────

fn func_na(
    args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !args.is_empty() {
        return Err("NA takes no arguments".into());
    }
    Ok(Variant::Error(ExcelError::NA))
}

fn func_n(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("N requires 1 argument".into());
    }
    Ok(match evaluate(&args[0], cells)? {
        Variant::Integer(n) => Variant::Integer(n),
        Variant::Float(f) => Variant::Float(f),
        Variant::Date(s) => Variant::Integer(s),
        Variant::Boolean(b) => Variant::Integer(if b { 1 } else { 0 }),
        Variant::Error(e) => Variant::Error(e),
        _ => Variant::Integer(0),
    })
}

fn func_type_fn(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("TYPE requires 1 argument".into());
    }
    let code = match evaluate(&args[0], cells)? {
        Variant::Integer(_)
        | Variant::Float(_)
        | Variant::Date(_)
        | Variant::Empty
        | Variant::Null => 1,
        Variant::Str(_) => 2,
        Variant::Boolean(_) => 4,
        Variant::Error(_) => 16,
        Variant::Array(_) | Variant::VbaArray(_) => 64,
        Variant::Record(_) => 64,
    };
    Ok(Variant::Integer(code))
}

fn func_error_type(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ERROR.TYPE requires 1 argument".into());
    }
    // Evaluation errors map to #VALUE! (code 3); non-errors return #N/A
    let v = evaluate(&args[0], cells).unwrap_or(Variant::Error(ExcelError::Value));
    let code: i64 = match v {
        Variant::Error(ExcelError::Null) => 1,
        Variant::Error(ExcelError::DivZero) => 2,
        Variant::Error(ExcelError::Value) => 3,
        Variant::Error(ExcelError::Ref) => 4,
        Variant::Error(ExcelError::Name) => 5,
        Variant::Error(ExcelError::Num) => 6,
        Variant::Error(ExcelError::NA) => 7,
        _ => return Ok(Variant::Error(ExcelError::NA)),
    };
    Ok(Variant::Integer(code))
}

fn func_formulatext(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("FORMULATEXT requires 1 argument".into());
    }
    let (row, col) = match &args[0] {
        FormulaExpr::CellRef { row, col, .. } => (*row, *col),
        FormulaExpr::Range { r1, c1, .. } => (*r1, *c1),
        _ => return Ok(Variant::Error(ExcelError::Value)),
    };
    match cells.get(&(row, col)).and_then(|c| c.formula.as_ref()) {
        Some(f) => Ok(Variant::Str(f.clone())),
        None => Ok(Variant::Error(ExcelError::NA)),
    }
}

pub(crate) fn col_to_letter(mut col: u32) -> String {
    let mut s = String::new();
    while col > 0 {
        col -= 1;
        s.insert(0, (b'A' + (col % 26) as u8) as char);
        col /= 26;
    }
    s
}

fn func_cell(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("CELL requires 1 or 2 arguments".into());
    }
    let info = to_str(&evaluate(&args[0], cells)?).to_lowercase();
    let (row, col) = if args.len() == 2 {
        match &args[1] {
            FormulaExpr::CellRef { row, col, .. } => (*row, *col),
            FormulaExpr::Range { r1, c1, .. } => (*r1, *c1),
            _ => (1u32, 1u32),
        }
    } else {
        (1u32, 1u32)
    };

    Ok(match info.as_str() {
        "address" => Variant::Str(format!("${}${}", col_to_letter(col), row)),
        "col" => Variant::Integer(col as i64),
        "row" => Variant::Integer(row as i64),
        "contents" => cell_val(cells, row, col),
        "type" => {
            let v = cell_val(cells, row, col);
            Variant::Str(match v {
                Variant::Empty => "b".into(),
                Variant::Str(_) => "l".into(),
                _ => "v".into(),
            })
        }
        "filename" | "prefix" | "format" => Variant::Str(String::new()),
        "protect" | "parentheses" | "color" => Variant::Integer(0),
        "width" => Variant::Integer(8),
        _ => Variant::Integer(0),
    })
}

// ── Statistics ────────────────────────────────────────────────────────────────

fn collect_nums(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Vec<f64>, String> {
    Ok(flatten_values(collect_all(args, cells)?)
        .into_iter()
        .filter_map(|value| as_f64(&value))
        .collect())
}

fn func_stdev_s(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = collect_nums(args, cells)?;
    if nums.len() < 2 {
        return Err("STDEV requires at least 2 values".into());
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let var = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64;
    Ok(Variant::Float(var.sqrt()))
}

fn func_stdev_p(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = collect_nums(args, cells)?;
    if nums.is_empty() {
        return Err("STDEVP requires at least 1 value".into());
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let var = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / nums.len() as f64;
    Ok(Variant::Float(var.sqrt()))
}

fn func_var_s(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = collect_nums(args, cells)?;
    if nums.len() < 2 {
        return Err("VAR requires at least 2 values".into());
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    Ok(Variant::Float(
        nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64,
    ))
}

fn func_var_p(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = collect_nums(args, cells)?;
    if nums.is_empty() {
        return Err("VARP requires at least 1 value".into());
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    Ok(Variant::Float(
        nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / nums.len() as f64,
    ))
}

fn func_var_a(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let values = match collect_a_values(args, cells) {
        Ok(values) => values,
        Err(error) => return Ok(Variant::Error(error)),
    };
    if values.len() < 2 {
        return Err("VARA requires at least 2 values".into());
    }
    let mean = values.iter().sum::<f64>() / values.len() as f64;
    Ok(Variant::Float(
        values.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (values.len() - 1) as f64,
    ))
}

fn func_var_pa(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let values = match collect_a_values(args, cells) {
        Ok(values) => values,
        Err(error) => return Ok(Variant::Error(error)),
    };
    if values.is_empty() {
        return Err("VARPA requires at least 1 value".into());
    }
    let mean = values.iter().sum::<f64>() / values.len() as f64;
    Ok(Variant::Float(
        values.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / values.len() as f64,
    ))
}

fn func_ftest(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("FTEST requires 2 arguments".into());
    }
    let first = collect_nums(std::slice::from_ref(&args[0]), cells)?;
    let second = collect_nums(std::slice::from_ref(&args[1]), cells)?;
    if first.len() < 2 || second.len() < 2 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let variance = |values: &[f64]| {
        let mean = values.iter().sum::<f64>() / values.len() as f64;
        values
            .iter()
            .map(|value| (value - mean).powi(2))
            .sum::<f64>()
            / (values.len() - 1) as f64
    };
    let first_variance = variance(&first);
    let second_variance = variance(&second);
    if !first_variance.is_finite()
        || !second_variance.is_finite()
        || first_variance <= 0.0
        || second_variance <= 0.0
    {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let ratio = (first_variance / second_variance).max(second_variance / first_variance);
    let p =
        (2.0 * (1.0 - f_cdf(ratio, (first.len() - 1) as f64, (second.len() - 1) as f64))).min(1.0);
    Ok(Variant::Float(p))
}

fn func_chitest(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("CHITEST requires 2 arguments".into());
    }
    let observed = collect_nums(std::slice::from_ref(&args[0]), cells)?;
    let expected = collect_nums(std::slice::from_ref(&args[1]), cells)?;
    if observed.is_empty() || observed.len() != expected.len() {
        return Ok(Variant::Error(ExcelError::NA));
    }
    let mut statistic = 0.0;
    for (actual, expected) in observed.iter().zip(expected.iter()) {
        if !actual.is_finite() || !expected.is_finite() || *expected <= 0.0 {
            return Ok(Variant::Error(ExcelError::Num));
        }
        statistic += (actual - expected).powi(2) / expected;
    }
    let dimensions = |expr: &FormulaExpr| match expr {
        FormulaExpr::Range { r1, r2, c1, c2, .. } => Some((
            ((*r1).max(*r2) - (*r1).min(*r2) + 1) as usize,
            ((*c1).max(*c2) - (*c1).min(*c2) + 1) as usize,
        )),
        _ => None,
    };
    let (rows, cols) = dimensions(&args[0]).unwrap_or((1, observed.len()));
    if dimensions(&args[1]).is_some_and(|shape| shape != (rows, cols)) {
        return Ok(Variant::Error(ExcelError::NA));
    }
    let degrees = (rows.saturating_sub(1) * cols.saturating_sub(1)) as f64;
    if degrees <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(1.0 - chisq_cdf(statistic, degrees)))
}

fn func_stdev_a(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    Ok(Variant::Float(to_float(&func_var_a(args, cells)?)?.sqrt()))
}

fn func_stdev_pa(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    Ok(Variant::Float(to_float(&func_var_pa(args, cells)?)?.sqrt()))
}

// ── Statistical: CORREL / COVARIANCE / NORM.DIST / NORM.INV / T.DIST ─────────

/// Collect paired numeric values from two equal-length range arguments.
fn collect_paired(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    fname: &str,
) -> Result<(Vec<f64>, Vec<f64>), String> {
    if args.len() != 2 {
        return Err(format!("{fname} requires 2 arguments"));
    }
    let a: Vec<f64> = flatten_values(collect_values(&args[0], cells)?)
        .into_iter()
        .filter_map(|v| as_f64(&v))
        .collect();
    let b: Vec<f64> = flatten_values(collect_values(&args[1], cells)?)
        .into_iter()
        .filter_map(|v| as_f64(&v))
        .collect();
    if a.len() != b.len() || a.is_empty() {
        return Err(format!("{fname}: arrays must have equal length"));
    }
    Ok((a, b))
}

fn func_correl(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (a, b) = collect_paired(args, cells, "CORREL")?;
    let n = a.len() as f64;
    let ma = a.iter().sum::<f64>() / n;
    let mb = b.iter().sum::<f64>() / n;
    let cov: f64 = a
        .iter()
        .zip(b.iter())
        .map(|(x, y)| (x - ma) * (y - mb))
        .sum();
    let sa: f64 = a.iter().map(|x| (x - ma).powi(2)).sum::<f64>().sqrt();
    let sb: f64 = b.iter().map(|y| (y - mb).powi(2)).sum::<f64>().sqrt();
    if sa == 0.0 || sb == 0.0 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    Ok(Variant::Float(cov / (sa * sb)))
}

fn regression_slope(a: &[f64], b: &[f64]) -> Result<f64, String> {
    if a.len() < 2 || a.len() != b.len() {
        return Err("regression requires two equal arrays with at least 2 values".into());
    }
    let ma = a.iter().sum::<f64>() / a.len() as f64;
    let mb = b.iter().sum::<f64>() / b.len() as f64;
    let denominator = b.iter().map(|x| (x - mb).powi(2)).sum::<f64>();
    if denominator == 0.0 {
        return Err("regression: known_x values must vary".into());
    }
    Ok(a.iter()
        .zip(b)
        .map(|(y, x)| (x - mb) * (y - ma))
        .sum::<f64>()
        / denominator)
}

fn func_slope(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (known_y, known_x) = collect_paired(args, cells, "SLOPE")?;
    regression_slope(&known_y, &known_x).map(Variant::Float)
}

fn func_intercept(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (known_y, known_x) = collect_paired(args, cells, "INTERCEPT")?;
    let slope = regression_slope(&known_y, &known_x)?;
    let my = known_y.iter().sum::<f64>() / known_y.len() as f64;
    let mx = known_x.iter().sum::<f64>() / known_x.len() as f64;
    Ok(Variant::Float(my - slope * mx))
}

fn func_rsq(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (known_y, known_x) = collect_paired(args, cells, "RSQ")?;
    let slope = regression_slope(&known_y, &known_x)?;
    let my = known_y.iter().sum::<f64>() / known_y.len() as f64;
    let mx = known_x.iter().sum::<f64>() / known_x.len() as f64;
    let ss_tot = known_y.iter().map(|y| (y - my).powi(2)).sum::<f64>();
    if ss_tot == 0.0 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let intercept = my - slope * mx;
    let ss_res = known_y
        .iter()
        .zip(&known_x)
        .map(|(y, x)| (y - (slope * x + intercept)).powi(2))
        .sum::<f64>();
    Ok(Variant::Float(1.0 - ss_res / ss_tot))
}

fn func_forecast_linear(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("FORECAST.LINEAR requires 3 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let (known_y, known_x) = collect_paired(&args[1..], cells, "FORECAST.LINEAR")?;
    let slope = regression_slope(&known_y, &known_x)?;
    let my = known_y.iter().sum::<f64>() / known_y.len() as f64;
    let mx = known_x.iter().sum::<f64>() / known_x.len() as f64;
    Ok(Variant::Float(my + slope * (x - mx)))
}

fn forecast_ets_series(
    values_expr: &FormulaExpr,
    timeline_expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<(Vec<f64>, Vec<f64>), String> {
    let values: Vec<f64> = collect_values(values_expr, cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    let timeline: Vec<f64> = collect_values(timeline_expr, cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
    if values.len() != timeline.len() || values.len() < 2 {
        return Err(format!(
            "{name}: values and timeline must have equal length >= 2"
        ));
    }
    let mut pairs: Vec<(f64, f64)> = timeline.into_iter().zip(values).collect();
    if pairs
        .iter()
        .any(|(time, value)| !time.is_finite() || !value.is_finite())
    {
        return Err(format!("{name}: values and timeline must be finite"));
    }
    pairs.sort_by(|left, right| left.0.partial_cmp(&right.0).unwrap_or(Ordering::Equal));
    if pairs.windows(2).any(|window| window[0].0 == window[1].0) {
        return Err(format!("{name}: timeline cannot contain duplicates"));
    }
    let step = pairs[1].0 - pairs[0].0;
    if step <= 0.0
        || pairs.windows(2).any(|window| {
            let current = window[1].0 - window[0].0;
            (current - step).abs() > step.abs().max(1.0) * 1e-9
        })
    {
        return Err(format!(
            "{name}: timeline must have a constant positive step"
        ));
    }
    Ok((
        pairs.iter().map(|(_, value)| *value).collect(),
        pairs.iter().map(|(time, _)| *time).collect(),
    ))
}

fn forecast_ets_seasonality(values: &[f64]) -> usize {
    if values.len() < 4 {
        return 1;
    }
    let mean = values.iter().sum::<f64>() / values.len() as f64;
    let variance = values
        .iter()
        .map(|value| (value - mean).powi(2))
        .sum::<f64>();
    if variance == 0.0 {
        return 1;
    }
    let max_lag = (values.len() / 2).min(24);
    let mut best = (1, f64::INFINITY);
    for lag in 2..=max_lag {
        let error = values
            .iter()
            .skip(lag)
            .zip(values.iter())
            .map(|(current, prior)| (current - prior).powi(2))
            .sum::<f64>();
        if error < best.1 {
            best = (lag, error);
        }
    }
    if best.1 <= variance * 0.2 { best.0 } else { 1 }
}

fn forecast_ets_optional(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<usize, String> {
    let seasonality = if let Some(arg) = args.get(3) {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=8760.0).contains(&value) {
            return Err("FORECAST.ETS: seasonality must be 0 or a positive integer <= 8760".into());
        }
        value as usize
    } else {
        1
    };
    if let Some(arg) = args.get(4) {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=1.0).contains(&value) {
            return Err("FORECAST.ETS: data_completion must be 0 or 1".into());
        }
    }
    if let Some(arg) = args.get(5) {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=6.0).contains(&value) {
            return Err("FORECAST.ETS: aggregation must be an integer from 0 to 6".into());
        }
    }
    Ok(seasonality)
}

fn func_forecast_ets(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(3..=6).contains(&args.len()) {
        return Err("FORECAST.ETS requires 3 to 6 arguments".into());
    }
    let target = to_float(&evaluate(&args[0], cells)?)?;
    let (values, timeline) = forecast_ets_series(&args[1], &args[2], cells, "FORECAST.ETS")?;
    let requested = forecast_ets_optional(args, cells)?;
    if requested > values.len() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let seasonality = if requested == 0 {
        0
    } else if args.len() >= 4 && requested > 1 {
        requested
    } else {
        forecast_ets_seasonality(&values)
    };
    let step = timeline[1] - timeline[0];
    let ahead = (target - timeline[2..].last().copied().unwrap_or(timeline[0])) / step;
    if !target.is_finite() || ahead <= 0.0 || (ahead - ahead.round()).abs() > 1e-8 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let horizon = ahead.round() as usize;
    let result = if seasonality == 0 {
        func_forecast_linear(
            &[
                FormulaExpr::Number(target),
                args[1].clone(),
                args[2].clone(),
            ],
            cells,
        )?
    } else if seasonality == 1 {
        Variant::Float(*values.last().unwrap_or(&0.0))
    } else {
        let index = values.len() - seasonality + (horizon - 1) % seasonality;
        Variant::Float(values[index])
    };
    Ok(result)
}

fn func_forecast_ets_seasonality(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(2..=4).contains(&args.len()) {
        return Err("FORECAST.ETS.SEASONALITY requires 2 to 4 arguments".into());
    }
    let (values, _) = forecast_ets_series(&args[0], &args[1], cells, "FORECAST.ETS.SEASONALITY")?;
    if let Some(arg) = args.get(2) {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=1.0).contains(&value) {
            return Err("FORECAST.ETS.SEASONALITY: data_completion must be 0 or 1".into());
        }
    }
    if let Some(arg) = args.get(3) {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=6.0).contains(&value) {
            return Err(
                "FORECAST.ETS.SEASONALITY: aggregation must be an integer from 0 to 6".into(),
            );
        }
    }
    Ok(Variant::Integer(forecast_ets_seasonality(&values) as i64))
}

fn forecast_ets_requested_seasonality(
    requested: usize,
    explicit: bool,
    values: &[f64],
    name: &str,
) -> Result<usize, String> {
    if requested > values.len() {
        return Err(format!("{name}: seasonality exceeds the values length"));
    }
    Ok(if requested == 0 {
        0
    } else if explicit && requested > 1 {
        requested
    } else {
        forecast_ets_seasonality(values)
    })
}

fn forecast_ets_point(
    target: f64,
    values: &[f64],
    timeline: &[f64],
    seasonality: usize,
    name: &str,
) -> Result<f64, String> {
    let step = timeline[1] - timeline[0];
    let ahead = (target - *timeline.last().unwrap_or(&timeline[0])) / step;
    if !target.is_finite() || ahead <= 0.0 || (ahead - ahead.round()).abs() > 1e-8 {
        return Err(format!(
            "{name}: target must be a future integral timeline step"
        ));
    }
    let horizon = ahead.round() as usize;
    if seasonality == 0 {
        let known_y = values.to_vec();
        let known_x = timeline.to_vec();
        let slope = regression_slope(&known_y, &known_x)?;
        let my = known_y.iter().sum::<f64>() / known_y.len() as f64;
        let mx = known_x.iter().sum::<f64>() / known_x.len() as f64;
        Ok(my + slope * (target - mx))
    } else if seasonality == 1 {
        Ok(*values.last().unwrap_or(&0.0))
    } else {
        let index = values.len() - seasonality + (horizon - 1) % seasonality;
        Ok(values[index])
    }
}

fn forecast_ets_metric_residuals(values: &[f64], seasonality: usize) -> Vec<f64> {
    if seasonality == 0 {
        let x: Vec<f64> = (0..values.len()).map(|i| i as f64).collect();
        let slope = regression_slope(values, &x).unwrap_or(0.0);
        let mean_x = (values.len() - 1) as f64 / 2.0;
        let mean_y = values.iter().sum::<f64>() / values.len() as f64;
        let intercept = mean_y - slope * mean_x;
        values
            .iter()
            .enumerate()
            .map(|(i, value)| value - (intercept + slope * i as f64))
            .collect()
    } else {
        values
            .iter()
            .enumerate()
            .skip(seasonality)
            .map(|(i, value)| value - values[i - seasonality])
            .collect()
    }
}

fn forecast_ets_optional_at(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    seasonality_index: usize,
    completion_index: usize,
    aggregation_index: usize,
    name: &str,
) -> Result<(usize, bool), String> {
    let requested = if let Some(arg) = args.get(seasonality_index) {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=8760.0).contains(&value) {
            return Err(format!(
                "{name}: seasonality must be 0 or a positive integer <= 8760"
            ));
        }
        value as usize
    } else {
        1
    };
    if let Some(arg) = args.get(completion_index) {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=1.0).contains(&value) {
            return Err(format!("{name}: data_completion must be 0 or 1"));
        }
    }
    if let Some(arg) = args.get(aggregation_index) {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=6.0).contains(&value) {
            return Err(format!(
                "{name}: aggregation must be an integer from 0 to 6"
            ));
        }
    }
    Ok((requested, args.len() > seasonality_index))
}

fn func_forecast_ets_confint(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(3..=7).contains(&args.len()) {
        return Err("FORECAST.ETS.CONFINT requires 3 to 7 arguments".into());
    }
    let target = to_float(&evaluate(&args[0], cells)?)?;
    let confidence = if let Some(arg) = args.get(3) {
        to_float(&evaluate(arg, cells)?)?
    } else {
        0.95
    };
    if !(0.0..1.0).contains(&confidence) || !confidence.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let (values, timeline) =
        forecast_ets_series(&args[1], &args[2], cells, "FORECAST.ETS.CONFINT")?;
    let (requested, explicit) =
        forecast_ets_optional_at(args, cells, 4, 5, 6, "FORECAST.ETS.CONFINT")?;
    let seasonality =
        forecast_ets_requested_seasonality(requested, explicit, &values, "FORECAST.ETS.CONFINT")?;
    let _point = forecast_ets_point(
        target,
        &values,
        &timeline,
        seasonality,
        "FORECAST.ETS.CONFINT",
    )?;
    let residuals = forecast_ets_metric_residuals(&values, seasonality);
    let rmse =
        (residuals.iter().map(|error| error * error).sum::<f64>() / residuals.len() as f64).sqrt();
    let z = norm_ppf(0.5 + confidence / 2.0);
    Ok(Variant::Float(
        z * rmse * (1.0 + 1.0 / values.len() as f64).sqrt(),
    ))
}

fn func_forecast_ets_stat(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(3..=6).contains(&args.len()) {
        return Err("FORECAST.ETS.STAT requires 3 to 6 arguments".into());
    }
    let stat_type = to_float(&evaluate(&args[2], cells)?)?;
    if !stat_type.is_finite() || stat_type.fract() != 0.0 || !(1.0..=8.0).contains(&stat_type) {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let (values, timeline) = forecast_ets_series(&args[0], &args[1], cells, "FORECAST.ETS.STAT")?;
    let (requested, explicit) =
        forecast_ets_optional_at(args, cells, 3, 4, 5, "FORECAST.ETS.STAT")?;
    let seasonality =
        forecast_ets_requested_seasonality(requested, explicit, &values, "FORECAST.ETS.STAT")?;
    let residuals = forecast_ets_metric_residuals(&values, seasonality);
    let mae = residuals.iter().map(|error| error.abs()).sum::<f64>() / residuals.len() as f64;
    let rmse =
        (residuals.iter().map(|error| error * error).sum::<f64>() / residuals.len() as f64).sqrt();
    let result = match stat_type as u8 {
        1 => 1.0,     // bounded model: level smoothing coefficient
        2 | 3 => 0.0, // no trend/seasonal smoothing in this implementation
        4 => {
            let scale = values
                .windows(2)
                .map(|pair| (pair[1] - pair[0]).abs())
                .sum::<f64>()
                / (values.len() - 1) as f64;
            if scale == 0.0 {
                return Ok(Variant::Error(ExcelError::DivZero));
            }
            mae / scale
        }
        5 => {
            residuals
                .iter()
                .enumerate()
                .map(|(i, error)| {
                    let actual = values[i + if seasonality == 0 { 0 } else { seasonality }];
                    let predicted = actual - error;
                    if actual.abs() + predicted.abs() == 0.0 {
                        0.0
                    } else {
                        2.0 * error.abs() / (actual.abs() + predicted.abs())
                    }
                })
                .sum::<f64>()
                / residuals.len() as f64
        }
        6 => mae,
        7 => rmse,
        8 => timeline[1] - timeline[0],
        _ => unreachable!(),
    };
    Ok(Variant::Float(result))
}

struct RegressionInputs {
    known_y: Vec<f64>,
    known_x: Vec<f64>,
    known_x_rows: usize,
    known_x_cols: usize,
    new_x: Vec<f64>,
    new_x_rows: usize,
    new_x_cols: usize,
    constant: bool,
}

fn regression_inputs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<RegressionInputs, String> {
    if !(2..=4).contains(&args.len()) {
        return Err(format!("{name} requires 2 to 4 arguments"));
    }
    let known_y = collect_values(&args[0], cells)?
        .into_iter()
        .filter_map(|v| as_f64(&v))
        .collect::<Vec<_>>();
    let known_x = collect_values(&args[1], cells)?
        .into_iter()
        .filter_map(|v| as_f64(&v))
        .collect::<Vec<_>>();
    if known_y.is_empty() {
        return Err(format!("{name}: known_y must have non-zero length"));
    }
    let (mut known_x_rows, mut known_x_cols) = array_shape_for_expr(&args[1], cells, known_x.len());
    if known_x_rows == 1 && known_x_cols == known_y.len() && known_y.len() > 1 {
        known_x_rows = known_y.len();
        known_x_cols = 1;
    }
    if known_x_rows != known_y.len()
        || known_x_cols == 0
        || known_x_rows.saturating_mul(known_x_cols) != known_x.len()
    {
        return Err(format!(
            "{name}: known arrays must have compatible dimensions"
        ));
    }
    let new_x = if args.len() >= 3 {
        collect_values(&args[2], cells)?
            .into_iter()
            .filter_map(|v| as_f64(&v))
            .collect::<Vec<_>>()
    } else {
        known_x.clone()
    };
    if new_x.is_empty() {
        return Err(format!("{name}: new_x must not be empty"));
    }
    let (mut new_x_rows, new_x_cols) = if args.len() >= 3 {
        array_shape_for_expr(&args[2], cells, new_x.len())
    } else {
        (known_x_rows, known_x_cols)
    };
    if new_x_rows == 1 && new_x_cols == known_x_cols && known_x_cols > 1 {
        new_x_rows = 1;
    }
    if new_x_cols == 0
        || new_x_cols != known_x_cols
        || new_x_rows.saturating_mul(new_x_cols) != new_x.len()
    {
        return Err(format!(
            "{name}: new_x must have the same column count as known_x"
        ));
    }
    let constant = if args.len() == 4 {
        is_truthy(&evaluate(&args[3], cells)?)
    } else {
        true
    };
    Ok(RegressionInputs {
        known_y,
        known_x,
        known_x_rows,
        known_x_cols,
        new_x,
        new_x_rows,
        new_x_cols,
        constant,
    })
}

fn func_trend(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let inputs = regression_inputs(args, cells, "TREND")?;
    if inputs.known_x_cols > 1 {
        let coefficients = regression_coefficients(
            &inputs.known_y,
            &inputs.known_x,
            inputs.known_x_rows,
            inputs.known_x_cols,
            inputs.constant,
            false,
            "TREND",
        )?;
        let intercept = if inputs.constant {
            coefficients[inputs.known_x_cols]
        } else {
            0.0
        };
        return Ok(Variant::Array(
            (0..inputs.new_x_rows)
                .map(|row| {
                    let value = (0..inputs.known_x_cols)
                        .map(|column| {
                            inputs.new_x[row * inputs.new_x_cols + column] * coefficients[column]
                        })
                        .sum::<f64>()
                        + intercept;
                    as_integer_if_whole(value)
                })
                .collect(),
        ));
    }
    let slope = regression_slope(&inputs.known_y, &inputs.known_x)?;
    let intercept = if inputs.constant {
        let mean_y = inputs.known_y.iter().sum::<f64>() / inputs.known_y.len() as f64;
        let mean_x = inputs.known_x.iter().sum::<f64>() / inputs.known_x.len() as f64;
        mean_y - slope * mean_x
    } else {
        0.0
    };
    Ok(Variant::Array(
        inputs
            .new_x
            .into_iter()
            .map(|x| as_integer_if_whole(intercept + slope * x))
            .collect(),
    ))
}

fn func_growth(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let inputs = regression_inputs(args, cells, "GROWTH")?;
    if inputs.known_y.iter().any(|value| *value <= 0.0)
        || inputs.known_x.iter().any(|value| *value <= 0.0)
        || inputs.new_x.iter().any(|value| *value <= 0.0)
    {
        return Err("GROWTH: values must be positive".into());
    }
    if inputs.known_x_cols > 1 {
        let coefficients = regression_coefficients(
            &inputs.known_y,
            &inputs.known_x,
            inputs.known_x_rows,
            inputs.known_x_cols,
            inputs.constant,
            true,
            "GROWTH",
        )?;
        let intercept = if inputs.constant {
            coefficients[inputs.known_x_cols]
        } else {
            0.0
        };
        return Ok(Variant::Array(
            (0..inputs.new_x_rows)
                .map(|row| {
                    let log_value = (0..inputs.known_x_cols)
                        .map(|column| {
                            inputs.new_x[row * inputs.new_x_cols + column] * coefficients[column]
                        })
                        .sum::<f64>()
                        + intercept;
                    as_integer_if_whole(log_value.exp())
                })
                .collect(),
        ));
    }
    let log_y = inputs
        .known_y
        .iter()
        .map(|value| value.ln())
        .collect::<Vec<_>>();
    let slope = regression_slope(&log_y, &inputs.known_x)?;
    let intercept = if inputs.constant {
        let mean_y = log_y.iter().sum::<f64>() / log_y.len() as f64;
        let mean_x = inputs.known_x.iter().sum::<f64>() / inputs.known_x.len() as f64;
        mean_y - slope * mean_x
    } else {
        0.0
    };
    Ok(Variant::Array(
        inputs
            .new_x
            .into_iter()
            .map(|x| as_integer_if_whole((intercept + slope * x).exp()))
            .collect(),
    ))
}

fn solve_regression_system(
    mut matrix: Vec<f64>,
    mut rhs: Vec<f64>,
    dimension: usize,
    name: &str,
) -> Result<Vec<f64>, String> {
    for pivot in 0..dimension {
        let Some(row) = (pivot..dimension).max_by(|left, right| {
            matrix[*left * dimension + pivot]
                .abs()
                .total_cmp(&matrix[*right * dimension + pivot].abs())
        }) else {
            return Err(format!("{name}: singular design matrix"));
        };
        let pivot_value = matrix[row * dimension + pivot];
        if !pivot_value.is_finite() || pivot_value.abs() < 1e-12 {
            return Err(format!("{name}: singular design matrix"));
        }
        if row != pivot {
            for column in 0..dimension {
                matrix.swap(pivot * dimension + column, row * dimension + column);
            }
            rhs.swap(pivot, row);
        }
        let diagonal = matrix[pivot * dimension + pivot];
        for column in pivot..dimension {
            matrix[pivot * dimension + column] /= diagonal;
        }
        rhs[pivot] /= diagonal;
        for row in 0..dimension {
            if row == pivot {
                continue;
            }
            let factor = matrix[row * dimension + pivot];
            if factor == 0.0 {
                continue;
            }
            for column in pivot..dimension {
                matrix[row * dimension + column] -= factor * matrix[pivot * dimension + column];
            }
            rhs[row] -= factor * rhs[pivot];
        }
    }
    if rhs.iter().all(|value| value.is_finite()) {
        Ok(rhs)
    } else {
        Err(format!("{name}: non-finite regression result"))
    }
}

#[allow(clippy::too_many_arguments)]
fn regression_coefficients(
    known_y: &[f64],
    known_x: &[f64],
    x_rows: usize,
    x_cols: usize,
    constant: bool,
    log_mode: bool,
    name: &str,
) -> Result<Vec<f64>, String> {
    if x_rows != known_y.len() || known_x.len() != x_rows.saturating_mul(x_cols) {
        return Err(format!(
            "{name}: known arrays must have compatible dimensions"
        ));
    }
    let transformed_y = if log_mode {
        if known_y
            .iter()
            .any(|value| *value <= 0.0 || !value.is_finite())
        {
            return Err(format!(
                "{name}: known_y must contain positive finite values"
            ));
        }
        known_y.iter().map(|value| value.ln()).collect::<Vec<_>>()
    } else {
        known_y.to_vec()
    };
    let dimension = x_cols + usize::from(constant);
    let mut normal = vec![0.0; dimension * dimension];
    let mut rhs = vec![0.0; dimension];
    for row in 0..x_rows {
        let mut design = Vec::with_capacity(dimension);
        design.extend_from_slice(&known_x[row * x_cols..(row + 1) * x_cols]);
        if constant {
            design.push(1.0);
        }
        for left in 0..dimension {
            rhs[left] += design[left] * transformed_y[row];
            for right in 0..dimension {
                normal[left * dimension + right] += design[left] * design[right];
            }
        }
    }
    solve_regression_system(normal, rhs, dimension, name)
}

#[allow(clippy::too_many_arguments)]
fn multivariate_regression(
    known_y: &[f64],
    known_x: &[f64],
    x_rows: usize,
    x_cols: usize,
    constant: bool,
    stats: bool,
    log_mode: bool,
    name: &str,
) -> Result<Variant, String> {
    if x_rows != known_y.len() || known_x.len() != x_rows.saturating_mul(x_cols) {
        return Err(format!(
            "{name}: known arrays must have compatible dimensions"
        ));
    }
    let transformed_y = if log_mode {
        if known_y
            .iter()
            .any(|value| *value <= 0.0 || !value.is_finite())
        {
            return Err(format!(
                "{name}: known_y must contain positive finite values"
            ));
        }
        known_y.iter().map(|value| value.ln()).collect::<Vec<_>>()
    } else {
        known_y.to_vec()
    };
    let dimension = x_cols + usize::from(constant);
    if transformed_y.len() < dimension {
        return Err(format!("{name}: insufficient observations"));
    }
    let mut normal = vec![0.0; dimension * dimension];
    let mut rhs = vec![0.0; dimension];
    for row in 0..x_rows {
        let mut design = Vec::with_capacity(dimension);
        design.extend_from_slice(&known_x[row * x_cols..(row + 1) * x_cols]);
        if constant {
            design.push(1.0);
        }
        for left in 0..dimension {
            rhs[left] += design[left] * transformed_y[row];
            for right in 0..dimension {
                normal[left * dimension + right] += design[left] * design[right];
            }
        }
    }
    let coefficients = solve_regression_system(normal.clone(), rhs, dimension, name)?;
    let predictions = (0..x_rows)
        .map(|row| {
            let mut prediction = (0..x_cols)
                .map(|column| known_x[row * x_cols + column] * coefficients[column])
                .sum::<f64>();
            if constant {
                prediction += coefficients[x_cols];
            }
            prediction
        })
        .collect::<Vec<_>>();
    let residuals = transformed_y
        .iter()
        .zip(&predictions)
        .map(|(value, prediction)| value - prediction)
        .collect::<Vec<_>>();
    let ss_resid = residuals.iter().map(|value| value * value).sum::<f64>();
    let mean_y = transformed_y.iter().sum::<f64>() / transformed_y.len() as f64;
    let ss_total = if constant {
        transformed_y
            .iter()
            .map(|value| (value - mean_y).powi(2))
            .sum::<f64>()
    } else {
        transformed_y.iter().map(|value| value * value).sum::<f64>()
    };
    let ss_reg = (ss_total - ss_resid).max(0.0);
    let degrees = x_rows as i64 - dimension as i64;
    if stats && degrees <= 0 {
        return Err(format!("{name}: insufficient degrees of freedom"));
    }
    let mut output = Vec::with_capacity(if stats { dimension * 5 } else { dimension });
    for column in (0..x_cols).rev() {
        output.push(as_integer_if_whole(if log_mode {
            coefficients[column].exp()
        } else {
            coefficients[column]
        }));
    }
    if constant {
        output.push(as_integer_if_whole(if log_mode {
            coefficients[x_cols].exp()
        } else {
            coefficients[x_cols]
        }));
    }
    if !stats {
        return Ok(Variant::Array(output));
    }
    let variance = ss_resid / degrees as f64;
    let standard_error = variance.sqrt();
    let mut standard_errors = vec![0.0; dimension];
    for column in 0..dimension {
        let mut unit = vec![0.0; dimension];
        unit[column] = 1.0;
        let inverse_column = solve_regression_system(normal.clone(), unit, dimension, name)?;
        standard_errors[column] = (variance * inverse_column[column]).max(0.0).sqrt();
    }
    for column in (0..x_cols).rev() {
        output.push(as_integer_if_whole(standard_errors[column]));
    }
    if constant {
        output.push(as_integer_if_whole(standard_errors[x_cols]));
    }
    let r_squared = if ss_total == 0.0 {
        0.0
    } else {
        1.0 - ss_resid / ss_total
    };
    let f_stat = if ss_resid == 0.0 {
        f64::INFINITY
    } else {
        ss_reg / (x_cols as f64 * variance)
    };
    output.push(as_integer_if_whole(r_squared));
    output.push(as_integer_if_whole(standard_error));
    output.extend(std::iter::repeat_n(
        Variant::Error(ExcelError::NA),
        dimension - 2,
    ));
    output.push(as_integer_if_whole(f_stat));
    output.push(as_integer_if_whole(degrees as f64));
    output.extend(std::iter::repeat_n(
        Variant::Error(ExcelError::NA),
        dimension - 2,
    ));
    output.push(as_integer_if_whole(ss_reg));
    output.push(as_integer_if_whole(ss_resid));
    output.extend(std::iter::repeat_n(
        Variant::Error(ExcelError::NA),
        dimension - 2,
    ));
    Ok(Variant::Array(output))
}

fn func_linest(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(1..=4).contains(&args.len()) {
        return Err("LINEST requires 1 to 4 arguments".into());
    }
    let known_y = collect_values(&args[0], cells)?
        .into_iter()
        .filter_map(|v| as_f64(&v))
        .collect::<Vec<_>>();
    if known_y.len() < 2 {
        return Err("LINEST: known_y must contain at least 2 numeric values".into());
    }
    let known_x = if args.len() >= 2 {
        collect_values(&args[1], cells)?
            .into_iter()
            .filter_map(|v| as_f64(&v))
            .collect::<Vec<_>>()
    } else {
        (1..=known_y.len()).map(|value| value as f64).collect()
    };
    if known_x.len() != known_y.len() {
        let compatible_multivariate = args.get(1).is_some_and(|expr| {
            let (x_rows, x_cols) = array_shape_for_expr(expr, cells, known_x.len());
            x_cols > 1 && x_rows > 1 && x_rows == known_y.len()
        });
        if !compatible_multivariate {
            return Err("LINEST: known arrays must have equal length".into());
        }
    }
    let constant = if args.len() >= 3 {
        is_truthy(&evaluate(&args[2], cells)?)
    } else {
        true
    };
    let stats = if args.len() == 4 {
        is_truthy(&evaluate(&args[3], cells)?)
    } else {
        false
    };
    if args.len() >= 2 {
        let (x_rows, x_cols) = array_shape_for_expr(&args[1], cells, known_x.len());
        if x_cols > 1 && x_rows > 1 {
            return multivariate_regression(
                &known_y, &known_x, x_rows, x_cols, constant, stats, false, "LINEST",
            );
        }
    }
    let n = known_y.len() as f64;
    let mean_y = known_y.iter().sum::<f64>() / n;
    let mean_x = known_x.iter().sum::<f64>() / n;
    let (slope, intercept, x_variation) = if constant {
        let x_variation = known_x.iter().map(|x| (x - mean_x).powi(2)).sum::<f64>();
        if x_variation == 0.0 {
            return Err("LINEST: known_x values must vary".into());
        }
        let slope = known_x
            .iter()
            .zip(&known_y)
            .map(|(x, y)| (x - mean_x) * (y - mean_y))
            .sum::<f64>()
            / x_variation;
        (slope, mean_y - slope * mean_x, x_variation)
    } else {
        let x_variation = known_x.iter().map(|x| x * x).sum::<f64>();
        if x_variation == 0.0 {
            return Err("LINEST: known_x values must not all be zero".into());
        }
        let slope = known_x
            .iter()
            .zip(&known_y)
            .map(|(x, y)| x * y)
            .sum::<f64>()
            / x_variation;
        (slope, 0.0, x_variation)
    };
    if !stats {
        return Ok(Variant::Array(vec![
            as_integer_if_whole(slope),
            as_integer_if_whole(intercept),
        ]));
    }
    let residuals = known_x
        .iter()
        .zip(&known_y)
        .map(|(x, y)| y - (slope * x + intercept))
        .collect::<Vec<_>>();
    let ss_resid = residuals.iter().map(|value| value * value).sum::<f64>();
    let degrees = known_y.len() as i64 - if constant { 2 } else { 1 };
    if degrees <= 0 {
        return Err("LINEST: insufficient degrees of freedom".into());
    }
    let standard_error = (ss_resid / degrees as f64).sqrt();
    let slope_se = standard_error / x_variation.sqrt();
    let intercept_se = if constant {
        standard_error * (1.0 / n + mean_x.powi(2) / x_variation).sqrt()
    } else {
        0.0
    };
    let ss_total = if constant {
        known_y.iter().map(|y| (y - mean_y).powi(2)).sum::<f64>()
    } else {
        known_y.iter().map(|y| y * y).sum::<f64>()
    };
    let ss_reg = (ss_total - ss_resid).max(0.0);
    let r_squared = if ss_total == 0.0 {
        0.0
    } else {
        1.0 - ss_resid / ss_total
    };
    let f_stat = if ss_resid == 0.0 {
        f64::INFINITY
    } else {
        ss_reg / (ss_resid / degrees as f64)
    };
    Ok(Variant::Array(
        [
            slope,
            intercept,
            slope_se,
            intercept_se,
            r_squared,
            standard_error,
            f_stat,
            degrees as f64,
            ss_reg,
            ss_resid,
        ]
        .into_iter()
        .map(as_integer_if_whole)
        .collect(),
    ))
}

fn func_logest(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(1..=4).contains(&args.len()) {
        return Err("LOGEST requires 1 to 4 arguments".into());
    }
    let known_y = collect_values(&args[0], cells)?
        .into_iter()
        .filter_map(|v| as_f64(&v))
        .collect::<Vec<_>>();
    if known_y.len() < 2 || known_y.iter().any(|value| *value <= 0.0) {
        return Err("LOGEST: known_y must contain at least 2 positive values".into());
    }
    let known_x = if args.len() >= 2 {
        collect_values(&args[1], cells)?
            .into_iter()
            .filter_map(|v| as_f64(&v))
            .collect::<Vec<_>>()
    } else {
        (1..=known_y.len()).map(|value| value as f64).collect()
    };
    let compatible_multivariate = args.get(1).is_some_and(|expr| {
        let (x_rows, x_cols) = array_shape_for_expr(expr, cells, known_x.len());
        x_cols > 1 && x_rows > 1 && x_rows == known_y.len()
    });
    if (known_x.len() != known_y.len() && !compatible_multivariate)
        || known_x.iter().any(|value| !value.is_finite())
    {
        return Err("LOGEST: known arrays must have equal length".into());
    }
    let constant = if args.len() >= 3 {
        is_truthy(&evaluate(&args[2], cells)?)
    } else {
        true
    };
    let stats = if args.len() == 4 {
        is_truthy(&evaluate(&args[3], cells)?)
    } else {
        false
    };
    if args.len() >= 2 {
        let (x_rows, x_cols) = array_shape_for_expr(&args[1], cells, known_x.len());
        if x_cols > 1 && x_rows > 1 {
            return multivariate_regression(
                &known_y, &known_x, x_rows, x_cols, constant, stats, true, "LOGEST",
            );
        }
    }
    let log_y = known_y.iter().map(|value| value.ln()).collect::<Vec<_>>();
    let n = log_y.len() as f64;
    let mean_y = log_y.iter().sum::<f64>() / n;
    let mean_x = known_x.iter().sum::<f64>() / n;
    let (slope, intercept, x_variation) = if constant {
        let variation = known_x.iter().map(|x| (x - mean_x).powi(2)).sum::<f64>();
        if variation == 0.0 {
            return Err("LOGEST: known_x values must vary".into());
        }
        let slope = known_x
            .iter()
            .zip(&log_y)
            .map(|(x, y)| (x - mean_x) * (y - mean_y))
            .sum::<f64>()
            / variation;
        (slope, mean_y - slope * mean_x, variation)
    } else {
        let variation = known_x.iter().map(|x| x * x).sum::<f64>();
        if variation == 0.0 {
            return Err("LOGEST: known_x values must not all be zero".into());
        }
        let slope = known_x.iter().zip(&log_y).map(|(x, y)| x * y).sum::<f64>() / variation;
        (slope, 0.0, variation)
    };
    let mut result = vec![Variant::Float(slope.exp()), Variant::Float(intercept.exp())];
    if !stats {
        return Ok(Variant::Array(result));
    }
    let residuals = known_x
        .iter()
        .zip(&log_y)
        .map(|(x, y)| y - (slope * x + intercept))
        .collect::<Vec<_>>();
    let ss_resid = residuals.iter().map(|value| value * value).sum::<f64>();
    let degrees = log_y.len() as i64 - if constant { 2 } else { 1 };
    if degrees <= 0 {
        return Err("LOGEST: insufficient degrees of freedom".into());
    }
    let standard_error = (ss_resid / degrees as f64).sqrt();
    let slope_se = standard_error / x_variation.sqrt();
    let intercept_se = if constant {
        standard_error * (1.0 / n + mean_x.powi(2) / x_variation).sqrt()
    } else {
        0.0
    };
    let ss_total = if constant {
        log_y.iter().map(|y| (y - mean_y).powi(2)).sum::<f64>()
    } else {
        log_y.iter().map(|y| y * y).sum::<f64>()
    };
    let ss_reg = (ss_total - ss_resid).max(0.0);
    let r_squared = if ss_total == 0.0 {
        0.0
    } else {
        1.0 - ss_resid / ss_total
    };
    let f_stat = if ss_resid == 0.0 {
        f64::INFINITY
    } else {
        ss_reg / (ss_resid / degrees as f64)
    };
    result.extend(
        [
            slope_se,
            intercept_se,
            r_squared,
            standard_error,
            f_stat,
            degrees as f64,
            ss_reg,
            ss_resid,
        ]
        .into_iter()
        .map(as_integer_if_whole),
    );
    Ok(Variant::Array(result))
}

fn func_steyx(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (known_y, known_x) = collect_paired(args, cells, "STEYX")?;
    if known_y.len() < 3 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let slope = regression_slope(&known_y, &known_x)?;
    let my = known_y.iter().sum::<f64>() / known_y.len() as f64;
    let mx = known_x.iter().sum::<f64>() / known_x.len() as f64;
    let intercept = my - slope * mx;
    let residuals = known_y
        .iter()
        .zip(&known_x)
        .map(|(y, x)| (y - (slope * x + intercept)).powi(2))
        .sum::<f64>();
    Ok(Variant::Float(
        (residuals / (known_y.len() - 2) as f64).sqrt(),
    ))
}

fn func_fisher(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("FISHER requires 1 argument".into());
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    if !value.is_finite() || value <= -1.0 || value >= 1.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(0.5 * ((1.0 + value) / (1.0 - value)).ln()))
}

fn func_fisherinv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("FISHERINV requires 1 argument".into());
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    if !value.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(value.tanh()))
}

fn func_standardize(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("STANDARDIZE requires 3 arguments".into());
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    let mean = to_float(&evaluate(&args[1], cells)?)?;
    let standard_dev = to_float(&evaluate(&args[2], cells)?)?;
    if !value.is_finite() || !mean.is_finite() || !standard_dev.is_finite() || standard_dev <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float((value - mean) / standard_dev))
}

fn func_z_test(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("Z.TEST requires 2 or 3 arguments".into());
    }
    let values = collect_nums(&args[..1], cells)?;
    if values.is_empty() {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let target = to_float(&evaluate(&args[1], cells)?)?;
    let mean = values.iter().sum::<f64>() / values.len() as f64;
    let standard_dev = if args.len() == 3 {
        to_float(&evaluate(&args[2], cells)?)?
    } else {
        if values.len() < 2 {
            return Ok(Variant::Error(ExcelError::DivZero));
        }
        let variance = values
            .iter()
            .map(|value| (value - mean).powi(2))
            .sum::<f64>()
            / (values.len() - 1) as f64;
        variance.sqrt()
    };
    if !target.is_finite() || !standard_dev.is_finite() || standard_dev <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let z = (mean - target) / (standard_dev / (values.len() as f64).sqrt());
    Ok(Variant::Float(1.0 - norm_cdf(z.abs(), 0.0, 1.0)))
}

fn func_confidence_norm(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("CONFIDENCE.NORM requires 3 arguments".into());
    }
    let alpha = to_float(&evaluate(&args[0], cells)?)?;
    let standard_dev = to_float(&evaluate(&args[1], cells)?)?;
    let size = to_float(&evaluate(&args[2], cells)?)?;
    if !alpha.is_finite()
        || !standard_dev.is_finite()
        || !size.is_finite()
        || !(0.0..1.0).contains(&alpha)
        || standard_dev <= 0.0
        || size < 1.0
        || size.fract() != 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let critical = norm_ppf(1.0 - alpha / 2.0);
    Ok(Variant::Float(critical * standard_dev / size.sqrt()))
}

fn func_confidence_t(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("CONFIDENCE.T requires 3 arguments".into());
    }
    let alpha = to_float(&evaluate(&args[0], cells)?)?;
    let standard_dev = to_float(&evaluate(&args[1], cells)?)?;
    let size = to_float(&evaluate(&args[2], cells)?)?;
    if !alpha.is_finite()
        || !standard_dev.is_finite()
        || !size.is_finite()
        || !(0.0..1.0).contains(&alpha)
        || standard_dev <= 0.0
        || size < 2.0
        || size.fract() != 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let critical = invert_t_probability(alpha, size - 1.0, true).abs();
    Ok(Variant::Float(critical * standard_dev / size.sqrt()))
}

fn func_covariance_s(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (a, b) = collect_paired(args, cells, "COVARIANCE.S")?;
    if a.len() < 2 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let n = a.len() as f64;
    let ma = a.iter().sum::<f64>() / n;
    let mb = b.iter().sum::<f64>() / n;
    let cov: f64 = a
        .iter()
        .zip(b.iter())
        .map(|(x, y)| (x - ma) * (y - mb))
        .sum::<f64>()
        / (n - 1.0);
    Ok(Variant::Float(cov))
}

fn func_covariance_p(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (a, b) = collect_paired(args, cells, "COVARIANCE.P")?;
    let n = a.len() as f64;
    let ma = a.iter().sum::<f64>() / n;
    let mb = b.iter().sum::<f64>() / n;
    let cov: f64 = a
        .iter()
        .zip(b.iter())
        .map(|(x, y)| (x - ma) * (y - mb))
        .sum::<f64>()
        / n;
    Ok(Variant::Float(cov))
}

/// Error function (Horner's method, Abramowitz & Stegun 7.1.26, max error 1.5e-7)
fn stat_erf(x: f64) -> f64 {
    if x == 0.0 {
        return 0.0;
    }
    let t = 1.0 / (1.0 + 0.3275911 * x.abs());
    let p = t
        * (0.254829592
            + t * (-0.284496736 + t * (1.421413741 + t * (-1.453152027 + t * 1.061405429))));
    let r = 1.0 - p * (-x * x).exp();
    if x >= 0.0 { r } else { -r }
}

/// Normal CDF: Φ(x) = 0.5*(1 + erf(x/√2))
fn norm_cdf(x: f64, mean: f64, std: f64) -> f64 {
    0.5 * (1.0 + stat_erf((x - mean) / (std * std::f64::consts::SQRT_2)))
}

/// Normal PDF
fn norm_pdf(x: f64, mean: f64, std: f64) -> f64 {
    let z = (x - mean) / std;
    (-0.5 * z * z).exp() / (std * (2.0 * std::f64::consts::PI).sqrt())
}

/// Rational approximation for inverse normal CDF (Abramowitz & Stegun 26.2.23)
/// then refined with one Newton step for accuracy ~1e-9
fn norm_ppf(p: f64) -> f64 {
    if p <= 0.0 || p >= 1.0 {
        return f64::NAN;
    }
    let (sign, q) = if p < 0.5 { (-1.0, p) } else { (1.0, 1.0 - p) };
    let t = (-2.0 * q.ln()).sqrt();
    let c = [2.515517_f64, 0.802853, 0.010328];
    let d = [1.432788_f64, 0.189269, 0.001308];
    let mut z =
        sign * (t - (c[0] + t * (c[1] + t * c[2])) / (1.0 + t * (d[0] + t * (d[1] + t * d[2]))));
    // Newton-Raphson refinement
    for _ in 0..3 {
        let err = norm_cdf(z, 0.0, 1.0) - p;
        z -= err / norm_pdf(z, 0.0, 1.0);
    }
    z
}

fn func_norm_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("NORM.DIST requires 4 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let mean = to_float(&evaluate(&args[1], cells)?)?;
    let std = to_float(&evaluate(&args[2], cells)?)?;
    let cum = is_truthy(&evaluate(&args[3], cells)?);
    if std <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(if cum {
        norm_cdf(x, mean, std)
    } else {
        norm_pdf(x, mean, std)
    }))
}

fn func_norm_inv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("NORM.INV requires 3 arguments".into());
    }
    let p = to_float(&evaluate(&args[0], cells)?)?;
    let mean = to_float(&evaluate(&args[1], cells)?)?;
    let std = to_float(&evaluate(&args[2], cells)?)?;
    if std <= 0.0 || p <= 0.0 || p >= 1.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(mean + std * norm_ppf(p)))
}

fn func_norm_s_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("NORM.S.DIST requires 2 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[1], cells)?);
    Ok(Variant::Float(if cumulative {
        norm_cdf(x, 0.0, 1.0)
    } else {
        norm_pdf(x, 0.0, 1.0)
    }))
}

fn func_phi(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("PHI requires 1 argument".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    Ok(Variant::Float(norm_pdf(x, 0.0, 1.0)))
}

fn func_gauss(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("GAUSS requires 1 argument".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    Ok(Variant::Float(norm_cdf(x, 0.0, 1.0) - 0.5))
}

fn func_norm_s_inv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("NORM.S.INV requires 1 argument".into());
    }
    let p = to_float(&evaluate(&args[0], cells)?)?;
    if !(0.0..1.0).contains(&p) {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(norm_ppf(p)))
}

fn binom_coeff(n: i64, k: i64) -> f64 {
    if k < 0 || k > n {
        return 0.0;
    }
    let k = k.min(n - k);
    (1..=k).fold(1.0, |value, i| value * (n - k + i) as f64 / i as f64)
}

fn func_binom_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("BINOM.DIST requires 4 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let n = to_float(&evaluate(&args[1], cells)?)?;
    let p = to_float(&evaluate(&args[2], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[3], cells)?);
    if !x.is_finite()
        || !n.is_finite()
        || x.fract() != 0.0
        || n.fract() != 0.0
        || x < 0.0
        || n < 0.0
        || x > n
        || !(0.0..=1.0).contains(&p)
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let x = x as i64;
    let n = n as i64;
    let probability = |successes: i64| {
        binom_coeff(n, successes)
            * p.powi(successes as i32)
            * (1.0 - p).powi((n - successes) as i32)
    };
    let result = if cumulative {
        (0..=x).map(probability).sum()
    } else {
        probability(x)
    };
    Ok(Variant::Float(result))
}

fn func_binom_dist_range(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 4 {
        return Err("BINOM.DIST.RANGE requires 3 or 4 arguments".into());
    }
    let trials = to_float(&evaluate(&args[0], cells)?)?;
    let probability = to_float(&evaluate(&args[1], cells)?)?;
    let lower = to_float(&evaluate(&args[2], cells)?)?;
    let upper = if args.len() == 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        lower
    };
    if !trials.is_finite()
        || !probability.is_finite()
        || !lower.is_finite()
        || !upper.is_finite()
        || trials.fract() != 0.0
        || lower.fract() != 0.0
        || upper.fract() != 0.0
        || trials < 0.0
        || lower < 0.0
        || upper < lower
        || upper > trials
        || !(0.0..=1.0).contains(&probability)
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let trials = trials as i64;
    let lower = lower as i64;
    let upper = upper as i64;
    let result = (lower..=upper)
        .map(|successes| {
            binom_coeff(trials, successes)
                * probability.powi(successes as i32)
                * (1.0 - probability).powi((trials - successes) as i32)
        })
        .sum::<f64>();
    Ok(Variant::Float(result))
}

fn func_binom_inv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("BINOM.INV requires 3 arguments".into());
    }
    let trials = to_float(&evaluate(&args[0], cells)?)?;
    let probability = to_float(&evaluate(&args[1], cells)?)?;
    let alpha = to_float(&evaluate(&args[2], cells)?)?;
    if !trials.is_finite()
        || trials.fract() != 0.0
        || trials < 0.0
        || !(0.0..=1.0).contains(&probability)
        || !alpha.is_finite()
        || !(0.0..=1.0).contains(&alpha)
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let trials = trials as i64;
    let mass = |successes: i64| {
        binom_coeff(trials, successes)
            * probability.powi(successes as i32)
            * (1.0 - probability).powi((trials - successes) as i32)
    };
    let mut cumulative = 0.0;
    for successes in 0..=trials {
        cumulative += mass(successes);
        if cumulative + 1e-15 >= alpha {
            return Ok(Variant::Integer(successes));
        }
    }
    Ok(Variant::Integer(trials))
}

fn func_negbinom_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("NEGBINOM.DIST requires 4 arguments".into());
    }
    let failures = to_float(&evaluate(&args[0], cells)?)?;
    let successes = to_float(&evaluate(&args[1], cells)?)?;
    let probability = to_float(&evaluate(&args[2], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[3], cells)?);
    if !failures.is_finite()
        || !successes.is_finite()
        || !probability.is_finite()
        || failures.fract() != 0.0
        || successes.fract() != 0.0
        || failures < 0.0
        || successes < 1.0
        || !(0.0..=1.0).contains(&probability)
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let failures = failures as i64;
    let successes = successes as i64;
    let pmf = |count: i64| {
        binom_coeff(count + successes - 1, count)
            * probability.powi(successes as i32)
            * (1.0 - probability).powi(count as i32)
    };
    let result = if cumulative {
        (0..=failures).map(pmf).sum()
    } else {
        pmf(failures)
    };
    Ok(Variant::Float(result))
}

fn func_hypgeom_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 5 {
        return Err("HYPGEOM.DIST requires 5 arguments".into());
    }
    let sample_success = to_float(&evaluate(&args[0], cells)?)?;
    let sample_size = to_float(&evaluate(&args[1], cells)?)?;
    let population_success = to_float(&evaluate(&args[2], cells)?)?;
    let population_size = to_float(&evaluate(&args[3], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[4], cells)?);
    let inputs = [
        sample_success,
        sample_size,
        population_success,
        population_size,
    ];
    if inputs
        .iter()
        .any(|value| !value.is_finite() || value.fract() != 0.0 || *value < 0.0)
        || sample_size > population_size
        || population_success > population_size
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let sample_success = sample_success as i64;
    let sample_size = sample_size as i64;
    let population_success = population_success as i64;
    let population_size = population_size as i64;
    let lower = 0.max(sample_size - (population_size - population_success));
    let upper = sample_size.min(population_success);
    if sample_success < lower || sample_success > upper {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let probability = |successes: i64| {
        binom_coeff(population_success, successes)
            * binom_coeff(
                population_size - population_success,
                sample_size - successes,
            )
            / binom_coeff(population_size, sample_size)
    };
    let result = if cumulative {
        (lower..=sample_success).map(probability).sum()
    } else {
        probability(sample_success)
    };
    Ok(Variant::Float(result))
}

fn func_poisson_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("POISSON.DIST requires 3 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let mean = to_float(&evaluate(&args[1], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[2], cells)?);
    if !x.is_finite() || x.fract() != 0.0 || x < 0.0 || !mean.is_finite() || mean <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let x = x as i64;
    let probability =
        |k: i64| (-mean).exp() * mean.powi(k as i32) / (1..=k).fold(1.0, |v, i| v * i as f64);
    let result = if cumulative {
        (0..=x).map(probability).sum()
    } else {
        probability(x)
    };
    Ok(Variant::Float(result))
}

fn func_gamma(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("GAMMA requires 1 argument".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let value = lgamma(x).exp();
    if !value.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(value))
}

fn func_gammaln(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("GAMMALN requires 1 argument".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    if x <= 0.0 || !x.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(lgamma(x)))
}

/// Regularized lower incomplete gamma P(a, x), using a series or continued fraction.
fn reg_inc_gamma(a: f64, x: f64) -> f64 {
    if x <= 0.0 || a <= 0.0 {
        return 0.0;
    }
    if x < a + 1.0 {
        let mut sum = 1.0 / a;
        let mut term = sum;
        for n in 1..=200 {
            term *= x / (a + n as f64);
            sum += term;
            if term.abs() < sum.abs() * 3e-14 {
                break;
            }
        }
        return sum * (-x + a * x.ln() - lgamma(a)).exp();
    }
    let mut b = x + 1.0 - a;
    let mut c = 1e300;
    let mut d = 1.0 / b;
    let mut h = d;
    for i in 1..=200 {
        let fi = i as f64;
        let an = -fi * (fi - a);
        b += 2.0;
        d = an * d + b;
        if d.abs() < 1e-300 {
            d = 1e-300;
        }
        c = b + an / c;
        if c.abs() < 1e-300 {
            c = 1e-300;
        }
        d = 1.0 / d;
        let delta = d * c;
        h *= delta;
        if (delta - 1.0).abs() < 3e-14 {
            break;
        }
    }
    1.0 - (-x + a * x.ln() - lgamma(a)).exp() * h
}

fn func_beta_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 6 {
        return Err("BETA.DIST requires 4 to 6 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let alpha = to_float(&evaluate(&args[1], cells)?)?;
    let beta = to_float(&evaluate(&args[2], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[3], cells)?);
    let lower = if args.len() >= 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let upper = if args.len() == 6 {
        to_float(&evaluate(&args[5], cells)?)?
    } else {
        1.0
    };
    if alpha <= 0.0 || beta <= 0.0 || lower >= upper || x < lower || x > upper {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let z = (x - lower) / (upper - lower);
    let result = if cumulative {
        reg_inc_beta(z, alpha, beta)
    } else {
        ((alpha - 1.0) * z.ln() + (beta - 1.0) * (1.0 - z).ln() - lgamma(alpha) - lgamma(beta)
            + lgamma(alpha + beta))
        .exp()
            / (upper - lower)
    };
    Ok(Variant::Float(result))
}

fn func_beta_inv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 {
        return Err("BETA.INV requires 3 to 5 arguments".into());
    }
    let probability = to_float(&evaluate(&args[0], cells)?)?;
    let alpha = to_float(&evaluate(&args[1], cells)?)?;
    let beta = to_float(&evaluate(&args[2], cells)?)?;
    let lower = if args.len() >= 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        0.0
    };
    let upper = if args.len() == 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        1.0
    };
    if !(0.0..=1.0).contains(&probability) || alpha <= 0.0 || beta <= 0.0 || lower >= upper {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut lo = 0.0;
    let mut hi = 1.0;
    for _ in 0..80 {
        let mid = (lo + hi) / 2.0;
        if reg_inc_beta(mid, alpha, beta) < probability {
            lo = mid;
        } else {
            hi = mid;
        }
    }
    Ok(Variant::Float(lower + (upper - lower) * (lo + hi) / 2.0))
}

fn func_chisq_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("CHISQ.DIST requires 3 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let df = to_float(&evaluate(&args[1], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[2], cells)?);
    if x < 0.0 || df <= 0.0 || !x.is_finite() || !df.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let a = df / 2.0;
    let result = if cumulative {
        reg_inc_gamma(a, x / 2.0)
    } else {
        ((a - 1.0) * x.ln() - x / 2.0 - a * 2.0_f64.ln() - lgamma(a)).exp()
    };
    Ok(Variant::Float(result))
}

fn func_f_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("F.DIST requires 4 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let df1 = to_float(&evaluate(&args[1], cells)?)?;
    let df2 = to_float(&evaluate(&args[2], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[3], cells)?);
    if x < 0.0 || df1 <= 0.0 || df2 <= 0.0 || !x.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let z = df1 * x / (df1 * x + df2);
    let result = if cumulative {
        reg_inc_beta(z, df1 / 2.0, df2 / 2.0)
    } else {
        let a = df1 / 2.0;
        let b = df2 / 2.0;
        (a * df1.ln() + b * df2.ln() + (a - 1.0) * x.ln() - (a + b) * (df2 + df1 * x).ln()
            + lgamma(a + b)
            - lgamma(a)
            - lgamma(b))
        .exp()
            * df1
    };
    Ok(Variant::Float(result))
}

fn gamma_pdf(x: f64, shape: f64, scale: f64) -> f64 {
    if x < 0.0 || shape <= 0.0 || scale <= 0.0 {
        return f64::NAN;
    }
    if x == 0.0 && shape < 1.0 {
        return f64::INFINITY;
    }
    ((shape - 1.0) * x.max(f64::MIN_POSITIVE).ln() - x / scale - lgamma(shape) - shape * scale.ln())
        .exp()
}

fn gamma_cdf(x: f64, shape: f64, scale: f64) -> f64 {
    if x < 0.0 || shape <= 0.0 || scale <= 0.0 {
        return f64::NAN;
    }
    reg_inc_gamma(shape, x / scale)
}

fn inverse_cdf<F>(probability: f64, cdf: F) -> f64
where
    F: Fn(f64) -> f64,
{
    let mut hi = 1.0;
    while cdf(hi) < probability && hi < 1.0e12 {
        hi *= 2.0;
    }
    let mut lo = 0.0;
    for _ in 0..100 {
        let mid = (lo + hi) / 2.0;
        if cdf(mid) < probability {
            lo = mid;
        } else {
            hi = mid;
        }
    }
    (lo + hi) / 2.0
}

fn func_gamma_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("GAMMA.DIST requires 4 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let alpha = to_float(&evaluate(&args[1], cells)?)?;
    let beta = to_float(&evaluate(&args[2], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[3], cells)?);
    if !x.is_finite()
        || !alpha.is_finite()
        || !beta.is_finite()
        || x < 0.0
        || alpha <= 0.0
        || beta <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let result = if cumulative {
        gamma_cdf(x, alpha, beta)
    } else {
        gamma_pdf(x, alpha, beta)
    };
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(result))
}

fn func_gamma_inv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("GAMMA.INV requires 3 arguments".into());
    }
    let probability = to_float(&evaluate(&args[0], cells)?)?;
    let alpha = to_float(&evaluate(&args[1], cells)?)?;
    let beta = to_float(&evaluate(&args[2], cells)?)?;
    if !probability.is_finite()
        || !alpha.is_finite()
        || !beta.is_finite()
        || !(0.0..1.0).contains(&probability)
        || alpha <= 0.0
        || beta <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(inverse_cdf(probability, |x| {
        gamma_cdf(x, alpha, beta)
    })))
}

fn chisq_cdf(x: f64, df: f64) -> f64 {
    reg_inc_gamma(df / 2.0, x / 2.0)
}

fn func_chisq_dist_rt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("CHISQ.DIST.RT requires 2 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let df = to_float(&evaluate(&args[1], cells)?)?;
    if !x.is_finite() || !df.is_finite() || x < 0.0 || df <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(1.0 - chisq_cdf(x, df)))
}

fn chisq_inverse(probability: f64, df: f64, right_tail: bool) -> f64 {
    let lower_probability = if right_tail {
        1.0 - probability
    } else {
        probability
    };
    inverse_cdf(lower_probability, |x| chisq_cdf(x, df))
}

fn func_chisq_inv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("CHISQ.INV requires 2 arguments".into());
    }
    let probability = to_float(&evaluate(&args[0], cells)?)?;
    let df = to_float(&evaluate(&args[1], cells)?)?;
    if !probability.is_finite()
        || !df.is_finite()
        || !(0.0..1.0).contains(&probability)
        || df <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(chisq_inverse(probability, df, false)))
}

fn func_chisq_inv_rt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("CHISQ.INV.RT requires 2 arguments".into());
    }
    let probability = to_float(&evaluate(&args[0], cells)?)?;
    let df = to_float(&evaluate(&args[1], cells)?)?;
    if !probability.is_finite()
        || !df.is_finite()
        || !(0.0..1.0).contains(&probability)
        || df <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(chisq_inverse(probability, df, true)))
}

fn f_cdf(x: f64, df1: f64, df2: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    let z = df1 * x / (df1 * x + df2);
    reg_inc_beta(z, df1 / 2.0, df2 / 2.0)
}

fn func_f_dist_rt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("F.DIST.RT requires 3 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let df1 = to_float(&evaluate(&args[1], cells)?)?;
    let df2 = to_float(&evaluate(&args[2], cells)?)?;
    if !x.is_finite() || !df1.is_finite() || !df2.is_finite() || x < 0.0 || df1 <= 0.0 || df2 <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(1.0 - f_cdf(x, df1, df2)))
}

fn func_f_dist_2t(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("F.DIST.2T requires 3 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let df1 = to_float(&evaluate(&args[1], cells)?)?;
    let df2 = to_float(&evaluate(&args[2], cells)?)?;
    if !x.is_finite() || !df1.is_finite() || !df2.is_finite() || x < 0.0 || df1 <= 0.0 || df2 <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float((2.0 * (1.0 - f_cdf(x, df1, df2))).min(1.0)))
}

fn f_inverse(probability: f64, df1: f64, df2: f64, right_tail: bool) -> f64 {
    let lower_probability = if right_tail {
        1.0 - probability
    } else {
        probability
    };
    inverse_cdf(lower_probability, |x| f_cdf(x, df1, df2))
}

fn func_f_inv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("F.INV requires 3 arguments".into());
    }
    let probability = to_float(&evaluate(&args[0], cells)?)?;
    let df1 = to_float(&evaluate(&args[1], cells)?)?;
    let df2 = to_float(&evaluate(&args[2], cells)?)?;
    if !probability.is_finite()
        || !df1.is_finite()
        || !df2.is_finite()
        || !(0.0..1.0).contains(&probability)
        || df1 <= 0.0
        || df2 <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(f_inverse(probability, df1, df2, false)))
}

fn func_f_inv_rt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("F.INV.RT requires 3 arguments".into());
    }
    let probability = to_float(&evaluate(&args[0], cells)?)?;
    let df1 = to_float(&evaluate(&args[1], cells)?)?;
    let df2 = to_float(&evaluate(&args[2], cells)?)?;
    if !probability.is_finite()
        || !df1.is_finite()
        || !df2.is_finite()
        || !(0.0..1.0).contains(&probability)
        || df1 <= 0.0
        || df2 <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(f_inverse(probability, df1, df2, true)))
}

fn func_weibull_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("WEIBULL.DIST requires 4 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let alpha = to_float(&evaluate(&args[1], cells)?)?;
    let beta = to_float(&evaluate(&args[2], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[3], cells)?);
    if !x.is_finite()
        || !alpha.is_finite()
        || !beta.is_finite()
        || x < 0.0
        || alpha <= 0.0
        || beta <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let ratio = x / beta;
    let result = if cumulative {
        1.0 - (-ratio.powf(alpha)).exp()
    } else if x == 0.0 && alpha < 1.0 {
        f64::INFINITY
    } else {
        alpha / beta * ratio.powf(alpha - 1.0) * (-ratio.powf(alpha)).exp()
    };
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(result))
}

fn func_expon_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("EXPON.DIST requires 3 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let lambda = to_float(&evaluate(&args[1], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[2], cells)?);
    if !x.is_finite() || !lambda.is_finite() || x < 0.0 || lambda <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let result = if cumulative {
        1.0 - (-lambda * x).exp()
    } else {
        lambda * (-lambda * x).exp()
    };
    Ok(Variant::Float(result))
}

fn func_lognorm_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("LOGNORM.DIST requires 4 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let mean = to_float(&evaluate(&args[1], cells)?)?;
    let std = to_float(&evaluate(&args[2], cells)?)?;
    let cumulative = is_truthy(&evaluate(&args[3], cells)?);
    if !x.is_finite() || !mean.is_finite() || !std.is_finite() || x <= 0.0 || std <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let log_x = x.ln();
    let result = if cumulative {
        norm_cdf(log_x, mean, std)
    } else {
        norm_pdf(log_x, mean, std) / x
    };
    Ok(Variant::Float(result))
}

fn func_lognorm_inv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("LOGNORM.INV requires 3 arguments".into());
    }
    let probability = to_float(&evaluate(&args[0], cells)?)?;
    let mean = to_float(&evaluate(&args[1], cells)?)?;
    let std = to_float(&evaluate(&args[2], cells)?)?;
    if !probability.is_finite()
        || !mean.is_finite()
        || !std.is_finite()
        || !(0.0..1.0).contains(&probability)
        || std <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float((mean + std * norm_ppf(probability)).exp()))
}

/// Natural log of gamma function.
/// Uses the reflection formula for x < 0.5, recurrence to shift x ≥ 7,
/// then the Stirling series for large x.
fn lgamma(x: f64) -> f64 {
    use std::f64::consts::PI;
    if x < 0.5 {
        PI.ln() - (PI * x).sin().ln() - lgamma(1.0 - x)
    } else {
        let mut y = x;
        let mut adj = 0.0;
        while y < 7.0 {
            adj -= y.ln();
            y += 1.0;
        }
        adj + (y - 0.5) * y.ln() - y + 0.5 * (2.0 * PI).ln() + 1.0 / (12.0 * y)
            - 1.0 / (360.0 * y.powi(3))
            + 1.0 / (1260.0 * y.powi(5))
    }
}

/// Regularized incomplete beta function I_x(a, b) via Lentz continued fraction
fn reg_inc_beta(x: f64, a: f64, b: f64) -> f64 {
    if x <= 0.0 {
        return 0.0;
    }
    if x >= 1.0 {
        return 1.0;
    }
    if x > (a + 1.0) / (a + b + 2.0) {
        return 1.0 - reg_inc_beta(1.0 - x, b, a);
    }
    let bt = (a * x.ln() + b * (1.0 - x).ln() + lgamma(a + b) - lgamma(a) - lgamma(b)).exp();
    const EPS: f64 = 3e-10;
    const FPMIN: f64 = 1e-300;
    let (qab, qap, qam) = (a + b, a + 1.0, a - 1.0);
    let (mut c, mut d) = (1.0, 1.0 - qab * x / qap);
    if d.abs() < FPMIN {
        d = FPMIN;
    }
    d = 1.0 / d;
    let mut h = d;
    for m in 1usize..=100 {
        let mf = m as f64;
        let m2 = 2.0 * mf;
        let aa = mf * (b - mf) * x / ((qam + m2) * (a + m2));
        d = 1.0 + aa * d;
        if d.abs() < FPMIN {
            d = FPMIN;
        }
        c = 1.0 + aa / c;
        if c.abs() < FPMIN {
            c = FPMIN;
        }
        d = 1.0 / d;
        h *= d * c;
        let aa = -(a + mf) * (qab + mf) * x / ((a + m2) * (qap + m2));
        d = 1.0 + aa * d;
        if d.abs() < FPMIN {
            d = FPMIN;
        }
        c = 1.0 + aa / c;
        if c.abs() < FPMIN {
            c = FPMIN;
        }
        d = 1.0 / d;
        let delta = d * c;
        h *= delta;
        if (delta - 1.0).abs() < EPS {
            break;
        }
    }
    bt * h / a
}

fn func_t_dist(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("T.DIST requires 3 arguments".into());
    }
    let t = to_float(&evaluate(&args[0], cells)?)?;
    let v = to_float(&evaluate(&args[1], cells)?)?;
    let cum = is_truthy(&evaluate(&args[2], cells)?);
    if !t.is_finite() || !v.is_finite() || v < 1.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(t_dist_value(t, v, cum)))
}

fn func_tdist_legacy(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("TDIST requires 3 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let degrees = to_float(&evaluate(&args[1], cells)?)?;
    let tails = to_float(&evaluate(&args[2], cells)?)?;
    if !x.is_finite()
        || !degrees.is_finite()
        || !tails.is_finite()
        || x < 0.0
        || degrees < 1.0
        || tails.fract() != 0.0
        || !(tails == 1.0 || tails == 2.0)
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let right_tail = 1.0 - t_dist_value(x, degrees, true);
    Ok(Variant::Float(if tails == 1.0 {
        right_tail
    } else {
        (2.0 * right_tail).min(1.0)
    }))
}

fn t_dist_value(t: f64, v: f64, cumulative: bool) -> f64 {
    if cumulative {
        let x = v / (v + t * t);
        let p = 0.5 * reg_inc_beta(x, v / 2.0, 0.5);
        if t >= 0.0 { 1.0 - p } else { p }
    } else {
        // PDF: Γ((v+1)/2) / (√(vπ) Γ(v/2)) * (1 + t²/v)^(-(v+1)/2)
        ((lgamma((v + 1.0) / 2.0) - lgamma(v / 2.0)).exp()) / (v * std::f64::consts::PI).sqrt()
            * (1.0 + t * t / v).powf(-(v + 1.0) / 2.0)
    }
}

fn func_t_dist_2t(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("T.DIST.2T requires 2 arguments".into());
    }
    let t = to_float(&evaluate(&args[0], cells)?)?;
    let v = to_float(&evaluate(&args[1], cells)?)?;
    if !t.is_finite() || !v.is_finite() || v < 1.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(2.0 * (1.0 - t_dist_value(t.abs(), v, true))))
}

fn func_t_dist_rt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("T.DIST.RT requires 2 arguments".into());
    }
    let t = to_float(&evaluate(&args[0], cells)?)?;
    let v = to_float(&evaluate(&args[1], cells)?)?;
    if !t.is_finite() || !v.is_finite() || v < 1.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(1.0 - t_dist_value(t, v, true)))
}

fn func_ttest(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("TTEST requires 4 arguments".into());
    }
    let first = collect_nums(std::slice::from_ref(&args[0]), cells)?;
    let second = collect_nums(std::slice::from_ref(&args[1]), cells)?;
    let tails = to_float(&evaluate(&args[2], cells)?)?;
    let test_type = to_float(&evaluate(&args[3], cells)?)?;
    if !tails.is_finite()
        || !test_type.is_finite()
        || tails.fract() != 0.0
        || test_type.fract() != 0.0
        || !(tails == 1.0 || tails == 2.0)
        || !(1.0..=3.0).contains(&test_type)
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    if first.len() < 2 || second.len() < 2 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let mean = |values: &[f64]| values.iter().sum::<f64>() / values.len() as f64;
    let variance = |values: &[f64]| {
        let average = mean(values);
        values
            .iter()
            .map(|value| (value - average).powi(2))
            .sum::<f64>()
            / (values.len() - 1) as f64
    };
    let (t, df) = match test_type as u8 {
        1 => {
            if first.len() != second.len() {
                return Ok(Variant::Error(ExcelError::NA));
            }
            let differences: Vec<f64> = first
                .iter()
                .zip(second.iter())
                .map(|(left, right)| left - right)
                .collect();
            let difference_mean = mean(&differences);
            let difference_variance = variance(&differences);
            if difference_variance <= 0.0 {
                return Ok(Variant::Error(ExcelError::DivZero));
            }
            (
                difference_mean / (difference_variance / differences.len() as f64).sqrt(),
                (differences.len() - 1) as f64,
            )
        }
        2 => {
            let first_variance = variance(&first);
            let second_variance = variance(&second);
            let df = (first.len() + second.len() - 2) as f64;
            let pooled = ((first.len() - 1) as f64 * first_variance
                + (second.len() - 1) as f64 * second_variance)
                / df;
            if pooled <= 0.0 {
                return Ok(Variant::Error(ExcelError::DivZero));
            }
            (
                (mean(&first) - mean(&second))
                    / (pooled * (1.0 / first.len() as f64 + 1.0 / second.len() as f64)).sqrt(),
                df,
            )
        }
        3 => {
            let first_variance = variance(&first);
            let second_variance = variance(&second);
            let first_term = first_variance / first.len() as f64;
            let second_term = second_variance / second.len() as f64;
            let denominator = (first_term * first_term) / (first.len() - 1) as f64
                + (second_term * second_term) / (second.len() - 1) as f64;
            if first_term + second_term <= 0.0 || denominator <= 0.0 {
                return Ok(Variant::Error(ExcelError::DivZero));
            }
            (
                (mean(&first) - mean(&second)) / (first_term + second_term).sqrt(),
                (first_term + second_term).powi(2) / denominator,
            )
        }
        _ => unreachable!(),
    };
    let p = if tails == 1.0 {
        1.0 - t_dist_value(t.abs(), df, true)
    } else {
        2.0 * (1.0 - t_dist_value(t.abs(), df, true))
    };
    Ok(Variant::Float(p.min(1.0)))
}

fn invert_t_probability(probability: f64, v: f64, two_tailed: bool) -> f64 {
    let target = if two_tailed {
        1.0 - probability / 2.0
    } else {
        probability
    };
    let mut lo = -1.0e6;
    let mut hi = 1.0e6;
    for _ in 0..100 {
        let mid = (lo + hi) / 2.0;
        if t_dist_value(mid, v, true) < target {
            lo = mid;
        } else {
            hi = mid;
        }
    }
    (lo + hi) / 2.0
}

fn func_t_inv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("T.INV requires 2 arguments".into());
    }
    let p = to_float(&evaluate(&args[0], cells)?)?;
    let v = to_float(&evaluate(&args[1], cells)?)?;
    if !(0.0..1.0).contains(&p) || !v.is_finite() || v < 1.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(invert_t_probability(p, v, false)))
}

fn func_t_inv_2t(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("T.INV.2T requires 2 arguments".into());
    }
    let p = to_float(&evaluate(&args[0], cells)?)?;
    let v = to_float(&evaluate(&args[1], cells)?)?;
    if !(0.0..=1.0).contains(&p) || p == 0.0 || !v.is_finite() || v < 1.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(invert_t_probability(p, v, true).abs()))
}

// ── Rounding ──────────────────────────────────────────────────────────────────

fn func_floor(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("FLOOR requires at least 1 argument".into());
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let sig = if args.len() >= 2 {
        to_float(&evaluate(&args[1], cells)?)?
    } else {
        1.0
    };
    if sig == 0.0 {
        return Ok(Variant::Integer(0));
    }
    Ok(as_integer_if_whole((num / sig).floor() * sig))
}

fn func_ceiling(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("CEILING requires at least 1 argument".into());
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let sig = if args.len() >= 2 {
        to_float(&evaluate(&args[1], cells)?)?
    } else {
        1.0
    };
    if sig == 0.0 {
        return Ok(Variant::Integer(0));
    }
    Ok(as_integer_if_whole((num / sig).ceil() * sig))
}

fn func_mround(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("MROUND requires 2 arguments".into());
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let mult = to_float(&evaluate(&args[1], cells)?)?;
    if mult == 0.0 {
        return Ok(Variant::Integer(0));
    }
    if (num < 0.0) != (mult < 0.0) {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole((num / mult).round() * mult))
}

fn func_precise_round(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    ceiling: bool,
) -> Result<Variant, String> {
    let name = if ceiling {
        "CEILING.PRECISE"
    } else {
        "FLOOR.PRECISE"
    };
    if args.is_empty() || args.len() > 2 {
        return Err(format!("{name} requires 1 or 2 arguments"));
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let significance = if args.len() == 2 {
        to_float(&evaluate(&args[1], cells)?)?.abs()
    } else {
        1.0
    };
    if !num.is_finite() || !significance.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    if significance == 0.0 || num == 0.0 {
        return Ok(Variant::Integer(0));
    }
    let quotient = num / significance;
    let magnitude = if ceiling {
        quotient.ceil() * significance
    } else {
        quotient.floor() * significance
    };
    Ok(as_integer_if_whole(magnitude))
}

fn func_even_odd(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    even: bool,
) -> Result<Variant, String> {
    let name = if even { "EVEN" } else { "ODD" };
    if args.len() != 1 {
        return Err(format!("{name} requires 1 argument"));
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    if !num.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut magnitude = num.abs().ceil();
    if even {
        if (magnitude as i64) % 2 != 0 {
            magnitude += 1.0;
        }
    } else if (magnitude as i64) % 2 == 0 {
        magnitude += 1.0;
    }
    Ok(as_integer_if_whole(num.signum() * magnitude))
}

// ── Math ──────────────────────────────────────────────────────────────────────

fn func_abs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ABS requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    Ok(as_integer_if_whole(n.abs()))
}

fn func_sqrt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("SQRT requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    if n < 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole(n.sqrt()))
}

fn func_sqrtpi(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("SQRTPI requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    if !n.is_finite() || n < 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole((n * std::f64::consts::PI).sqrt()))
}

fn func_power(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("POWER requires 2 arguments".into());
    }
    let base = to_float(&evaluate(&args[0], cells)?)?;
    let exp = to_float(&evaluate(&args[1], cells)?)?;
    Ok(as_integer_if_whole(base.powf(exp)))
}

fn func_exp(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("EXP requires 1 argument".into());
    }
    Ok(Variant::Float(to_float(&evaluate(&args[0], cells)?)?.exp()))
}

fn func_log(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("LOG requires at least 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    if n <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let base = if args.len() >= 2 {
        to_float(&evaluate(&args[1], cells)?)?
    } else {
        10.0
    };
    if base <= 0.0 || base == 1.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(n.log(base)))
}

fn func_log10(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("LOG10 requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    if n <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(n.log10()))
}

fn func_ln(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("LN requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    if n <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(n.ln()))
}

// ── Engineering ─────────────────────────────────────────────────────────────

const MAX_BITWISE_VALUE: u64 = (1u64 << 48) - 1;

fn bitwise_integer(value: &Variant, name: &str) -> Result<u64, Variant> {
    let number = match to_float(value) {
        Ok(number) => number,
        Err(_) => return Err(Variant::Error(ExcelError::Value)),
    };
    if !number.is_finite() || number < 0.0 || number > MAX_BITWISE_VALUE as f64 {
        return Err(Variant::Error(ExcelError::Num));
    }
    let integer = number.trunc();
    if integer > i64::MAX as f64 {
        return Err(Variant::Error(ExcelError::Num));
    }
    let result = integer as u64;
    if result > MAX_BITWISE_VALUE {
        return Err(Variant::Error(ExcelError::Num));
    }
    let _ = name;
    Ok(result)
}

fn func_bitwise(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err(format!("{name} requires 2 arguments"));
    }
    let left = match bitwise_integer(&evaluate(&args[0], cells)?, name) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    let right = match bitwise_integer(&evaluate(&args[1], cells)?, name) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    let result = match name {
        "BITAND" => left & right,
        "BITOR" => left | right,
        "BITXOR" => left ^ right,
        _ => unreachable!(),
    };
    Ok(Variant::Integer(result as i64))
}

fn func_bitshift(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    left: bool,
) -> Result<Variant, String> {
    let name = if left { "BITLSHIFT" } else { "BITRSHIFT" };
    if args.len() != 2 {
        return Err(format!("{name} requires 2 arguments"));
    }
    let number = match bitwise_integer(&evaluate(&args[0], cells)?, name) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    let shift_value = match to_float(&evaluate(&args[1], cells)?) {
        Ok(value) => value,
        Err(_) => return Ok(Variant::Error(ExcelError::Value)),
    };
    if !shift_value.is_finite() || shift_value.trunc() != shift_value || shift_value.abs() > 53.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let shift = shift_value as i32;
    let effective_left = if shift < 0 { !left } else { left };
    let amount = shift.unsigned_abs();
    if effective_left && (amount >= 48 || number > (MAX_BITWISE_VALUE >> amount)) {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let result = if effective_left {
        number << amount
    } else {
        number >> amount
    };
    if result > MAX_BITWISE_VALUE {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Integer(result as i64))
}

fn func_delta(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("DELTA requires 1 or 2 arguments".into());
    }
    let left = to_float(&evaluate(&args[0], cells)?)?;
    let right = if args.len() == 2 {
        to_float(&evaluate(&args[1], cells)?)?
    } else {
        0.0
    };
    if !left.is_finite() || !right.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Integer((left == right) as i64))
}

fn func_gestep(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("GESTEP requires 1 or 2 arguments".into());
    }
    let number = to_float(&evaluate(&args[0], cells)?)?;
    let step = if args.len() == 2 {
        to_float(&evaluate(&args[1], cells)?)?
    } else {
        0.0
    };
    if !number.is_finite() || !step.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Integer((number >= step) as i64))
}

fn func_erf(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("ERF requires 1 or 2 arguments".into());
    }
    let lower = to_float(&evaluate(&args[0], cells)?)?;
    let result = if args.len() == 2 {
        let upper = to_float(&evaluate(&args[1], cells)?)?;
        stat_erf(upper) - stat_erf(lower)
    } else {
        stat_erf(lower)
    };
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(result))
}

fn func_erfc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ERFC requires 1 argument".into());
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    let result = 1.0 - stat_erf(value);
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(result))
}

fn func_bessel(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    modified: bool,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err(format!(
            "{} requires 2 arguments",
            if modified { "BESSELI" } else { "BESSELJ" }
        ));
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let order = to_float(&evaluate(&args[1], cells)?)?;
    if !x.is_finite() || !order.is_finite() || order < 0.0 || order.fract() != 0.0 || order > 100.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let order = order as usize;
    let result = bessel_first_kind(x, order, modified)?;
    Ok(as_integer_if_whole(result))
}

fn bessel_first_kind(x: f64, order: usize, modified: bool) -> Result<f64, String> {
    let mut factorial = 1.0;
    for value in 2..=order {
        factorial *= value as f64;
    }
    let mut term = (x / 2.0).powi(order as i32) / factorial;
    let x_squared_quarter = x * x / 4.0;
    let mut result = 0.0;
    for m in 0..10_000_u32 {
        let signed_term = if modified || m % 2 == 0 { term } else { -term };
        result += signed_term;
        if !result.is_finite() {
            return Err("Bessel series overflow".to_string());
        }
        let next = term * x_squared_quarter / ((m + 1) as f64 * (m as usize + order + 1) as f64);
        if next.abs() <= 1e-15 * result.abs().max(1.0) {
            break;
        }
        term = next;
    }
    Ok(result)
}

fn func_bessel_second_kind(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    modified: bool,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err(format!(
            "{} requires 2 arguments",
            if modified { "BESSELK" } else { "BESSELY" }
        ));
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let order = to_float(&evaluate(&args[1], cells)?)?;
    if !x.is_finite()
        || !order.is_finite()
        || x <= 0.0
        || order < 0.0
        || order.fract() != 0.0
        || order > 100.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let order = order as usize;
    let result = if modified {
        bessel_k(x, order)
    } else {
        bessel_y(x, order)
    }?;
    if result.is_finite() {
        Ok(as_integer_if_whole(result))
    } else {
        Ok(Variant::Error(ExcelError::Num))
    }
}

fn bessel_y(x: f64, order: usize) -> Result<f64, String> {
    const EULER_GAMMA: f64 = 0.5772156649015329;
    let j0 = bessel_first_kind(x, 0, false)?;
    let j1 = bessel_first_kind(x, 1, false)?;
    let a = (x / 2.0).ln() + EULER_GAMMA;
    let mut term = x * x / 4.0;
    let mut harmonic = 1.0;
    let mut sum = 0.0;
    let mut derivative_sum = 0.0;
    for m in 1..=10_000_u32 {
        let signed = if m % 2 == 1 { term } else { -term };
        sum += harmonic * signed;
        derivative_sum += (2.0 * m as f64 / x) * harmonic * signed;
        if signed.abs() <= 1e-15 * sum.abs().max(1.0) {
            break;
        }
        term *= x * x / (4.0 * (m + 1) as f64 * (m + 1) as f64);
        harmonic += 1.0 / (m + 1) as f64;
    }
    let factor = 2.0 / std::f64::consts::PI;
    let y0 = factor * (a * j0 + sum);
    if order == 0 {
        return Ok(y0);
    }
    let y1 = -factor * (j0 / x - a * j1 + derivative_sum);
    if order == 1 {
        return Ok(y1);
    }
    let mut previous = y0;
    let mut current = y1;
    for n in 1..order {
        let next = 2.0 * n as f64 / x * current - previous;
        previous = current;
        current = next;
    }
    Ok(current)
}

fn bessel_k(x: f64, order: usize) -> Result<f64, String> {
    const EULER_GAMMA: f64 = 0.5772156649015329;
    let i0 = bessel_first_kind(x, 0, true)?;
    let i1 = bessel_first_kind(x, 1, true)?;
    let a = (x / 2.0).ln() + EULER_GAMMA;
    let mut term = x * x / 4.0;
    let mut harmonic = 1.0;
    let mut sum = 0.0;
    let mut derivative_sum = 0.0;
    for m in 1..=10_000_u32 {
        sum += harmonic * term;
        derivative_sum += (2.0 * m as f64 / x) * harmonic * term;
        if term.abs() <= 1e-15 * sum.abs().max(1.0) {
            break;
        }
        term *= x * x / (4.0 * (m + 1) as f64 * (m + 1) as f64);
        harmonic += 1.0 / (m + 1) as f64;
    }
    let k0 = -a * i0 + sum;
    if order == 0 {
        return Ok(k0);
    }
    let k1 = i0 / x + a * i1 - derivative_sum;
    if order == 1 {
        return Ok(k1);
    }
    let mut previous = k0;
    let mut current = k1;
    for n in 1..order {
        let next = 2.0 * n as f64 / x * current + previous;
        previous = current;
        current = next;
    }
    Ok(current)
}

fn radix_limits(base: u32) -> (usize, u32, i64, i64) {
    match base {
        2 => (10, 10, -512, 511),
        8 => (10, 30, -(1i64 << 29), (1i64 << 29) - 1),
        16 => (10, 40, -(1i64 << 39), (1i64 << 39) - 1),
        _ => unreachable!(),
    }
}

fn radix_digit(value: char) -> Option<u32> {
    value.to_digit(16)
}

fn parse_radix_value(value: &Variant, base: u32) -> Result<i64, Variant> {
    let text = to_str(value);
    let text = text.trim();
    let (max_digits, bits, min_value, max_value) = radix_limits(base);
    if text.is_empty() {
        return Err(Variant::Error(ExcelError::Num));
    }
    let (negative, digits) = if let Some(rest) = text.strip_prefix('-') {
        (true, rest)
    } else {
        (false, text)
    };
    if digits.is_empty() || digits.len() > max_digits {
        return Err(Variant::Error(ExcelError::Num));
    }
    let mut magnitude = 0u64;
    for digit in digits.chars() {
        let digit = match radix_digit(digit.to_ascii_uppercase()) {
            Some(digit) if digit < base => digit as u64,
            _ => return Err(Variant::Error(ExcelError::Num)),
        };
        magnitude = match magnitude
            .checked_mul(base as u64)
            .and_then(|value| value.checked_add(digit))
        {
            Some(value) => value,
            None => return Err(Variant::Error(ExcelError::Num)),
        };
    }
    if negative {
        let signed = -(magnitude as i128);
        if signed < min_value as i128 {
            return Err(Variant::Error(ExcelError::Num));
        }
        return Ok(signed as i64);
    }
    // Excel uses a fixed-width two's-complement representation for a
    // max-width negative input (10 binary/octal/hex digits).
    let first_digit = radix_digit(digits.chars().next().unwrap().to_ascii_uppercase()).unwrap();
    if digits.len() == max_digits && first_digit >= base / 2 {
        let signed = magnitude as i128 - (1i128 << bits);
        if signed < min_value as i128 || signed > max_value as i128 {
            return Err(Variant::Error(ExcelError::Num));
        }
        Ok(signed as i64)
    } else if magnitude > max_value as u64 {
        Err(Variant::Error(ExcelError::Num))
    } else {
        Ok(magnitude as i64)
    }
}

fn format_radix_value(value: i64, base: u32, places: Option<usize>) -> Result<Variant, String> {
    let (max_digits, bits, min_value, max_value) = radix_limits(base);
    if value < min_value || value > max_value {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let negative = value < 0;
    let magnitude = if negative {
        ((1i128 << bits) + value as i128) as u64
    } else {
        value as u64
    };
    let mut digits = if magnitude == 0 {
        String::from("0")
    } else {
        let mut reversed = String::new();
        let mut remaining = magnitude;
        while remaining != 0 {
            let digit = (remaining % base as u64) as u32;
            reversed.push(
                std::char::from_digit(digit, base)
                    .unwrap()
                    .to_ascii_uppercase(),
            );
            remaining /= base as u64;
        }
        reversed.chars().rev().collect()
    };
    let required_digits = if negative { max_digits } else { digits.len() };
    if let Some(places) = places {
        if places == 0 || places < required_digits || places > max_digits {
            return Ok(Variant::Error(ExcelError::Num));
        }
        digits = format!("{digits:0>places$}");
    } else if negative {
        digits = format!("{digits:0>max_digits$}");
    }
    Ok(Variant::Str(digits))
}

fn optional_places(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Option<usize>, String> {
    if args.len() < 2 {
        return Ok(None);
    }
    let places = to_float(&evaluate(&args[1], cells)?)?;
    if !places.is_finite() || places < 0.0 || places.fract() != 0.0 {
        return Ok(Some(usize::MAX));
    }
    Ok(Some(places as usize))
}

fn func_dec_to_radix(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    base: u32,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err(format!("DEC2{base} requires 1 or 2 arguments"));
    }
    let number = to_float(&evaluate(&args[0], cells)?)?;
    let (_, _, min_value, max_value) = radix_limits(base);
    if !number.is_finite() || number.trunc() < min_value as f64 || number.trunc() > max_value as f64
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let places = optional_places(args, cells)?;
    if places == Some(usize::MAX) {
        return Ok(Variant::Error(ExcelError::Num));
    }
    format_radix_value(number.trunc() as i64, base, places)
}

fn func_radix_to_dec(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    base: u32,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err(format!("{base}2DEC requires 1 argument"));
    }
    match parse_radix_value(&evaluate(&args[0], cells)?, base) {
        Ok(value) => Ok(Variant::Integer(value)),
        Err(error) => Ok(error),
    }
}

fn func_radix_to_radix(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    source_base: u32,
    target_base: u32,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("base conversion requires 1 or 2 arguments".into());
    }
    let value = match parse_radix_value(&evaluate(&args[0], cells)?, source_base) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    let places = optional_places(args, cells)?;
    if places == Some(usize::MAX) {
        return Ok(Variant::Error(ExcelError::Num));
    }
    format_radix_value(value, target_base, places)
}

#[derive(Clone, Copy)]
struct UnitDef {
    category: &'static str,
    scale: f64,
    offset: f64,
}

fn unit_def(unit: &str) -> Option<UnitDef> {
    let unit = unit.to_ascii_lowercase().replace('^', "");
    let unit = unit.as_str();
    let def = match unit {
        "m" | "meter" | "meters" => UnitDef {
            category: "length",
            scale: 1.0,
            offset: 0.0,
        },
        "km" | "kilometer" | "kilometers" => UnitDef {
            category: "length",
            scale: 1_000.0,
            offset: 0.0,
        },
        "cm" | "centimeter" | "centimeters" => UnitDef {
            category: "length",
            scale: 0.01,
            offset: 0.0,
        },
        "mm" | "millimeter" | "millimeters" => UnitDef {
            category: "length",
            scale: 0.001,
            offset: 0.0,
        },
        "um" | "micron" | "microns" => UnitDef {
            category: "length",
            scale: 1e-6,
            offset: 0.0,
        },
        "in" | "inch" | "inches" => UnitDef {
            category: "length",
            scale: 0.0254,
            offset: 0.0,
        },
        "ft" | "foot" | "feet" => UnitDef {
            category: "length",
            scale: 0.3048,
            offset: 0.0,
        },
        "yd" | "yard" | "yards" => UnitDef {
            category: "length",
            scale: 0.9144,
            offset: 0.0,
        },
        "mi" | "mile" | "miles" => UnitDef {
            category: "length",
            scale: 1_609.344,
            offset: 0.0,
        },
        "nmi" | "nauticalmile" | "nauticalmiles" => UnitDef {
            category: "length",
            scale: 1_852.0,
            offset: 0.0,
        },
        "ang" | "angstrom" | "angstroms" => UnitDef {
            category: "length",
            scale: 1e-10,
            offset: 0.0,
        },
        "pica" | "picas" => UnitDef {
            category: "length",
            scale: 0.0003528,
            offset: 0.0,
        },
        "g" | "gram" | "grams" => UnitDef {
            category: "mass",
            scale: 0.001,
            offset: 0.0,
        },
        "kg" | "kilogram" | "kilograms" => UnitDef {
            category: "mass",
            scale: 1.0,
            offset: 0.0,
        },
        "mg" | "milligram" | "milligrams" => UnitDef {
            category: "mass",
            scale: 1e-6,
            offset: 0.0,
        },
        "lbm" | "lb" | "pound" | "pounds" => UnitDef {
            category: "mass",
            scale: 0.45359237,
            offset: 0.0,
        },
        "ozm" | "oz" | "ounce" | "ounces" => UnitDef {
            category: "mass",
            scale: 0.028349523125,
            offset: 0.0,
        },
        "stone" | "st" => UnitDef {
            category: "mass",
            scale: 6.35029318,
            offset: 0.0,
        },
        "ton" | "tonne" | "tonnes" => UnitDef {
            category: "mass",
            scale: 1_000.0,
            offset: 0.0,
        },
        "sg" | "slug" | "slugs" => UnitDef {
            category: "mass",
            scale: 14.59390294,
            offset: 0.0,
        },
        "u" | "amu" => UnitDef {
            category: "mass",
            scale: 1.66053906660e-27,
            offset: 0.0,
        },
        "s" | "sec" | "second" | "seconds" => UnitDef {
            category: "time",
            scale: 1.0,
            offset: 0.0,
        },
        "min" | "minute" | "minutes" => UnitDef {
            category: "time",
            scale: 60.0,
            offset: 0.0,
        },
        "h" | "hr" | "hour" | "hours" => UnitDef {
            category: "time",
            scale: 3_600.0,
            offset: 0.0,
        },
        "d" | "day" | "days" => UnitDef {
            category: "time",
            scale: 86_400.0,
            offset: 0.0,
        },
        "week" | "weeks" => UnitDef {
            category: "time",
            scale: 604_800.0,
            offset: 0.0,
        },
        "m2" | "squaremeter" | "squaremeters" => UnitDef {
            category: "area",
            scale: 1.0,
            offset: 0.0,
        },
        "km2" | "squarekilometer" | "squarekilometers" => UnitDef {
            category: "area",
            scale: 1e6,
            offset: 0.0,
        },
        "cm2" | "squarecentimeter" | "squarecentimeters" => UnitDef {
            category: "area",
            scale: 1e-4,
            offset: 0.0,
        },
        "ft2" | "squarefoot" | "squarefeet" => UnitDef {
            category: "area",
            scale: 0.09290304,
            offset: 0.0,
        },
        "in2" | "squareinch" | "squareinches" => UnitDef {
            category: "area",
            scale: 0.00064516,
            offset: 0.0,
        },
        "acre" | "acres" => UnitDef {
            category: "area",
            scale: 4_046.8564224,
            offset: 0.0,
        },
        "ha" | "hectare" | "hectares" => UnitDef {
            category: "area",
            scale: 10_000.0,
            offset: 0.0,
        },
        "l" | "liter" | "liters" | "litre" | "litres" => UnitDef {
            category: "volume",
            scale: 0.001,
            offset: 0.0,
        },
        "ml" | "milliliter" | "milliliters" | "millilitre" | "millilitres" => UnitDef {
            category: "volume",
            scale: 1e-6,
            offset: 0.0,
        },
        "m3" | "cubicmeter" | "cubicmeters" => UnitDef {
            category: "volume",
            scale: 1.0,
            offset: 0.0,
        },
        "cm3" | "cubiccentimeter" | "cubiccentimeters" => UnitDef {
            category: "volume",
            scale: 1e-6,
            offset: 0.0,
        },
        "gal" | "gallon" | "gallons" => UnitDef {
            category: "volume",
            scale: 0.003785411784,
            offset: 0.0,
        },
        "qt" | "quart" | "quarts" => UnitDef {
            category: "volume",
            scale: 0.000946352946,
            offset: 0.0,
        },
        "pt" | "pint" | "pints" => UnitDef {
            category: "volume",
            scale: 0.000473176473,
            offset: 0.0,
        },
        "cup" | "cups" => UnitDef {
            category: "volume",
            scale: 0.0002365882365,
            offset: 0.0,
        },
        "tsp" | "teaspoon" | "teaspoons" => UnitDef {
            category: "volume",
            scale: 4.92892159375e-6,
            offset: 0.0,
        },
        "tbsp" | "tbs" | "tablespoon" | "tablespoons" => UnitDef {
            category: "volume",
            scale: 1.478676478125e-5,
            offset: 0.0,
        },
        "in3" | "cubicinch" | "cubicinches" => UnitDef {
            category: "volume",
            scale: 1.6387064e-5,
            offset: 0.0,
        },
        "ft3" | "cubicfoot" | "cubicfeet" => UnitDef {
            category: "volume",
            scale: 0.028316846592,
            offset: 0.0,
        },
        "yd3" | "cubicyard" | "cubicyards" => UnitDef {
            category: "volume",
            scale: 0.764554857984,
            offset: 0.0,
        },
        "n" | "force_n" | "newton" | "newtons" => UnitDef {
            category: "force",
            scale: 1.0,
            offset: 0.0,
        },
        "dyn" | "dyne" | "dynes" => UnitDef {
            category: "force",
            scale: 1e-5,
            offset: 0.0,
        },
        "lbf" | "poundforce" => UnitDef {
            category: "force",
            scale: 4.4482216152605,
            offset: 0.0,
        },
        "pa" | "pascal" | "pascals" => UnitDef {
            category: "pressure",
            scale: 1.0,
            offset: 0.0,
        },
        "kpa" => UnitDef {
            category: "pressure",
            scale: 1_000.0,
            offset: 0.0,
        },
        "bar" => UnitDef {
            category: "pressure",
            scale: 100_000.0,
            offset: 0.0,
        },
        "atm" | "atmosphere" | "atmospheres" => UnitDef {
            category: "pressure",
            scale: 101_325.0,
            offset: 0.0,
        },
        "psi" => UnitDef {
            category: "pressure",
            scale: 6_894.757293168,
            offset: 0.0,
        },
        "mmhg" => UnitDef {
            category: "pressure",
            scale: 133.322387415,
            offset: 0.0,
        },
        "torr" => UnitDef {
            category: "pressure",
            scale: 133.322368421,
            offset: 0.0,
        },
        "j" | "joule" | "joules" => UnitDef {
            category: "energy",
            scale: 1.0,
            offset: 0.0,
        },
        "kj" => UnitDef {
            category: "energy",
            scale: 1_000.0,
            offset: 0.0,
        },
        "cal" | "calorie" | "calories" => UnitDef {
            category: "energy",
            scale: 4.184,
            offset: 0.0,
        },
        "kcal" => UnitDef {
            category: "energy",
            scale: 4_184.0,
            offset: 0.0,
        },
        "wh" => UnitDef {
            category: "energy",
            scale: 3_600.0,
            offset: 0.0,
        },
        "kwh" => UnitDef {
            category: "energy",
            scale: 3_600_000.0,
            offset: 0.0,
        },
        "btu" => UnitDef {
            category: "energy",
            scale: 1_055.05585262,
            offset: 0.0,
        },
        "ev" | "electronvolt" | "electronvolts" => UnitDef {
            category: "energy",
            scale: 1.602176634e-19,
            offset: 0.0,
        },
        "hz" | "hertz" => UnitDef {
            category: "frequency",
            scale: 1.0,
            offset: 0.0,
        },
        "khz" => UnitDef {
            category: "frequency",
            scale: 1e3,
            offset: 0.0,
        },
        "mhz" => UnitDef {
            category: "frequency",
            scale: 1e6,
            offset: 0.0,
        },
        "ghz" => UnitDef {
            category: "frequency",
            scale: 1e9,
            offset: 0.0,
        },
        "bit" | "bits" => UnitDef {
            category: "information",
            scale: 1.0,
            offset: 0.0,
        },
        "byte" | "bytes" => UnitDef {
            category: "information",
            scale: 8.0,
            offset: 0.0,
        },
        "kbit" => UnitDef {
            category: "information",
            scale: 1e3,
            offset: 0.0,
        },
        "kbyte" => UnitDef {
            category: "information",
            scale: 8e3,
            offset: 0.0,
        },
        "w" | "watt" | "watts" => UnitDef {
            category: "power",
            scale: 1.0,
            offset: 0.0,
        },
        "kw" => UnitDef {
            category: "power",
            scale: 1_000.0,
            offset: 0.0,
        },
        "mw" => UnitDef {
            category: "power",
            scale: 1_000_000.0,
            offset: 0.0,
        },
        "hp" | "horsepower" => UnitDef {
            category: "power",
            scale: 745.699871582,
            offset: 0.0,
        },
        "m/s" | "mps" => UnitDef {
            category: "speed",
            scale: 1.0,
            offset: 0.0,
        },
        "km/h" | "kph" => UnitDef {
            category: "speed",
            scale: 1.0 / 3.6,
            offset: 0.0,
        },
        "mph" => UnitDef {
            category: "speed",
            scale: 0.44704,
            offset: 0.0,
        },
        "knot" | "knots" => UnitDef {
            category: "speed",
            scale: 0.514444444444,
            offset: 0.0,
        },
        "c" | "celsius" => UnitDef {
            category: "temperature",
            scale: 1.0,
            offset: 0.0,
        },
        "f" | "fahrenheit" => UnitDef {
            category: "temperature",
            scale: 5.0 / 9.0,
            offset: -32.0,
        },
        "k" | "kelvin" => UnitDef {
            category: "temperature",
            scale: 1.0,
            offset: -273.15,
        },
        _ => return None,
    };
    Some(def)
}

fn func_convert(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("CONVERT requires 3 arguments".into());
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    let from = to_str(&evaluate(&args[1], cells)?);
    let to = to_str(&evaluate(&args[2], cells)?);
    let from_def = match unit_def(&from) {
        Some(def) => def,
        None => return Ok(Variant::Error(ExcelError::NA)),
    };
    let to_def = match unit_def(&to) {
        Some(def) => def,
        None => return Ok(Variant::Error(ExcelError::NA)),
    };
    if !value.is_finite() || from_def.category != to_def.category {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let base_value = (value + from_def.offset) * from_def.scale;
    let result = base_value / to_def.scale - to_def.offset;
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole(result))
}

fn roman_value(number: i64) -> String {
    const VALUES: &[(i64, &str)] = &[
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut remaining = number;
    let mut result = String::new();
    for &(value, symbol) in VALUES {
        while remaining >= value {
            result.push_str(symbol);
            remaining -= value;
        }
    }
    result
}

fn func_roman(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("ROMAN requires 1 or 2 arguments".into());
    }
    let number = to_float(&evaluate(&args[0], cells)?)?;
    if !number.is_finite() || number.fract() != 0.0 || !(1.0..=3999.0).contains(&number) {
        return Ok(Variant::Error(ExcelError::Value));
    }
    if args.len() == 2 {
        let form = to_float(&evaluate(&args[1], cells)?)?;
        if !form.is_finite() || form.fract() != 0.0 || !(0.0..=4.0).contains(&form) {
            return Ok(Variant::Error(ExcelError::Value));
        }
    }
    Ok(Variant::Str(roman_value(number as i64)))
}

fn roman_digit(value: char) -> Option<i64> {
    match value {
        'I' => Some(1),
        'V' => Some(5),
        'X' => Some(10),
        'L' => Some(50),
        'C' => Some(100),
        'D' => Some(500),
        'M' => Some(1000),
        _ => None,
    }
}

fn func_arabic(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("ARABIC requires 1 argument".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?).to_ascii_uppercase();
    if text.is_empty() {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let digits: Vec<i64> = match text.chars().map(roman_digit).collect() {
        Some(digits) => digits,
        None => return Ok(Variant::Error(ExcelError::Value)),
    };
    let mut total = 0;
    for i in 0..digits.len() {
        total += if i + 1 < digits.len() && digits[i] < digits[i + 1] {
            -digits[i]
        } else {
            digits[i]
        };
    }
    if !(1..=3999).contains(&total) || roman_value(total) != text {
        return Ok(Variant::Error(ExcelError::Value));
    }
    Ok(Variant::Integer(total))
}

#[derive(Clone, Copy)]
struct ComplexValue {
    re: f64,
    im: f64,
}

fn parse_complex(text: &str) -> Result<ComplexValue, String> {
    let normalized = text.trim().replace(' ', "");
    if normalized.is_empty() {
        return Err("complex value is empty".into());
    }
    let (body, has_suffix) = match normalized.chars().last() {
        Some('i' | 'j') => (&normalized[..normalized.len() - 1], true),
        _ => (normalized.as_str(), false),
    };
    if !has_suffix {
        return Ok(ComplexValue {
            re: body
                .parse()
                .map_err(|_| "invalid complex real part".to_string())?,
            im: 0.0,
        });
    }
    if body.is_empty() || body == "+" {
        return Ok(ComplexValue { re: 0.0, im: 1.0 });
    }
    if body == "-" {
        return Ok(ComplexValue { re: 0.0, im: -1.0 });
    }
    let split = body
        .char_indices()
        .skip(1)
        .find(|(_, ch)| *ch == '+' || *ch == '-')
        .map(|(index, _)| index);
    if let Some(index) = split {
        let re = body[..index]
            .parse()
            .map_err(|_| "invalid complex real part".to_string())?;
        let imaginary = &body[index..];
        let im = match imaginary {
            "+" => 1.0,
            "-" => -1.0,
            _ => imaginary
                .parse()
                .map_err(|_| "invalid complex imaginary part".to_string())?,
        };
        Ok(ComplexValue { re, im })
    } else {
        Ok(ComplexValue {
            re: 0.0,
            im: body
                .parse()
                .map_err(|_| "invalid complex imaginary part".to_string())?,
        })
    }
}

fn complex_number(value: f64) -> String {
    if value == 0.0 {
        "0".into()
    } else if value.fract() == 0.0 {
        (value as i64).to_string()
    } else {
        value.to_string()
    }
}

fn complex_text(value: ComplexValue, suffix: char) -> String {
    if value.im == 0.0 {
        return complex_number(value.re);
    }
    let imaginary = value.im.abs();
    let imaginary_text = if imaginary == 1.0 {
        suffix.to_string()
    } else {
        format!("{}{}", complex_number(imaginary), suffix)
    };
    if value.re == 0.0 {
        return if value.im < 0.0 {
            format!("-{}", imaginary_text)
        } else {
            imaginary_text
        };
    }
    format!(
        "{}{}{}",
        complex_number(value.re),
        if value.im < 0.0 { "-" } else { "+" },
        imaginary_text
    )
}

fn complex_argument(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<ComplexValue, String> {
    let value = evaluate(expr, cells)?;
    parse_complex(&to_str(&value)).map_err(|error| format!("{name}: {error}"))
}

fn func_complex(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("COMPLEX requires 2 or 3 arguments".into());
    }
    let re = to_float(&evaluate(&args[0], cells)?)?;
    let im = to_float(&evaluate(&args[1], cells)?)?;
    let suffix = if args.len() == 3 {
        let value = to_str(&evaluate(&args[2], cells)?).to_ascii_lowercase();
        match value.as_str() {
            "i" => 'i',
            "j" => 'j',
            _ => return Err("COMPLEX: suffix must be i or j".into()),
        }
    } else {
        'i'
    };
    if !re.is_finite() || !im.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Str(complex_text(ComplexValue { re, im }, suffix)))
}

fn func_imreal(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMREAL requires 1 argument".into());
    }
    Ok(as_integer_if_whole(
        complex_argument(&args[0], cells, "IMREAL")?.re,
    ))
}

fn func_imaginary(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMAGINARY requires 1 argument".into());
    }
    Ok(as_integer_if_whole(
        complex_argument(&args[0], cells, "IMAGINARY")?.im,
    ))
}

fn func_imargument(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMARGUMENT requires 1 argument".into());
    }
    let value = complex_argument(&args[0], cells, "IMARGUMENT")?;
    if value.re == 0.0 && value.im == 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(value.im.atan2(value.re)))
}

fn func_imabs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMABS requires 1 argument".into());
    }
    let value = complex_argument(&args[0], cells, "IMABS")?;
    Ok(as_integer_if_whole(value.re.hypot(value.im)))
}

fn complex_suffix(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<char, String> {
    let text = to_str(&evaluate(expr, cells)?);
    Ok(match text.chars().last() {
        Some('j' | 'J') => 'j',
        _ => 'i',
    })
}

fn func_imsum(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("IMSUM requires at least 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let mut result = ComplexValue { re: 0.0, im: 0.0 };
    for arg in args {
        let value = complex_argument(arg, cells, "IMSUM")?;
        result.re += value.re;
        result.im += value.im;
    }
    Ok(Variant::Str(complex_text(result, suffix)))
}

fn func_imsub(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("IMSUB requires 2 arguments".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let left = complex_argument(&args[0], cells, "IMSUB")?;
    let right = complex_argument(&args[1], cells, "IMSUB")?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: left.re - right.re,
            im: left.im - right.im,
        },
        suffix,
    )))
}

fn func_improduct(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("IMPRODUCT requires at least 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let mut result = ComplexValue { re: 1.0, im: 0.0 };
    for arg in args {
        let value = complex_argument(arg, cells, "IMPRODUCT")?;
        result = ComplexValue {
            re: result.re * value.re - result.im * value.im,
            im: result.re * value.im + result.im * value.re,
        };
    }
    Ok(Variant::Str(complex_text(result, suffix)))
}

fn func_imdiv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("IMDIV requires 2 arguments".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let left = complex_argument(&args[0], cells, "IMDIV")?;
    let right = complex_argument(&args[1], cells, "IMDIV")?;
    let denominator = right.re * right.re + right.im * right.im;
    if denominator == 0.0 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: (left.re * right.re + left.im * right.im) / denominator,
            im: (left.im * right.re - left.re * right.im) / denominator,
        },
        suffix,
    )))
}

fn func_imconjugate(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMCONJUGATE requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMCONJUGATE")?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: value.re,
            im: -value.im,
        },
        suffix,
    )))
}

fn func_imexp(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMEXP requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMEXP")?;
    let scale = value.re.exp();
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: scale * value.im.cos(),
            im: scale * value.im.sin(),
        },
        suffix,
    )))
}

fn func_imln(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMLN requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMLN")?;
    if value.re == 0.0 && value.im == 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: value.re.hypot(value.im).ln(),
            im: value.im.atan2(value.re),
        },
        suffix,
    )))
}

fn func_imlog10(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMLOG10 requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMLOG10")?;
    if value.re == 0.0 && value.im == 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let scale = 10.0_f64.ln();
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: value.re.hypot(value.im).ln() / scale,
            im: value.im.atan2(value.re) / scale,
        },
        suffix,
    )))
}

fn func_imlog2(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMLOG2 requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMLOG2")?;
    let logarithm = complex_ln_value(value).map_err(|error| format!("IMLOG2: {error:?}"))?;
    let scale = 2.0_f64.ln();
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: logarithm.re / scale,
            im: logarithm.im / scale,
        },
        suffix,
    )))
}

fn func_imsqrt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMSQRT requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMSQRT")?;
    let magnitude = value.re.hypot(value.im);
    let re = ((magnitude + value.re) / 2.0).sqrt();
    let im = ((magnitude - value.re) / 2.0).sqrt().copysign(value.im);
    Ok(Variant::Str(complex_text(ComplexValue { re, im }, suffix)))
}

fn func_impower(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("IMPOWER requires 2 arguments".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let base = complex_argument(&args[0], cells, "IMPOWER")?;
    let exponent = complex_argument(&args[1], cells, "IMPOWER")?;
    if base.re == 0.0 && base.im == 0.0 {
        if exponent.re > 0.0 && exponent.im == 0.0 {
            return Ok(Variant::Str("0".into()));
        }
        return Ok(Variant::Error(ExcelError::Num));
    }
    let ln_base = ComplexValue {
        re: base.re.hypot(base.im).ln(),
        im: base.im.atan2(base.re),
    };
    let log_product = ComplexValue {
        re: exponent.re * ln_base.re - exponent.im * ln_base.im,
        im: exponent.re * ln_base.im + exponent.im * ln_base.re,
    };
    let scale = log_product.re.exp();
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: scale * log_product.im.cos(),
            im: scale * log_product.im.sin(),
        },
        suffix,
    )))
}

fn complex_div_values(left: ComplexValue, right: ComplexValue) -> Result<ComplexValue, ExcelError> {
    let denominator = right.re * right.re + right.im * right.im;
    if denominator == 0.0 {
        return Err(ExcelError::DivZero);
    }
    Ok(ComplexValue {
        re: (left.re * right.re + left.im * right.im) / denominator,
        im: (left.im * right.re - left.re * right.im) / denominator,
    })
}

fn func_imsin(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMSIN requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMSIN")?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: value.re.sin() * value.im.cosh(),
            im: value.re.cos() * value.im.sinh(),
        },
        suffix,
    )))
}

fn func_imcos(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMCOS requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMCOS")?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: value.re.cos() * value.im.cosh(),
            im: -value.re.sin() * value.im.sinh(),
        },
        suffix,
    )))
}

fn func_imtan(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMTAN requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMTAN")?;
    let sin = ComplexValue {
        re: value.re.sin() * value.im.cosh(),
        im: value.re.cos() * value.im.sinh(),
    };
    let cos = ComplexValue {
        re: value.re.cos() * value.im.cosh(),
        im: -value.re.sin() * value.im.sinh(),
    };
    let result = complex_div_values(sin, cos).map_err(|error| format!("IMTAN: {error:?}"))?;
    Ok(Variant::Str(complex_text(result, suffix)))
}

fn func_imsinh(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMSINH requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMSINH")?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: value.re.sinh() * value.im.cos(),
            im: value.re.cosh() * value.im.sin(),
        },
        suffix,
    )))
}

fn func_imcosh(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMCOSH requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMCOSH")?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: value.re.cosh() * value.im.cos(),
            im: value.re.sinh() * value.im.sin(),
        },
        suffix,
    )))
}

fn func_imtanh(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMTANH requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMTANH")?;
    let sinh = ComplexValue {
        re: value.re.sinh() * value.im.cos(),
        im: value.re.cosh() * value.im.sin(),
    };
    let cosh = ComplexValue {
        re: value.re.cosh() * value.im.cos(),
        im: value.re.sinh() * value.im.sin(),
    };
    let result = complex_div_values(sinh, cosh).map_err(|error| format!("IMTANH: {error:?}"))?;
    Ok(Variant::Str(complex_text(result, suffix)))
}

fn func_imsec(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMSEC requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMSEC")?;
    let cos = ComplexValue {
        re: value.re.cos() * value.im.cosh(),
        im: -value.re.sin() * value.im.sinh(),
    };
    let result = complex_div_values(ComplexValue { re: 1.0, im: 0.0 }, cos)
        .map_err(|error| format!("IMSEC: {error:?}"))?;
    Ok(Variant::Str(complex_text(result, suffix)))
}

fn func_imcsc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMCSC requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMCSC")?;
    let sin = ComplexValue {
        re: value.re.sin() * value.im.cosh(),
        im: value.re.cos() * value.im.sinh(),
    };
    let result = complex_div_values(ComplexValue { re: 1.0, im: 0.0 }, sin)
        .map_err(|error| format!("IMCSC: {error:?}"))?;
    Ok(Variant::Str(complex_text(result, suffix)))
}

fn func_imcot(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMCOT requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMCOT")?;
    let sin = ComplexValue {
        re: value.re.sin() * value.im.cosh(),
        im: value.re.cos() * value.im.sinh(),
    };
    let cos = ComplexValue {
        re: value.re.cos() * value.im.cosh(),
        im: -value.re.sin() * value.im.sinh(),
    };
    let result = complex_div_values(cos, sin).map_err(|error| format!("IMCOT: {error:?}"))?;
    Ok(Variant::Str(complex_text(result, suffix)))
}

fn complex_sqrt_value(value: ComplexValue) -> ComplexValue {
    let magnitude = value.re.hypot(value.im);
    ComplexValue {
        re: ((magnitude + value.re) / 2.0).sqrt(),
        im: ((magnitude - value.re) / 2.0).sqrt().copysign(value.im),
    }
}

fn complex_ln_value(value: ComplexValue) -> Result<ComplexValue, ExcelError> {
    if value.re == 0.0 && value.im == 0.0 {
        return Err(ExcelError::Num);
    }
    Ok(ComplexValue {
        re: value.re.hypot(value.im).ln(),
        im: value.im.atan2(value.re),
    })
}

fn complex_mul_values(left: ComplexValue, right: ComplexValue) -> ComplexValue {
    ComplexValue {
        re: left.re * right.re - left.im * right.im,
        im: left.re * right.im + left.im * right.re,
    }
}

fn func_imasin(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMASIN requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMASIN")?;
    let square = complex_mul_values(value, value);
    let root = complex_sqrt_value(ComplexValue {
        re: 1.0 - square.re,
        im: -square.im,
    });
    let logarithm = complex_ln_value(ComplexValue {
        re: root.re - value.im,
        im: root.im + value.re,
    })
    .map_err(|error| format!("IMASIN: {error:?}"))?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: logarithm.im,
            im: -logarithm.re,
        },
        suffix,
    )))
}

fn func_imacos(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMACOS requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMACOS")?;
    let square = complex_mul_values(value, value);
    let root = complex_sqrt_value(ComplexValue {
        re: 1.0 - square.re,
        im: -square.im,
    });
    let logarithm = complex_ln_value(ComplexValue {
        re: root.re - value.im,
        im: root.im + value.re,
    })
    .map_err(|error| format!("IMACOS: {error:?}"))?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: std::f64::consts::FRAC_PI_2 - logarithm.im,
            im: logarithm.re,
        },
        suffix,
    )))
}

fn func_imatan(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMATAN requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMATAN")?;
    let left = complex_ln_value(ComplexValue {
        re: 1.0 + value.im,
        im: -value.re,
    })
    .map_err(|error| format!("IMATAN: {error:?}"))?;
    let right = complex_ln_value(ComplexValue {
        re: 1.0 - value.im,
        im: value.re,
    })
    .map_err(|error| format!("IMATAN: {error:?}"))?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: -(left.im - right.im) / 2.0,
            im: (left.re - right.re) / 2.0,
        },
        suffix,
    )))
}

fn func_imacot(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMACOT requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMACOT")?;
    let left = complex_ln_value(ComplexValue {
        re: 1.0 + value.im,
        im: -value.re,
    })
    .map_err(|error| format!("IMACOT: {error:?}"))?;
    let right = complex_ln_value(ComplexValue {
        re: 1.0 - value.im,
        im: value.re,
    })
    .map_err(|error| format!("IMACOT: {error:?}"))?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: std::f64::consts::FRAC_PI_2 + (left.im - right.im) / 2.0,
            im: -(left.re - right.re) / 2.0,
        },
        suffix,
    )))
}

fn func_imasinh(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMASINH requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMASINH")?;
    let square = complex_mul_values(value, value);
    let root = complex_sqrt_value(ComplexValue {
        re: square.re + 1.0,
        im: square.im,
    });
    let logarithm = complex_ln_value(ComplexValue {
        re: value.re + root.re,
        im: value.im + root.im,
    })
    .map_err(|error| format!("IMASINH: {error:?}"))?;
    Ok(Variant::Str(complex_text(logarithm, suffix)))
}

fn func_imacosh(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMACOSH requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMACOSH")?;
    let root_plus = complex_sqrt_value(ComplexValue {
        re: value.re + 1.0,
        im: value.im,
    });
    let root_minus = complex_sqrt_value(ComplexValue {
        re: value.re - 1.0,
        im: value.im,
    });
    let product = complex_mul_values(root_plus, root_minus);
    let logarithm = complex_ln_value(ComplexValue {
        re: value.re + product.re,
        im: value.im + product.im,
    })
    .map_err(|error| format!("IMACOSH: {error:?}"))?;
    Ok(Variant::Str(complex_text(logarithm, suffix)))
}

fn func_imatanh(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMATANH requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMATANH")?;
    let left = complex_ln_value(ComplexValue {
        re: 1.0 + value.re,
        im: value.im,
    })
    .map_err(|error| format!("IMATANH: {error:?}"))?;
    let right = complex_ln_value(ComplexValue {
        re: 1.0 - value.re,
        im: -value.im,
    })
    .map_err(|error| format!("IMATANH: {error:?}"))?;
    Ok(Variant::Str(complex_text(
        ComplexValue {
            re: (left.re - right.re) / 2.0,
            im: (left.im - right.im) / 2.0,
        },
        suffix,
    )))
}

fn func_imsech(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMSECH requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMSECH")?;
    let cosh = ComplexValue {
        re: value.re.cosh() * value.im.cos(),
        im: value.re.sinh() * value.im.sin(),
    };
    let result = complex_div_values(ComplexValue { re: 1.0, im: 0.0 }, cosh)
        .map_err(|error| format!("IMSECH: {error:?}"))?;
    Ok(Variant::Str(complex_text(result, suffix)))
}

fn func_imcsch(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMCSCH requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMCSCH")?;
    let sinh = ComplexValue {
        re: value.re.sinh() * value.im.cos(),
        im: value.re.cosh() * value.im.sin(),
    };
    let result = complex_div_values(ComplexValue { re: 1.0, im: 0.0 }, sinh)
        .map_err(|error| format!("IMCSCH: {error:?}"))?;
    Ok(Variant::Str(complex_text(result, suffix)))
}

fn func_imcoth(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("IMCOTH requires 1 argument".into());
    }
    let suffix = complex_suffix(&args[0], cells)?;
    let value = complex_argument(&args[0], cells, "IMCOTH")?;
    let sinh = ComplexValue {
        re: value.re.sinh() * value.im.cos(),
        im: value.re.cosh() * value.im.sin(),
    };
    let cosh = ComplexValue {
        re: value.re.cosh() * value.im.cos(),
        im: value.re.sinh() * value.im.sin(),
    };
    let result = complex_div_values(cosh, sinh).map_err(|error| format!("IMCOTH: {error:?}"))?;
    Ok(Variant::Str(complex_text(result, suffix)))
}

// ── Trigonometry ──────────────────────────────────────────────────────────────

fn func_pi(
    args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !args.is_empty() {
        return Err("PI takes no arguments".into());
    }
    Ok(Variant::Float(std::f64::consts::PI))
}

fn func_trig1(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    f: fn(f64) -> f64,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("Trig function requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    Ok(as_integer_if_whole(f(n)))
}

fn func_atan2(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("ATAN2 requires 2 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?; // Excel: x-coordinate first
    let y = to_float(&evaluate(&args[1], cells)?)?; // then y-coordinate
    if x == 0.0 && y == 0.0 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    Ok(Variant::Float(y.atan2(x)))
}

fn func_extended_trig(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err(format!("{name} requires 1 argument"));
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    if !value.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let result = match name {
        "SINH" => value.sinh(),
        "COSH" => value.cosh(),
        "TANH" => value.tanh(),
        "ASINH" => value.asinh(),
        "ACOSH" if value >= 1.0 => value.acosh(),
        "ACOSH" => return Ok(Variant::Error(ExcelError::Num)),
        "ATANH" if value.abs() < 1.0 => value.atanh(),
        "ATANH" => return Ok(Variant::Error(ExcelError::Num)),
        "SEC" if value.cos() != 0.0 => 1.0 / value.cos(),
        "SEC" => return Ok(Variant::Error(ExcelError::DivZero)),
        "SECH" => 1.0 / value.cosh(),
        "CSC" if value.sin() != 0.0 => 1.0 / value.sin(),
        "CSC" => return Ok(Variant::Error(ExcelError::DivZero)),
        "CSCH" if value != 0.0 => 1.0 / value.sinh(),
        "CSCH" => return Ok(Variant::Error(ExcelError::DivZero)),
        "COT" if value.tan() != 0.0 => 1.0 / value.tan(),
        "COT" => return Ok(Variant::Error(ExcelError::DivZero)),
        "COTH" if value != 0.0 => 1.0 / value.tanh(),
        "COTH" => return Ok(Variant::Error(ExcelError::DivZero)),
        "ACOT" => (std::f64::consts::FRAC_PI_2 - value.atan()).rem_euclid(std::f64::consts::PI),
        "ACOTH" if value.abs() > 1.0 => 0.5 * ((value + 1.0) / (value - 1.0)).ln(),
        "ACOTH" => return Ok(Variant::Error(ExcelError::Num)),
        _ => unreachable!(),
    };
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole(result))
}

// ── Info ──────────────────────────────────────────────────────────────────────

fn func_countblank(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let n = collect_all(args, cells)?
        .into_iter()
        .filter(|v| matches!(v, Variant::Empty) || matches!(v, Variant::Str(s) if s.is_empty()))
        .count();
    Ok(Variant::Integer(n as i64))
}

fn func_address(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("ADDRESS requires at least 2 arguments".into());
    }
    let row = to_float(&evaluate(&args[0], cells)?)? as u32;
    let col = to_float(&evaluate(&args[1], cells)?)? as u32;
    let abs_num = if args.len() >= 3 {
        to_float(&evaluate(&args[2], cells)?)? as i32
    } else {
        1
    };
    let col_str = num_to_col_letter(col);
    let addr = match abs_num {
        1 => format!("${}${}", col_str, row),
        2 => format!("{}${}", col_str, row),
        3 => format!("${}{}", col_str, row),
        _ => format!("{}{}", col_str, row),
    };
    Ok(Variant::Str(addr))
}

fn num_to_col_letter(mut col: u32) -> String {
    let mut s = String::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        s.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    s.chars().rev().collect()
}

// ── LET ───────────────────────────────────────────────────────────────────────

fn func_let(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    // =LET(x, val1, y, val2, ..., result_expr)
    // Must have odd number of args: 2*n+1
    if args.len() < 3 || args.len().is_multiple_of(2) {
        return Err("LET requires an odd number of arguments: LET(name, val, ..., result)".into());
    }
    let mut frame: HashMap<String, BindingValue> = HashMap::new();
    let mut i = 0;
    while i < args.len() - 1 {
        let name = match &args[i] {
            FormulaExpr::FuncCall { name, args } if args.is_empty() => name.clone(),
            _ => return Err("LET: name arguments must be identifiers".into()),
        };
        let val = evaluate(&args[i + 1], cells)?; // evaluated with current scope
        frame.insert(name, BindingValue::Value(val));
        i += 2;
    }
    push_bindings(frame);
    let result = evaluate(&args[args.len() - 1], cells);
    pop_bindings();
    result
}

// ── LAMBDA (structural — arg is inspected as AST, not evaluated to a value) ──

/// Extract `(params, body)` from a `LAMBDA(p1, p2, ..., body)` expression node.
fn extract_lambda(expr: &FormulaExpr) -> Result<(Vec<String>, &FormulaExpr), String> {
    match expr {
        FormulaExpr::FuncCall { name, args } if name.to_uppercase() == "LAMBDA" => {
            if args.len() < 2 {
                return Err("LAMBDA requires at least 2 arguments".into());
            }
            let params: Result<Vec<String>, String> = args[..args.len() - 1]
                .iter()
                .map(|a| match a {
                    FormulaExpr::FuncCall { name, args } if args.is_empty() => Ok(name.clone()),
                    _ => Err("LAMBDA: parameter names must be identifiers".into()),
                })
                .collect();
            Ok((params?, &args[args.len() - 1]))
        }
        _ => Err("expected a LAMBDA expression".into()),
    }
}

/// LAMBDA(...) in expression context just returns a sentinel — use extract_lambda at call sites.
fn func_lambda(
    _args: &[FormulaExpr],
    _cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    Ok(Variant::Str("#LAMBDA".into())) // placeholder (not useful as a value by itself)
}

/// Call a LAMBDA with the given argument values.
fn call_lambda(
    lambda_expr: &FormulaExpr,
    arg_vals: Vec<Variant>,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    call_lambda_bindings(
        lambda_expr,
        arg_vals.into_iter().map(BindingValue::Value).collect(),
        cells,
    )
}

fn evaluate_lambda_call(
    lambda_expr: &FormulaExpr,
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let arg_bindings = args
        .iter()
        .map(|arg| match arg {
            FormulaExpr::Omitted => Ok(BindingValue::Omitted),
            _ => evaluate(arg, cells).map(BindingValue::Value),
        })
        .collect::<Result<Vec<_>, _>>()?;
    call_lambda_bindings(lambda_expr, arg_bindings, cells)
}

fn call_lambda_bindings(
    lambda_expr: &FormulaExpr,
    arg_bindings: Vec<BindingValue>,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (params, body) = extract_lambda(lambda_expr)?;
    if arg_bindings.len() > params.len() {
        return Err(format!(
            "LAMBDA: expected at most {} args, got {}",
            params.len(),
            arg_bindings.len()
        ));
    }
    let frame: HashMap<String, BindingValue> = params
        .into_iter()
        .enumerate()
        .map(|(index, name)| {
            (
                name,
                arg_bindings
                    .get(index)
                    .cloned()
                    .unwrap_or(BindingValue::Omitted),
            )
        })
        .collect();
    push_bindings(frame);
    let result = evaluate(body, cells);
    pop_bindings();
    result
}

fn func_makearray(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("MAKEARRAY requires 3 arguments".into());
    }
    let rows = to_float(&evaluate(&args[0], cells)?)?;
    let cols = to_float(&evaluate(&args[1], cells)?)?;
    if !rows.is_finite()
        || !cols.is_finite()
        || rows < 1.0
        || cols < 1.0
        || rows.fract() != 0.0
        || cols.fract() != 0.0
    {
        return Err("MAKEARRAY dimensions must be positive integers".into());
    }
    let rows = rows as usize;
    let cols = cols as usize;
    let element_count = rows
        .checked_mul(cols)
        .ok_or_else(|| "MAKEARRAY dimensions are too large".to_string())?;
    if element_count > MAX_ARRAY_ELEMENTS {
        return Err("MAKEARRAY dimensions exceed the element limit".into());
    }
    let lambda = &args[2];
    let mut result = Vec::with_capacity(element_count);
    for row in 1..=rows {
        for col in 1..=cols {
            result.push(call_lambda(
                lambda,
                vec![Variant::Integer(row as i64), Variant::Integer(col as i64)],
                cells,
            )?);
        }
    }
    Ok(wrap_array(result))
}

// ── MAP ───────────────────────────────────────────────────────────────────────

fn func_map(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("MAP requires at least 2 arguments".into());
    }
    let lambda_expr = &args[args.len() - 1];
    // Single-array case: MAP(array, LAMBDA(x, body))
    // Multi-array case: MAP(a1, a2, ..., LAMBDA(x,y,...,body))
    let arrays: Vec<Vec<Variant>> = (0..args.len() - 1)
        .map(|i| collect_values(&args[i], cells))
        .collect::<Result<_, _>>()?;
    let len = arrays.first().map(|a| a.len()).unwrap_or(0);
    if arrays.iter().any(|a| a.len() != len) {
        return Err("MAP: all array arguments must have equal length".into());
    }
    let result: Result<Vec<Variant>, String> = (0..len)
        .map(|i| {
            let vals: Vec<Variant> = arrays.iter().map(|a| a[i].clone()).collect();
            call_lambda(lambda_expr, vals, cells)
        })
        .collect();
    Ok(wrap_array(result?))
}

// ── REDUCE ────────────────────────────────────────────────────────────────────

fn func_reduce(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 {
        return Err(
            "REDUCE requires 3 arguments: REDUCE(initial, array, LAMBDA(acc,x,body))".into(),
        );
    }
    let mut acc = evaluate(&args[0], cells)?;
    let data = collect_values(&args[1], cells)?;
    let lambda = &args[2];
    for val in data {
        acc = call_lambda(lambda, vec![acc, val], cells)?;
    }
    Ok(acc)
}

// ── SCAN ──────────────────────────────────────────────────────────────────────

fn func_scan(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 {
        return Err("SCAN requires 3 arguments: SCAN(initial, array, LAMBDA(acc,x,body))".into());
    }
    let mut acc = evaluate(&args[0], cells)?;
    let data = collect_values(&args[1], cells)?;
    let lambda = &args[2];
    let mut result = vec![];
    for val in data {
        acc = call_lambda(lambda, vec![acc.clone(), val], cells)?;
        result.push(acc.clone());
    }
    Ok(wrap_array(result))
}

// ── BYROW / BYCOL ─────────────────────────────────────────────────────────────

fn func_byrow(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("BYROW requires 2 arguments".into());
    }
    let lambda = &args[1];
    match &args[0] {
        FormulaExpr::Range { c1, r1, c2, r2, .. } => {
            let result: Result<Vec<Variant>, String> = (*r1..=*r2)
                .map(|row| {
                    let row_vals: Vec<Variant> =
                        (*c1..=*c2).map(|col| cell_val(cells, row, col)).collect();
                    call_lambda(lambda, row_vals, cells)
                })
                .collect();
            Ok(wrap_array(result?))
        }
        _ => {
            let val = evaluate(&args[0], cells)?;
            call_lambda(lambda, vec![val], cells)
        }
    }
}

fn func_bycol(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("BYCOL requires 2 arguments".into());
    }
    let lambda = &args[1];
    match &args[0] {
        FormulaExpr::Range { c1, r1, c2, r2, .. } => {
            let result: Result<Vec<Variant>, String> = (*c1..=*c2)
                .map(|col| {
                    let col_vals: Vec<Variant> =
                        (*r1..=*r2).map(|row| cell_val(cells, row, col)).collect();
                    call_lambda(lambda, col_vals, cells)
                })
                .collect();
            Ok(wrap_array(result?))
        }
        _ => {
            let val = evaluate(&args[0], cells)?;
            call_lambda(lambda, vec![val], cells)
        }
    }
}

// ── INDIRECT ──────────────────────────────────────────────────────────────────

fn func_indirect(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("INDIRECT requires 1 or 2 arguments".into());
    }
    let addr_str = match evaluate(&args[0], cells)? {
        Variant::Str(s) => s,
        other => return Err(format!("INDIRECT: expected string, got {}", other)),
    };
    let a1 = args
        .get(1)
        .map(|arg| evaluate(arg, cells).map(|value| is_truthy(&value)))
        .transpose()?
        .unwrap_or(true);
    let reference = if a1 {
        crate::types::parse_range_addr(addr_str.trim())
    } else {
        parse_r1c1_range(addr_str.trim())
    }
    .ok_or_else(|| format!("INDIRECT: invalid reference '{}'", addr_str))?;
    let ((r1, c1), (r2, c2)) = reference;
    let mut values = Vec::new();
    for row in r1.min(r2)..=r1.max(r2) {
        for col in c1.min(c2)..=c1.max(c2) {
            values.push(cell_val(cells, row, col));
        }
    }
    Ok(wrap_array(values))
}

/// Parse the absolute R1C1 form accepted by INDIRECT(...,FALSE).
/// Relative R\[delta\]C\[delta\] references need a caller cell context and are
/// deliberately rejected by this context-free formula evaluator.
fn parse_r1c1_range(addr: &str) -> Option<((u32, u32), (u32, u32))> {
    fn parse_piece(piece: &str) -> Option<(u32, u32)> {
        let piece = piece.trim().to_ascii_uppercase();
        let (row_text, col_text) = piece.strip_prefix('R')?.split_once('C')?;
        if row_text.is_empty()
            || col_text.is_empty()
            || row_text.starts_with('[')
            || col_text.starts_with('[')
        {
            return None;
        }
        let row = row_text.parse::<u32>().ok()?;
        let col = col_text.parse::<u32>().ok()?;
        if row == 0 || col == 0 {
            return None;
        }
        Some((row, col))
    }
    if let Some((left, right)) = addr.split_once(':') {
        Some((parse_piece(left)?, parse_piece(right)?))
    } else {
        let cell = parse_piece(addr)?;
        Some((cell, cell))
    }
}

// ── OFFSET ────────────────────────────────────────────────────────────────────

fn func_offset(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let values = match offset_values(args, cells) {
        Ok(values) => values,
        Err(error) if error.contains("outside the worksheet") => {
            return Ok(Variant::Error(ExcelError::Ref));
        }
        Err(error) => return Err(error),
    };
    Ok(wrap_array(values))
}

// ── Array / spill helpers ─────────────────────────────────────────────────────

/// Element-wise boolean comparison op (for array FILTER conditions).
fn compare_element(op: &BinOpKind, l: &Variant, r: &Variant) -> bool {
    match op {
        BinOpKind::Eq => variant_eq(l, r),
        BinOpKind::Ne => !variant_eq(l, r),
        BinOpKind::Lt => variant_cmp(l, r)
            .map(|o| o == Ordering::Less)
            .unwrap_or(false),
        BinOpKind::Le => variant_cmp(l, r)
            .map(|o| o != Ordering::Greater)
            .unwrap_or(false),
        BinOpKind::Gt => variant_cmp(l, r)
            .map(|o| o == Ordering::Greater)
            .unwrap_or(false),
        BinOpKind::Ge => variant_cmp(l, r)
            .map(|o| o != Ordering::Less)
            .unwrap_or(false),
        _ => is_truthy(l),
    }
}

/// Evaluate a formula expression to a `Vec<bool>`:
/// - Range → truthy check on each cell
/// - BinOp with a Range lhs → element-wise comparison
/// - Scalar → single-element vec
fn eval_as_bool_array(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Vec<bool>, String> {
    match expr {
        FormulaExpr::Range { .. } => {
            Ok(collect_values(expr, cells)?.iter().map(is_truthy).collect())
        }
        FormulaExpr::BinOp { op, lhs, rhs } => {
            if matches!(op, BinOpKind::Add | BinOpKind::Mul) {
                let lhs_flags = eval_as_bool_array(lhs, cells)?;
                let rhs_flags = eval_as_bool_array(rhs, cells)?;
                if lhs_flags.len() != 1
                    && rhs_flags.len() != 1
                    && lhs_flags.len() != rhs_flags.len()
                {
                    return Err("composite FILTER conditions must have equal length".into());
                }
                let len = lhs_flags.len().max(rhs_flags.len());
                return Ok((0..len)
                    .map(|index| {
                        let left = if lhs_flags.len() == 1 {
                            lhs_flags[0]
                        } else {
                            lhs_flags[index]
                        };
                        let right = if rhs_flags.len() == 1 {
                            rhs_flags[0]
                        } else {
                            rhs_flags[index]
                        };
                        if matches!(op, BinOpKind::Mul) {
                            left && right
                        } else {
                            left || right
                        }
                    })
                    .collect());
            }
            let lhs_vals = flatten_array_vals(collect_values(lhs, cells)?);
            let rhs_vals = flatten_array_vals(collect_values(rhs, cells)?);
            if lhs_vals.len() > 1 || rhs_vals.len() > 1 {
                if lhs_vals.len() != 1 && rhs_vals.len() != 1 && lhs_vals.len() != rhs_vals.len() {
                    return Err("array comparison operands must have equal length".into());
                }
                let len = lhs_vals.len().max(rhs_vals.len());
                Ok((0..len)
                    .map(|index| {
                        let left = if lhs_vals.len() == 1 {
                            &lhs_vals[0]
                        } else {
                            &lhs_vals[index]
                        };
                        let right = if rhs_vals.len() == 1 {
                            &rhs_vals[0]
                        } else {
                            &rhs_vals[index]
                        };
                        compare_element(op, left, right)
                    })
                    .collect())
            } else {
                Ok(vec![is_truthy(&evaluate(expr, cells)?)])
            }
        }
        _ => match evaluate(expr, cells)? {
            Variant::Array(values) => Ok(values.iter().map(is_truthy).collect()),
            value => Ok(vec![is_truthy(&value)]),
        },
    }
}

fn wrap_array(mut vals: Vec<Variant>) -> Variant {
    match vals.len() {
        0 => Variant::Empty,
        1 => vals.remove(0),
        _ => Variant::Array(vals),
    }
}

// ── FILTER ────────────────────────────────────────────────────────────────────

#[derive(Default)]
struct FilterXmlNode {
    name: String,
    text: String,
    attrs: HashMap<String, String>,
}

fn filterxml_name_and_attrs(raw: &str) -> Option<(String, HashMap<String, String>)> {
    let mut parts = raw.split_whitespace();
    let name = parts.next()?.trim().to_string();
    if name.is_empty() {
        return None;
    }
    let mut attrs = HashMap::new();
    for part in parts {
        let (key, value) = part.split_once('=')?;
        let value = value.trim_matches(['"', '\'']);
        attrs.insert(key.to_string(), value.to_string());
    }
    Some((name, attrs))
}

fn filterxml_xpath_match(path: &[String], query: &[&str], descendant: bool) -> bool {
    if path.len() < query.len() {
        return false;
    }
    let offset = if descendant {
        path.len() - query.len()
    } else {
        0
    };
    (!descendant && path.len() != query.len())
        || path[offset..]
            .iter()
            .zip(query)
            .all(|(name, expected)| *expected == "*" || name == expected)
}

fn func_filterxml(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("FILTERXML requires 2 arguments".into());
    }
    let xml = to_str(&evaluate(&args[0], cells)?);
    let xpath = to_str(&evaluate(&args[1], cells)?);
    let descendant = xpath.starts_with("//");
    let normalized = xpath.trim_start_matches('/');
    let mut query_parts = normalized
        .split('/')
        .filter(|part| !part.is_empty())
        .collect::<Vec<_>>();
    if query_parts.is_empty() || (!descendant && !xpath.starts_with('/')) {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let attribute = query_parts
        .last()
        .and_then(|part| part.strip_prefix('@'))
        .map(str::to_string);
    if attribute.is_some() {
        query_parts.pop();
    }
    if query_parts.last() == Some(&"text()") {
        query_parts.pop();
    }
    if query_parts.is_empty() {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let mut stack: Vec<FilterXmlNode> = Vec::new();
    let mut path: Vec<String> = Vec::new();
    let mut values = Vec::new();
    let mut cursor = 0;
    while cursor < xml.len() {
        let Some(relative_open) = xml[cursor..].find('<') else {
            if let Some(node) = stack.last_mut() {
                node.text.push_str(&xml[cursor..]);
            }
            break;
        };
        if relative_open > 0
            && let Some(node) = stack.last_mut()
        {
            node.text.push_str(&xml[cursor..cursor + relative_open]);
        }
        let open = cursor + relative_open;
        let Some(relative_end) = xml[open..].find('>') else {
            return Ok(Variant::Error(ExcelError::Value));
        };
        let end = open + relative_end;
        let token = xml[open + 1..end].trim();
        if token.starts_with('?') || token.starts_with('!') {
            cursor = end + 1;
            continue;
        }
        if let Some(close_name) = token.strip_prefix('/') {
            let Some(node) = stack.pop() else {
                return Ok(Variant::Error(ExcelError::Value));
            };
            if node.name != close_name.trim() {
                return Ok(Variant::Error(ExcelError::Value));
            }
            if filterxml_xpath_match(&path, &query_parts, descendant) {
                let value = attribute
                    .as_ref()
                    .and_then(|name| node.attrs.get(name).cloned())
                    .unwrap_or_else(|| node.text.trim().to_string());
                values.push(Variant::Str(value));
            }
            path.pop();
        } else {
            let self_closing = token.ends_with('/');
            let start = token.trim_end_matches('/').trim();
            let Some((name, attrs)) = filterxml_name_and_attrs(start) else {
                return Ok(Variant::Error(ExcelError::Value));
            };
            path.push(name.clone());
            stack.push(FilterXmlNode {
                name,
                text: String::new(),
                attrs,
            });
            if self_closing {
                let node = stack.pop().expect("just pushed");
                if filterxml_xpath_match(&path, &query_parts, descendant) {
                    let value = attribute
                        .as_ref()
                        .and_then(|name| node.attrs.get(name).cloned())
                        .unwrap_or_default();
                    values.push(Variant::Str(value));
                }
                path.pop();
            }
        }
        cursor = end + 1;
    }
    if !stack.is_empty() {
        return Ok(Variant::Error(ExcelError::Value));
    }
    if values.is_empty() {
        Ok(Variant::Error(ExcelError::NA))
    } else {
        Ok(wrap_array(values))
    }
}

fn groupby_aggregate(name: &str, values: &[Variant]) -> Result<Variant, String> {
    let nums: Vec<f64> = values.iter().filter_map(as_f64).collect();
    match name.to_ascii_uppercase().as_str() {
        "SUM" => Ok(as_integer_if_whole(nums.iter().sum())),
        "AVERAGE" => {
            if nums.is_empty() {
                return Ok(Variant::Error(ExcelError::DivZero));
            }
            Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
        }
        "COUNT" => Ok(Variant::Integer(nums.len() as i64)),
        "COUNTA" => Ok(Variant::Integer(
            values
                .iter()
                .filter(|value| !matches!(value, Variant::Empty))
                .count() as i64,
        )),
        "MAX" => nums
            .iter()
            .copied()
            .reduce(f64::max)
            .map(as_integer_if_whole)
            .ok_or_else(|| "GROUPBY: MAX has no numeric values".into()),
        "MIN" => nums
            .iter()
            .copied()
            .reduce(f64::min)
            .map(as_integer_if_whole)
            .ok_or_else(|| "GROUPBY: MIN has no numeric values".into()),
        "PRODUCT" => Ok(as_integer_if_whole(nums.iter().product())),
        _ => Err(format!("GROUPBY: unsupported aggregate {}", name)),
    }
}

fn groupby_reduce(
    reducer: &FormulaExpr,
    values: &[Variant],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    groupby_reduce_with_total(reducer, values, values, cells)
}

fn groupby_reduce_with_total(
    reducer: &FormulaExpr,
    values: &[Variant],
    total: &[Variant],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if matches!(reducer, FormulaExpr::FuncCall { name, .. } if name.eq_ignore_ascii_case("LAMBDA"))
    {
        let (params, _) = extract_lambda(reducer)?;
        let args = if params.len() >= 2 {
            vec![
                Variant::Array(values.to_vec()),
                Variant::Array(total.to_vec()),
            ]
        } else {
            vec![Variant::Array(values.to_vec())]
        };
        return call_lambda(reducer, args, cells);
    }
    let name = match reducer {
        FormulaExpr::FuncCall {
            name,
            args: call_args,
        } if call_args.is_empty() => name,
        FormulaExpr::Str(name) => name,
        _ => return Ok(Variant::Error(ExcelError::Value)),
    };
    groupby_aggregate(name, values)
}

fn func_groupby(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 8 {
        return Err("GROUPBY requires 3 to 8 arguments".into());
    }
    let mut group_values = flatten_array_vals(collect_values(&args[0], cells)?);
    let mut data_values = flatten_array_vals(collect_values(&args[1], cells)?);
    let (mut group_rows, group_cols) = array_shape_for_expr(&args[0], cells, group_values.len());
    let (data_rows, data_cols) = array_shape_for_expr(&args[1], cells, data_values.len());
    if group_cols == 0 || data_cols == 0 || group_rows != data_rows {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let reducer = &args[2];
    let reducer_vertical = matches!(
        reducer,
        FormulaExpr::FuncCall { name, args }
            if name.eq_ignore_ascii_case("VSTACK") && !args.is_empty()
    );
    let reducers: Vec<&FormulaExpr> = match reducer {
        FormulaExpr::FuncCall { name, args }
            if (name.eq_ignore_ascii_case("HSTACK") || name.eq_ignore_ascii_case("VSTACK")) =>
        {
            if args.is_empty() {
                return Ok(Variant::Error(ExcelError::Value));
            }
            args.iter().collect()
        }
        _ => vec![reducer],
    };
    if reducers.iter().any(|reducer| {
        !matches!(reducer, FormulaExpr::Str(_))
            && !matches!(reducer, FormulaExpr::FuncCall { name, .. } if name.eq_ignore_ascii_case("LAMBDA"))
            && !matches!(reducer, FormulaExpr::FuncCall { args, .. } if args.is_empty())
    }) {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let reducer_count = reducers.len();
    let output_value_count = if reducer_vertical {
        data_cols
    } else {
        data_cols.saturating_mul(reducer_count)
    };
    let integer_option = |index: usize, default: i64| -> Result<i64, String> {
        let Some(expr) = args.get(index) else {
            return Ok(default);
        };
        match evaluate(expr, cells)? {
            Variant::Integer(value) => Ok(value),
            Variant::Float(value) if value.is_finite() && value.fract() == 0.0 => Ok(value as i64),
            _ => Ok(i64::MIN),
        }
    };
    let field_headers = integer_option(3, 0)?;
    if !(0..=3).contains(&field_headers) {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let total_depth = integer_option(4, 0)?;
    if !(-2..=2).contains(&total_depth) {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let generated_headers = field_headers == 2;
    let input_headers = field_headers == 1 || field_headers == 3;
    let header_values = if input_headers {
        if group_rows < 2 {
            return Ok(Variant::Error(ExcelError::Value));
        }
        let mut headers = group_values[..group_cols].to_vec();
        if reducer_vertical {
            headers.extend(data_values[..data_cols].iter().cloned());
        } else {
            for _ in &reducers {
                headers.extend(data_values[..data_cols].iter().cloned());
            }
        }
        Some(headers)
    } else if generated_headers {
        let mut headers = (1..=group_cols)
            .map(|index| Variant::Str(format!("Field{index}")))
            .collect::<Vec<_>>();
        if reducer_vertical {
            headers.extend((1..=data_cols).map(|index| {
                if data_cols == 1 {
                    Variant::Str("Values".into())
                } else {
                    Variant::Str(format!("Value{index}"))
                }
            }));
        } else {
            for reducer_index in 0..reducer_count {
                headers.extend((1..=data_cols).map(|index| {
                    if data_cols == 1 && reducer_count == 1 {
                        Variant::Str("Values".into())
                    } else {
                        Variant::Str(format!("Value{}-{}", reducer_index + 1, index))
                    }
                }));
            }
        }
        Some(headers)
    } else {
        None
    };
    if input_headers {
        group_values.drain(..group_cols);
        data_values.drain(..data_cols);
        group_rows -= 1;
    }
    let filter = if let Some(filter_expr) = args.get(6) {
        let flags = eval_as_bool_array(filter_expr, cells)?;
        if flags.len() != group_rows {
            return Ok(Variant::Error(ExcelError::Value));
        }
        Some(flags)
    } else {
        None
    };
    let mut groups: Vec<(Vec<Variant>, Vec<Vec<Variant>>)> = Vec::new();
    let mut included_values = vec![Vec::new(); data_cols];
    for index in 0..group_rows {
        if filter.as_ref().is_some_and(|flags| !flags[index]) {
            continue;
        }
        let key = group_values[index * group_cols..(index + 1) * group_cols].to_vec();
        let row_values = data_values[index * data_cols..(index + 1) * data_cols].to_vec();
        for (column, value) in row_values.iter().enumerate() {
            included_values[column].push(value.clone());
        }
        if let Some((_, group_data)) = groups.iter_mut().find(|(candidate, _)| {
            candidate.len() == key.len()
                && candidate
                    .iter()
                    .zip(&key)
                    .all(|(left, right)| variant_eq(left, right))
        }) {
            group_data.push(row_values);
        } else {
            groups.push((key, vec![row_values]));
        }
    }
    let sort_orders = if let Some(sort_expr) = args.get(5) {
        let value = evaluate(sort_expr, cells)?;
        let scalar_order = |value: &Variant| match value {
            Variant::Integer(value) if *value == 1 || *value == -1 => Some(*value),
            Variant::Float(value) if *value == 1.0 || *value == -1.0 => Some(*value as i64),
            _ => None,
        };
        match value {
            Variant::Array(values) => {
                if values.len() != group_cols {
                    return Ok(Variant::Error(ExcelError::Value));
                }
                let Some(orders) = values.iter().map(scalar_order).collect::<Option<Vec<_>>>()
                else {
                    return Ok(Variant::Error(ExcelError::Value));
                };
                orders
            }
            value => {
                let Some(order) = scalar_order(&value) else {
                    return Ok(Variant::Error(ExcelError::Value));
                };
                vec![order; group_cols]
            }
        }
    } else {
        vec![1; group_cols]
    };
    if args.get(5).is_some() {
        groups.sort_by(|left, right| {
            left.0
                .iter()
                .zip(&right.0)
                .zip(&sort_orders)
                .map(|((left, right), order)| {
                    let ordering = variant_cmp(left, right)
                        .unwrap_or_else(|_| to_str(left).cmp(&to_str(right)));
                    if *order < 0 {
                        ordering.reverse()
                    } else {
                        ordering
                    }
                })
                .find(|ordering| *ordering != Ordering::Equal)
                .unwrap_or(Ordering::Equal)
        });
    }
    let width = group_cols + output_value_count;
    let reduce_rows =
        |reducer: &FormulaExpr, rows: &[Vec<Variant>]| -> Result<Vec<Variant>, String> {
            (0..data_cols)
                .map(|column| {
                    let values = rows
                        .iter()
                        .map(|row| row[column].clone())
                        .collect::<Vec<_>>();
                    groupby_reduce_with_total(reducer, &values, &included_values[column], cells)
                })
                .collect()
        };
    let totals = if total_depth != 0 {
        Some(
            reducers
                .iter()
                .map(|reducer| {
                    included_values
                        .iter()
                        .map(|values| groupby_reduce_with_total(reducer, values, values, cells))
                        .collect::<Result<Vec<_>, _>>()
                })
                .collect::<Result<Vec<_>, _>>()?
                .into_iter()
                .flatten()
                .collect::<Vec<_>>(),
        )
    } else {
        None
    };
    let mut result = Vec::with_capacity(
        header_values.as_ref().map_or(0, Vec::len)
            + (groups.len() + usize::from(totals.is_some())) * width,
    );
    if let Some(headers) = header_values {
        result.extend(headers);
    }
    if total_depth < 0 {
        result.push(Variant::Str("Total".into()));
        result.extend((1..group_cols).map(|_| Variant::Empty));
        result.extend(totals.clone().expect("total requested"));
    }
    for index in 0..groups.len() {
        let (key, group_data) = &groups[index];
        if total_depth == -2 && group_cols >= 2 {
            let prefix = &key[..group_cols.saturating_sub(1)];
            let parent_appeared_before = groups[..index]
                .iter()
                .any(|(candidate_key, _)| candidate_key[..group_cols.saturating_sub(1)] == *prefix);
            if !parent_appeared_before {
                let mut subtotal_rows = Vec::new();
                for (candidate_key, candidate_data) in &groups {
                    if candidate_key[..group_cols.saturating_sub(1)] == *prefix {
                        subtotal_rows.extend(candidate_data.iter().cloned());
                    }
                }
                result.extend(prefix.iter().cloned());
                result.push(Variant::Str("Subtotal".into()));
                for reducer in &reducers {
                    result.extend(reduce_rows(reducer, &subtotal_rows)?);
                }
            }
        }
        result.extend(key.clone());
        for reducer in &reducers {
            result.extend(reduce_rows(reducer, group_data)?);
        }
        if total_depth == 2 && group_cols >= 2 {
            let prefix = &key[..group_cols.saturating_sub(1)];
            let parent_appears_later = groups[index + 1..]
                .iter()
                .any(|(candidate_key, _)| candidate_key[..group_cols.saturating_sub(1)] == *prefix);
            if !parent_appears_later {
                let mut subtotal_rows = Vec::new();
                for (candidate_key, candidate_data) in &groups {
                    if candidate_key[..group_cols.saturating_sub(1)] == *prefix {
                        subtotal_rows.extend(candidate_data.iter().cloned());
                    }
                }
                result.extend(prefix.iter().cloned());
                result.push(Variant::Str("Subtotal".into()));
                for reducer in &reducers {
                    result.extend(reduce_rows(reducer, &subtotal_rows)?);
                }
            }
        }
    }
    if total_depth > 0 {
        result.push(Variant::Str("Total".into()));
        result.extend((1..group_cols).map(|_| Variant::Empty));
        result.extend(totals.expect("total requested"));
    }
    Ok(wrap_array(result))
}

fn func_pivotby(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(4..=11).contains(&args.len()) {
        return Err("PIVOTBY requires 4 to 11 arguments".into());
    }
    let mut row_values = flatten_array_vals(collect_values(&args[0], cells)?);
    let mut col_values = flatten_array_vals(collect_values(&args[1], cells)?);
    let mut data_values = flatten_array_vals(collect_values(&args[2], cells)?);
    let (mut row_rows, row_cols) = array_shape_for_expr(&args[0], cells, row_values.len());
    let (col_rows, col_cols) = array_shape_for_expr(&args[1], cells, col_values.len());
    let (data_rows, data_cols) = array_shape_for_expr(&args[2], cells, data_values.len());
    if row_cols == 0
        || col_cols == 0
        || data_cols == 0
        || row_rows != col_rows
        || row_rows != data_rows
        || data_values.len() != data_rows * data_cols
    {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let explicit_field_headers = args.get(4).is_some();
    let field_headers = match args.get(4).map(|expr| evaluate(expr, cells)).transpose()? {
        None => 3,
        Some(Variant::Integer(value)) => value,
        Some(Variant::Float(value)) if value.is_finite() && value.fract() == 0.0 => value as i64,
        Some(_) => i64::MIN,
    };
    if !(0..=3).contains(&field_headers) {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let automatic_input_headers = !explicit_field_headers
        && data_rows >= 2
        && matches!(data_values.first(), Some(Variant::Str(_)))
        && matches!(
            data_values.get(data_cols),
            Some(Variant::Integer(_) | Variant::Float(_))
        );
    let input_headers = if explicit_field_headers {
        matches!(field_headers, 1 | 3)
    } else {
        automatic_input_headers
    };
    if input_headers {
        if row_rows < 2 || col_rows < 2 || data_rows < 2 {
            return Ok(Variant::Error(ExcelError::Value));
        }
        row_values.drain(..row_cols);
        col_values.drain(..col_cols);
        data_values.drain(..data_cols);
        row_rows -= 1;
    }
    // Automatic (omitted) keeps the historical layout; explicit 0/1 omit
    // the generated header row, while 2/3 retain it.
    let show_field_header_row = if explicit_field_headers {
        matches!(field_headers, 2 | 3)
    } else if automatic_input_headers {
        row_cols > 1 || col_cols > 1
    } else {
        true
    };
    if let Some(filter_expr) = args.get(9) {
        let flags = eval_as_bool_array(filter_expr, cells)?;
        if flags.len() != row_rows {
            return Ok(Variant::Error(ExcelError::Value));
        }
        let selected = flags
            .into_iter()
            .enumerate()
            .filter_map(|(index, keep)| keep.then_some(index))
            .collect::<Vec<_>>();
        row_values = selected
            .iter()
            .flat_map(|&index| {
                row_values[index * row_cols..(index + 1) * row_cols]
                    .iter()
                    .cloned()
            })
            .collect();
        col_values = selected
            .iter()
            .flat_map(|&index| {
                col_values[index * col_cols..(index + 1) * col_cols]
                    .iter()
                    .cloned()
            })
            .collect();
        data_values = selected
            .iter()
            .flat_map(|&index| {
                data_values[index * data_cols..(index + 1) * data_cols]
                    .iter()
                    .cloned()
            })
            .collect();
    }
    let data_rows = row_values.len() / row_cols;
    let reducer = &args[3];
    let reducer_vertical = matches!(
        reducer,
        FormulaExpr::FuncCall { name, args }
            if name.eq_ignore_ascii_case("VSTACK") && !args.is_empty()
    );
    let reducers: Vec<&FormulaExpr> = match reducer {
        FormulaExpr::FuncCall { name, args }
            if name.eq_ignore_ascii_case("HSTACK") || name.eq_ignore_ascii_case("VSTACK") =>
        {
            if args.is_empty() {
                return Ok(Variant::Error(ExcelError::Value));
            }
            args.iter().collect()
        }
        _ => vec![reducer],
    };
    if reducers.iter().any(|reducer| {
        !matches!(reducer, FormulaExpr::Str(_))
            && !matches!(reducer, FormulaExpr::FuncCall { name, .. } if name.eq_ignore_ascii_case("LAMBDA"))
            && !matches!(reducer, FormulaExpr::FuncCall { args, .. } if args.is_empty())
    }) {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let reducer_count = reducers.len();
    let output_value_count = if reducer_vertical {
        data_cols
    } else {
        data_cols.saturating_mul(reducer_count)
    };
    let total_depth = |index: usize| -> Result<i64, String> {
        let Some(expr) = args.get(index) else {
            return Ok(0);
        };
        match evaluate(expr, cells)? {
            Variant::Integer(value) => Ok(value),
            Variant::Float(value) if value.is_finite() && value.fract() == 0.0 => Ok(value as i64),
            _ => Ok(i64::MIN),
        }
    };
    let row_total_depth = total_depth(5)?;
    let col_total_depth = total_depth(7)?;
    if !(-2..=2).contains(&row_total_depth) || !(-2..=2).contains(&col_total_depth) {
        return Ok(Variant::Error(ExcelError::Value));
    }
    if row_total_depth.unsigned_abs() >= 2 && row_cols < 2 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    if col_total_depth.unsigned_abs() >= 2 && col_cols < 2 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let relative_to = match args.get(10) {
        None | Some(FormulaExpr::Omitted) => 0,
        Some(expr) => match evaluate(expr, cells)? {
            Variant::Integer(value) => value,
            Variant::Float(value) if value.is_finite() && value.fract() == 0.0 => value as i64,
            _ => i64::MIN,
        },
    };
    if !(0..=4).contains(&relative_to) {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let mut row_keys: Vec<Vec<Variant>> = Vec::new();
    let mut col_keys: Vec<Vec<Variant>> = Vec::new();
    let key_equal = |left: &[Variant], right: &[Variant]| {
        left.len() == right.len()
            && left
                .iter()
                .zip(right)
                .all(|(left, right)| variant_eq(left, right))
    };
    for row in 0..data_rows {
        let key = row_values[row * row_cols..(row + 1) * row_cols].to_vec();
        if !row_keys.iter().any(|candidate| key_equal(candidate, &key)) {
            row_keys.push(key);
        }
        let key = col_values[row * col_cols..(row + 1) * col_cols].to_vec();
        if !col_keys.iter().any(|candidate| key_equal(candidate, &key)) {
            col_keys.push(key);
        }
    }
    let sort_order = |index: usize| -> Result<Option<Vec<i64>>, String> {
        let Some(expr) = args.get(index) else {
            return Ok(None);
        };
        let parse = |value: Variant| -> i64 {
            match value {
                Variant::Integer(value) => value,
                Variant::Float(value) if value.is_finite() && value.fract() == 0.0 => value as i64,
                _ => i64::MIN,
            }
        };
        Ok(Some(match evaluate(expr, cells)? {
            Variant::Array(values) => values.into_iter().map(parse).collect(),
            value => vec![parse(value)],
        }))
    };
    let row_sort_order = sort_order(6)?;
    let col_sort_order = sort_order(8)?;
    let valid_sort_order = |orders: &Option<Vec<i64>>, width: usize| {
        orders.as_ref().is_none_or(|orders| {
            !orders.is_empty()
                && orders.iter().all(|order| {
                    *order != 0
                        && if orders.len() == 1 {
                            order.unsigned_abs() as usize <= width + data_cols
                        } else {
                            order.unsigned_abs() as usize <= width
                        }
                })
        })
    };
    if !valid_sort_order(&row_sort_order, row_cols) || !valid_sort_order(&col_sort_order, col_cols)
    {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let sort_aggregate =
        |row_key: Option<&[Variant]>, col_key: Option<&[Variant]>, data_col: usize| {
            let group = row_values
                .chunks(row_cols)
                .zip(col_values.chunks(col_cols))
                .enumerate()
                .filter(|(index, (row, col))| {
                    row_key.is_none_or(|key| key_equal(key, row))
                        && col_key.is_none_or(|key| key_equal(key, col))
                        && *index < data_rows
                })
                .map(|(index, _)| data_values[index * data_cols + data_col].clone())
                .collect::<Vec<_>>();
            groupby_reduce(reducer, &group, cells).unwrap_or(Variant::Empty)
        };
    let sort_key_values = |keys: &[Vec<Variant>], axis_width: usize, rows_axis: bool| {
        keys.iter()
            .map(|key| {
                (0..data_cols)
                    .map(|data_col| {
                        if rows_axis {
                            sort_aggregate(Some(key), None, data_col)
                        } else {
                            sort_aggregate(None, Some(key), data_col)
                        }
                    })
                    .collect::<Vec<_>>()
            })
            .map(|values| (axis_width, values))
            .collect::<Vec<_>>()
    };
    let compare_keys = |left: (&[Variant], &[Variant]),
                        right: (&[Variant], &[Variant]),
                        orders: &[i64],
                        width: usize| {
        orders
            .iter()
            .map(|order| {
                let index = order.unsigned_abs() as usize - 1;
                let (left_value, right_value) = if index < width {
                    (&left.0[index], &right.0[index])
                } else {
                    (&left.1[index - width], &right.1[index - width])
                };
                let comparison = variant_cmp(left_value, right_value)
                    .unwrap_or_else(|_| to_str(left_value).cmp(&to_str(right_value)));
                if *order < 0 {
                    comparison.reverse()
                } else {
                    comparison
                }
            })
            .find(|ordering| *ordering != Ordering::Equal)
            .unwrap_or(Ordering::Equal)
    };
    if let Some(orders) = row_sort_order {
        let values = sort_key_values(&row_keys, row_cols, true);
        let mut paired = row_keys.into_iter().zip(values).collect::<Vec<_>>();
        paired.sort_by(|left, right| {
            compare_keys(
                (&left.0, &left.1.1),
                (&right.0, &right.1.1),
                &orders,
                row_cols,
            )
        });
        row_keys = paired.into_iter().map(|(key, _)| key).collect();
    }
    if let Some(orders) = col_sort_order {
        let values = sort_key_values(&col_keys, col_cols, false);
        let mut paired = col_keys.into_iter().zip(values).collect::<Vec<_>>();
        paired.sort_by(|left, right| {
            compare_keys(
                (&left.0, &left.1.1),
                (&right.0, &right.1.1),
                &orders,
                col_cols,
            )
        });
        col_keys = paired.into_iter().map(|(key, _)| key).collect();
    }
    let mut col_segments: Vec<(Vec<Variant>, bool)> = Vec::new();
    for (index, col_key) in col_keys.iter().enumerate() {
        let parent = &col_key[..col_cols.saturating_sub(1)];
        let first_in_parent =
            index == 0 || col_keys[index - 1][..col_cols.saturating_sub(1)] != *parent;
        let last_in_parent = index + 1 == col_keys.len()
            || col_keys[index + 1][..col_cols.saturating_sub(1)] != *parent;
        if col_total_depth == -2 && first_in_parent {
            col_segments.push((parent.to_vec(), true));
        }
        col_segments.push((col_key.clone(), false));
        if col_total_depth == 2 && last_in_parent {
            col_segments.push((parent.to_vec(), true));
        }
    }
    let total_column = row_total_depth != 0 || col_total_depth != 0;
    let row_total_cols = usize::from(total_column) * output_value_count;
    let output_cols = row_cols + col_segments.len() * output_value_count + row_total_cols;
    let row_subtotal_count = if row_total_depth.unsigned_abs() >= 2 {
        row_keys
            .windows(2)
            .filter(|keys| keys[0][..row_cols - 1] != keys[1][..row_cols - 1])
            .count()
            + usize::from(!row_keys.is_empty())
    } else {
        0
    };
    let output_rows = row_keys.len() + row_subtotal_count + 1 + usize::from(col_total_depth != 0);
    let mut result = Vec::with_capacity(output_rows * output_cols);
    if show_field_header_row {
        if col_total_depth < 0 {
            result.push(Variant::Str("Total".into()));
            result.extend(std::iter::repeat_n(Variant::Empty, row_cols - 1));
        } else {
            result.extend(std::iter::repeat_n(Variant::Empty, row_cols));
        }
    } else if col_total_depth < 0 {
        result.extend(std::iter::repeat_n(Variant::Empty, row_cols));
    }
    let collect_matching = |row_key: Option<&[Variant]>,
                            row_prefix: bool,
                            col_key: Option<&[Variant]>,
                            col_prefix: bool,
                            data_col: usize| {
        row_values
            .chunks(row_cols)
            .zip(col_values.chunks(col_cols))
            .enumerate()
            .filter(|(index, (row, col))| {
                let row_matches = row_key.is_none_or(|key| {
                    if row_prefix {
                        row.starts_with(key)
                    } else {
                        key_equal(key, row)
                    }
                });
                let col_matches = col_key.is_none_or(|key| {
                    if col_prefix {
                        col.starts_with(key)
                    } else {
                        key_equal(key, col)
                    }
                });
                row_matches && col_matches && *index < data_rows
            })
            .map(|(index, _)| data_values[index * data_cols + data_col].clone())
            .collect::<Vec<_>>()
    };
    let reduce_values = |reducer: &FormulaExpr, subset: &[Variant], total: &[Variant]| {
        if matches!(
            reducer,
            FormulaExpr::FuncCall { name, .. } if name.eq_ignore_ascii_case("LAMBDA")
        ) {
            let (params, _) = extract_lambda(reducer)?;
            let args = if params.len() >= 2 {
                vec![
                    Variant::Array(subset.to_vec()),
                    Variant::Array(total.to_vec()),
                ]
            } else {
                vec![Variant::Array(subset.to_vec())]
            };
            return call_lambda(reducer, args, cells);
        }
        let is_percentof = matches!(
            reducer,
            FormulaExpr::FuncCall { name, args }
                if name.eq_ignore_ascii_case("PERCENTOF") && args.is_empty()
        );
        if !is_percentof {
            return groupby_reduce(reducer, subset, cells);
        }
        let sum = |values: &[Variant]| values.iter().filter_map(as_f64).sum::<f64>();
        let numerator = sum(subset);
        let denominator = sum(total);
        if !numerator.is_finite() || !denominator.is_finite() {
            return Ok(Variant::Error(ExcelError::Num));
        }
        if denominator == 0.0 {
            return Ok(Variant::Error(ExcelError::DivZero));
        }
        Ok(as_integer_if_whole(numerator / denominator))
    };
    let reduce_matching = |reducer: &FormulaExpr,
                           row_key: Option<&[Variant]>,
                           row_prefix: bool,
                           col_key: Option<&[Variant]>,
                           col_prefix: bool,
                           data_col: usize| {
        let subset = collect_matching(row_key, row_prefix, col_key, col_prefix, data_col);
        let total = match relative_to {
            0 => collect_matching(None, false, col_key, col_prefix, data_col),
            1 => collect_matching(row_key, row_prefix, None, false, data_col),
            2 => collect_matching(None, false, None, false, data_col),
            3 => {
                let parent = col_key.map(|key| &key[..key.len().saturating_sub(1)]);
                collect_matching(None, false, parent, true, data_col)
            }
            4 => {
                let parent = row_key.map(|key| &key[..key.len().saturating_sub(1)]);
                collect_matching(parent, true, None, false, data_col)
            }
            _ => unreachable!(),
        };
        reduce_values(reducer, &subset, &total)
    };
    let aggregate = |reducer: &FormulaExpr,
                     row_key: Option<&[Variant]>,
                     col_key: Option<&[Variant]>,
                     data_col: usize| {
        reduce_matching(reducer, row_key, false, col_key, false, data_col)
    };
    let aggregate_row_prefix = |reducer: &FormulaExpr,
                                prefix: &[Variant],
                                col_key: Option<&[Variant]>,
                                data_col: usize| {
        reduce_matching(reducer, Some(prefix), true, col_key, false, data_col)
    };
    let aggregate_col_prefix = |reducer: &FormulaExpr,
                                row_key: Option<&[Variant]>,
                                prefix: &[Variant],
                                data_col: usize| {
        reduce_matching(reducer, row_key, false, Some(prefix), true, data_col)
    };
    let aggregate_segment = |reducer: &FormulaExpr,
                             row_key: Option<&[Variant]>,
                             segment: &(Vec<Variant>, bool),
                             data_col: usize| {
        if segment.1 {
            aggregate_col_prefix(reducer, row_key, &segment.0, data_col)
        } else {
            aggregate(reducer, row_key, Some(segment.0.as_slice()), data_col)
        }
    };
    let aggregate_row_segment = |reducer: &FormulaExpr,
                                 prefix: &[Variant],
                                 segment: &(Vec<Variant>, bool),
                                 data_col: usize| {
        reduce_matching(
            reducer,
            Some(prefix),
            true,
            Some(segment.0.as_slice()),
            segment.1,
            data_col,
        )
    };
    if show_field_header_row {
        for (col_key, subtotal) in &col_segments {
            let label = if *subtotal {
                Variant::Str("Subtotal".into())
            } else {
                col_key.last().cloned().unwrap_or(Variant::Empty)
            };
            if reducer_vertical {
                result.extend(std::iter::repeat_n(label, data_cols));
            } else {
                for _ in &reducers {
                    result.extend(std::iter::repeat_n(label.clone(), data_cols));
                }
            }
        }
        if total_column {
            result.extend(std::iter::repeat_n(
                Variant::Str("Total".into()),
                output_value_count,
            ));
        }
    }
    if col_total_depth < 0 {
        if reducer_vertical {
            for reducer in &reducers {
                result.extend(std::iter::repeat_n(Variant::Empty, row_cols));
                for segment in &col_segments {
                    for data_col in 0..data_cols {
                        result.push(aggregate_segment(reducer, None, segment, data_col)?);
                    }
                }
                for data_col in 0..data_cols {
                    result.push(aggregate(reducer, None, None, data_col)?);
                }
            }
        } else {
            for segment in &col_segments {
                for reducer in &reducers {
                    for data_col in 0..data_cols {
                        result.push(aggregate_segment(reducer, None, segment, data_col)?);
                    }
                }
            }
            for reducer in &reducers {
                for data_col in 0..data_cols {
                    result.push(aggregate(reducer, None, None, data_col)?);
                }
            }
        }
    }
    let emit_subtotal = |result: &mut Vec<Variant>, prefix: &[Variant]| -> Result<(), String> {
        result.extend(prefix.iter().cloned());
        result.push(Variant::Str("Subtotal".into()));
        if reducer_vertical {
            for reducer in &reducers {
                for segment in &col_segments {
                    for data_col in 0..data_cols {
                        result.push(aggregate_row_segment(reducer, prefix, segment, data_col)?);
                    }
                }
                if row_total_depth != 0 {
                    for data_col in 0..data_cols {
                        result.push(aggregate_row_prefix(reducer, prefix, None, data_col)?);
                    }
                } else if total_column {
                    result.extend(std::iter::repeat_n(Variant::Empty, output_value_count));
                }
            }
        } else {
            for segment in &col_segments {
                for reducer in &reducers {
                    for data_col in 0..data_cols {
                        result.push(aggregate_row_segment(reducer, prefix, segment, data_col)?);
                    }
                }
            }
            if row_total_depth != 0 {
                for reducer in &reducers {
                    for data_col in 0..data_cols {
                        result.push(aggregate_row_prefix(reducer, prefix, None, data_col)?);
                    }
                }
            } else if total_column {
                result.extend(std::iter::repeat_n(Variant::Empty, output_value_count));
            }
        }
        Ok(())
    };
    for (row_index, row_key) in row_keys.iter().enumerate() {
        if row_total_depth == -2
            && (row_index == 0
                || row_keys[row_index - 1][..row_cols - 1] != row_key[..row_cols - 1])
        {
            emit_subtotal(result.as_mut(), &row_key[..row_cols - 1])?;
        }
        if reducer_vertical {
            for reducer in &reducers {
                result.extend(row_key.clone());
                for segment in &col_segments {
                    for data_col in 0..data_cols {
                        result.push(aggregate_segment(
                            reducer,
                            Some(row_key.as_slice()),
                            segment,
                            data_col,
                        )?);
                    }
                }
                if row_total_depth != 0 {
                    for data_col in 0..data_cols {
                        result.push(aggregate(
                            reducer,
                            Some(row_key.as_slice()),
                            None,
                            data_col,
                        )?);
                    }
                } else if total_column {
                    result.extend(std::iter::repeat_n(Variant::Empty, output_value_count));
                }
            }
        } else {
            result.extend(row_key.clone());
            for segment in &col_segments {
                for reducer in &reducers {
                    for data_col in 0..data_cols {
                        result.push(aggregate_segment(
                            reducer,
                            Some(row_key.as_slice()),
                            segment,
                            data_col,
                        )?);
                    }
                }
            }
            if row_total_depth != 0 {
                for reducer in &reducers {
                    for data_col in 0..data_cols {
                        result.push(aggregate(
                            reducer,
                            Some(row_key.as_slice()),
                            None,
                            data_col,
                        )?);
                    }
                }
            } else if total_column {
                result.extend(std::iter::repeat_n(Variant::Empty, output_value_count));
            }
        }
        if row_total_depth == 2
            && (row_index + 1 == row_keys.len()
                || row_keys[row_index + 1][..row_cols - 1] != row_key[..row_cols - 1])
        {
            emit_subtotal(result.as_mut(), &row_key[..row_cols - 1])?;
        }
    }
    if col_total_depth > 0 {
        if reducer_vertical {
            for reducer in &reducers {
                if show_field_header_row {
                    result.push(Variant::Str("Total".into()));
                    result.extend(std::iter::repeat_n(Variant::Empty, row_cols - 1));
                } else {
                    result.extend(std::iter::repeat_n(Variant::Empty, row_cols));
                }
                for segment in &col_segments {
                    for data_col in 0..data_cols {
                        result.push(aggregate_segment(reducer, None, segment, data_col)?);
                    }
                }
                for data_col in 0..data_cols {
                    result.push(aggregate(reducer, None, None, data_col)?);
                }
            }
        } else {
            if show_field_header_row {
                result.push(Variant::Str("Total".into()));
                result.extend(std::iter::repeat_n(Variant::Empty, row_cols - 1));
            } else {
                result.extend(std::iter::repeat_n(Variant::Empty, row_cols));
            }
            for segment in &col_segments {
                for reducer in &reducers {
                    for data_col in 0..data_cols {
                        result.push(aggregate_segment(reducer, None, segment, data_col)?);
                    }
                }
            }
            for reducer in &reducers {
                for data_col in 0..data_cols {
                    result.push(aggregate(reducer, None, None, data_col)?);
                }
            }
        }
    }
    Ok(wrap_array(result))
}

fn func_filter(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("FILTER requires at least 2 arguments".into());
    }
    let include = eval_as_bool_array(&args[1], cells)?;

    let data_values = flatten_array_vals(collect_values(&args[0], cells)?);
    let (data_rows, data_cols) = array_shape_for_expr(&args[0], cells, data_values.len());
    let (include_rows, include_cols) = array_shape_for_expr(&args[1], cells, include.len());
    let row_include = include_rows == data_rows && include_cols == 1;
    let column_include = include_rows == 1 && include_cols == data_cols;
    if !matches!(args[0], FormulaExpr::Range { .. })
        && data_rows > 1
        && data_cols > 1
        && (row_include || column_include)
    {
        let mut result = Vec::new();
        if row_include {
            for (row, keep) in include.iter().copied().enumerate() {
                if keep {
                    result.extend(
                        data_values[row * data_cols..(row + 1) * data_cols]
                            .iter()
                            .cloned(),
                    );
                }
            }
        } else {
            for row in 0..data_rows {
                for (col, keep) in include.iter().copied().enumerate() {
                    if keep {
                        result.push(data_values[row * data_cols + col].clone());
                    }
                }
            }
        }
        if result.is_empty() {
            return if args.len() >= 3 {
                evaluate(&args[2], cells)
            } else {
                Ok(Variant::Error(ExcelError::NA))
            };
        }
        return Ok(wrap_array(result));
    }
    if !matches!(args[0], FormulaExpr::Range { .. }) && data_rows > 1 && data_cols > 1 {
        return Err(format!(
            "FILTER: 2D data requires a {}-row or {}-column include vector",
            data_rows, data_cols
        ));
    }

    match &args[0] {
        FormulaExpr::Range { c1, r1, c2, r2, .. } => {
            let data_rows = (*r2 - *r1 + 1) as usize;
            if include.len() != data_rows {
                return Err(format!(
                    "FILTER: data has {} rows but include has {} elements",
                    data_rows,
                    include.len()
                ));
            }
            let mut result = vec![];
            for (i, inc) in include.iter().enumerate() {
                if *inc {
                    let row = *r1 + i as u32;
                    for col in *c1..=*c2 {
                        result.push(cell_val(cells, row, col));
                    }
                }
            }
            if result.is_empty() {
                return if args.len() >= 3 {
                    evaluate(&args[2], cells)
                } else {
                    Ok(Variant::Error(ExcelError::NA))
                };
            }
            Ok(wrap_array(result))
        }
        _ => evaluate(&args[0], cells),
    }
}

// ── UNIQUE ────────────────────────────────────────────────────────────────────

fn func_unique(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("UNIQUE requires 1 argument".into());
    }
    let values = flatten_array_vals(collect_values(&args[0], cells)?);
    let (rows, cols) = array_shape_for_expr(&args[0], cells, values.len());
    if rows > 1 && cols > 1 {
        let exactly_once = args
            .get(1)
            .map(|arg| evaluate(arg, cells))
            .transpose()?
            .is_some_and(|value| is_truthy(&value));
        let by_col = args
            .get(2)
            .map(|arg| evaluate(arg, cells))
            .transpose()?
            .is_some_and(|value| is_truthy(&value));
        let outer = if by_col { cols } else { rows };
        let inner = if by_col { rows } else { cols };
        let mut groups: Vec<Vec<Variant>> = Vec::new();
        let mut counts: Vec<usize> = Vec::new();
        for index in 0..outer {
            let group: Vec<Variant> = if by_col {
                (0..inner)
                    .map(|offset| values[offset * cols + index].clone())
                    .collect()
            } else {
                values[index * cols..(index + 1) * cols].to_vec()
            };
            if let Some(existing) = groups.iter().position(|candidate| {
                candidate.len() == group.len()
                    && candidate
                        .iter()
                        .zip(&group)
                        .all(|(left, right)| variant_eq(left, right))
            }) {
                counts[existing] += 1;
            } else {
                groups.push(group);
                counts.push(1);
            }
        }
        let selected: Vec<Vec<Variant>> = groups
            .into_iter()
            .zip(counts)
            .filter(|(_, count)| !exactly_once || *count == 1)
            .map(|(group, _)| group)
            .collect();
        let mut result = Vec::new();
        if by_col {
            for row in 0..selected.first().map_or(0, Vec::len) {
                for group in &selected {
                    result.push(group[row].clone());
                }
            }
        } else {
            for group in selected {
                result.extend(group);
            }
        }
        return Ok(wrap_array(result));
    }
    // Collect non-empty values with their original positions
    let mut indexed: Vec<(Variant, usize)> = values
        .into_iter()
        .enumerate()
        .filter(|(_, v)| !matches!(v, Variant::Empty))
        .map(|(i, v)| (v, i))
        .collect();
    // Sort by (value, original_index) so equal values have lowest index first
    indexed.sort_by(|(a, ai), (b, bi)| {
        variant_cmp(a, b)
            .unwrap_or(Ordering::Equal)
            .then(ai.cmp(bi))
    });
    // Dedup by value — keeps first occurrence (lowest original index) of each group
    indexed.dedup_by(|(a, _), (b, _)| variant_eq(a, b));
    // Restore original insertion order
    indexed.sort_by_key(|(_, i)| *i);
    Ok(wrap_array(indexed.into_iter().map(|(v, _)| v).collect()))
}

// ── SORT ──────────────────────────────────────────────────────────────────────

fn func_sort(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("SORT requires 1 argument".into());
    }
    let mut vals = flatten_array_vals(collect_values(&args[0], cells)?);
    let (rows, cols) = array_shape_for_expr(&args[0], cells, vals.len());
    if rows > 1 && cols > 1 {
        let sort_index = if args.len() >= 2 {
            match evaluate(&args[1], cells)? {
                Variant::Integer(value) => value,
                Variant::Float(value)
                    if value.is_finite()
                        && value.fract() == 0.0
                        && value >= i64::MIN as f64
                        && value <= i64::MAX as f64 =>
                {
                    value as i64
                }
                _ => return Ok(Variant::Error(ExcelError::Value)),
            }
        } else {
            1
        };
        let order = if args.len() >= 3 {
            match evaluate(&args[2], cells)? {
                Variant::Integer(value) if value == 1 || value == -1 => value,
                Variant::Float(value) if value == 1.0 || value == -1.0 => value as i64,
                _ => return Ok(Variant::Error(ExcelError::Value)),
            }
        } else {
            1
        };
        let by_col = if args.len() >= 4 {
            is_truthy(&evaluate(&args[3], cells)?)
        } else {
            false
        };
        let limit = if by_col { rows } else { cols };
        if sort_index < 1 || sort_index > limit as i64 {
            return Ok(Variant::Error(ExcelError::Value));
        }
        let key = (sort_index - 1) as usize;
        let mut indices: Vec<usize> = (0..if by_col { cols } else { rows }).collect();
        indices.sort_by(|&left, &right| {
            let (left_index, right_index) = if by_col {
                (key * cols + left, key * cols + right)
            } else {
                (left * cols + key, right * cols + key)
            };
            let cmp = variant_cmp(&vals[left_index], &vals[right_index])
                .unwrap_or_else(|_| to_str(&vals[left_index]).cmp(&to_str(&vals[right_index])));
            if order < 0 { cmp.reverse() } else { cmp }
        });
        let mut result = Vec::with_capacity(vals.len());
        if by_col {
            for row in 0..rows {
                for &col in &indices {
                    result.push(vals[row * cols + col].clone());
                }
            }
        } else {
            for &row in &indices {
                result.extend(vals[row * cols..(row + 1) * cols].iter().cloned());
            }
        }
        return Ok(wrap_array(result));
    }
    let order = if args.len() >= 3 {
        match evaluate(&args[2], cells)? {
            Variant::Integer(value) if value == 1 || value == -1 => value,
            Variant::Float(value) if value == 1.0 || value == -1.0 => value as i64,
            _ => return Ok(Variant::Error(ExcelError::Value)),
        }
    } else {
        1
    };
    vals.sort_by(|a, b| {
        let af = to_float(a).unwrap_or(f64::INFINITY);
        let bf = to_float(b).unwrap_or(f64::INFINITY);
        let o = af.partial_cmp(&bf).unwrap_or(Ordering::Equal);
        if order < 0 { o.reverse() } else { o }
    });
    Ok(wrap_array(vals))
}

// ── SORTBY ────────────────────────────────────────────────────────────────────

fn func_sortby(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("SORTBY requires at least 2 arguments".into());
    }
    let data = flatten_array_vals(collect_values(&args[0], cells)?);
    let (rows, cols) = array_shape_for_expr(&args[0], cells, data.len());
    if rows > 1 && cols > 1 {
        let mut sort_keys = Vec::new();
        let mut key_arg = 1;
        while key_arg < args.len() {
            let key_values = flatten_array_vals(collect_values(&args[key_arg], cells)?);
            let (key_rows, key_cols) =
                array_shape_for_expr(&args[key_arg], cells, key_values.len());
            if key_rows != rows || key_cols != 1 {
                return Err(
                    "SORTBY: 2D data requires one-column sort-by arrays with equal rows".into(),
                );
            }
            let key_order = match args.get(key_arg + 1) {
                Some(arg) => match evaluate(arg, cells)? {
                    Variant::Integer(value) if value == 1 || value == -1 => value,
                    Variant::Float(value) if value == 1.0 || value == -1.0 => value as i64,
                    _ => return Ok(Variant::Error(ExcelError::Value)),
                },
                None => 1,
            };
            sort_keys.push((key_values, key_order));
            key_arg += 2;
        }
        let mut indices: Vec<usize> = (0..rows).collect();
        indices.sort_by(|&left, &right| {
            for (key_values, key_order) in &sort_keys {
                let cmp = variant_cmp(&key_values[left], &key_values[right])
                    .unwrap_or_else(|_| to_str(&key_values[left]).cmp(&to_str(&key_values[right])));
                let cmp = if *key_order < 0 { cmp.reverse() } else { cmp };
                if cmp != Ordering::Equal {
                    return cmp;
                }
            }
            Ordering::Equal
        });
        let mut result = Vec::with_capacity(data.len());
        for row in indices {
            result.extend(data[row * cols..(row + 1) * cols].iter().cloned());
        }
        return Ok(wrap_array(result));
    }
    let mut sort_keys = Vec::new();
    let mut key_arg = 1;
    while key_arg < args.len() {
        let key_values = flatten_array_vals(collect_values(&args[key_arg], cells)?);
        if data.len() != key_values.len() {
            return Err("SORTBY: data and sort-by arrays must have equal length".into());
        }
        let key_order = match args.get(key_arg + 1) {
            Some(arg) => match evaluate(arg, cells)? {
                Variant::Integer(value) if value == 1 || value == -1 => value,
                Variant::Float(value) if value == 1.0 || value == -1.0 => value as i64,
                _ => return Ok(Variant::Error(ExcelError::Value)),
            },
            None => 1,
        };
        sort_keys.push((key_values, key_order));
        key_arg += 2;
    }
    let mut indexed: Vec<usize> = (0..data.len()).collect();
    indexed.sort_by(|&a, &b| {
        for (key_values, key_order) in &sort_keys {
            let ordering = variant_cmp(&key_values[a], &key_values[b])
                .unwrap_or_else(|_| to_str(&key_values[a]).cmp(&to_str(&key_values[b])));
            let ordering = if *key_order < 0 {
                ordering.reverse()
            } else {
                ordering
            };
            if ordering != Ordering::Equal {
                return ordering;
            }
        }
        Ordering::Equal
    });
    let result: Vec<Variant> = indexed.iter().map(|&i| data[i].clone()).collect();
    Ok(wrap_array(result))
}

// ── SEQUENCE ──────────────────────────────────────────────────────────────────

fn func_sequence(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("SEQUENCE requires at least 1 argument".into());
    }
    let dimension = |arg: &FormulaExpr| -> Result<usize, String> {
        match evaluate(arg, cells)? {
            Variant::Integer(value) if value > 0 => {
                usize::try_from(value).map_err(|_| "SEQUENCE dimension is too large".to_string())
            }
            Variant::Float(value) if value.is_finite() && value.fract() == 0.0 && value > 0.0 => {
                if value > usize::MAX as f64 {
                    Err("SEQUENCE dimension is too large".into())
                } else {
                    Ok(value as usize)
                }
            }
            _ => Err("SEQUENCE dimensions must be positive integers".into()),
        }
    };
    let rows = dimension(&args[0])?;
    let cols = if args.len() >= 2 {
        dimension(&args[1])?
    } else {
        1usize
    };
    let start = if args.len() >= 3 {
        to_float(&evaluate(&args[2], cells)?)?
    } else {
        1.0
    };
    let step = if args.len() >= 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        1.0
    };
    let count = rows
        .checked_mul(cols)
        .ok_or_else(|| "SEQUENCE dimensions are too large".to_string())?;
    let result: Vec<Variant> = (0..count)
        .map(|i| as_integer_if_whole(start + i as f64 * step))
        .collect();
    Ok(wrap_array(result))
}

// ── TRANSPOSE ─────────────────────────────────────────────────────────────────

fn func_transpose(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("TRANSPOSE requires 1 argument".into());
    }
    match &args[0] {
        FormulaExpr::Range { c1, r1, c2, r2, .. } => {
            let rows = (*r2 - *r1 + 1) as usize;
            let cols = (*c2 - *c1 + 1) as usize;
            let mut result = vec![Variant::Empty; rows * cols];
            for (ri, row) in (*r1..=*r2).enumerate() {
                for (ci, col) in (*c1..=*c2).enumerate() {
                    result[ci * rows + ri] = cell_val(cells, row, col);
                }
            }
            Ok(wrap_array(result))
        }
        _ => evaluate(&args[0], cells),
    }
}

fn matrix_arg(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<(Vec<Variant>, usize, usize), String> {
    let values = flatten_array_vals(collect_values(expr, cells)?);
    let (rows, cols) = array_shape_for_expr(expr, cells, values.len());
    if rows == 0 || cols == 0 || rows.checked_mul(cols) != Some(values.len()) {
        return Err(format!("{name}: array must be rectangular"));
    }
    Ok((values, rows, cols))
}

fn func_munit(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("MUNIT requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    if !n.is_finite() || n.fract() != 0.0 || n <= 0.0 || n > 4096.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let n = n as usize;
    let mut result = vec![Variant::Integer(0); n * n];
    for i in 0..n {
        result[i * n + i] = Variant::Integer(1);
    }
    Ok(wrap_array(result))
}

fn func_mmult(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("MMULT requires 2 arguments".into());
    }
    let (left, rows, inner) = matrix_arg(&args[0], cells, "MMULT")?;
    let (right, inner_right, cols) = matrix_arg(&args[1], cells, "MMULT")?;
    if inner != inner_right {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let mut result = Vec::with_capacity(rows * cols);
    for r in 0..rows {
        for c in 0..cols {
            let mut value = 0.0;
            for k in 0..inner {
                value += to_float(&left[r * inner + k])? * to_float(&right[k * cols + c])?;
            }
            result.push(as_integer_if_whole(value));
        }
    }
    Ok(wrap_array(result))
}

fn matrix_numeric(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<(Vec<f64>, usize), String> {
    let (values, rows, cols) = matrix_arg(expr, cells, name)?;
    if rows != cols {
        return Err(format!("{name}: matrix must be square"));
    }
    Ok((
        values.iter().map(to_float).collect::<Result<Vec<_>, _>>()?,
        rows,
    ))
}

fn func_mdeterm(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("MDETERM requires 1 argument".into());
    }
    let (mut matrix, n) = matrix_numeric(&args[0], cells, "MDETERM")?;
    let mut determinant = 1.0;
    for pivot in 0..n {
        let Some(row) = (pivot..n).max_by(|a, b| {
            matrix[*a * n + pivot]
                .abs()
                .total_cmp(&matrix[*b * n + pivot].abs())
        }) else {
            unreachable!()
        };
        if matrix[row * n + pivot].abs() < 1e-14 {
            return Ok(Variant::Integer(0));
        }
        if row != pivot {
            for col in 0..n {
                matrix.swap(pivot * n + col, row * n + col);
            }
            determinant = -determinant;
        }
        let diagonal = matrix[pivot * n + pivot];
        determinant *= diagonal;
        for row in (pivot + 1)..n {
            let factor = matrix[row * n + pivot] / diagonal;
            for col in (pivot + 1)..n {
                matrix[row * n + col] -= factor * matrix[pivot * n + col];
            }
        }
    }
    Ok(as_integer_if_whole(determinant))
}

fn func_minverse(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("MINVERSE requires 1 argument".into());
    }
    let (matrix, n) = matrix_numeric(&args[0], cells, "MINVERSE")?;
    let mut augmented = vec![0.0; n * n * 2];
    for r in 0..n {
        for c in 0..n {
            augmented[r * 2 * n + c] = matrix[r * n + c];
            augmented[r * 2 * n + n + c] = if r == c { 1.0 } else { 0.0 };
        }
    }
    for pivot in 0..n {
        let Some(row) = (pivot..n).max_by(|a, b| {
            augmented[*a * 2 * n + pivot]
                .abs()
                .total_cmp(&augmented[*b * 2 * n + pivot].abs())
        }) else {
            unreachable!()
        };
        if augmented[row * 2 * n + pivot].abs() < 1e-14 {
            return Ok(Variant::Error(ExcelError::Num));
        }
        if row != pivot {
            for c in 0..2 * n {
                augmented.swap(pivot * 2 * n + c, row * 2 * n + c);
            }
        }
        let diagonal = augmented[pivot * 2 * n + pivot];
        for c in 0..2 * n {
            augmented[pivot * 2 * n + c] /= diagonal;
        }
        for r in 0..n {
            if r != pivot {
                let factor = augmented[r * 2 * n + pivot];
                for c in 0..2 * n {
                    augmented[r * 2 * n + c] -= factor * augmented[pivot * 2 * n + c];
                }
            }
        }
    }
    let mut result = Vec::with_capacity(n * n);
    for r in 0..n {
        for c in 0..n {
            result.push(as_integer_if_whole(augmented[r * 2 * n + n + c]));
        }
    }
    Ok(wrap_array(result))
}

// ── RANDARRAY ─────────────────────────────────────────────────────────────────
//
// Thread-local xorshift64 PRNG — no external crate needed.

use std::cell::Cell;
thread_local! {
    static RAND_STATE: Cell<u64> = const { Cell::new(0) };
}

fn next_rand_f64() -> f64 {
    RAND_STATE.with(|state| {
        let mut s = state.get();
        if s == 0 {
            // Seed once from the system clock (nanosecond resolution).
            s = std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .map(|d| {
                    let ns = d.subsec_nanos() as u64;
                    let secs = d.as_secs().wrapping_mul(6_364_136_223_846_793_005);
                    ns ^ secs ^ 0x9e37_79b9_7f4a_7c15
                })
                .unwrap_or(0xdead_beef_cafe_1234);
            if s == 0 {
                s = 1;
            }
        }
        // xorshift64
        s ^= s << 13;
        s ^= s >> 7;
        s ^= s << 17;
        state.set(s);
        // Map to [0, 1)
        (s >> 11) as f64 / (1u64 << 53) as f64
    })
}

fn func_randarray(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let dimension = |arg: &FormulaExpr| -> Result<usize, String> {
        match evaluate(arg, cells)? {
            Variant::Integer(value) if value > 0 => {
                usize::try_from(value).map_err(|_| "RANDARRAY dimension is too large".to_string())
            }
            Variant::Float(value) if value.is_finite() && value.fract() == 0.0 && value > 0.0 => {
                if value > usize::MAX as f64 {
                    Err("RANDARRAY dimension is too large".into())
                } else {
                    Ok(value as usize)
                }
            }
            _ => Err("RANDARRAY dimensions must be positive integers".into()),
        }
    };
    let rows = if !args.is_empty() {
        dimension(&args[0])?
    } else {
        1
    };
    let cols = if args.len() >= 2 {
        dimension(&args[1])?
    } else {
        1
    };
    let min = if args.len() >= 3 {
        to_float(&evaluate(&args[2], cells)?)?
    } else {
        0.0
    };
    let max = if args.len() >= 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        1.0
    };
    let whole = if args.len() >= 5 {
        is_truthy(&evaluate(&args[4], cells)?)
    } else {
        false
    };
    if max < min {
        return Err("RANDARRAY: max must be >= min".into());
    }
    let n = rows
        .checked_mul(cols)
        .ok_or_else(|| "RANDARRAY dimensions are too large".to_string())?;
    let result: Vec<Variant> = (0..n)
        .map(|_| {
            let v = min + next_rand_f64() * (max - min);
            if whole {
                as_integer_if_whole(v.floor())
            } else {
                Variant::Float(v)
            }
        })
        .collect();
    Ok(wrap_array(result))
}

// ── WORKDAY ──────────────────────────────────────────────────────────────────

fn func_workday(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("WORKDAY requires 2 or 3 arguments".into());
    }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let days = to_float(&evaluate(&args[1], cells)?)? as i64;
    let mask = [false, false, false, false, false, true, true]; // Sat+Sun
    let holidays: std::collections::HashSet<i64> = if args.len() == 3 {
        collect_values(&args[2], cells)?
            .iter()
            .filter_map(|v| to_float(v).ok().map(|f| f as i64))
            .collect()
    } else {
        std::collections::HashSet::new()
    };
    let mut current = start;
    let mut remaining = days.abs();
    let step = if days >= 0 { 1i64 } else { -1 };
    while remaining > 0 {
        current += step;
        if !is_weekend_intl(current, &mask) && !holidays.contains(&current) {
            remaining -= 1;
        }
    }
    Ok(Variant::Date(current))
}

// ── PMT ──────────────────────────────────────────────────────────────────────

fn func_pmt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 {
        return Err("PMT requires 3 to 5 arguments".into());
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let nper = to_float(&evaluate(&args[1], cells)?)?;
    let pv = to_float(&evaluate(&args[2], cells)?)?;
    let fv = if args.len() >= 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        0.0
    };
    let typ = if args.len() >= 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    if nper == 0.0 {
        return Err("PMT: nper cannot be 0".into());
    }
    let result = if rate == 0.0 {
        -(pv + fv) / nper
    } else {
        let factor = (1.0 + rate).powf(nper);
        -(rate * (pv * factor + fv)) / ((factor - 1.0) * (1.0 + rate * typ))
    };
    Ok(Variant::Float(result))
}

// ── Financial functions (FV / PV / NPER / RATE / IPMT / PPMT / NPV / IRR / MIRR / XNPV / XIRR) ──

/// Future value: FV(rate, nper, pmt, [pv=0], [type=0])
fn annuity_fv(rate: f64, nper: f64, pmt: f64, pv: f64, typ: f64) -> f64 {
    if rate == 0.0 {
        -(pv + pmt * nper)
    } else {
        let factor = (1.0 + rate).powf(nper);
        -(pv * factor + pmt * (1.0 + rate * typ) * (factor - 1.0) / rate)
    }
}

/// PMT helper (same logic as func_pmt, usable internally)
fn compute_pmt(rate: f64, nper: f64, pv: f64, fv: f64, typ: f64) -> f64 {
    if nper == 0.0 {
        return f64::NAN;
    }
    if rate == 0.0 {
        -(pv + fv) / nper
    } else {
        let factor = (1.0 + rate).powf(nper);
        -(rate * (pv * factor + fv)) / ((factor - 1.0) * (1.0 + rate * typ))
    }
}

fn func_fv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 {
        return Err("FV requires 3 to 5 arguments".into());
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let nper = to_float(&evaluate(&args[1], cells)?)?;
    let pmt = to_float(&evaluate(&args[2], cells)?)?;
    let pv = if args.len() >= 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        0.0
    };
    let typ = if args.len() >= 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    Ok(Variant::Float(annuity_fv(rate, nper, pmt, pv, typ)))
}

fn func_pv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 {
        return Err("PV requires 3 to 5 arguments".into());
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let nper = to_float(&evaluate(&args[1], cells)?)?;
    let pmt = to_float(&evaluate(&args[2], cells)?)?;
    let fv = if args.len() >= 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        0.0
    };
    let typ = if args.len() >= 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let result = if rate == 0.0 {
        -(fv + pmt * nper)
    } else {
        let factor = (1.0 + rate).powf(nper);
        -(fv / factor + pmt * (1.0 + rate * typ) * (1.0 - 1.0 / factor) / rate)
    };
    Ok(Variant::Float(result))
}

fn func_nper(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 {
        return Err("NPER requires 3 to 5 arguments".into());
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let pmt = to_float(&evaluate(&args[1], cells)?)?;
    let pv = to_float(&evaluate(&args[2], cells)?)?;
    let fv = if args.len() >= 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        0.0
    };
    let typ = if args.len() >= 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let result = if rate == 0.0 {
        if pmt == 0.0 {
            return Ok(Variant::Error(ExcelError::DivZero));
        }
        -(pv + fv) / pmt
    } else {
        let z = pmt * (1.0 + rate * typ) / rate;
        let numer = z - fv;
        let denom = z + pv;
        if denom == 0.0 || numer / denom <= 0.0 {
            return Ok(Variant::Error(ExcelError::Num));
        }
        (numer / denom).ln() / (1.0 + rate).ln()
    };
    Ok(Variant::Float(result))
}

fn func_rate(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 6 {
        return Err("RATE requires 3 to 6 arguments".into());
    }
    let nper = to_float(&evaluate(&args[0], cells)?)?;
    let pmt = to_float(&evaluate(&args[1], cells)?)?;
    let pv = to_float(&evaluate(&args[2], cells)?)?;
    let fv = if args.len() >= 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        0.0
    };
    let typ = if args.len() >= 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let guess = if args.len() >= 6 {
        to_float(&evaluate(&args[5], cells)?)?
    } else {
        0.1
    };
    // Newton-Raphson
    let mut r = guess;
    for _ in 0..100 {
        let a = (1.0 + r).powf(nper);
        let da = nper * (1.0 + r).powf(nper - 1.0);
        let f = pv * a + pmt * (1.0 + r * typ) * (a - 1.0) / r + fv;
        let df = pv * da
            + pmt * (typ * (a - 1.0) / r + (1.0 + r * typ) * (da * r - (a - 1.0)) / (r * r));
        let delta = f / df;
        r -= delta;
        if delta.abs() < 1e-10 {
            return Ok(Variant::Float(r));
        }
    }
    Ok(Variant::Error(ExcelError::Num)) // did not converge
}

fn func_ispmt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("ISPMT requires 4 arguments".into());
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let per = to_float(&evaluate(&args[1], cells)?)?;
    let nper = to_float(&evaluate(&args[2], cells)?)?;
    let pv = to_float(&evaluate(&args[3], cells)?)?;
    if !rate.is_finite()
        || !per.is_finite()
        || !nper.is_finite()
        || !pv.is_finite()
        || per.fract() != 0.0
        || nper.fract() != 0.0
        || nper <= 0.0
        || per < 1.0
        || per > nper
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(pv * rate * (per / nper - 1.0)))
}

fn func_ipmt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 6 {
        return Err("IPMT requires 4 to 6 arguments".into());
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let per_f = to_float(&evaluate(&args[1], cells)?)?;
    if per_f.fract() != 0.0 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let per = per_f as i64;
    let nper = to_float(&evaluate(&args[2], cells)?)?;
    let pv = to_float(&evaluate(&args[3], cells)?)?;
    let fv = if args.len() >= 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let typ = if args.len() >= 6 {
        to_float(&evaluate(&args[5], cells)?)?
    } else {
        0.0
    };
    if per < 1 || per as f64 > nper {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let pmt_val = compute_pmt(rate, nper, pv, fv, typ);
    // Remaining balance after (per-1) periods = annuity_fv(rate, per-1, pmt, pv, typ)
    let balance = annuity_fv(rate, (per - 1) as f64, pmt_val, pv, typ);
    let mut ipmt = balance * rate;
    if typ == 1.0 {
        if per == 1 {
            return Ok(Variant::Float(0.0));
        }
        ipmt /= 1.0 + rate;
    }
    Ok(Variant::Float(ipmt))
}

fn func_ppmt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 6 {
        return Err("PPMT requires 4 to 6 arguments".into());
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let per_f = to_float(&evaluate(&args[1], cells)?)?;
    if per_f.fract() != 0.0 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let per = per_f as i64;
    let nper = to_float(&evaluate(&args[2], cells)?)?;
    let pv = to_float(&evaluate(&args[3], cells)?)?;
    let fv = if args.len() >= 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let typ = if args.len() >= 6 {
        to_float(&evaluate(&args[5], cells)?)?
    } else {
        0.0
    };
    if per < 1 || per as f64 > nper {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let pmt_val = compute_pmt(rate, nper, pv, fv, typ);
    let balance = annuity_fv(rate, (per - 1) as f64, pmt_val, pv, typ);
    let mut ipmt = balance * rate;
    if typ == 1.0 {
        if per == 1 {
            return Ok(Variant::Float(pmt_val));
        }
        ipmt /= 1.0 + rate;
    }
    Ok(Variant::Float(pmt_val - ipmt))
}

fn cumulative_period_args(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<(f64, f64, f64, i64, i64, f64), String> {
    if args.len() != 6 {
        return Err(format!("{name} requires 6 arguments"));
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let nper = to_float(&evaluate(&args[1], cells)?)?;
    let pv = to_float(&evaluate(&args[2], cells)?)?;
    let start = to_float(&evaluate(&args[3], cells)?)?;
    let end = to_float(&evaluate(&args[4], cells)?)?;
    let typ = to_float(&evaluate(&args[5], cells)?)?;
    if !rate.is_finite()
        || !nper.is_finite()
        || !pv.is_finite()
        || !start.is_finite()
        || !end.is_finite()
        || !typ.is_finite()
        || nper < 1.0
        || nper.fract() != 0.0
        || start.fract() != 0.0
        || end.fract() != 0.0
        || start < 1.0
        || end < start
        || end > nper
        || (typ != 0.0 && typ != 1.0)
    {
        return Err(format!("{name}: invalid period or payment type"));
    }
    Ok((rate, nper, pv, start as i64, end as i64, typ))
}

fn cumulative_interest_or_principal(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    principal: bool,
) -> Result<Variant, String> {
    let name = if principal { "CUMPRINC" } else { "CUMIPMT" };
    let (rate, nper, pv, start, end, typ) = cumulative_period_args(args, cells, name)?;
    let mut total = 0.0;
    let pmt = compute_pmt(rate, nper, pv, 0.0, typ);
    for period in start..=end {
        let balance = annuity_fv(rate, (period - 1) as f64, pmt, pv, 0.0);
        let mut interest = balance * rate;
        if typ == 1.0 {
            if period == 1 {
                interest = 0.0;
            } else {
                interest /= 1.0 + rate;
            }
        }
        total += if principal { pmt - interest } else { interest };
    }
    if !total.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(total))
}

fn func_cumipmt(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    cumulative_interest_or_principal(args, cells, false)
}

fn func_cumprinc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    cumulative_interest_or_principal(args, cells, true)
}

fn func_fvschedule(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("FVSCHEDULE requires 2 arguments".into());
    }
    let principal = to_float(&evaluate(&args[0], cells)?)?;
    let rates = collect_nums(&args[1..], cells)?;
    if !principal.is_finite() || rates.iter().any(|rate| !rate.is_finite() || *rate <= -1.0) {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let result = rates
        .iter()
        .fold(principal, |value, rate| value * (1.0 + rate));
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(result))
}

fn func_dollarde(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("DOLLARDE requires 2 arguments".into());
    }
    let dollar = to_float(&evaluate(&args[0], cells)?)?;
    let fraction = to_float(&evaluate(&args[1], cells)?)?;
    if !dollar.is_finite() || !fraction.is_finite() || fraction < 1.0 || fraction.fract() != 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let sign = dollar.signum();
    let magnitude = dollar.abs();
    let whole = magnitude.trunc();
    let result = sign * (whole + (magnitude - whole) * 100.0 / fraction);
    Ok(Variant::Float(result))
}

fn func_dollarfr(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("DOLLARFR requires 2 arguments".into());
    }
    let dollar = to_float(&evaluate(&args[0], cells)?)?;
    let fraction = to_float(&evaluate(&args[1], cells)?)?;
    if !dollar.is_finite() || !fraction.is_finite() || fraction < 1.0 || fraction.fract() != 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let sign = dollar.signum();
    let magnitude = dollar.abs();
    let whole = magnitude.trunc();
    let result = sign * (whole + (magnitude - whole) * fraction / 100.0);
    Ok(Variant::Float(result))
}

fn euro_currency(code: &str) -> Option<(f64, u32)> {
    // Fixed EU conversion rates: units of the legacy currency per EUR and
    // the currency-specific calculation/display precision used by Excel.
    Some(match code.to_ascii_uppercase().as_str() {
        "BEF" => (40.3399, 0),
        "LUF" => (40.3399, 0),
        "DEM" => (1.95583, 2),
        "ESP" => (166.386, 0),
        "FRF" => (6.55957, 2),
        "IEP" => (0.787564, 2),
        "ITL" => (1936.27, 0),
        "NLG" => (2.20371, 2),
        "ATS" => (13.7603, 2),
        "PTE" => (200.482, 0),
        "FIM" => (5.94573, 2),
        "GRD" => (340.750, 0),
        "SIT" => (239.640, 2),
        "EUR" => (1.0, 2),
        _ => return None,
    })
}

fn round_significant(value: f64, digits: i32) -> f64 {
    if value == 0.0 {
        return value;
    }
    let scale = 10f64.powi(digits - 1 - value.abs().log10().floor() as i32);
    (value * scale).round() / scale
}

fn func_euroconvert(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 {
        return Err("EUROCONVERT requires 3 to 5 arguments".into());
    }
    let number = to_float(&evaluate(&args[0], cells)?)?;
    let source = to_str(&evaluate(&args[1], cells)?).to_ascii_uppercase();
    let target = to_str(&evaluate(&args[2], cells)?).to_ascii_uppercase();
    let full_precision = if args.len() >= 4 {
        to_float(&evaluate(&args[3], cells)?)? != 0.0
    } else {
        false
    };
    let triangulation_precision = if args.len() == 5 {
        let value = to_float(&evaluate(&args[4], cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || value < 3.0 {
            return Ok(Variant::Error(ExcelError::Value));
        }
        Some(value as i32)
    } else {
        None
    };
    let Some((source_rate, _)) = euro_currency(&source) else {
        return Ok(Variant::Error(ExcelError::Value));
    };
    let Some((target_rate, target_precision)) = euro_currency(&target) else {
        return Ok(Variant::Error(ExcelError::Value));
    };
    if !number.is_finite() {
        return Ok(Variant::Error(ExcelError::Value));
    }
    if source == target {
        return Ok(as_integer_if_whole(number));
    }
    let mut euros = number / source_rate;
    if source != "EUR"
        && target != "EUR"
        && let Some(precision) = triangulation_precision
    {
        euros = round_significant(euros, precision);
    }
    let mut result = euros * target_rate;
    if !full_precision {
        result = round_significant(result, target_precision as i32);
    }
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole(result))
}

fn bond_day_fraction(settlement: f64, maturity: f64, basis: f64) -> Result<f64, Variant> {
    if !settlement.is_finite()
        || !maturity.is_finite()
        || !basis.is_finite()
        || maturity <= settlement
        || basis.fract() != 0.0
        || !(0.0..=4.0).contains(&basis)
    {
        return Err(Variant::Error(ExcelError::Num));
    }
    let basis = basis as i32;
    let start = settlement.trunc() as i64;
    let end = maturity.trunc() as i64;
    let days = if basis == 0 || basis == 4 {
        days360_serial(start, end, basis == 4) as f64
    } else {
        (maturity - settlement).trunc()
    };
    let denominator = match basis {
        0 | 2 | 4 => 360.0,
        1 | 3 => 365.0,
        _ => unreachable!(),
    };
    if days <= 0.0 {
        return Err(Variant::Error(ExcelError::Num));
    }
    Ok(days / denominator)
}

fn coupon_month_shift(serial: i64, months: i32) -> i64 {
    let (year, month, day) = serial_to_ymd(serial);
    let total = year * 12 + month as i32 - 1 + months;
    let shifted_year = total.div_euclid(12);
    let shifted_month = (total.rem_euclid(12) + 1) as u32;
    let shifted_day = day.min(days_in_month(shifted_year, shifted_month));
    date_to_serial(shifted_year, shifted_month, shifted_day)
}

fn coupon_arguments(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<(i64, i64, i32, i32), String> {
    if args.len() < 3 || args.len() > 4 {
        return Err(format!("{name} requires 3 or 4 arguments"));
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let frequency = to_float(&evaluate(&args[2], cells)?)?;
    let basis = if args.len() == 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        0.0
    };
    if !settlement.is_finite()
        || !maturity.is_finite()
        || !frequency.is_finite()
        || !basis.is_finite()
        || settlement.fract() != 0.0
        || maturity.fract() != 0.0
        || frequency.fract() != 0.0
        || basis.fract() != 0.0
        || maturity <= settlement
        || !matches!(frequency as i32, 1 | 2 | 4)
        || !(0.0..=4.0).contains(&basis)
    {
        return Err(format!(
            "{name}: invalid settlement, maturity, frequency, or basis"
        ));
    }
    Ok((
        settlement as i64,
        maturity as i64,
        frequency as i32,
        basis as i32,
    ))
}

fn coupon_boundaries(settlement: i64, maturity: i64, frequency: i32) -> (i64, i64) {
    let months = 12 / frequency;
    let mut next = maturity;
    while next > settlement {
        let previous = coupon_month_shift(next, -months);
        if previous <= settlement {
            return (previous, next);
        }
        next = previous;
    }
    (next, coupon_month_shift(next, months))
}

fn coupon_day_count(start: i64, end: i64, basis: i32) -> f64 {
    if basis == 0 || basis == 4 {
        days360_serial(start, end, basis == 4) as f64
    } else {
        (end - start) as f64
    }
}

fn func_accrint(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 6 || args.len() > 8 {
        return Err("ACCRINT requires 6 to 8 arguments".into());
    }
    let issue = to_float(&evaluate(&args[0], cells)?)?;
    let first_interest = to_float(&evaluate(&args[1], cells)?)?;
    let settlement = to_float(&evaluate(&args[2], cells)?)?;
    let rate = to_float(&evaluate(&args[3], cells)?)?;
    let par = to_float(&evaluate(&args[4], cells)?)?;
    let frequency = to_float(&evaluate(&args[5], cells)?)?;
    let basis = if args.len() >= 7 {
        to_float(&evaluate(&args[6], cells)?)?
    } else {
        0.0
    };
    let calc_method = if args.len() == 8 {
        to_float(&evaluate(&args[7], cells)?)?
    } else {
        1.0
    };
    if !issue.is_finite()
        || !first_interest.is_finite()
        || !settlement.is_finite()
        || !rate.is_finite()
        || !par.is_finite()
        || !frequency.is_finite()
        || !basis.is_finite()
        || !calc_method.is_finite()
        || issue.fract() != 0.0
        || first_interest.fract() != 0.0
        || settlement.fract() != 0.0
        || first_interest <= issue
        || settlement <= issue
        || rate < 0.0
        || par <= 0.0
        || frequency.fract() != 0.0
        || !matches!(frequency as i32, 1 | 2 | 4)
        || basis.fract() != 0.0
        || !(0.0..=4.0).contains(&basis)
        || (calc_method != 0.0 && calc_method != 1.0)
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let issue = issue as i64;
    let first_interest = first_interest as i64;
    let settlement = settlement as i64;
    let frequency = frequency as i32;
    let basis = basis as i32;
    let months = 12 / frequency;
    let mut next_coupon = first_interest;
    while next_coupon <= settlement {
        next_coupon = coupon_month_shift(next_coupon, months);
    }
    let previous_coupon = coupon_month_shift(next_coupon, -months);
    let period_days = coupon_day_count(previous_coupon, next_coupon, basis);
    let accrual_start = issue.max(previous_coupon);
    let accrued_days = coupon_day_count(accrual_start, settlement, basis);
    if period_days <= 0.0 || accrued_days < 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(
        par * rate / frequency as f64 * accrued_days / period_days,
    ))
}

fn func_accrintm(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 {
        return Err("ACCRINTM requires 3 to 5 arguments".into());
    }
    let issue = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let rate = to_float(&evaluate(&args[2], cells)?)?;
    let par = if args.len() >= 4 {
        to_float(&evaluate(&args[3], cells)?)?
    } else {
        1000.0
    };
    let basis = if args.len() == 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let year_fraction = match bond_day_fraction(issue, maturity, basis) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    if !rate.is_finite() || !par.is_finite() || !basis.is_finite() || rate < 0.0 || par <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(par * rate * year_fraction))
}

fn func_coupdaybs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (settlement, maturity, frequency, basis) = coupon_arguments(args, cells, "COUPDAYBS")?;
    let (previous, _) = coupon_boundaries(settlement, maturity, frequency);
    Ok(Variant::Float(coupon_day_count(
        previous, settlement, basis,
    )))
}

fn func_coupdays(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (settlement, maturity, frequency, basis) = coupon_arguments(args, cells, "COUPDAYS")?;
    let (previous, next) = coupon_boundaries(settlement, maturity, frequency);
    Ok(Variant::Float(coupon_day_count(previous, next, basis)))
}

fn func_coupdaysnc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (settlement, maturity, frequency, basis) = coupon_arguments(args, cells, "COUPDAYSNC")?;
    let (_, next) = coupon_boundaries(settlement, maturity, frequency);
    Ok(Variant::Float(coupon_day_count(settlement, next, basis)))
}

fn func_coupncd(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (settlement, maturity, frequency, _) = coupon_arguments(args, cells, "COUPNCD")?;
    let (_, next) = coupon_boundaries(settlement, maturity, frequency);
    Ok(Variant::Date(next))
}

fn func_coupnum(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (settlement, maturity, frequency, _) = coupon_arguments(args, cells, "COUPNUM")?;
    let months = 12 / frequency;
    let (_, mut next) = coupon_boundaries(settlement, maturity, frequency);
    let mut count = 0_i64;
    while next <= maturity {
        count += 1;
        next = coupon_month_shift(next, months);
        if count > 10_000 {
            return Ok(Variant::Error(ExcelError::Num));
        }
    }
    Ok(Variant::Integer(count))
}

fn func_couppcd(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (settlement, maturity, frequency, _) = coupon_arguments(args, cells, "COUPPCD")?;
    let (previous, _) = coupon_boundaries(settlement, maturity, frequency);
    Ok(Variant::Date(previous))
}

fn maturity_security_arguments(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<(f64, f64, f64, f64, i32), String> {
    if args.len() < 4 || args.len() > 5 {
        return Err(format!("{name} requires 4 or 5 arguments"));
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let investment = to_float(&evaluate(&args[2], cells)?)?;
    let redemption = to_float(&evaluate(&args[3], cells)?)?;
    let basis = if args.len() == 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    if !settlement.is_finite()
        || !maturity.is_finite()
        || !investment.is_finite()
        || !redemption.is_finite()
        || !basis.is_finite()
        || settlement.fract() != 0.0
        || maturity.fract() != 0.0
        || maturity <= settlement
        || investment <= 0.0
        || redemption <= 0.0
        || basis.fract() != 0.0
        || !(0.0..=4.0).contains(&basis)
    {
        return Err(format!("{name}: invalid security arguments"));
    }
    Ok((settlement, maturity, investment, redemption, basis as i32))
}

fn func_intrate(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (settlement, maturity, investment, redemption, basis) =
        maturity_security_arguments(args, cells, "INTRATE")?;
    let year_fraction = match bond_day_fraction(settlement, maturity, basis as f64) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    Ok(Variant::Float(
        (redemption / investment - 1.0) / year_fraction,
    ))
}

fn func_pricemat(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 5 || args.len() > 6 {
        return Err("PRICEMAT requires 5 or 6 arguments".into());
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let issue = to_float(&evaluate(&args[2], cells)?)?;
    let rate = to_float(&evaluate(&args[3], cells)?)?;
    let yield_rate = to_float(&evaluate(&args[4], cells)?)?;
    let basis = if args.len() == 6 {
        to_float(&evaluate(&args[5], cells)?)?
    } else {
        0.0
    };
    let issue_fraction = match bond_day_fraction(issue, maturity, basis) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    let settlement_fraction = match bond_day_fraction(settlement, maturity, basis) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    let denominator = 1.0 + yield_rate * settlement_fraction;
    if !rate.is_finite()
        || !yield_rate.is_finite()
        || !basis.is_finite()
        || denominator <= 0.0
        || issue >= settlement
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(
        (100.0 + 100.0 * rate * issue_fraction) / denominator,
    ))
}

fn func_yieldmat(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 5 || args.len() > 6 {
        return Err("YIELDMAT requires 5 or 6 arguments".into());
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let issue = to_float(&evaluate(&args[2], cells)?)?;
    let rate = to_float(&evaluate(&args[3], cells)?)?;
    let price = to_float(&evaluate(&args[4], cells)?)?;
    let basis = if args.len() == 6 {
        to_float(&evaluate(&args[5], cells)?)?
    } else {
        0.0
    };
    let issue_fraction = match bond_day_fraction(issue, maturity, basis) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    let settlement_fraction = match bond_day_fraction(settlement, maturity, basis) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    if !rate.is_finite()
        || !price.is_finite()
        || !basis.is_finite()
        || price <= 0.0
        || issue >= settlement
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(
        ((100.0 + 100.0 * rate * issue_fraction) / price - 1.0) / settlement_fraction,
    ))
}

fn func_duration(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    modified: bool,
) -> Result<Variant, String> {
    if args.len() < 5 || args.len() > 6 {
        return Err(format!(
            "{} requires 5 or 6 arguments",
            if modified { "MDURATION" } else { "DURATION" }
        ));
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let coupon_rate = to_float(&evaluate(&args[2], cells)?)?;
    let yield_rate = to_float(&evaluate(&args[3], cells)?)?;
    let frequency = to_float(&evaluate(&args[4], cells)?)?;
    let basis = if args.len() == 6 {
        to_float(&evaluate(&args[5], cells)?)?
    } else {
        0.0
    };
    if !settlement.is_finite()
        || !maturity.is_finite()
        || !coupon_rate.is_finite()
        || !yield_rate.is_finite()
        || !frequency.is_finite()
        || !basis.is_finite()
        || settlement.fract() != 0.0
        || maturity.fract() != 0.0
        || maturity <= settlement
        || frequency.fract() != 0.0
        || !matches!(frequency as i32, 1 | 2 | 4)
        || basis.fract() != 0.0
        || !(0.0..=4.0).contains(&basis)
        || yield_rate <= -frequency
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let frequency = frequency as i32;
    let basis = basis as i32;
    let maturity_serial = maturity as i64;
    let (previous, next_coupon) = coupon_boundaries(settlement as i64, maturity_serial, frequency);
    let period_days = coupon_day_count(previous, next_coupon, basis);
    let first_days = coupon_day_count(settlement as i64, next_coupon, basis);
    if period_days <= 0.0 || first_days <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let periods_per_year = frequency as f64;
    let coupon = 100.0 * coupon_rate / periods_per_year;
    let rate_per_period = yield_rate / periods_per_year;
    let mut payment_date = next_coupon;
    let mut period_index = 0_i64;
    let mut price = 0.0;
    let mut weighted = 0.0;
    while payment_date <= maturity_serial {
        period_index += 1;
        let periods_from_settlement = (period_index - 1) as f64 + first_days / period_days;
        let discount = (1.0 + rate_per_period).powf(periods_from_settlement);
        let cashflow = coupon
            + if payment_date == maturity_serial {
                100.0
            } else {
                0.0
            };
        price += cashflow / discount;
        weighted += (periods_from_settlement / periods_per_year) * cashflow / discount;
        payment_date = coupon_month_shift(payment_date, 12 / frequency);
        if period_index > 10_000 {
            return Ok(Variant::Error(ExcelError::Num));
        }
    }
    if !price.is_finite() || price <= 0.0 || !weighted.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let duration = weighted / price;
    let result = if modified {
        duration / (1.0 + rate_per_period)
    } else {
        duration
    };
    Ok(Variant::Float(result))
}

fn bond_price_for_yield(
    settlement: i64,
    maturity: i64,
    coupon_rate: f64,
    yield_rate: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
) -> Result<f64, Variant> {
    if !coupon_rate.is_finite()
        || !yield_rate.is_finite()
        || !redemption.is_finite()
        || redemption <= 0.0
        || yield_rate <= -(frequency as f64)
    {
        return Err(Variant::Error(ExcelError::Num));
    }
    let (previous, next_coupon) = coupon_boundaries(settlement, maturity, frequency);
    let period_days = coupon_day_count(previous, next_coupon, basis);
    let first_days = coupon_day_count(settlement, next_coupon, basis);
    if period_days <= 0.0 || first_days <= 0.0 {
        return Err(Variant::Error(ExcelError::Num));
    }
    let frequency_f = frequency as f64;
    let coupon = redemption * coupon_rate / frequency_f;
    let rate_per_period = yield_rate / frequency_f;
    let mut payment_date = next_coupon;
    let mut period_index = 0_i64;
    let mut dirty_price = 0.0;
    while payment_date <= maturity {
        period_index += 1;
        let periods_from_settlement = (period_index - 1) as f64 + first_days / period_days;
        let discount = (1.0 + rate_per_period).powf(periods_from_settlement);
        let cashflow = coupon
            + if payment_date == maturity {
                redemption
            } else {
                0.0
            };
        dirty_price += cashflow / discount;
        payment_date = coupon_month_shift(payment_date, 12 / frequency);
        if period_index > 10_000 {
            return Err(Variant::Error(ExcelError::Num));
        }
    }
    let accrued = coupon * coupon_day_count(previous, settlement, basis) / period_days;
    let clean_price = dirty_price - accrued;
    if !clean_price.is_finite() || clean_price <= 0.0 {
        return Err(Variant::Error(ExcelError::Num));
    }
    Ok(clean_price)
}

fn bond_pricing_arguments(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
) -> Result<(i64, i64, f64, f64, i32, i32), String> {
    if args.len() < 6 || args.len() > 7 {
        return Err(format!("{name} requires 6 or 7 arguments"));
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let coupon_rate = to_float(&evaluate(&args[2], cells)?)?;
    let yield_rate = to_float(&evaluate(&args[3], cells)?)?;
    let _redemption = to_float(&evaluate(&args[4], cells)?)?;
    let frequency = to_float(&evaluate(&args[5], cells)?)?;
    let basis = if args.len() == 7 {
        to_float(&evaluate(&args[6], cells)?)?
    } else {
        0.0
    };
    if !settlement.is_finite()
        || !maturity.is_finite()
        || !frequency.is_finite()
        || !basis.is_finite()
        || settlement.fract() != 0.0
        || maturity.fract() != 0.0
        || maturity <= settlement
        || frequency.fract() != 0.0
        || !matches!(frequency as i32, 1 | 2 | 4)
        || basis.fract() != 0.0
        || !(0.0..=4.0).contains(&basis)
    {
        return Err(format!("{name}: invalid bond arguments"));
    }
    Ok((
        settlement as i64,
        maturity as i64,
        coupon_rate,
        yield_rate,
        frequency as i32,
        basis as i32,
    ))
}

fn func_price(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let (settlement, maturity, coupon_rate, yield_rate, frequency, basis) =
        bond_pricing_arguments(args, cells, "PRICE")?;
    let redemption = to_float(&evaluate(&args[4], cells)?)?;
    match bond_price_for_yield(
        settlement,
        maturity,
        coupon_rate,
        yield_rate,
        redemption,
        frequency,
        basis,
    ) {
        Ok(value) => Ok(Variant::Float(value)),
        Err(error) => Ok(error),
    }
}

fn func_yield(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 6 || args.len() > 7 {
        return Err("YIELD requires 6 or 7 arguments".into());
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let coupon_rate = to_float(&evaluate(&args[2], cells)?)?;
    let target_price = to_float(&evaluate(&args[3], cells)?)?;
    let redemption = to_float(&evaluate(&args[4], cells)?)?;
    let frequency = to_float(&evaluate(&args[5], cells)?)?;
    let basis = if args.len() == 7 {
        to_float(&evaluate(&args[6], cells)?)?
    } else {
        0.0
    };
    if !settlement.is_finite()
        || !maturity.is_finite()
        || !coupon_rate.is_finite()
        || !target_price.is_finite()
        || !redemption.is_finite()
        || !frequency.is_finite()
        || !basis.is_finite()
        || settlement.fract() != 0.0
        || maturity.fract() != 0.0
        || maturity <= settlement
        || target_price <= 0.0
        || redemption <= 0.0
        || frequency.fract() != 0.0
        || !matches!(frequency as i32, 1 | 2 | 4)
        || basis.fract() != 0.0
        || !(0.0..=4.0).contains(&basis)
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let frequency = frequency as i32;
    let basis = basis as i32;
    let settlement_serial = settlement as i64;
    let maturity_serial = maturity as i64;
    let low = -(frequency as f64) + 1e-10;
    let high = 10.0;
    let low_price = bond_price_for_yield(
        settlement_serial,
        maturity_serial,
        coupon_rate,
        low,
        redemption,
        frequency,
        basis,
    )
    .unwrap_or(0.0);
    let high_price = bond_price_for_yield(
        settlement_serial,
        maturity_serial,
        coupon_rate,
        high,
        redemption,
        frequency,
        basis,
    )
    .unwrap_or(0.0);
    if target_price > low_price || target_price < high_price {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut lo = low;
    let mut hi = high;
    for _ in 0..100 {
        let mid = (lo + hi) / 2.0;
        let price = bond_price_for_yield(
            settlement_serial,
            maturity_serial,
            coupon_rate,
            mid,
            redemption,
            frequency,
            basis,
        )
        .unwrap_or(0.0);
        if price > target_price {
            lo = mid;
        } else {
            hi = mid;
        }
    }
    Ok(Variant::Float((lo + hi) / 2.0))
}

struct OddFirstBondArgs {
    settlement: i64,
    maturity: i64,
    issue: i64,
    first_coupon: i64,
    rate: f64,
    price_or_yield: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
}

fn odd_first_price_for_yield(args: &OddFirstBondArgs, yield_rate: f64) -> Result<f64, Variant> {
    if args.issue >= args.settlement
        || args.settlement >= args.first_coupon
        || args.first_coupon > args.maturity
    {
        return Err(Variant::Error(ExcelError::Num));
    }
    if !args.rate.is_finite()
        || !yield_rate.is_finite()
        || !args.redemption.is_finite()
        || args.redemption <= 0.0
        || yield_rate <= -(args.frequency as f64)
    {
        return Err(Variant::Error(ExcelError::Num));
    }
    let months = 12 / args.frequency;
    let previous_coupon = coupon_month_shift(args.first_coupon, -months);
    let period_days = coupon_day_count(previous_coupon, args.first_coupon, args.basis);
    let odd_days = coupon_day_count(args.issue, args.first_coupon, args.basis);
    let accrued_days = coupon_day_count(args.issue, args.settlement, args.basis);
    if period_days <= 0.0 || odd_days <= 0.0 || accrued_days < 0.0 {
        return Err(Variant::Error(ExcelError::Num));
    }
    let coupon = args.redemption * args.rate / args.frequency as f64;
    let first_payment = coupon * odd_days / period_days;
    let rate_per_period = yield_rate / args.frequency as f64;
    let mut payment = args.first_coupon;
    let mut dirty = 0.0;
    let mut period_index = 0_u32;
    while payment <= args.maturity {
        let amount = if payment == args.first_coupon {
            first_payment
        } else {
            coupon
        } + if payment == args.maturity {
            args.redemption
        } else {
            0.0
        };
        let periods_from_settlement =
            coupon_day_count(args.settlement, payment, args.basis) / period_days;
        dirty += amount / (1.0 + rate_per_period).powf(periods_from_settlement);
        payment = coupon_month_shift(payment, months);
        period_index += 1;
        if period_index > 10_000 {
            return Err(Variant::Error(ExcelError::Num));
        }
    }
    let clean = dirty - coupon * accrued_days / period_days;
    if clean.is_finite() && clean > 0.0 {
        Ok(clean)
    } else {
        Err(Variant::Error(ExcelError::Num))
    }
}

fn odd_first_arguments(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
    price_form: bool,
) -> Result<OddFirstBondArgs, String> {
    if args.len() < 8 || args.len() > 9 {
        return Err(format!("{name} requires 8 or 9 arguments"));
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let issue = to_float(&evaluate(&args[2], cells)?)?;
    let first_coupon = to_float(&evaluate(&args[3], cells)?)?;
    let rate = to_float(&evaluate(&args[4], cells)?)?;
    let price_or_yield = to_float(&evaluate(&args[5], cells)?)?;
    let redemption = to_float(&evaluate(&args[6], cells)?)?;
    let frequency = to_float(&evaluate(&args[7], cells)?)?;
    let basis = if args.len() == 9 {
        to_float(&evaluate(&args[8], cells)?)?
    } else {
        0.0
    };
    if !settlement.is_finite()
        || !maturity.is_finite()
        || !issue.is_finite()
        || !first_coupon.is_finite()
        || !rate.is_finite()
        || !price_or_yield.is_finite()
        || !redemption.is_finite()
        || !frequency.is_finite()
        || !basis.is_finite()
        || settlement.fract() != 0.0
        || maturity.fract() != 0.0
        || issue.fract() != 0.0
        || first_coupon.fract() != 0.0
        || frequency.fract() != 0.0
        || basis.fract() != 0.0
        || (!price_form && price_or_yield <= 0.0)
        || (price_form && price_or_yield <= -(frequency.max(1.0)))
        || redemption <= 0.0
        || !matches!(frequency as i32, 1 | 2 | 4)
        || !(0.0..=4.0).contains(&basis)
    {
        return Err(format!("{name}: invalid odd-first bond arguments"));
    }
    Ok(OddFirstBondArgs {
        settlement: settlement as i64,
        maturity: maturity as i64,
        issue: issue as i64,
        first_coupon: first_coupon as i64,
        rate,
        price_or_yield,
        redemption,
        frequency: frequency as i32,
        basis: basis as i32,
    })
}

fn func_oddfprice(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let bond = odd_first_arguments(args, cells, "ODDFPRICE", true)?;
    match odd_first_price_for_yield(&bond, bond.price_or_yield) {
        Ok(value) => Ok(Variant::Float(value)),
        Err(error) => Ok(error),
    }
}

fn func_oddfyield(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let bond = odd_first_arguments(args, cells, "ODDFYIELD", false)?;
    let low = -(bond.frequency as f64) + 1e-10;
    let high = 10.0;
    let low_price = odd_first_price_for_yield(&bond, low).unwrap_or(0.0);
    let high_price = odd_first_price_for_yield(&bond, high).unwrap_or(0.0);
    if bond.price_or_yield > low_price || bond.price_or_yield < high_price {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut lo = low;
    let mut hi = high;
    for _ in 0..100 {
        let mid = (lo + hi) / 2.0;
        let price = odd_first_price_for_yield(&bond, mid).unwrap_or(0.0);
        if price > bond.price_or_yield {
            lo = mid;
        } else {
            hi = mid;
        }
    }
    Ok(Variant::Float((lo + hi) / 2.0))
}

struct OddLastBondArgs {
    settlement: i64,
    maturity: i64,
    last_interest: i64,
    rate: f64,
    price_or_yield: f64,
    redemption: f64,
    frequency: i32,
    basis: i32,
}

fn odd_last_arguments(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
    price_form: bool,
) -> Result<OddLastBondArgs, String> {
    if args.len() < 7 || args.len() > 8 {
        return Err(format!("{name} requires 7 or 8 arguments"));
    }
    let values = args[..7]
        .iter()
        .map(|arg| evaluate(arg, cells).and_then(|value| to_float(&value)))
        .collect::<Result<Vec<_>, _>>()?;
    let basis = if args.len() == 8 {
        to_float(&evaluate(&args[7], cells)?)?
    } else {
        0.0
    };
    let [
        settlement,
        maturity,
        last_interest,
        rate,
        price_or_yield,
        redemption,
        frequency,
    ] = values.as_slice()
    else {
        return Err(format!("{name}: invalid argument count"));
    };
    if !values.iter().all(|value| value.is_finite())
        || !basis.is_finite()
        || settlement.fract() != 0.0
        || maturity.fract() != 0.0
        || last_interest.fract() != 0.0
        || frequency.fract() != 0.0
        || (!price_form && *price_or_yield <= 0.0)
        || (price_form && *price_or_yield <= -frequency.max(1.0))
        || *redemption <= 0.0
        || !matches!(*frequency as i32, 1 | 2 | 4)
        || !(0.0..=4.0).contains(&basis)
    {
        return Err(format!("{name}: invalid odd-last bond arguments"));
    }
    Ok(OddLastBondArgs {
        settlement: *settlement as i64,
        maturity: *maturity as i64,
        last_interest: *last_interest as i64,
        rate: *rate,
        price_or_yield: *price_or_yield,
        redemption: *redemption,
        frequency: *frequency as i32,
        basis: basis as i32,
    })
}

fn odd_last_price_for_yield(args: &OddLastBondArgs, yield_rate: f64) -> Result<f64, Variant> {
    if args.settlement >= args.maturity
        || args.last_interest >= args.maturity
        || args.settlement > args.last_interest
        || !yield_rate.is_finite()
        || yield_rate <= -(args.frequency as f64)
    {
        return Err(Variant::Error(ExcelError::Num));
    }
    let months = 12 / args.frequency;
    let previous_coupon = coupon_month_shift(args.last_interest, -months);
    let period_days = coupon_day_count(previous_coupon, args.last_interest, args.basis);
    let odd_days = coupon_day_count(args.last_interest, args.maturity, args.basis);
    let accrued_days = coupon_day_count(previous_coupon, args.settlement, args.basis);
    if period_days <= 0.0 || odd_days <= 0.0 || accrued_days < 0.0 {
        return Err(Variant::Error(ExcelError::Num));
    }
    let coupon = args.redemption * args.rate / args.frequency as f64;
    let rate_per_period = yield_rate / args.frequency as f64;
    let mut payment = coupon_month_shift(args.last_interest, months);
    let mut dirty = 0.0;
    let mut period_index = 0_u32;
    while payment < args.maturity {
        let periods = coupon_day_count(args.settlement, payment, args.basis) / period_days;
        dirty += coupon / (1.0 + rate_per_period).powf(periods);
        payment = coupon_month_shift(payment, months);
        period_index += 1;
        if period_index > 10_000 {
            return Err(Variant::Error(ExcelError::Num));
        }
    }
    let final_coupon = coupon * odd_days / period_days;
    let final_periods = coupon_day_count(args.settlement, args.maturity, args.basis) / period_days;
    dirty += (final_coupon + args.redemption) / (1.0 + rate_per_period).powf(final_periods);
    let clean = dirty - coupon * accrued_days / period_days;
    if clean.is_finite() && clean > 0.0 {
        Ok(clean)
    } else {
        Err(Variant::Error(ExcelError::Num))
    }
}

fn func_oddlprice(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let bond = odd_last_arguments(args, cells, "ODDLPRICE", true)?;
    match odd_last_price_for_yield(&bond, bond.price_or_yield) {
        Ok(value) => Ok(Variant::Float(value)),
        Err(error) => Ok(error),
    }
}

fn func_oddlyield(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let bond = odd_last_arguments(args, cells, "ODDLYIELD", false)?;
    let low = -(bond.frequency as f64) + 1e-10;
    let high = 10.0;
    let low_price = odd_last_price_for_yield(&bond, low).unwrap_or(0.0);
    let high_price = odd_last_price_for_yield(&bond, high).unwrap_or(0.0);
    if bond.price_or_yield > low_price || bond.price_or_yield < high_price {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut lo = low;
    let mut hi = high;
    for _ in 0..100 {
        let mid = (lo + hi) / 2.0;
        if odd_last_price_for_yield(&bond, mid).unwrap_or(0.0) > bond.price_or_yield {
            lo = mid;
        } else {
            hi = mid;
        }
    }
    Ok(Variant::Float((lo + hi) / 2.0))
}

fn func_pduration(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("PDURATION requires 3 arguments".into());
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let present = to_float(&evaluate(&args[1], cells)?)?;
    let future = to_float(&evaluate(&args[2], cells)?)?;
    if !rate.is_finite()
        || !present.is_finite()
        || !future.is_finite()
        || rate <= -1.0
        || present <= 0.0
        || future <= 0.0
        || rate == 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let result = (future / present).ln() / (1.0 + rate).ln();
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(result))
}

fn func_pricedisc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 5 {
        return Err("PRICEDISC requires 4 or 5 arguments".into());
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let discount = to_float(&evaluate(&args[2], cells)?)?;
    let redemption = to_float(&evaluate(&args[3], cells)?)?;
    let basis = if args.len() == 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let year_fraction = match bond_day_fraction(settlement, maturity, basis) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    if !discount.is_finite()
        || !redemption.is_finite()
        || !(0.0..1.0).contains(&discount)
        || redemption <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(
        redemption * (1.0 - discount * year_fraction),
    ))
}

fn func_disc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 5 {
        return Err("DISC requires 4 or 5 arguments".into());
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let price = to_float(&evaluate(&args[2], cells)?)?;
    let redemption = to_float(&evaluate(&args[3], cells)?)?;
    let basis = if args.len() == 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let year_fraction = match bond_day_fraction(settlement, maturity, basis) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    if !price.is_finite() || !redemption.is_finite() || price <= 0.0 || redemption <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(
        (redemption - price) / redemption / year_fraction,
    ))
}

fn func_received(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 5 {
        return Err("RECEIVED requires 4 or 5 arguments".into());
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let investment = to_float(&evaluate(&args[2], cells)?)?;
    let discount = to_float(&evaluate(&args[3], cells)?)?;
    let basis = if args.len() == 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let year_fraction = match bond_day_fraction(settlement, maturity, basis) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    let denominator = 1.0 - discount * year_fraction;
    if !investment.is_finite()
        || !discount.is_finite()
        || investment <= 0.0
        || !(0.0..1.0).contains(&discount)
        || denominator <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(investment / denominator))
}

fn func_yielddisc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 5 {
        return Err("YIELDDISC requires 4 or 5 arguments".into());
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let price = to_float(&evaluate(&args[2], cells)?)?;
    let redemption = to_float(&evaluate(&args[3], cells)?)?;
    let basis = if args.len() == 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        0.0
    };
    let year_fraction = match bond_day_fraction(settlement, maturity, basis) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    if !price.is_finite() || !redemption.is_finite() || price <= 0.0 || redemption <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float((redemption / price - 1.0) / year_fraction))
}

fn func_tbillprice(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("TBILLPRICE requires 3 arguments".into());
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let discount = to_float(&evaluate(&args[2], cells)?)?;
    let year_fraction = match bond_day_fraction(settlement, maturity, 2.0) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    if !discount.is_finite() || !(0.0..1.0).contains(&discount) {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(100.0 * (1.0 - discount * year_fraction)))
}

fn func_tbillyield(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("TBILLYIELD requires 3 arguments".into());
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let price = to_float(&evaluate(&args[2], cells)?)?;
    let year_fraction = match bond_day_fraction(settlement, maturity, 2.0) {
        Ok(value) => value,
        Err(error) => return Ok(error),
    };
    if !price.is_finite() || price <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float((100.0 - price) / price / year_fraction))
}

fn func_tbilleq(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("TBILLEQ requires 3 arguments".into());
    }
    let settlement = to_float(&evaluate(&args[0], cells)?)?;
    let maturity = to_float(&evaluate(&args[1], cells)?)?;
    let discount = to_float(&evaluate(&args[2], cells)?)?;
    let days = match bond_day_fraction(settlement, maturity, 2.0) {
        Ok(value) => value * 360.0,
        Err(error) => return Ok(error),
    };
    let denominator = 360.0 - discount * days;
    if !discount.is_finite() || !(0.0..1.0).contains(&discount) || denominator <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(365.0 * discount / denominator))
}

fn func_npv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("NPV requires at least 2 arguments".into());
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let values = collect_all(&args[1..], cells)?;
    let nums: Vec<f64> = values.iter().filter_map(as_f64).collect();
    let result = nums
        .iter()
        .enumerate()
        .map(|(i, &v)| v / (1.0 + rate).powf((i + 1) as f64))
        .sum::<f64>();
    Ok(Variant::Float(result))
}

fn func_irr(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 {
        return Err("IRR requires 1 to 2 arguments".into());
    }
    let values = collect_values(&args[0], cells)?;
    let nums: Vec<f64> = values.iter().filter_map(as_f64).collect();
    if nums.is_empty() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let guess = if args.len() >= 2 {
        to_float(&evaluate(&args[1], cells)?)?
    } else {
        0.1
    };
    // Newton-Raphson: find r where NPV = sum(v[i]/(1+r)^(i+1)) = 0
    let mut r = guess;
    for _ in 0..100 {
        // IRR: find r where sum(v[i]/(1+r)^i) = 0, i=0,1,...
        let f: f64 = nums
            .iter()
            .enumerate()
            .map(|(i, &v)| v / (1.0 + r).powf(i as f64))
            .sum();
        let df: f64 = nums
            .iter()
            .enumerate()
            .map(|(i, &v)| -(i as f64) * v / (1.0 + r).powf((i + 1) as f64))
            .sum();
        if df == 0.0 {
            break;
        }
        let delta = f / df;
        r -= delta;
        if delta.abs() < 1e-10 {
            return Ok(Variant::Float(r));
        }
    }
    Ok(Variant::Error(ExcelError::Num))
}

fn func_mirr(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("MIRR requires 3 arguments".into());
    }
    let values = collect_values(&args[0], cells)?;
    let nums: Vec<f64> = values.iter().filter_map(as_f64).collect();
    let n = nums.len();
    if n < 2 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let finance_rate = to_float(&evaluate(&args[1], cells)?)?;
    let reinvest_rate = to_float(&evaluate(&args[2], cells)?)?;
    // PV of negative cash flows at finance_rate
    let pv_neg: f64 = nums
        .iter()
        .enumerate()
        .map(|(i, &v)| {
            if v < 0.0 {
                v / (1.0 + finance_rate).powf(i as f64)
            } else {
                0.0
            }
        })
        .sum();
    // FV of positive cash flows at reinvest_rate
    let fv_pos: f64 = nums
        .iter()
        .enumerate()
        .map(|(i, &v)| {
            if v > 0.0 {
                v * (1.0 + reinvest_rate).powf((n - 1 - i) as f64)
            } else {
                0.0
            }
        })
        .sum();
    if pv_neg == 0.0 || fv_pos == 0.0 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let result = (fv_pos / (-pv_neg)).powf(1.0 / (n - 1) as f64) - 1.0;
    Ok(Variant::Float(result))
}

fn func_xnpv(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("XNPV requires 3 arguments".into());
    }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let values = collect_values(&args[1], cells)?;
    let dates = collect_values(&args[2], cells)?;
    if values.len() != dates.len() || values.is_empty() {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let d0 = match as_f64(&dates[0]) {
        Some(d) => d,
        None => return Ok(Variant::Error(ExcelError::Value)),
    };
    let mut result = 0.0;
    for (v, d) in values.iter().zip(dates.iter()) {
        let val = match as_f64(v) {
            Some(x) => x,
            None => return Ok(Variant::Error(ExcelError::Value)),
        };
        let date = match as_f64(d) {
            Some(x) => x,
            None => return Ok(Variant::Error(ExcelError::Value)),
        };
        let exp = (date - d0) / 365.0;
        result += val / (1.0 + rate).powf(exp);
    }
    Ok(Variant::Float(result))
}

fn func_xirr(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("XIRR requires 2 to 3 arguments".into());
    }
    let values = collect_values(&args[0], cells)?;
    let dates = collect_values(&args[1], cells)?;
    if values.len() != dates.len() || values.is_empty() {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let guess = if args.len() >= 3 {
        to_float(&evaluate(&args[2], cells)?)?
    } else {
        0.1
    };
    let d0 = match as_f64(&dates[0]) {
        Some(d) => d,
        None => return Ok(Variant::Error(ExcelError::Value)),
    };
    // Collect numeric pairs only (skip rows where either value is non-numeric, matching Excel)
    let pairs: Vec<(f64, f64)> = values
        .iter()
        .zip(dates.iter())
        .filter_map(|(v, d)| Some((as_f64(v)?, as_f64(d)?)))
        .collect();
    if pairs.is_empty() {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let nums: Vec<f64> = pairs.iter().map(|(v, _)| *v).collect();
    let days: Vec<f64> = pairs.iter().map(|(_, d)| (d - d0) / 365.0).collect();
    let mut r = guess;
    for _ in 0..100 {
        let f: f64 = nums
            .iter()
            .zip(days.iter())
            .map(|(&v, &t)| v / (1.0 + r).powf(t))
            .sum();
        let df: f64 = nums
            .iter()
            .zip(days.iter())
            .map(|(&v, &t)| -t * v / (1.0 + r).powf(t + 1.0))
            .sum();
        if df == 0.0 {
            break;
        }
        let delta = f / df;
        r -= delta;
        if delta.abs() < 1e-10 {
            return Ok(Variant::Float(r));
        }
    }
    Ok(Variant::Error(ExcelError::Num))
}

fn func_sln(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("SLN requires 3 arguments".into());
    }
    let cost = to_float(&evaluate(&args[0], cells)?)?;
    let salvage = to_float(&evaluate(&args[1], cells)?)?;
    let life = to_float(&evaluate(&args[2], cells)?)?;
    if life <= 0.0 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    Ok(Variant::Float((cost - salvage) / life))
}

fn func_syd(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("SYD requires 4 arguments".into());
    }
    let cost = to_float(&evaluate(&args[0], cells)?)?;
    let salvage = to_float(&evaluate(&args[1], cells)?)?;
    let life = to_float(&evaluate(&args[2], cells)?)?;
    let period = to_float(&evaluate(&args[3], cells)?)?;
    if life <= 0.0 || period <= 0.0 || period > life {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(
        (cost - salvage) * (life - period + 1.0) * 2.0 / (life * (life + 1.0)),
    ))
}

fn func_ddb(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 5 {
        return Err("DDB requires 4 or 5 arguments".into());
    }
    let cost = to_float(&evaluate(&args[0], cells)?)?;
    let salvage = to_float(&evaluate(&args[1], cells)?)?;
    let life = to_float(&evaluate(&args[2], cells)?)?;
    let period = to_float(&evaluate(&args[3], cells)?)?;
    let factor = if args.len() == 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        2.0
    };
    if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || period > life || factor <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut book = cost;
    let mut depreciation = 0.0;
    for _ in 1..=(period.ceil() as i64) {
        depreciation = (book * factor / life).min((book - salvage).max(0.0));
        book -= depreciation;
    }
    Ok(Variant::Float(depreciation))
}

fn func_vdb(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 5 || args.len() > 7 {
        return Err("VDB requires 5 to 7 arguments".into());
    }
    let cost = to_float(&evaluate(&args[0], cells)?)?;
    let salvage = to_float(&evaluate(&args[1], cells)?)?;
    let life = to_float(&evaluate(&args[2], cells)?)?;
    let start_period = to_float(&evaluate(&args[3], cells)?)?;
    let end_period = to_float(&evaluate(&args[4], cells)?)?;
    let factor = if args.len() >= 6 {
        to_float(&evaluate(&args[5], cells)?)?
    } else {
        2.0
    };
    let no_switch = if args.len() == 7 {
        let value = to_float(&evaluate(&args[6], cells)?)?;
        if !value.is_finite() || (value != 0.0 && value != 1.0) {
            return Ok(Variant::Error(ExcelError::Num));
        }
        value == 1.0
    } else {
        false
    };
    if !cost.is_finite()
        || !salvage.is_finite()
        || !life.is_finite()
        || !start_period.is_finite()
        || !end_period.is_finite()
        || !factor.is_finite()
        || cost < 0.0
        || salvage < 0.0
        || salvage > cost
        || life <= 0.0
        || start_period < 0.0
        || end_period <= start_period
        || end_period > life
        || factor <= 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut book = cost;
    let mut depreciation = 0.0;
    let last_segment = end_period.ceil() as usize;
    for period in 0..last_segment {
        let period_start = period as f64;
        let period_end = period_start + 1.0;
        let overlap = (end_period.min(period_end) - start_period.max(period_start)).max(0.0);
        if overlap == 0.0 {
            let full_declining = (book * factor / life).min((book - salvage).max(0.0));
            book -= full_declining;
            continue;
        }
        let declining = book * factor / life;
        let straight = if life > period_start {
            (book - salvage).max(0.0) / (life - period_start)
        } else {
            0.0
        };
        let full_depreciation = if no_switch || declining >= straight {
            declining
        } else {
            straight
        }
        .min((book - salvage).max(0.0));
        depreciation += full_depreciation * overlap;
        book -= full_depreciation;
    }
    if depreciation.is_finite() {
        Ok(Variant::Float(depreciation))
    } else {
        Ok(Variant::Error(ExcelError::Num))
    }
}

fn func_db(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 5 {
        return Err("DB requires 4 or 5 arguments".into());
    }
    let cost = to_float(&evaluate(&args[0], cells)?)?;
    let salvage = to_float(&evaluate(&args[1], cells)?)?;
    let life = to_float(&evaluate(&args[2], cells)?)?;
    let period = to_float(&evaluate(&args[3], cells)?)?;
    let month = if args.len() == 5 {
        to_float(&evaluate(&args[4], cells)?)?
    } else {
        12.0
    };
    if cost < 0.0 || salvage < 0.0 || life <= 0.0 || period <= 0.0 || month <= 0.0 || month > 12.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let factor = ((1.0 - (salvage / cost).powf(1.0 / life)) * 1000.0).round() / 1000.0;
    let mut book = cost;
    let mut depreciation = 0.0;
    let years = period.ceil() as i64;
    for year in 1..=years {
        let months = if year == 1 { month } else { 12.0 };
        depreciation = (book * factor * months / 12.0).min((book - salvage).max(0.0));
        book -= depreciation;
    }
    Ok(Variant::Float(depreciation))
}

fn func_amorlinc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 7 {
        return Err("AMORLINC requires 7 arguments".into());
    }
    let cost = to_float(&evaluate(&args[0], cells)?)?;
    let purchased = to_float(&evaluate(&args[1], cells)?)?;
    let first_period = to_float(&evaluate(&args[2], cells)?)?;
    let salvage = to_float(&evaluate(&args[3], cells)?)?;
    let period = to_float(&evaluate(&args[4], cells)?)?;
    let rate = to_float(&evaluate(&args[5], cells)?)?;
    let basis = to_float(&evaluate(&args[6], cells)?)?;
    if !cost.is_finite()
        || !purchased.is_finite()
        || !first_period.is_finite()
        || !salvage.is_finite()
        || !period.is_finite()
        || !rate.is_finite()
        || !basis.is_finite()
        || purchased.fract() != 0.0
        || first_period.fract() != 0.0
        || purchased >= first_period
        || cost < 0.0
        || salvage < 0.0
        || salvage > cost
        || period < 0.0
        || period.fract() != 0.0
        || rate < 0.0
        || basis.fract() != 0.0
        || !(0.0..=4.0).contains(&basis)
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let purchased = purchased as i64;
    let first_period = first_period as i64;
    let period = period as u64;
    let basis = basis as i32;
    let year_days = match basis {
        0 | 2 | 4 => 360.0,
        1 | 3 => 365.0,
        _ => unreachable!(),
    };
    let first_days = coupon_day_count(purchased, first_period, basis);
    let first_depreciation = cost * rate * first_days / year_days;
    let regular_depreciation = cost * rate;
    let prior = if period == 0 {
        0.0
    } else {
        first_depreciation + regular_depreciation * (period - 1) as f64
    };
    let depreciation = if period == 0 {
        first_depreciation
    } else {
        regular_depreciation
    };
    Ok(Variant::Float(
        depreciation.min((cost - salvage - prior).max(0.0)),
    ))
}

fn func_amordegrc(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 7 {
        return Err("AMORDEGRC requires 7 arguments".into());
    }
    let cost = to_float(&evaluate(&args[0], cells)?)?;
    let purchased = to_float(&evaluate(&args[1], cells)?)?;
    let first_period = to_float(&evaluate(&args[2], cells)?)?;
    let salvage = to_float(&evaluate(&args[3], cells)?)?;
    let period = to_float(&evaluate(&args[4], cells)?)?;
    let rate = to_float(&evaluate(&args[5], cells)?)?;
    let basis = to_float(&evaluate(&args[6], cells)?)?;
    if !cost.is_finite()
        || !purchased.is_finite()
        || !first_period.is_finite()
        || !salvage.is_finite()
        || !period.is_finite()
        || !rate.is_finite()
        || !basis.is_finite()
        || purchased.fract() != 0.0
        || first_period.fract() != 0.0
        || purchased >= first_period
        || cost < 0.0
        || salvage < 0.0
        || salvage > cost
        || period < 0.0
        || period.fract() != 0.0
        || rate <= 0.0
        || basis.fract() != 0.0
        || !(0.0..=4.0).contains(&basis)
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let purchased = purchased as i64;
    let first_period = first_period as i64;
    let period = period as u64;
    let basis = basis as i32;
    let year_days = match basis {
        0 | 2 | 4 => 360.0,
        1 | 3 => 365.0,
        _ => unreachable!(),
    };
    let first_days = coupon_day_count(purchased, first_period, basis);
    let first_depreciation = cost * rate * first_days / year_days;
    let life = 1.0 / rate;
    let factor = if life < 3.0 {
        1.0
    } else if life < 5.0 {
        1.5
    } else {
        2.0
    };
    let mut book = cost;
    let mut depreciation = 0.0;
    for current in 0..=period {
        let remaining = (book - salvage).max(0.0);
        depreciation = if current == 0 {
            first_depreciation.min(remaining)
        } else {
            (book * rate * factor).min(remaining)
        };
        book -= depreciation;
    }
    Ok(Variant::Float(depreciation))
}

fn func_effect(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("EFFECT requires 2 arguments".into());
    }
    let nominal = to_float(&evaluate(&args[0], cells)?)?;
    let periods = to_float(&evaluate(&args[1], cells)?)?;
    if periods < 1.0 || periods.fract() != 0.0 || !nominal.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(
        (1.0 + nominal / periods).powf(periods) - 1.0,
    ))
}

fn func_nominal(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("NOMINAL requires 2 arguments".into());
    }
    let effective = to_float(&evaluate(&args[0], cells)?)?;
    let periods = to_float(&evaluate(&args[1], cells)?)?;
    if periods < 1.0 || periods.fract() != 0.0 || effective <= -1.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float(
        periods * ((1.0 + effective).powf(1.0 / periods) - 1.0),
    ))
}

fn func_rri(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("RRI requires 3 arguments".into());
    }
    let nper = to_float(&evaluate(&args[0], cells)?)?;
    let pv = to_float(&evaluate(&args[1], cells)?)?;
    let fv = to_float(&evaluate(&args[2], cells)?)?;
    if nper <= 0.0 || pv <= 0.0 || fv <= 0.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Float((fv / pv).powf(1.0 / nper) - 1.0))
}

// ── TEXTSPLIT ─────────────────────────────────────────────────────────────────

fn func_textsplit(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(2..=6).contains(&args.len()) {
        return Err("TEXTSPLIT requires 2 to 6 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let column_delims = textsplit_delimiters(&args[1], cells, true)?;
    // An empty third argument is commonly used as an explicit placeholder
    // for the optional row delimiter; treat it like an omitted delimiter.
    let row_delims = if let Some(arg) = args.get(2) {
        let delimiters = textsplit_delimiters(arg, cells, true)?;
        (!delimiters.is_empty()).then_some(delimiters)
    } else {
        None
    };
    if column_delims.is_empty() && row_delims.is_none() {
        return Err("TEXTSPLIT: at least one delimiter is required".into());
    }
    let ignore_empty = args.len() >= 4 && is_truthy(&evaluate(&args[3], cells)?);
    let case_insensitive = if args.len() >= 5 {
        let mode = to_float(&evaluate(&args[4], cells)?)?;
        if !mode.is_finite() || mode.fract() != 0.0 || !(0.0..=1.0).contains(&mode) {
            return Ok(Variant::Error(ExcelError::Value));
        }
        mode == 1.0
    } else {
        false
    };
    let pad_with = if args.len() >= 6 {
        evaluate(&args[5], cells)?
    } else {
        Variant::Error(ExcelError::NA)
    };

    let split = |source: &str, delimiters: &[String]| -> Vec<String> {
        let search_source = if case_insensitive {
            source.to_lowercase()
        } else {
            source.to_string()
        };
        let mut parts = vec![];
        let mut start = 0usize;
        while start <= search_source.len() {
            let next = delimiters
                .iter()
                .filter_map(|delimiter| {
                    let search_delimiter = if case_insensitive {
                        delimiter.to_lowercase()
                    } else {
                        delimiter.clone()
                    };
                    search_source[start..]
                        .find(&search_delimiter)
                        .map(|relative| (start + relative, search_delimiter.len()))
                })
                .min_by(|(left_pos, left_len), (right_pos, right_len)| {
                    left_pos
                        .cmp(right_pos)
                        .then_with(|| right_len.cmp(left_len))
                });
            let Some((end, delimiter_len)) = next else {
                break;
            };
            let part = &source[start..end];
            if !ignore_empty || !part.is_empty() {
                parts.push(part.to_string());
            }
            start = end + delimiter_len;
        }
        let part = &source[start..];
        if !ignore_empty || !part.is_empty() {
            parts.push(part.to_string());
        }
        parts
    };

    let rows = row_delims
        .as_deref()
        .map_or_else(|| vec![text.clone()], |delimiters| split(&text, delimiters));
    let mut split_rows: Vec<Vec<String>> =
        rows.iter().map(|row| split(row, &column_delims)).collect();
    let width = split_rows.iter().map(Vec::len).max().unwrap_or(0);
    let mut result = vec![];
    for row in &mut split_rows {
        row.resize(width, to_str(&pad_with));
        result.extend(row.iter().cloned().map(Variant::Str));
    }
    Ok(wrap_array(result))
}

fn textsplit_delimiters(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
    allow_empty_placeholder: bool,
) -> Result<Vec<String>, String> {
    let values = collect_values(expr, cells)?;
    let single_value = values.len() == 1;
    let mut delimiters = Vec::with_capacity(values.len());
    for value in values {
        let delimiter = to_str(&value);
        if delimiter.is_empty() {
            if allow_empty_placeholder && single_value {
                continue;
            }
            return Err("TEXTSPLIT: delimiter cannot be empty".into());
        }
        delimiters.push(delimiter);
    }
    if delimiters.is_empty() && !allow_empty_placeholder {
        return Err("TEXTSPLIT: column delimiter cannot be empty".into());
    }
    Ok(delimiters)
}

// ── TEXTBEFORE / TEXTAFTER ────────────────────────────────────────────────────

/// Collect byte offsets of every occurrence of `delim` in `haystack`.
fn find_all_occurrences(haystack: &str, delim: &str) -> Vec<usize> {
    let mut positions = vec![];
    let mut start = 0;
    while let Some(p) = haystack[start..].find(delim) {
        positions.push(start + p);
        start += p + delim.len();
    }
    positions
}

fn text_before_after(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    before: bool,
    fname: &str,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err(format!("{} requires at least 2 arguments", fname));
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let delim = to_str(&evaluate(&args[1], cells)?);
    let instance_num: i64 = if args.len() >= 3 {
        let value = to_float(&evaluate(&args[2], cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 {
            return Ok(Variant::Error(ExcelError::Value));
        }
        if value < i64::MIN as f64 || value > i64::MAX as f64 {
            return Ok(Variant::Error(ExcelError::Value));
        }
        value as i64
    } else {
        1
    };
    let case_insensitive = if args.len() >= 4 {
        let value = to_float(&evaluate(&args[3], cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=1.0).contains(&value) {
            return Ok(Variant::Error(ExcelError::Value));
        }
        value == 1.0
    } else {
        false
    };
    let match_end = if args.len() >= 5 {
        let value = to_float(&evaluate(&args[4], cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=1.0).contains(&value) {
            return Ok(Variant::Error(ExcelError::Value));
        }
        value == 1.0
    } else {
        false
    };

    // Excel treats an empty delimiter as an immediate boundary. This is
    // useful for explicitly selecting the first/last side of a string and
    // must not be confused with an invalid delimiter.
    if delim.is_empty() {
        let result = if before {
            if instance_num < 0 {
                text.clone()
            } else if instance_num > 0 {
                String::new()
            } else {
                return Ok(Variant::Error(ExcelError::Value));
            }
        } else if instance_num < 0 {
            String::new()
        } else if instance_num > 0 {
            text.clone()
        } else {
            return Ok(Variant::Error(ExcelError::Value));
        };
        return Ok(Variant::Str(result));
    }

    let (search_text, search_delim) = if case_insensitive {
        (text.to_lowercase(), delim.to_lowercase())
    } else {
        (text.clone(), delim.clone())
    };

    let positions = find_all_occurrences(&search_text, &search_delim);
    let mut positions = positions;
    if match_end && positions.last().copied() != Some(text.len()) {
        positions.push(text.len());
    }

    // Resolve instance_num to an index
    let idx: Option<usize> = if instance_num > 0 {
        let n = (instance_num - 1) as usize;
        if n < positions.len() { Some(n) } else { None }
    } else if instance_num < 0 {
        let n = (-instance_num) as usize;
        if n <= positions.len() {
            Some(positions.len() - n)
        } else {
            None
        }
    } else {
        None // 0 is invalid
    };

    let not_found = || -> Result<Variant, String> {
        if args.len() >= 6 {
            Ok(evaluate(&args[5], cells)?)
        } else {
            Ok(Variant::Error(ExcelError::NA))
        }
    };

    match idx {
        None => not_found(),
        Some(i) => {
            let byte_pos = positions[i];
            if before {
                Ok(Variant::Str(text[..byte_pos].to_string()))
            } else if byte_pos == text.len() {
                Ok(Variant::Str(String::new()))
            } else {
                Ok(Variant::Str(text[byte_pos + delim.len()..].to_string()))
            }
        }
    }
}

fn func_textbefore(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    text_before_after(args, cells, true, "TEXTBEFORE")
}

fn func_textafter(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    text_before_after(args, cells, false, "TEXTAFTER")
}

// ── VALUETOTEXT ───────────────────────────────────────────────────────────────

fn func_valuetotext(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("VALUETOTEXT requires at least 1 argument".into());
    }
    let val = evaluate(&args[0], cells)?;
    let format = if args.len() >= 2 {
        to_float(&evaluate(&args[1], cells)?)? as i32
    } else {
        0
    };
    let s = match &val {
        Variant::Str(s) => {
            if format == 1 {
                format!("\"{}\"", s.replace('"', "\"\""))
            } else {
                s.clone()
            }
        }
        Variant::Integer(n) => n.to_string(),
        Variant::Float(f) => format!("{}", f),
        Variant::Boolean(b) => {
            if *b {
                "TRUE".into()
            } else {
                "FALSE".into()
            }
        }
        Variant::Empty => String::new(),
        _ => to_str(&val),
    };
    Ok(Variant::Str(s))
}

fn arraytotext_atom(value: &Variant, format: i64) -> String {
    match value {
        Variant::Str(text) if format == 1 => format!("\"{}\"", text.replace('"', "\"\"")),
        Variant::Array(values) => values
            .iter()
            .map(|value| arraytotext_atom(value, format))
            .collect::<Vec<_>>()
            .join(","),
        _ => to_str(value),
    }
}

fn func_arraytotext(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if !(1..=2).contains(&args.len()) {
        return Err("ARRAYTOTEXT requires 1 or 2 arguments".into());
    }
    let format = if let Some(arg) = args.get(1) {
        let value = to_float(&evaluate(arg, cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=1.0).contains(&value) {
            return Ok(Variant::Error(ExcelError::Value));
        }
        value as i64
    } else {
        0
    };
    let value = evaluate(&args[0], cells)?;
    if let Variant::Error(error) = value {
        return Ok(Variant::Error(error));
    }
    let text = match &value {
        Variant::Array(values) => {
            let body = values
                .iter()
                .map(|item| arraytotext_atom(item, format))
                .collect::<Vec<_>>()
                .join(if format == 1 { "," } else { ", " });
            if format == 1 {
                format!("{{{body}}}")
            } else {
                body
            }
        }
        _ => arraytotext_atom(&value, format),
    };
    Ok(Variant::Str(text))
}

fn func_single(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("SINGLE requires 1 argument".into());
    }
    match evaluate(&args[0], cells)? {
        Variant::Array(values) => values
            .into_iter()
            .next()
            .ok_or_else(|| "SINGLE: empty array".into()),
        value => Ok(value),
    }
}

// ── TAKE / DROP ───────────────────────────────────────────────────────────────

fn func_take(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("TAKE requires 2 or 3 arguments".into());
    }
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    let value_len = vals.len();
    slice_array(
        vals,
        array_shape_for_expr(&args[0], cells, value_len),
        args,
        cells,
        true,
    )
}

fn func_drop(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("DROP requires 2 or 3 arguments".into());
    }
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    let value_len = vals.len();
    slice_array(
        vals,
        array_shape_for_expr(&args[0], cells, value_len),
        args,
        cells,
        false,
    )
}

fn slice_array(
    values: Vec<Variant>,
    (rows, cols): (usize, usize),
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    take: bool,
) -> Result<Variant, String> {
    let count = |arg: &FormulaExpr, size: usize| -> Result<(usize, usize), String> {
        let n = match evaluate(arg, cells)? {
            Variant::Integer(value) => value,
            Variant::Float(value)
                if value.is_finite()
                    && value.fract() == 0.0
                    && value >= i64::MIN as f64
                    && value <= i64::MAX as f64 =>
            {
                value as i64
            }
            _ => return Err("TAKE/DROP row or column count must be an integer".into()),
        };
        if n == 0 {
            return Err("TAKE/DROP row or column count cannot be zero".into());
        }
        let amount = n.unsigned_abs() as usize;
        let amount = amount.min(size);
        if take {
            Ok(if n > 0 {
                (0, amount)
            } else {
                (size.saturating_sub(amount), size)
            })
        } else {
            Ok(if n > 0 {
                (amount, size)
            } else {
                (0, size.saturating_sub(amount))
            })
        }
    };
    let (row_start, row_end) = count(&args[1], rows)?;
    let (col_start, col_end) = if let Some(arg) = args.get(2) {
        count(arg, cols)?
    } else {
        (0, cols)
    };
    let mut result = Vec::with_capacity((row_end - row_start) * (col_end - col_start));
    for row in row_start..row_end {
        for col in col_start..col_end {
            result.push(values[row * cols + col].clone());
        }
    }
    Ok(wrap_array(result))
}

// ── VSTACK / HSTACK ───────────────────────────────────────────────────────────

fn func_vstack(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("VSTACK requires at least 1 argument".into());
    }
    let parts: Vec<(Vec<Variant>, usize, usize)> = args
        .iter()
        .map(|arg| {
            let values = flatten_array_vals(collect_values(arg, cells)?);
            let (rows, cols) = array_shape_for_expr(arg, cells, values.len());
            Ok((values, rows, cols))
        })
        .collect::<Result<_, String>>()?;
    let width = parts.iter().map(|(_, _, cols)| *cols).max().unwrap_or(0);
    let mut result = Vec::new();
    for (values, rows, cols) in parts {
        for row in 0..rows {
            for col in 0..width {
                result.push(if col < cols {
                    values[row * cols + col].clone()
                } else {
                    Variant::Error(ExcelError::NA)
                });
            }
        }
    }
    Ok(wrap_array(result))
}

/// Returns the worksheet shape for the array forms whose legacy value is a
/// flat row-major vector. Unknown expressions intentionally fall back to one
/// row, preserving the historical evaluator contract.
fn array_shape_for_expr(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
    value_len: usize,
) -> (usize, usize) {
    match expr {
        FormulaExpr::Range { c1, r1, c2, r2, .. } => (
            (r1.max(r2) - r1.min(r2) + 1) as usize,
            (c1.max(c2) - c1.min(c2) + 1) as usize,
        ),
        FormulaExpr::FuncCall { name, args }
            if name.eq_ignore_ascii_case("SEQUENCE") || name.eq_ignore_ascii_case("RANDARRAY") =>
        {
            let dimension = |arg: Option<&FormulaExpr>| {
                arg.and_then(|arg| evaluate(arg, cells).ok())
                    .and_then(|value| match value {
                        Variant::Integer(n) if n >= 0 => Some(n as usize),
                        Variant::Float(n) if n.is_finite() && n >= 0.0 => Some(n as usize),
                        _ => None,
                    })
                    .unwrap_or(1)
                    .max(1)
            };
            (dimension(args.first()), dimension(args.get(1)))
        }
        FormulaExpr::FuncCall { name, args } if name.eq_ignore_ascii_case("MUNIT") => {
            let n = args
                .first()
                .and_then(|arg| evaluate(arg, cells).ok())
                .and_then(|value| match value {
                    Variant::Integer(n) if n > 0 => Some(n as usize),
                    Variant::Float(n) if n.is_finite() && n.fract() == 0.0 && n > 0.0 => {
                        Some(n as usize)
                    }
                    _ => None,
                })
                .unwrap_or(1);
            (n, n)
        }
        FormulaExpr::FuncCall { name, args } if name.eq_ignore_ascii_case("VSTACK") => {
            let mut rows = 0usize;
            let mut cols = 0usize;
            for arg in args {
                let values = flatten_array_vals(collect_values(arg, cells).unwrap_or_default());
                let (arg_rows, arg_cols) = array_shape_for_expr(arg, cells, values.len());
                rows = rows.saturating_add(arg_rows);
                cols = cols.max(arg_cols);
            }
            (rows.max(1), cols.max(1))
        }
        FormulaExpr::FuncCall { name, args } if name.eq_ignore_ascii_case("HSTACK") => {
            let mut rows = 0usize;
            let mut cols = 0usize;
            for arg in args {
                let values = flatten_array_vals(collect_values(arg, cells).unwrap_or_default());
                let (arg_rows, arg_cols) = array_shape_for_expr(arg, cells, values.len());
                rows = rows.max(arg_rows);
                cols = cols.saturating_add(arg_cols);
            }
            (rows.max(1), cols.max(1))
        }
        FormulaExpr::FuncCall { name, args } if name.eq_ignore_ascii_case("TRANSPOSE") => {
            let Some(arg) = args.first() else {
                return (1, value_len);
            };
            let value = evaluate(arg, cells).ok();
            let (rows, cols) = array_shape_for_expr(
                arg,
                cells,
                value.as_ref().map_or(1, |v| match v {
                    Variant::Array(values) => values.len(),
                    _ => 1,
                }),
            );
            (cols, rows)
        }
        FormulaExpr::FuncCall { name, args }
            if matches!(name.to_ascii_uppercase().as_str(), "LINEST" | "LOGEST") =>
        {
            let known_x = args
                .get(1)
                .map(|arg| {
                    let values = collect_values(arg, cells).unwrap_or_default();
                    array_shape_for_expr(arg, cells, values.len()).1
                })
                .unwrap_or(1);
            let stats = args
                .get(3)
                .and_then(|arg| evaluate(arg, cells).ok())
                .is_some_and(|value| is_truthy(&value));
            (if stats { 5 } else { 1 }, known_x.saturating_add(1))
        }
        FormulaExpr::FuncCall { name, args }
            if matches!(name.to_ascii_uppercase().as_str(), "TREND" | "GROWTH") =>
        {
            args.get(2)
                .map(|arg| {
                    let values = collect_values(arg, cells).unwrap_or_default();
                    array_shape_for_expr(arg, cells, values.len())
                })
                .unwrap_or((1, 1))
        }
        FormulaExpr::BinOp { lhs, rhs, .. } => {
            let operand_shape = |operand: &FormulaExpr| {
                let values = flatten_array_vals(collect_values(operand, cells).ok()?);
                (values.len() > 1).then(|| array_shape_for_expr(operand, cells, values.len()))
            };
            operand_shape(lhs)
                .or_else(|| operand_shape(rhs))
                .unwrap_or((1, value_len))
        }
        _ => (1, value_len),
    }
}

fn func_hstack(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("HSTACK requires at least 1 argument".into());
    }
    let parts: Vec<(Vec<Variant>, usize, usize)> = args
        .iter()
        .map(|arg| {
            let values = flatten_array_vals(collect_values(arg, cells)?);
            let (rows, cols) = array_shape_for_expr(arg, cells, values.len());
            Ok((values, rows, cols))
        })
        .collect::<Result<_, String>>()?;
    let two_dimensional = parts.iter().any(|(_, rows, _)| *rows > 1);
    let mut result = vec![];
    if two_dimensional {
        let height = parts.iter().map(|(_, rows, _)| *rows).max().unwrap_or(0);
        for row in 0..height {
            for (values, _, cols) in &parts {
                if row < values.len() / (*cols).max(1) {
                    result.extend(values[row * cols..(row + 1) * cols].iter().cloned());
                } else {
                    result.extend(std::iter::repeat_n(Variant::Error(ExcelError::NA), *cols));
                }
            }
        }
    } else {
        for (values, _, _) in parts {
            result.extend(values);
        }
    }
    Ok(wrap_array(result))
}

// ── CHOOSECOLS / CHOOSEROWS ───────────────────────────────────────────────────

fn choose_elements(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    fname: &str,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err(format!("{} requires at least 2 arguments", fname));
    }
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    let value_len = vals.len();
    let (rows, cols) = array_shape_for_expr(&args[0], cells, value_len);
    let integer_index = |arg: &FormulaExpr| -> Result<i64, String> {
        match evaluate(arg, cells)? {
            Variant::Integer(value) => Ok(value),
            Variant::Float(value)
                if value.is_finite()
                    && value.fract() == 0.0
                    && value >= i64::MIN as f64
                    && value <= i64::MAX as f64 =>
            {
                Ok(value as i64)
            }
            _ => Err(format!("{} index must be an integer", fname)),
        }
    };
    if rows > 1 && cols > 1 {
        let mut indices = Vec::with_capacity(args.len() - 1);
        for arg in &args[1..] {
            let n = match integer_index(arg) {
                Ok(value) => value,
                Err(_) => return Ok(Variant::Error(ExcelError::Value)),
            };
            let size = if fname == "CHOOSECOLS" { cols } else { rows };
            let idx = if n > 0 { n - 1 } else { size as i64 + n };
            if idx < 0 || idx >= size as i64 {
                return Ok(Variant::Error(ExcelError::Value));
            }
            indices.push(idx as usize);
        }
        let mut result = Vec::new();
        if fname == "CHOOSECOLS" {
            for row in 0..rows {
                for &col in &indices {
                    result.push(vals[row * cols + col].clone());
                }
            }
        } else {
            for &row in &indices {
                result.extend(vals[row * cols..(row + 1) * cols].iter().cloned());
            }
        }
        return Ok(wrap_array(result));
    }
    let len = vals.len() as i64;
    let mut result = vec![];
    for arg in &args[1..] {
        let n = match integer_index(arg) {
            Ok(value) => value,
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        };
        let idx = if n > 0 { n - 1 } else { len + n };
        if idx < 0 || idx >= len {
            result.push(Variant::Error(ExcelError::Value));
        } else {
            result.push(vals[idx as usize].clone());
        }
    }
    Ok(wrap_array(result))
}

fn func_choosecols(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    choose_elements(args, cells, "CHOOSECOLS")
}

fn func_chooserows(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    choose_elements(args, cells, "CHOOSEROWS")
}

fn func_expand(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 4 {
        return Err("EXPAND requires 2 to 4 arguments".into());
    }
    let values = flatten_array_vals(collect_values(&args[0], cells)?);
    let (source_rows, source_cols) = array_shape_for_expr(&args[0], cells, values.len());
    let integer_arg = |arg: &FormulaExpr, name: &str| -> Result<usize, String> {
        match evaluate(arg, cells)? {
            Variant::Integer(value) if value > 0 => Ok(value as usize),
            Variant::Float(value) if value.is_finite() && value.fract() == 0.0 && value > 0.0 => {
                Ok(value as usize)
            }
            _ => Err(format!("EXPAND {name} must be a positive integer")),
        }
    };
    let target_rows = integer_arg(&args[1], "rows")?;
    let target_cols = if let Some(columns) = args.get(2) {
        integer_arg(columns, "columns")?
    } else {
        source_cols
    };
    if target_rows < source_rows || target_cols < source_cols {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let pad = if let Some(pad_with) = args.get(3) {
        evaluate(pad_with, cells)?
    } else {
        Variant::Error(ExcelError::NA)
    };
    let mut result = Vec::with_capacity(target_rows.saturating_mul(target_cols));
    for row in 0..target_rows {
        for col in 0..target_cols {
            result.push(if row < source_rows && col < source_cols {
                values[row * source_cols + col].clone()
            } else {
                pad.clone()
            });
        }
    }
    Ok(wrap_array(result))
}

fn func_trimrange(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 3 {
        return Err("TRIMRANGE requires 1 to 3 arguments".into());
    }
    let values = flatten_array_vals(collect_values(&args[0], cells)?);
    let (rows, cols) = array_shape_for_expr(&args[0], cells, values.len());
    if rows == 0 || cols == 0 || values.len() != rows.saturating_mul(cols) {
        return Ok(Variant::Empty);
    }
    let trim_mode = |arg: Option<&FormulaExpr>, name: &str| -> Result<u8, String> {
        let value = if let Some(arg) = arg {
            to_float(&evaluate(arg, cells)?)?
        } else {
            3.0
        };
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=3.0).contains(&value) {
            return Err(format!("TRIMRANGE {name} must be an integer from 0 to 3"));
        }
        Ok(value as u8)
    };
    let row_mode = trim_mode(args.get(1), "trim_rows")?;
    let col_mode = trim_mode(args.get(2), "trim_cols")?;
    let blank = |row: usize, col: usize| matches!(values[row * cols + col], Variant::Empty);
    let mut row_start = 0;
    let mut row_end = rows;
    if row_mode & 1 != 0 {
        while row_start < row_end && (0..cols).all(|col| blank(row_start, col)) {
            row_start += 1;
        }
    }
    if row_mode & 2 != 0 {
        while row_end > row_start && (0..cols).all(|col| blank(row_end - 1, col)) {
            row_end -= 1;
        }
    }
    let mut col_start = 0;
    let mut col_end = cols;
    if col_mode & 1 != 0 {
        while col_start < col_end && (row_start..row_end).all(|row| blank(row, col_start)) {
            col_start += 1;
        }
    }
    if col_mode & 2 != 0 {
        while col_end > col_start && (row_start..row_end).all(|row| blank(row, col_end - 1)) {
            col_end -= 1;
        }
    }
    let mut result = Vec::with_capacity((row_end - row_start) * (col_end - col_start));
    for row in row_start..row_end {
        for col in col_start..col_end {
            result.push(values[row * cols + col].clone());
        }
    }
    Ok(wrap_array(result))
}

// ── COMBIN ───────────────────────────────────────────────────────────────────

fn func_combin(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("COMBIN requires 2 arguments".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)? as i64;
    let k = to_float(&evaluate(&args[1], cells)?)? as i64;
    if n < 0 || k < 0 || k > n {
        return Ok(Variant::Error(ExcelError::Num));
    }
    // Multiplicative formula: ∏ (n-i)/(i+1) for i in 0..k  — avoids factorial overflow
    let mut result = 1f64;
    for i in 0..k {
        result = result * (n - i) as f64 / (i + 1) as f64;
    }
    Ok(as_integer_if_whole(result.round()))
}

fn func_combina(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("COMBINA requires 2 arguments".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    let k = to_float(&evaluate(&args[1], cells)?)?;
    if !n.is_finite()
        || !k.is_finite()
        || n < 0.0
        || k < 0.0
        || n.fract() != 0.0
        || k.fract() != 0.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let n = n as u64;
    let k = k as u64;
    if n == 0 {
        return Ok(Variant::Integer(if k == 0 { 1 } else { 0 }));
    }
    let total = n.saturating_add(k).saturating_sub(1);
    if total > i64::MAX as u64 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let choose_k = k.min(total.saturating_sub(k));
    let mut result = 1.0;
    for i in 0..choose_k {
        result = result * (total - i) as f64 / (i + 1) as f64;
    }
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole(result.round()))
}

// ── Math & Combinatorics ──────────────────────────────────────────────────────

fn func_fact(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("FACT requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)? as i64;
    if n < 0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    if n > 170 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut result = 1f64;
    for i in 2..=n {
        result *= i as f64;
    }
    Ok(as_integer_if_whole(result))
}

fn func_factdouble(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("FACTDOUBLE requires 1 argument".into());
    }
    let value = to_float(&evaluate(&args[0], cells)?)?;
    if !value.is_finite() || value < 0.0 || value.fract() != 0.0 || value > 300.0 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut result = 1.0;
    let mut n = value as i64;
    while n > 1 {
        result *= n as f64;
        n -= 2;
    }
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole(result))
}

fn func_permut(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("PERMUT requires 2 arguments".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)? as i64;
    let k = to_float(&evaluate(&args[1], cells)?)? as i64;
    if n < 0 || k < 0 || k > n {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut result = 1f64;
    for i in (n - k + 1)..=n {
        result *= i as f64;
    }
    Ok(as_integer_if_whole(result))
}

fn func_permutationa(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("PERMUTATIONA requires 2 arguments".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    let k = to_float(&evaluate(&args[1], cells)?)?;
    if !n.is_finite()
        || !k.is_finite()
        || n < 0.0
        || k < 0.0
        || n.fract() != 0.0
        || k.fract() != 0.0
        || n > 1.0e6
        || k > 1024.0
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let result = n.powf(k);
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole(result))
}

fn gcd_two(mut a: u64, mut b: u64) -> u64 {
    while b != 0 {
        let t = b;
        b = a % b;
        a = t;
    }
    a
}

fn func_gcd(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("GCD requires at least 1 argument".into());
    }
    let vals = collect_all(args, cells)?;
    let mut result = 0u64;
    for v in &vals {
        let n = to_float(v)? as i64;
        if n < 0 {
            return Ok(Variant::Error(ExcelError::Num));
        }
        result = gcd_two(result, n as u64);
    }
    Ok(Variant::Integer(result as i64))
}

fn func_lcm(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("LCM requires at least 1 argument".into());
    }
    let vals = collect_all(args, cells)?;
    let mut result = 1u64;
    for v in &vals {
        let n = to_float(v)? as i64;
        if n < 0 {
            return Ok(Variant::Error(ExcelError::Num));
        }
        let n = n as u64;
        let g = gcd_two(result, n);
        if g == 0 {
            result = 0;
            break;
        }
        result = match (result / g).checked_mul(n) {
            Some(r) => r,
            None => return Ok(Variant::Error(ExcelError::Num)),
        };
    }
    if result > i64::MAX as u64 {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(Variant::Integer(result as i64))
}

fn func_quotient(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("QUOTIENT requires 2 arguments".into());
    }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let den = to_float(&evaluate(&args[1], cells)?)?;
    if den == 0.0 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    Ok(Variant::Integer((num / den).trunc() as i64))
}

fn func_sign(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 1 {
        return Err("SIGN requires 1 argument".into());
    }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    Ok(Variant::Integer(if n > 0.0 {
        1
    } else if n < 0.0 {
        -1
    } else {
        0
    }))
}

fn func_multinomial(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("MULTINOMIAL requires at least 1 argument".into());
    }
    let values = collect_all(args, cells)?;
    let mut total = 0u64;
    let mut denominator = 1.0;
    for value in values {
        let number = to_float(&value)?;
        if !number.is_finite() || number < 0.0 || number.fract() != 0.0 {
            return Ok(Variant::Error(ExcelError::Num));
        }
        let number = number as u64;
        total = match total.checked_add(number) {
            Some(total) if total <= 170 => total,
            _ => return Ok(Variant::Error(ExcelError::Num)),
        };
        for factor in 2..=number {
            denominator *= factor as f64;
        }
    }
    let mut numerator = 1.0;
    for factor in 2..=total {
        numerator *= factor as f64;
    }
    let result = numerator / denominator;
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole(result.round()))
}

fn paired_sum<F>(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    name: &str,
    operation: F,
) -> Result<Variant, String>
where
    F: Fn(f64, f64) -> f64,
{
    let (left, right) = collect_paired(args, cells, name)?;
    let result = left
        .iter()
        .zip(right.iter())
        .map(|(a, b)| operation(*a, *b))
        .sum::<f64>();
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole(result))
}

fn func_sumx2my2(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    paired_sum(args, cells, "SUMX2MY2", |a, b| a * a - b * b)
}

fn func_sumx2py2(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    paired_sum(args, cells, "SUMX2PY2", |a, b| a * a + b * b)
}

fn func_sumxmy2(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    paired_sum(args, cells, "SUMXMY2", |a, b| (a - b) * (a - b))
}

fn func_seriessum(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 4 {
        return Err("SERIESSUM requires 4 arguments".into());
    }
    let x = to_float(&evaluate(&args[0], cells)?)?;
    let n = to_float(&evaluate(&args[1], cells)?)?;
    let m = to_float(&evaluate(&args[2], cells)?)?;
    let coefficients = collect_values(&args[3], cells)?
        .into_iter()
        .filter_map(|value| as_f64(&value))
        .collect::<Vec<_>>();
    if !x.is_finite()
        || !n.is_finite()
        || !m.is_finite()
        || coefficients.is_empty()
        || coefficients.iter().any(|value| !value.is_finite())
    {
        return Ok(Variant::Error(ExcelError::Num));
    }
    let mut result = 0.0;
    for (index, coefficient) in coefficients.into_iter().enumerate() {
        result += coefficient * x.powf(n + m * index as f64);
    }
    if !result.is_finite() {
        return Ok(Variant::Error(ExcelError::Num));
    }
    Ok(as_integer_if_whole(result))
}

// ── DGET ─────────────────────────────────────────────────────────────────────

/// Resolve the `field` argument of a database function to an absolute column number.
/// `field` may be a 1-based column index (Integer/Float) or a header name (Str).
fn resolve_db_field(
    field_val: &Variant,
    cells: &HashMap<(u32, u32), CellContent>,
    c1: u32,
    c2: u32,
    header_row: u32,
) -> Result<u32, String> {
    match field_val {
        Variant::Integer(n) => {
            if *n < 1 {
                return Err("DGET: field index must be >= 1".into());
            }
            Ok(c1 + (*n as u32) - 1)
        }
        Variant::Float(f) => {
            let n = *f as i64;
            if n < 1 {
                return Err("DGET: field index must be >= 1".into());
            }
            Ok(c1 + (n as u32) - 1)
        }
        Variant::Str(s) => (c1..=c2)
            .find(|&c| to_str(&cell_val(cells, header_row, c)).eq_ignore_ascii_case(s))
            .ok_or_else(|| format!("DGET: field '{}' not found in database header", s)),
        _ => Err("DGET: field argument must be a column number or field name string".into()),
    }
}

/// Test whether a single database row satisfies the criteria range.
/// Criteria rows (excluding header) are OR-combined; columns within a row are AND-combined.
/// An empty criteria cell is treated as "match all" (wildcard). `db_range`/`criteria_range`
/// are `(c1, r1, c2, r2)`, matching `require_range`'s own return shape unchanged -- the
/// caller already has both as one tuple each from `require_range`, so this avoids
/// destructuring just to re-pass four scalars each (`db_range`'s own `r2` -- the database's
/// last row -- goes unused here; the caller's own loop bound needs it, this function
/// doesn't).
fn db_row_matches_criteria(
    cells: &HashMap<(u32, u32), CellContent>,
    data_row: u32,
    db_range: (u32, u32, u32, u32),
    criteria_range: (u32, u32, u32, u32),
) -> bool {
    let (db_c1, db_header_row, db_c2, _db_r2) = db_range;
    let (cr_c1, cr_r1, cr_c2, cr_r2) = criteria_range;
    if cr_r2 == cr_r1 {
        return true;
    }
    // Each criteria row (cr_r1+1 .. cr_r2) is one OR-branch
    for cr_row in (cr_r1 + 1)..=cr_r2 {
        let mut row_match = true;
        for cr_col in cr_c1..=cr_c2 {
            let crit = cell_val(cells, cr_row, cr_col);
            if matches!(crit, Variant::Empty) {
                continue;
            } // blank criteria = match all
            // Find the database column this criteria column corresponds to
            let header_name = to_str(&cell_val(cells, cr_r1, cr_col));
            if header_name.is_empty() {
                continue;
            }
            // Locate the matching database column
            let db_col = match (db_c1..=db_c2).find(|&c| {
                to_str(&cell_val(cells, db_header_row, c)).eq_ignore_ascii_case(&header_name)
            }) {
                Some(c) => c,
                None => {
                    row_match = false;
                    break;
                }
            };
            let data_val = cell_val(cells, data_row, db_col);
            if !matches_criteria(&data_val, &crit) {
                row_match = false;
                break;
            }
        }
        if row_match {
            return true;
        }
    }
    false
}

fn func_dget(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("DGET requires 3 arguments".into());
    }

    let (db_c1, db_r1, db_c2, db_r2) = require_range(&args[0], "DGET")?;
    let field_val = evaluate(&args[1], cells)?;
    let field_col = resolve_db_field(&field_val, cells, db_c1, db_c2, db_r1)?;
    if field_col > db_c2 {
        return Ok(Variant::Error(ExcelError::Ref));
    }

    let (cr_c1, cr_r1, cr_c2, cr_r2) = require_range(&args[2], "DGET")?;

    let db_range = (db_c1, db_r1, db_c2, db_r2);
    let criteria_range = (cr_c1, cr_r1, cr_c2, cr_r2);
    let mut matched: Vec<Variant> = vec![];
    for row in (db_r1 + 1)..=db_r2 {
        if db_row_matches_criteria(cells, row, db_range, criteria_range) {
            matched.push(cell_val(cells, row, field_col));
        }
    }

    match matched.len() {
        0 => Ok(Variant::Error(ExcelError::Value)), // no match
        1 => Ok(matched.remove(0)),
        _ => Ok(Variant::Error(ExcelError::Num)), // multiple matches
    }
}

// ── DSUM / DAVERAGE / DCOUNT / DCOUNTA / DMAX / DMIN ─────────────────────────

/// Pre-compiled criteria for one database function call.
/// `criteria[i]` = one OR-branch; each entry is `(db_col, ParsedCriteria)` (AND within branch).
/// Built once in `db_resolve_args`; matched per data row without header scanning.
struct DbCtx {
    db_r1: u32,
    db_r2: u32,
    field_col: u32,
    criteria: Vec<Vec<(u32, ParsedCriteria)>>,
}

fn db_resolve_args(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    fname: &str,
) -> Result<Option<DbCtx>, String> {
    if args.len() != 3 {
        return Err(format!("{fname} requires 3 arguments"));
    }
    let (db_c1, db_r1, db_c2, db_r2) = require_range(&args[0], fname)?;
    let field_val = evaluate(&args[1], cells)?;
    let field_col = resolve_db_field(&field_val, cells, db_c1, db_c2, db_r1)?;
    if field_col > db_c2 {
        return Ok(None);
    }
    let (cr_c1, cr_r1, cr_c2, cr_r2) = require_range(&args[2], fname)?;

    // Build header name → db column map once
    let header_map: std::collections::HashMap<String, u32> = (db_c1..=db_c2)
        .map(|c| (to_str(&cell_val(cells, db_r1, c)).to_lowercase(), c))
        .collect();

    // Pre-build criteria: outer = OR rows, inner = (db_col, ParsedCriteria)
    let criteria: Vec<Vec<(u32, ParsedCriteria)>> = if cr_r2 == cr_r1 {
        vec![Vec::new()]
    } else {
        (cr_r1 + 1..=cr_r2)
            .map(|cr_row| {
                let mut branch: Vec<(u32, ParsedCriteria)> = Vec::new();
                let mut valid = true;
                for cr_col in cr_c1..=cr_c2 {
                    let crit_val = cell_val(cells, cr_row, cr_col);
                    if matches!(crit_val, Variant::Empty) {
                        continue;
                    }
                    let header_name = to_str(&cell_val(cells, cr_r1, cr_col)).to_lowercase();
                    if header_name.is_empty() {
                        continue;
                    }
                    match header_map.get(&header_name) {
                        Some(&db_col) => branch.push((db_col, parse_criteria(&crit_val))),
                        None => {
                            valid = false;
                            break;
                        }
                    }
                }
                if !valid {
                    vec![(0, ParsedCriteria::CompNum(CompOp::Gt, f64::NAN))]
                } else {
                    branch
                }
            })
            .collect()
    };

    Ok(Some(DbCtx {
        db_r1,
        db_r2,
        field_col,
        criteria,
    }))
}

fn db_matched_vals(ctx: &DbCtx, cells: &HashMap<(u32, u32), CellContent>) -> Vec<Variant> {
    const EMPTY: Variant = Variant::Empty;
    (ctx.db_r1 + 1..=ctx.db_r2)
        .filter(|&row| {
            ctx.criteria.iter().any(|or_branch| {
                or_branch.iter().all(|(db_col, pcrit)| {
                    matches_parsed(cell_ref(cells, row, *db_col).unwrap_or(&EMPTY), pcrit)
                })
            })
        })
        .map(|row| cell_val(cells, row, ctx.field_col))
        .collect()
}

fn func_dsum(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DSUM")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    let sum: f64 = db_matched_vals(&ctx, cells).iter().filter_map(as_f64).sum();
    Ok(as_integer_if_whole(sum))
}

fn func_daverage(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DAVERAGE")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    let nums: Vec<f64> = db_matched_vals(&ctx, cells)
        .iter()
        .filter_map(as_f64)
        .collect();
    if nums.is_empty() {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
}

fn func_dcount(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DCOUNT")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    let count = db_matched_vals(&ctx, cells)
        .iter()
        .filter(|v| as_f64(v).is_some())
        .count();
    Ok(Variant::Integer(count as i64))
}

fn func_dcounta(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DCOUNTA")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    let count = db_matched_vals(&ctx, cells)
        .iter()
        .filter(|v| !matches!(v, Variant::Empty))
        .count();
    Ok(Variant::Integer(count as i64))
}

fn func_dmax(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DMAX")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    let max = db_matched_vals(&ctx, cells)
        .iter()
        .filter_map(as_f64)
        .reduce(f64::max);
    Ok(as_integer_if_whole(max.unwrap_or(0.0)))
}

fn func_dmin(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DMIN")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    let min = db_matched_vals(&ctx, cells)
        .iter()
        .filter_map(as_f64)
        .reduce(f64::min);
    Ok(as_integer_if_whole(min.unwrap_or(0.0)))
}

fn db_numeric_vals(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    fname: &str,
) -> Result<Option<Vec<f64>>, String> {
    let Some(ctx) = db_resolve_args(args, cells, fname)? else {
        return Ok(None);
    };
    Ok(Some(
        db_matched_vals(&ctx, cells)
            .iter()
            .filter_map(as_f64)
            .collect(),
    ))
}

fn func_dproduct(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(nums) = db_numeric_vals(args, cells, "DPRODUCT")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    Ok(as_integer_if_whole(nums.into_iter().product::<f64>()))
}

fn func_dstdev(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(nums) = db_numeric_vals(args, cells, "DSTDEV")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    if nums.len() < 2 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let variance =
        nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64;
    Ok(Variant::Float(variance.sqrt()))
}

fn func_dstdevp(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(nums) = db_numeric_vals(args, cells, "DSTDEVP")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    if nums.is_empty() {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let variance = nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>() / nums.len() as f64;
    Ok(Variant::Float(variance.sqrt()))
}

fn func_dvar(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(nums) = db_numeric_vals(args, cells, "DVAR")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    if nums.len() < 2 {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    Ok(Variant::Float(
        nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64,
    ))
}

fn func_dvarp(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let Some(nums) = db_numeric_vals(args, cells, "DVARP")? else {
        return Ok(Variant::Error(ExcelError::Ref));
    };
    if nums.is_empty() {
        return Ok(Variant::Error(ExcelError::DivZero));
    }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    Ok(Variant::Float(
        nums.iter().map(|value| (value - mean).powi(2)).sum::<f64>() / nums.len() as f64,
    ))
}

// ── TOCOL / TOROW ─────────────────────────────────────────────────────────────

/// Flatten any `Variant::Array` items in the collected values so that functions
/// like `WRAPCOLS(SEQUENCE(6), 2)` work correctly even when the first argument
/// returns an Array variant rather than a cell range.
fn flatten_array_vals(vals: Vec<Variant>) -> Vec<Variant> {
    vals.into_iter()
        .flat_map(|v| match v {
            Variant::Array(inner) => inner,
            other => vec![other],
        })
        .collect()
}

fn ignore_filter(vals: Vec<Variant>, ignore: u8) -> Vec<Variant> {
    vals.into_iter()
        .filter(|v| {
            let skip_blank = ignore == 1 || ignore == 3;
            let skip_error = ignore == 2 || ignore == 3;
            !(skip_blank && matches!(v, Variant::Empty))
                && !(skip_error && matches!(v, Variant::Error(_)))
        })
        .collect()
}

fn func_tocol(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 3 {
        return Err("TOCOL requires 1 to 3 arguments".into());
    }
    let ignore = if args.len() >= 2 {
        let value = to_float(&evaluate(&args[1], cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=3.0).contains(&value) {
            return Ok(Variant::Error(ExcelError::Value));
        }
        value as u8
    } else {
        0
    };
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    let (rows, cols) = array_shape_for_expr(&args[0], cells, vals.len());
    let scan_by_column = args.len() == 3 && is_truthy(&evaluate(&args[2], cells)?);
    Ok(wrap_array(ignore_filter(
        scan_array_values(vals, rows, cols, scan_by_column),
        ignore,
    )))
}

fn func_torow(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 3 {
        return Err("TOROW requires 1 to 3 arguments".into());
    }
    let ignore = if args.len() >= 2 {
        let value = to_float(&evaluate(&args[1], cells)?)?;
        if !value.is_finite() || value.fract() != 0.0 || !(0.0..=3.0).contains(&value) {
            return Ok(Variant::Error(ExcelError::Value));
        }
        value as u8
    } else {
        0
    };
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    let (rows, cols) = array_shape_for_expr(&args[0], cells, vals.len());
    let scan_by_column = args.len() == 3 && is_truthy(&evaluate(&args[2], cells)?);
    Ok(wrap_array(ignore_filter(
        scan_array_values(vals, rows, cols, scan_by_column),
        ignore,
    )))
}

fn scan_array_values(
    values: Vec<Variant>,
    rows: usize,
    cols: usize,
    scan_by_column: bool,
) -> Vec<Variant> {
    if !scan_by_column || rows <= 1 || cols <= 1 || rows.checked_mul(cols) != Some(values.len()) {
        return values;
    }
    let mut scanned = Vec::with_capacity(values.len());
    for col in 0..cols {
        for row in 0..rows {
            scanned.push(values[row * cols + col].clone());
        }
    }
    scanned
}

// ── WRAPCOLS / WRAPROWS ───────────────────────────────────────────────────────

fn func_wrapcols(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("WRAPCOLS requires 2 arguments".into());
    }
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    let wrap_count = to_float(&evaluate(&args[1], cells)?)? as usize;
    if wrap_count == 0 {
        return Err("WRAPCOLS: wrap_count must be > 0".into());
    }
    let pad = if args.len() >= 3 {
        evaluate(&args[2], cells)?
    } else {
        Variant::Empty
    };
    let n_cols = vals.len().div_ceil(wrap_count);
    // Result is row-major of a (wrap_count × n_cols) 2-D table filled column-by-column:
    //   cell (row, col) = vals[col * wrap_count + row]
    let mut result = Vec::with_capacity(wrap_count * n_cols);
    for row in 0..wrap_count {
        for col in 0..n_cols {
            let idx = col * wrap_count + row;
            result.push(if idx < vals.len() {
                vals[idx].clone()
            } else {
                pad.clone()
            });
        }
    }
    Ok(wrap_array(result))
}

fn func_wraprows(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("WRAPROWS requires 2 arguments".into());
    }
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    let wrap_count = to_float(&evaluate(&args[1], cells)?)? as usize;
    if wrap_count == 0 {
        return Err("WRAPROWS: wrap_count must be > 0".into());
    }
    let pad = if args.len() >= 3 {
        evaluate(&args[2], cells)?
    } else {
        Variant::Empty
    };
    let n_rows = vals.len().div_ceil(wrap_count);
    // Fill row-by-row; pad the last (partial) row if needed.
    let total = n_rows * wrap_count;
    let mut result = Vec::with_capacity(total);
    for i in 0..total {
        result.push(if i < vals.len() {
            vals[i].clone()
        } else {
            pad.clone()
        });
    }
    Ok(wrap_array(result))
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::formula::parser::parse as fparse;
    use crate::vm::ExcelError;

    fn cells_from(pairs: &[((u32, u32), Variant)]) -> HashMap<(u32, u32), CellContent> {
        pairs
            .iter()
            .map(|(k, v)| {
                (
                    *k,
                    CellContent {
                        formula: None,
                        value: v.clone(),
                    },
                )
            })
            .collect()
    }

    fn calc(formula: &str, cells: &HashMap<(u32, u32), CellContent>) -> Variant {
        evaluate(&fparse(formula).unwrap(), cells).unwrap()
    }

    #[test]
    fn test_arithmetic_ops() {
        let c = HashMap::new();
        assert_eq!(calc("=1+2", &c), Variant::Integer(3));
        assert_eq!(calc("=10-3", &c), Variant::Integer(7));
        assert_eq!(calc("=4*5", &c), Variant::Integer(20));
        assert_eq!(calc("=10/4", &c), Variant::Float(2.5));
        assert_eq!(calc("=(1+2)*3", &c), Variant::Integer(9));
    }

    #[test]
    fn test_engineering_bitwise_and_error_functions() {
        let c = HashMap::new();
        assert_eq!(calc("=BITAND(13,7)", &c), Variant::Integer(5));
        assert_eq!(calc("=BITOR(8,3)", &c), Variant::Integer(11));
        assert_eq!(calc("=BITXOR(15,6)", &c), Variant::Integer(9));
        assert_eq!(calc("=BITLSHIFT(3,2)", &c), Variant::Integer(12));
        assert_eq!(calc("=BITRSHIFT(12,2)", &c), Variant::Integer(3));
        assert_eq!(calc("=BITLSHIFT(8,-2)", &c), Variant::Integer(2));
        assert_eq!(calc("=DELTA(4,4)", &c), Variant::Integer(1));
        assert_eq!(calc("=DELTA(4)", &c), Variant::Integer(0));
        assert_eq!(calc("=GESTEP(5,5)", &c), Variant::Integer(1));
        assert_eq!(calc("=GESTEP(-1)", &c), Variant::Integer(0));
        match calc("=ERF(0)", &c) {
            Variant::Float(value) => assert!(value.abs() < 1e-12),
            other => panic!("ERF: {:?}", other),
        }
        match calc("=ERFC(0)", &c) {
            Variant::Float(value) => assert!((value - 1.0).abs() < 1e-12),
            other => panic!("ERFC: {:?}", other),
        }
        assert_eq!(calc("=BESSELJ(0,0)", &c), Variant::Integer(1));
        match calc("=BESSELJ(1,0)", &c) {
            Variant::Float(value) => assert!((value - 0.7651976865579666).abs() < 1e-12),
            other => panic!("BESSELJ: {:?}", other),
        }
        assert_eq!(calc("=BESSELI(0,0)", &c), Variant::Integer(1));
        match calc("=BESSELI(1,0)", &c) {
            Variant::Float(value) => assert!((value - 1.2660658777520084).abs() < 1e-12),
            other => panic!("BESSELI: {:?}", other),
        }
        match calc("=BESSELY(1,0)", &c) {
            Variant::Float(value) => assert!((value - 0.08825696421567696).abs() < 1e-12),
            other => panic!("BESSELY: {:?}", other),
        }
        match calc("=BESSELK(1,0)", &c) {
            Variant::Float(value) => assert!((value - 0.42102443824070834).abs() < 1e-12),
            other => panic!("BESSELK: {:?}", other),
        }
        assert_eq!(
            calc("=BITLSHIFT(281474976710655,1)", &c),
            Variant::Error(ExcelError::Num)
        );
    }

    #[test]
    fn test_engineering_radix_conversions() {
        let c = HashMap::new();
        assert_eq!(calc("=DEC2BIN(10)", &c), Variant::Str("1010".into()));
        assert_eq!(calc("=DEC2BIN(-1)", &c), Variant::Str("1111111111".into()));
        assert_eq!(calc("=BIN2DEC(\"1111111111\")", &c), Variant::Integer(-1));
        assert_eq!(calc("=DEC2HEX(255)", &c), Variant::Str("FF".into()));
        assert_eq!(calc("=HEX2DEC(\"FF\")", &c), Variant::Integer(255));
        assert_eq!(calc("=BIN2HEX(\"1010\")", &c), Variant::Str("A".into()));
        assert_eq!(calc("=HEX2BIN(\"A\")", &c), Variant::Str("1010".into()));
        assert_eq!(calc("=DEC2OCT(8)", &c), Variant::Str("10".into()));
        assert_eq!(calc("=OCT2DEC(\"10\")", &c), Variant::Integer(8));
        assert_eq!(calc("=BIN2OCT(\"1010\")", &c), Variant::Str("12".into()));
        assert_eq!(calc("=OCT2BIN(\"12\")", &c), Variant::Str("1010".into()));
        assert_eq!(calc("=HEX2OCT(\"FF\")", &c), Variant::Str("377".into()));
        assert_eq!(calc("=OCT2HEX(\"377\")", &c), Variant::Str("FF".into()));
        assert_eq!(calc("=DEC2BIN(512)", &c), Variant::Error(ExcelError::Num));
        assert_eq!(calc("=DEC2BIN(10,3)", &c), Variant::Error(ExcelError::Num));
    }

    #[test]
    fn test_complex_number_functions() {
        let c = HashMap::new();
        assert_eq!(calc("=COMPLEX(3,4)", &c), Variant::Str("3+4i".into()));
        assert_eq!(calc("=COMPLEX(3,-1,\"j\")", &c), Variant::Str("3-j".into()));
        assert_eq!(calc("=IMREAL(\"-3-4j\")", &c), Variant::Integer(-3));
        assert_eq!(calc("=IMAGINARY(\"-3-4j\")", &c), Variant::Integer(-4));
        match calc("=IMARGUMENT(\"0+1i\")", &c) {
            Variant::Float(value) => {
                assert!((value - std::f64::consts::FRAC_PI_2).abs() < 1e-12)
            }
            other => panic!("IMARGUMENT unexpected: {:?}", other),
        }
        assert_eq!(
            calc("=IMARGUMENT(\"0\")", &c),
            Variant::Error(ExcelError::Num)
        );
        assert_eq!(calc("=IMABS(\"3+4i\")", &c), Variant::Integer(5));
        assert_eq!(
            calc("=IMSUM(\"1+2j\",\"3+4j\")", &c),
            Variant::Str("4+6j".into())
        );
        assert_eq!(
            calc("=IMSUB(\"5+4i\",\"2+1i\")", &c),
            Variant::Str("3+3i".into())
        );
        assert_eq!(
            calc("=IMPRODUCT(\"1+2i\",\"3+4i\")", &c),
            Variant::Str("-5+10i".into())
        );
        assert_eq!(
            calc("=IMDIV(\"1+2i\",\"1+1i\")", &c),
            Variant::Str("1.5+0.5i".into())
        );
        assert_eq!(
            calc("=IMDIV(\"1+2i\",\"0\")", &c),
            Variant::Error(ExcelError::DivZero)
        );
        assert_eq!(
            calc("=IMCONJUGATE(\"3+4i\")", &c),
            Variant::Str("3-4i".into())
        );
        assert_eq!(calc("=IMEXP(\"0\")", &c), Variant::Str("1".into()));
        assert_eq!(calc("=IMLN(\"1\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMLOG10(\"1\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMLOG2(\"2\")", &c), Variant::Str("1".into()));
        assert_eq!(calc("=IMSQRT(\"-1\")", &c), Variant::Str("i".into()));
        assert_eq!(calc("=IMPOWER(\"2\",\"2\")", &c), Variant::Str("4".into()));
        assert_eq!(calc("=IMSIN(\"0\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMCOS(\"0\")", &c), Variant::Str("1".into()));
        assert_eq!(calc("=IMTAN(\"0\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMSINH(\"0\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMCOSH(\"0\")", &c), Variant::Str("1".into()));
        assert_eq!(calc("=IMTANH(\"0\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMSEC(\"0\")", &c), Variant::Str("1".into()));
        assert!(evaluate(&fparse("=IMCSC(\"0\")").unwrap(), &c).is_err());
        assert!(evaluate(&fparse("=IMCOT(\"0\")").unwrap(), &c).is_err());
        assert_eq!(calc("=IMASIN(\"0\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMACOS(\"1\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMATAN(\"0\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMASINH(\"0\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMACOSH(\"1\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMATANH(\"0\")", &c), Variant::Str("0".into()));
        assert_eq!(calc("=IMSECH(\"0\")", &c), Variant::Str("1".into()));
        assert!(evaluate(&fparse("=IMCSCH(\"0\")").unwrap(), &c).is_err());
        assert!(evaluate(&fparse("=IMCOTH(\"0\")").unwrap(), &c).is_err());
    }

    #[test]
    fn test_series_and_reference_functions() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(3)),
        ]);
        assert_eq!(calc("=SERIESSUM(2,0,1,A1:A3)", &c), Variant::Integer(17));
        assert_eq!(calc("=ISREF(A1)", &c), Variant::Boolean(true));
        assert_eq!(calc("=ISREF(A1:A3)", &c), Variant::Boolean(true));
        assert_eq!(calc("=ISREF(1+2)", &c), Variant::Boolean(false));
        assert_eq!(calc("=ISREF(INDIRECT(\"A1\"))", &c), Variant::Boolean(true));
        assert_eq!(calc("=SHEET()", &c), Variant::Integer(1));
        assert_eq!(calc("=SHEET(A1)", &c), Variant::Integer(1));
        assert_eq!(calc("=SHEETS()", &c), Variant::Integer(1));
        assert_eq!(calc("=SHEETS(A1:A3)", &c), Variant::Integer(1));
        assert_eq!(calc("=ISOMITTED(42)", &c), Variant::Boolean(false));
        assert_eq!(calc("=PERCENTOF(A1:A2,A1:A3)", &c), Variant::Float(0.5));
        assert_eq!(
            with_sheet_context(3, 5, || calc("=SHEET()", &c)),
            Variant::Integer(3)
        );
        assert_eq!(
            with_sheet_context(3, 5, || calc("=SHEETS()", &c)),
            Variant::Integer(5)
        );
        let mut formula_cells = c.clone();
        formula_cells.insert(
            (2, 1),
            CellContent {
                formula: Some("=A1*2".into()),
                value: Variant::Integer(2),
            },
        );
        assert_eq!(
            calc("=ISFORMULA(A1)", &formula_cells),
            Variant::Boolean(false)
        );
        assert_eq!(
            calc("=ISFORMULA(A2)", &formula_cells),
            Variant::Boolean(true)
        );
        assert_eq!(
            calc("=ISFORMULA(A1:A2)", &formula_cells),
            Variant::Array(vec![Variant::Boolean(false), Variant::Boolean(true)])
        );
    }

    #[test]
    fn test_getpivotdata_on_bounded_rendered_pivot_grid() {
        let c = cells_from(&[
            ((1, 1), Variant::Str("Region".into())),
            ((1, 2), Variant::Str("Sales".into())),
            ((2, 1), Variant::Str("East".into())),
            ((2, 2), Variant::Integer(10)),
            ((3, 1), Variant::Str("West".into())),
            ((3, 2), Variant::Integer(20)),
            ((4, 1), Variant::Str("Grand Total".into())),
            ((4, 2), Variant::Integer(30)),
        ]);
        assert_eq!(
            calc("=GETPIVOTDATA(\"Sales\",A1,\"Region\",\"East\")", &c),
            Variant::Integer(10)
        );
        assert_eq!(
            calc("=GETPIVOTDATA(\"Sales\",A1)", &c),
            Variant::Integer(30)
        );
        assert_eq!(
            calc("=GETPIVOTDATA(\"Missing\",A1)", &c),
            Variant::Error(ExcelError::Ref)
        );

        let c = cells_from(&[
            ((1, 1), Variant::Str("Sum of Sales".into())),
            ((1, 2), Variant::Str("Region".into())),
            ((1, 3), Variant::Str("East".into())),
            ((1, 4), Variant::Str("West".into())),
            ((1, 5), Variant::Str("Grand Total".into())),
            ((2, 1), Variant::Str("Product".into())),
            ((3, 1), Variant::Str("A".into())),
            ((3, 3), Variant::Integer(12)),
            ((3, 4), Variant::Integer(8)),
            ((3, 5), Variant::Integer(20)),
            ((4, 1), Variant::Str("Grand Total".into())),
            ((4, 3), Variant::Integer(30)),
            ((4, 4), Variant::Integer(25)),
            ((4, 5), Variant::Integer(55)),
        ]);
        assert_eq!(
            calc("=GETPIVOTDATA(\"Sales\",A1,\"Region\",\"East\")", &c),
            Variant::Integer(30)
        );
        assert_eq!(
            calc("=GETPIVOTDATA(\"Sales\",A1,\"Region\",\"West\")", &c),
            Variant::Integer(25)
        );
        assert_eq!(
            calc("=GETPIVOTDATA(\"Sales\",A1)", &c),
            Variant::Integer(55)
        );
    }

    #[test]
    fn test_extended_trigonometry() {
        let c = HashMap::new();
        assert_eq!(calc("=SINH(0)", &c), Variant::Integer(0));
        assert_eq!(calc("=COSH(0)", &c), Variant::Integer(1));
        assert_eq!(calc("=TANH(0)", &c), Variant::Integer(0));
        assert_eq!(calc("=SECH(0)", &c), Variant::Integer(1));
        match calc("=ACOT(1)", &c) {
            Variant::Float(value) => assert!((value - std::f64::consts::FRAC_PI_4).abs() < 1e-12),
            other => panic!("ACOT: {:?}", other),
        }
        match calc("=ACOTH(2)", &c) {
            Variant::Float(value) => assert!((value - 0.5493061443340549).abs() < 1e-12),
            other => panic!("ACOTH: {:?}", other),
        }
        assert_eq!(calc("=ACOSH(0.5)", &c), Variant::Error(ExcelError::Num));
        assert_eq!(calc("=ATANH(1)", &c), Variant::Error(ExcelError::Num));
        assert_eq!(calc("=COT(0)", &c), Variant::Error(ExcelError::DivZero));
        assert_eq!(calc("=COTH(0)", &c), Variant::Error(ExcelError::DivZero));
    }

    #[test]
    fn test_combinatorics_and_pairwise_sums() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(3)),
            ((1, 2), Variant::Integer(3)),
            ((2, 2), Variant::Integer(2)),
            ((3, 2), Variant::Integer(1)),
        ]);
        assert_eq!(calc("=COMBINA(5,2)", &c), Variant::Integer(15));
        assert_eq!(calc("=FACTDOUBLE(7)", &c), Variant::Integer(105));
        assert_eq!(calc("=MULTINOMIAL(2,3,1)", &c), Variant::Integer(60));
        assert_eq!(calc("=SUMX2MY2(A1:A3,B1:B3)", &c), Variant::Integer(0));
        assert_eq!(calc("=SUMX2PY2(A1:A3,B1:B3)", &c), Variant::Integer(28));
        assert_eq!(calc("=SUMXMY2(A1:A3,B1:B3)", &c), Variant::Integer(8));
    }

    #[test]
    fn test_rounding_family() {
        let c = HashMap::new();
        assert_eq!(calc("=EVEN(3.1)", &c), Variant::Integer(4));
        assert_eq!(calc("=EVEN(-3.1)", &c), Variant::Integer(-4));
        assert_eq!(calc("=ODD(2.1)", &c), Variant::Integer(3));
        assert_eq!(calc("=ODD(-2.1)", &c), Variant::Integer(-3));
        assert_eq!(calc("=CEILING.PRECISE(-4.3,2)", &c), Variant::Integer(-4));
        assert_eq!(calc("=FLOOR.PRECISE(-4.3,2)", &c), Variant::Integer(-6));
        assert_eq!(calc("=ISO.CEILING(4.1,2)", &c), Variant::Integer(6));
        assert_eq!(
            calc("=ECMA.CEILING(-4.3,2)", &c),
            calc("=CEILING.PRECISE(-4.3,2)", &c)
        );
        assert_eq!(calc("=ISO.FLOOR(4.1,2)", &c), Variant::Integer(4));
    }

    #[test]
    fn test_cell_ref() {
        let c = cells_from(&[((1, 1), Variant::Integer(42))]);
        assert_eq!(calc("=A1", &c), Variant::Integer(42));
        assert_eq!(calc("=B1", &c), Variant::Empty);
    }

    // ── 0.14.0-A2: cross-sheet references parse but never silently evaluate ──

    #[test]
    fn evaluating_a_qualified_cell_ref_is_an_explicit_error_not_a_wrong_answer() {
        // Sheet2!A1 must never be silently read as if it were this sheet's
        // own A1 -- `cells` here has no sheet dimension at all, so reading it
        // that way would be a real correctness bug, not a missing feature.
        let c = cells_from(&[((1, 1), Variant::Integer(999))]);
        let expr = fparse("=Sheet2!A1").unwrap();
        let err = evaluate(&expr, &c).unwrap_err();
        assert!(err.contains("cross-sheet"), "unexpected error text: {err}");
    }

    #[test]
    fn a_qualified_ref_nested_inside_a_function_call_still_trips_the_guard() {
        let c = HashMap::new();
        let expr = fparse("=SUM(Sheet2!A1:A3)").unwrap();
        assert!(evaluate(&expr, &c).is_err());
    }

    #[test]
    fn a_qualified_ref_nested_inside_a_binop_still_trips_the_guard() {
        let c = HashMap::new();
        let expr = fparse("=1+Sheet2!A1").unwrap();
        assert!(evaluate(&expr, &c).is_err());
    }

    #[test]
    fn a_formula_with_no_qualified_ref_anywhere_still_evaluates_normally() {
        // Regression: the guard must not over-trigger on ordinary formulas.
        let c = cells_from(&[((1, 1), Variant::Integer(1)), ((2, 1), Variant::Integer(2))]);
        assert_eq!(calc("=SUM(A1:A2)+1", &c), Variant::Integer(4));
    }

    #[test]
    fn test_sum() {
        // (row, col): A1=(1,1), A2=(2,1), A3=(3,1)
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(3)),
        ]);
        assert_eq!(calc("=SUM(A1:A3)", &c), Variant::Integer(6));
        assert_eq!(calc("=SUM(SEQUENCE(2,3))", &c), Variant::Integer(21));
        assert_eq!(calc("=SUM(SEQUENCE(2,3)+1)", &c), Variant::Integer(27));
        assert_eq!(calc("=SUM(-SEQUENCE(2,3))", &c), Variant::Integer(-21));
        assert_eq!(calc("=SUM(-SEQUENCE(2,3)+1)", &c), Variant::Integer(-15));
        assert_eq!(
            calc("=SUM(SEQUENCE(2,1)+SEQUENCE(1,2))", &c),
            Variant::Integer(12)
        );
        assert_eq!(
            calc("=SUM(SEQUENCE(2)+SEQUENCE(3))", &c),
            Variant::Error(ExcelError::Value)
        );
        let with_error = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Error(ExcelError::DivZero)),
        ]);
        assert_eq!(
            calc("=SUM(A1:A2)", &with_error),
            Variant::Error(ExcelError::DivZero)
        );
        assert_eq!(calc("=SUM(1/0,2)", &c), Variant::Error(ExcelError::DivZero));
        assert_eq!(
            calc("=SUM(CHOOSE(SEQUENCE(2),1,1/0))", &c),
            Variant::Error(ExcelError::DivZero)
        );
        assert_eq!(
            calc("=SUM(SEQUENCE(2,3)+1/0)", &c),
            Variant::Error(ExcelError::DivZero)
        );
        assert_eq!(
            calc("=SUM(IF(SEQUENCE(2),1,1/0))", &c),
            Variant::Integer(2)
        );
    }

    #[test]
    fn test_average() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(20)),
            ((3, 1), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=AVERAGE(A1:A3)", &c), Variant::Float(20.0));
        assert_eq!(
            calc("=AVERAGE(CHOOSE(SEQUENCE(2),1,1/0))", &c),
            Variant::Error(ExcelError::DivZero)
        );
        assert_eq!(calc("=MIN(\"1\",TRUE)", &c), Variant::Integer(1));
        assert_eq!(calc("=MAX(\"1\",TRUE)", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNT(\"1\")", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNT(TRUE)", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNT(\"x\")", &c), Variant::Integer(0));
        let with_error = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Error(ExcelError::NA)),
        ]);
        assert_eq!(
            calc("=AVERAGE(A1:A2)", &with_error),
            Variant::Error(ExcelError::NA)
        );
        match calc("=AVERAGEA(1,TRUE,\"x\")", &c) {
            Variant::Float(value) => assert!((value - 2.0 / 3.0).abs() < 1e-12),
            other => panic!("AVERAGEA: {:?}", other),
        }
        assert_eq!(calc("=MINA(1,TRUE,\"x\")", &c), Variant::Integer(0));
        assert_eq!(calc("=MAXA(1,TRUE,\"x\")", &c), Variant::Integer(1));
        assert_eq!(
            calc("=AVERAGEA(CHOOSE(SEQUENCE(2),1,2))", &c),
            Variant::Float(1.5)
        );
        assert_eq!(
            calc("=MINA(CHOOSE(SEQUENCE(2),3,1))", &c),
            Variant::Integer(1)
        );
        assert_eq!(
            calc("=MAXA(CHOOSE(SEQUENCE(2),3,1))", &c),
            Variant::Integer(3)
        );
        assert_eq!(
            calc("=AVERAGEA(CHOOSE(SEQUENCE(2),1,1/0))", &c),
            Variant::Error(ExcelError::DivZero)
        );
        assert!(matches!(
            calc("=VARA(1,2,TRUE,\"x\")", &c),
            Variant::Float(value) if (value - (2.0 / 3.0)).abs() < 1e-12
        ));
        assert!(matches!(
            calc("=VARPA(1,2,TRUE,\"x\")", &c),
            Variant::Float(value) if (value - 0.5).abs() < 1e-12
        ));
        assert_eq!(calc("=ISEVEN(3.9)", &c), Variant::Boolean(false));
        assert_eq!(calc("=ISODD(-3.9)", &c), Variant::Boolean(true));
    }

    #[test]
    fn test_min_max() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(5)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(8)),
        ]);
        assert_eq!(calc("=MIN(A1:A3)", &c), Variant::Integer(2));
        assert_eq!(calc("=MAX(A1:A3)", &c), Variant::Integer(8));
        let with_error = cells_from(&[
            ((1, 1), Variant::Integer(5)),
            ((2, 1), Variant::Error(ExcelError::Value)),
        ]);
        assert_eq!(
            calc("=MIN(A1:A2)", &with_error),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=MAX(A1:A2)", &with_error),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_count_counta() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Str("hi".into())),
            // (3,1) is absent → Empty by default
        ]);
        assert_eq!(calc("=COUNT(A1:A3)", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNTA(A1:A3)", &c), Variant::Integer(2));
        assert_eq!(calc("=COUNT(\"1\")", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNT(TRUE)", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNT(\"x\")", &c), Variant::Integer(0));
        assert_eq!(calc("=COUNTA(\"\")", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNTA(1/0)", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNT(SEQUENCE(2,3))", &c), Variant::Integer(6));
        assert_eq!(calc("=COUNTA(SEQUENCE(2,3))", &c), Variant::Integer(6));
        assert_eq!(calc("=COUNT(SEQUENCE(3))", &c), Variant::Integer(3));
        assert_eq!(
            calc("=COUNTA(CHOOSE(SEQUENCE(2),\"x\",\"\"))", &c),
            Variant::Integer(2)
        );
        assert_eq!(
            calc("=COUNT(CHOOSE(SEQUENCE(2),1,1/0))", &c),
            Variant::Integer(1)
        );
        assert_eq!(
            calc("=COUNTA(CHOOSE(SEQUENCE(2),1,1/0))", &c),
            Variant::Integer(2)
        );
    }

    #[test]
    fn test_if() {
        let c = cells_from(&[((1, 1), Variant::Integer(5))]);
        assert_eq!(
            calc("=IF(A1>3,\"yes\",\"no\")", &c),
            Variant::Str("yes".into())
        );
        assert_eq!(
            calc("=IF(A1>10,\"yes\",\"no\")", &c),
            Variant::Str("no".into())
        );
        assert_eq!(
            calc("=IF(1/0,\"yes\",\"no\")", &c),
            Variant::Error(ExcelError::DivZero)
        );
    }

    #[test]
    fn test_unit_and_roman_conversions() {
        let c = HashMap::new();
        assert_eq!(calc("=CONVERT(1,\"km\",\"m\")", &c), Variant::Integer(1000));
        assert_eq!(calc("=CONVERT(32,\"F\",\"C\")", &c), Variant::Integer(0));
        assert!(matches!(
            calc("=CONVERT(1,\"kWh\",\"J\")", &c),
            Variant::Integer(3600000)
        ));
        assert!(matches!(
            calc("=CONVERT(1,\"N\",\"dyn\")", &c),
            Variant::Float(value) if (value - 100000.0).abs() < 1e-9
        ));
        assert_eq!(
            calc("=CONVERT(1,\"GHz\",\"MHz\")", &c),
            Variant::Integer(1000)
        );
        assert_eq!(
            calc("=CONVERT(1,\"byte\",\"bit\")", &c),
            Variant::Integer(8)
        );
        assert!(matches!(
            calc("=CONVERT(1,\"m\",\"N\")", &c),
            Variant::Error(ExcelError::Num)
        ));
        assert_eq!(
            calc("=CONVERT(1,\"m\",\"s\")", &c),
            Variant::Error(ExcelError::Num)
        );
        assert_eq!(calc("=ROMAN(1999)", &c), Variant::Str("MCMXCIX".into()));
        assert_eq!(calc("=ARABIC(\"mcmxcix\")", &c), Variant::Integer(1999));
        assert_eq!(calc("=ROMAN(4000)", &c), Variant::Error(ExcelError::Value));
        assert_eq!(
            calc("=ARABIC(\"IC\")", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_and_or_not() {
        let c = HashMap::new();
        assert_eq!(calc("=TRUE()", &c), Variant::Boolean(true));
        assert_eq!(calc("=FALSE()", &c), Variant::Boolean(false));
        assert_eq!(calc("=AND(TRUE,TRUE)", &c), Variant::Boolean(true));
        assert_eq!(calc("=AND(TRUE,FALSE)", &c), Variant::Boolean(false));
        assert_eq!(calc("=OR(FALSE,TRUE)", &c), Variant::Boolean(true));
        assert_eq!(calc("=NOT(TRUE)", &c), Variant::Boolean(false));
        assert_eq!(
            calc("=AND(CHOOSE(SEQUENCE(2),TRUE,FALSE),TRUE)", &c),
            Variant::Boolean(false)
        );
        assert_eq!(
            calc("=OR(CHOOSE(SEQUENCE(2),FALSE,TRUE),FALSE)", &c),
            Variant::Boolean(true)
        );
        assert_eq!(
            calc("=NOT(CHOOSE(SEQUENCE(2),TRUE,FALSE))", &c),
            Variant::Array(vec![Variant::Boolean(false), Variant::Boolean(true)])
        );
        assert_eq!(
            calc("=XOR(CHOOSE(SEQUENCE(3),TRUE,FALSE,TRUE),FALSE)", &c),
            Variant::Boolean(false)
        );
        assert_eq!(
            calc("=IFERROR(NOT(CHOOSE(SEQUENCE(2),1/0,TRUE)),99)", &c),
            Variant::Array(vec![Variant::Integer(99), Variant::Boolean(false)])
        );
        assert_eq!(calc("=NOT(1/0)", &c), Variant::Error(ExcelError::DivZero));
        assert_eq!(
            calc("=AND(FALSE,1/0)", &c),
            Variant::Error(ExcelError::DivZero)
        );
        assert_eq!(
            calc("=OR(TRUE,1/0)", &c),
            Variant::Error(ExcelError::DivZero)
        );
        assert_eq!(
            calc("=IFERROR(AND(FALSE,1/0),99)", &c),
            Variant::Integer(99)
        );
    }

    #[test]
    fn test_detectlanguage_is_local_and_conservative() {
        let c = HashMap::new();
        assert_eq!(
            calc("=DETECTLANGUAGE(\"Hello world and the test\")", &c),
            Variant::Str("en".into())
        );
        assert_eq!(
            calc("=DETECTLANGUAGE(\"こんにちは世界\")", &c),
            Variant::Str("ja".into())
        );
        assert_eq!(
            calc("=DETECTLANGUAGE(\"12345\")", &c),
            Variant::Error(ExcelError::NA)
        );
        assert_eq!(
            calc("=DETECTLANGUAGE(42)", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_iferror() {
        let c = HashMap::new();
        assert_eq!(calc("=IFERROR(1/0,99)", &c), Variant::Integer(99));
        assert_eq!(calc("=IFERROR(10,99)", &c), Variant::Integer(10));
    }

    #[test]
    fn test_ifna_only_handles_na_and_keeps_other_errors() {
        let c = HashMap::new();
        assert_eq!(calc("=IFNA(NA(),99)", &c), Variant::Integer(99));
        assert_eq!(
            calc("=IFNA(1/0,99)", &c),
            Variant::Error(ExcelError::DivZero)
        );
        assert_eq!(calc("=IFNA(10,99)", &c), Variant::Integer(10));
    }

    #[test]
    fn test_iferror_and_ifna_map_array_errors_elementwise() {
        let c = HashMap::new();
        assert_eq!(
            calc("=IFERROR(CHOOSE(SEQUENCE(2),1/0,NA()),99)", &c),
            Variant::Array(vec![Variant::Integer(99), Variant::Integer(99)])
        );
        assert_eq!(
            calc("=IFNA(CHOOSE(SEQUENCE(2),1/0,NA()),99)", &c),
            Variant::Array(vec![
                Variant::Error(ExcelError::DivZero),
                Variant::Integer(99)
            ])
        );
    }

    #[test]
    fn test_string_functions() {
        let c = HashMap::new();
        assert_eq!(calc("=LEFT(\"hello\",3)", &c), Variant::Str("hel".into()));
        assert_eq!(calc("=RIGHT(\"hello\",2)", &c), Variant::Str("lo".into()));
        assert_eq!(calc("=MID(\"hello\",2,3)", &c), Variant::Str("ell".into()));
        assert_eq!(calc("=LEN(\"hello\")", &c), Variant::Integer(5));
        assert_eq!(
            calc("=CONCATENATE(\"A\",\"B\",\"C\")", &c),
            Variant::Str("ABC".into())
        );
        assert_eq!(
            calc("=HYPERLINK(\"https://example.test\",\"open\")", &c),
            Variant::Str("open".into())
        );
        assert_eq!(
            calc("=HYPERLINK(\"Sheet1!A1\")", &c),
            Variant::Str("Sheet1!A1".into())
        );
        assert_eq!(calc("=PHONETIC(\"東京\")", &c), Variant::Str("東京".into()));
        assert_eq!(calc("=PHONETIC(42)", &c), Variant::Error(ExcelError::Value));
        assert_eq!(
            calc("=ENCODEURL(\"hello world/a?x=1\")", &c),
            Variant::Str("hello%20world%2Fa%3Fx%3D1".into())
        );
        assert_eq!(
            calc("=ENCODEURL(\"東京\")", &c),
            Variant::Str("%E6%9D%B1%E4%BA%AC".into())
        );
        assert_eq!(
            calc("=REGEXTEST(\"Invoice-42\",\"[0-9]+\")", &c),
            Variant::Boolean(true)
        );
        assert_eq!(
            calc("=REGEXTEST(\"Invoice-AB\",\"[0-9]+\")", &c),
            Variant::Boolean(false)
        );
        assert_eq!(
            calc(
                "=REGEXEXTRACT(\"Invoice-42\",\"([A-Za-z]+)-([0-9]+)\",2)",
                &c
            ),
            Variant::Array(vec![
                Variant::Str("Invoice".into()),
                Variant::Str("42".into())
            ])
        );
        assert_eq!(
            calc("=REGEXREPLACE(\"a1 b22\",\"[0-9]+\",\"#\")", &c),
            Variant::Str("a# b#".into())
        );
        assert_eq!(
            calc("=ARRAYTOTEXT(SEQUENCE(3))", &c),
            Variant::Str("1, 2, 3".into())
        );
        assert_eq!(
            calc("=ARRAYTOTEXT(TEXTSPLIT(\"a,b\",\",\"),1)", &c),
            Variant::Str("{\"a\",\"b\"}".into())
        );
    }

    #[test]
    fn test_concat_operator() {
        let c = HashMap::new();
        assert_eq!(
            calc("=\"Hello\"&\" \"&\"World\"", &c),
            Variant::Str("Hello World".into())
        );
    }

    #[test]
    fn test_text_format() {
        let c = HashMap::new();
        assert_eq!(
            calc("=TEXT(3.14159,\"0.00\")", &c),
            Variant::Str("3.14".into())
        );
        assert_eq!(calc("=TEXT(0.5,\"0%\")", &c), Variant::Str("50%".into()));
    }

    #[test]
    fn test_vlookup_exact() {
        // (row, col): A1=(1,1), B1=(1,2), A2=(2,1), B2=(2,2), A3=(3,1), B3=(3,2)
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Str("one".into())),
            ((2, 1), Variant::Integer(2)),
            ((2, 2), Variant::Str("two".into())),
            ((3, 1), Variant::Integer(3)),
            ((3, 2), Variant::Str("three".into())),
        ]);
        assert_eq!(
            calc("=VLOOKUP(2,A1:B3,2,FALSE)", &c),
            Variant::Str("two".into())
        );
    }

    #[test]
    fn test_index() {
        // (row, col): A1=(1,1)=10, B1=(1,2)=20, A2=(2,1)=30, B2=(2,2)=40
        let c = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((1, 2), Variant::Integer(20)),
            ((2, 1), Variant::Integer(30)),
            ((2, 2), Variant::Integer(40)),
        ]);
        // INDEX(A1:B2, 2, 1) = row 2 col 1 of range = A2 = 30
        assert_eq!(calc("=INDEX(A1:B2,2,1)", &c), Variant::Integer(30));
        assert_eq!(
            calc("=INDEX(SEQUENCE(2,2),2,2)", &c),
            Variant::Integer(4)
        );
        assert_eq!(
            calc("=INDEX(A1:B2,-1,1)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=INDEX(A1:B2,3,1)", &c),
            Variant::Error(ExcelError::Ref)
        );
        assert_eq!(
            calc("=INDEX(A1:B2,1,3)", &c),
            Variant::Error(ExcelError::Ref)
        );
        assert_eq!(
            calc("=INDEX(A1:B2,0,2)", &c),
            Variant::Array(vec![Variant::Integer(20), Variant::Integer(40)])
        );
        assert_eq!(
            calc("=INDEX(A1:B2,2,0)", &c),
            Variant::Array(vec![Variant::Integer(30), Variant::Integer(40)])
        );
        assert_eq!(
            calc("=INDEX(A1:B2,0,0)", &c),
            Variant::Array(vec![
                Variant::Integer(10),
                Variant::Integer(20),
                Variant::Integer(30),
                Variant::Integer(40),
            ])
        );
        assert_eq!(
            calc("=INDEX(A1:B2,1.5,1)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=INDEX(A1:B2,1,1.5)", &c),
            Variant::Error(ExcelError::Value)
        );
        // Bounded reference-form support: this parser has one explicit range
        // per INDEX call, so area 1 is valid and other areas fail closed.
        assert_eq!(calc("=INDEX(A1:B2,2,2,1)", &c), Variant::Integer(40));
        assert_eq!(
            calc("=INDEX(A1:B2,1,1,2)", &c),
            Variant::Error(ExcelError::Ref)
        );
        assert_eq!(
            calc("=INDEX(A1:B2,1,1,1.5)", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_match_exact() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(20)),
            ((3, 1), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=MATCH(20,A1:A3,0)", &c), Variant::Integer(2));
        let text = cells_from(&[
            ((1, 1), Variant::Str("Alpha".into())),
            ((2, 1), Variant::Str("Beta".into())),
        ]);
        assert_eq!(calc("=MATCH(\"a*\",A1:A2,0)", &text), Variant::Integer(1));
        let literal = cells_from(&[
            ((1, 1), Variant::Str("rate*".into())),
            ((2, 1), Variant::Str("rate?".into())),
            ((3, 1), Variant::Str("rate~".into())),
        ]);
        assert_eq!(
            calc("=MATCH(\"rate~*\",A1:A3,0)", &literal),
            Variant::Integer(1)
        );
        assert_eq!(
            calc("=MATCH(\"rate~?\",A1:A3,0)", &literal),
            Variant::Integer(2)
        );
        assert_eq!(
            calc("=MATCH(\"rate~~\",A1:A3,0)", &literal),
            Variant::Integer(3)
        );
        assert_eq!(
            calc("=MATCH(\"Alpha\",A1:A2,9)", &text),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=MATCH(\"Alpha\",A1:A2,0.5)", &text),
            Variant::Error(ExcelError::Value)
        );

        let ascending = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(20)),
            ((3, 1), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=MATCH(25,A1:A3,1)", &ascending), Variant::Integer(2));
        assert_eq!(
            calc("=MATCH(25,A1:A3,-1)", &ascending),
            Variant::Error(ExcelError::NA)
        );

        let descending = cells_from(&[
            ((1, 1), Variant::Integer(30)),
            ((2, 1), Variant::Integer(20)),
            ((3, 1), Variant::Integer(10)),
        ]);
        assert_eq!(
            calc("=MATCH(25,A1:A3,-1)", &descending),
            Variant::Integer(1)
        );
        assert_eq!(
            calc("=MATCH(25,A1:A3,1)", &descending),
            Variant::Error(ExcelError::NA)
        );

        let incomparable = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Error(ExcelError::DivZero)),
        ]);
        assert_eq!(
            calc("=MATCH(20,A1:A2,1)", &incomparable),
            Variant::Error(ExcelError::Value)
        );

        let matrix = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Integer(2)),
            ((2, 1), Variant::Integer(3)),
            ((2, 2), Variant::Integer(4)),
        ]);
        assert_eq!(
            calc("=MATCH(2,A1:B2,0)", &matrix),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_lenb() {
        let c = HashMap::new();
        assert_eq!(calc("=LENB(\"ABC\")", &c), Variant::Integer(3));
        assert_eq!(calc("=LENB(\"日本語\")", &c), Variant::Integer(6));
        assert_eq!(calc("=LENB(\"A日\")", &c), Variant::Integer(3));
        assert_eq!(calc("=LENB(\"\")", &c), Variant::Integer(0));
    }

    #[test]
    fn test_leftb() {
        let c = HashMap::new();
        assert_eq!(calc("=LEFTB(\"ABC\",2)", &c), Variant::Str("AB".into()));
        assert_eq!(
            calc("=LEFTB(\"日本語\",4)", &c),
            Variant::Str("日本".into())
        );
        assert_eq!(calc("=LEFTB(\"日本語\",3)", &c), Variant::Str("日".into()));
        assert_eq!(calc("=LEFTB(\"ABC\",0)", &c), Variant::Str("".into()));
    }

    #[test]
    fn test_rightb() {
        let c = HashMap::new();
        assert_eq!(calc("=RIGHTB(\"ABC\",2)", &c), Variant::Str("BC".into()));
        assert_eq!(
            calc("=RIGHTB(\"日本語\",4)", &c),
            Variant::Str("本語".into())
        );
        assert_eq!(calc("=RIGHTB(\"日本語\",3)", &c), Variant::Str("語".into()));
        assert_eq!(calc("=RIGHTB(\"ABC\",0)", &c), Variant::Str("".into()));
    }

    #[test]
    fn test_midb() {
        let c = HashMap::new();
        assert_eq!(calc("=MIDB(\"日本語\",3,2)", &c), Variant::Str("本".into()));
        assert_eq!(calc("=MIDB(\"ABC\",2,2)", &c), Variant::Str("BC".into()));
        assert_eq!(
            calc("=MIDB(\"日本語\",1,4)", &c),
            Variant::Str("日本".into())
        );
    }

    #[test]
    fn test_round() {
        let c = HashMap::new();
        assert_eq!(calc("=ROUND(2.5,0)", &c), Variant::Integer(3));
        assert_eq!(calc("=ROUND(-2.5,0)", &c), Variant::Integer(-3));
        assert_eq!(calc("=ROUND(2.15,1)", &c), Variant::Float(2.2));
        assert_eq!(calc("=ROUND(1234,-2)", &c), Variant::Integer(1200));
        assert_eq!(calc("=ROUND(3.0,0)", &c), Variant::Integer(3));
    }

    #[test]
    fn test_roundup() {
        let c = HashMap::new();
        assert_eq!(calc("=ROUNDUP(2.1,0)", &c), Variant::Integer(3));
        assert_eq!(calc("=ROUNDUP(-2.1,0)", &c), Variant::Integer(-3));
        assert_eq!(calc("=ROUNDUP(2.0,0)", &c), Variant::Integer(2));
        assert_eq!(calc("=ROUNDUP(1.23,1)", &c), Variant::Float(1.3));
    }

    #[test]
    fn test_rounddown() {
        let c = HashMap::new();
        assert_eq!(calc("=ROUNDDOWN(2.9,0)", &c), Variant::Integer(2));
        assert_eq!(calc("=ROUNDDOWN(-2.9,0)", &c), Variant::Integer(-2));
        assert_eq!(calc("=ROUNDDOWN(1.99,1)", &c), Variant::Float(1.9));
        assert_eq!(calc("=ROUNDDOWN(1234,-2)", &c), Variant::Integer(1200));
    }

    #[test]
    fn test_countif() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(20)),
            ((3, 1), Variant::Integer(10)),
            ((4, 1), Variant::Str("apple".into())),
        ]);
        assert_eq!(calc("=COUNTIF(A1:A3,10)", &c), Variant::Integer(2));
        assert_eq!(calc("=COUNTIF(A1:A3,\">10\")", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNTIF(A1:A4,\"apple\")", &c), Variant::Integer(1));
        assert_eq!(
            calc("=COUNTIF(SEQUENCE(3),\">1\")", &c),
            Variant::Integer(2)
        );
    }

    #[test]
    fn test_sumif() {
        let c = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((1, 2), Variant::Integer(10)),
            ((2, 1), Variant::Str("b".into())),
            ((2, 2), Variant::Integer(20)),
            ((3, 1), Variant::Str("a".into())),
            ((3, 2), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=SUMIF(A1:A3,\"a\",B1:B3)", &c), Variant::Integer(40));
        assert_eq!(calc("=SUMIF(B1:B3,\">10\")", &c), Variant::Integer(50));
        assert_eq!(
            calc("=SUMIF(SEQUENCE(3),\">1\",SEQUENCE(3))", &c),
            Variant::Integer(5)
        );
        let with_error = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((1, 2), Variant::Error(ExcelError::DivZero)),
        ]);
        assert_eq!(
            calc("=SUMIF(A1:A1,\"a\",B1:B1)", &with_error),
            Variant::Error(ExcelError::DivZero)
        );
    }

    #[test]
    fn test_sumifs_countifs() {
        let c = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((1, 2), Variant::Integer(10)),
            ((2, 1), Variant::Str("b".into())),
            ((2, 2), Variant::Integer(20)),
            ((3, 1), Variant::Str("a".into())),
            ((3, 2), Variant::Integer(30)),
        ]);
        assert_eq!(
            calc("=SUMIFS(B1:B3,A1:A3,\"a\",B1:B3,\">10\")", &c),
            Variant::Integer(30)
        );
        assert_eq!(
            calc("=SUMIFS(SEQUENCE(3),SEQUENCE(3),\">1\")", &c),
            Variant::Integer(5)
        );
        let with_error = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((1, 2), Variant::Error(ExcelError::NA)),
        ]);
        assert_eq!(
            calc("=SUMIFS(B1:B1,A1:A1,\"a\")", &with_error),
            Variant::Error(ExcelError::NA)
        );
        assert_eq!(
            calc("=COUNTIFS(A1:A3,\"a\",B1:B3,\">10\")", &c),
            Variant::Integer(1)
        );
        // Equal element counts are not enough: Excel requires matching
        // rectangular dimensions for every criteria range.
        assert_eq!(
            calc("=SUMIFS(B1:C2,A1:A4,\"a\")", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=COUNTIFS(A1:B2,\"a\",A1:A4,\"a\")", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=AVERAGEIFS(B1:B3,A1:B2,\"a\")", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=MAXIFS(B1:B3,A1:B2,\"a\")", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=MINIFS(B1:B3,A1:B2,\"a\")", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_median() {
        let c = HashMap::new();
        assert_eq!(calc("=MEDIAN(1,3,2)", &c), Variant::Integer(2));
        assert_eq!(calc("=MEDIAN(1,2,3,4)", &c), Variant::Float(2.5));
        assert_eq!(
            calc("=MEDIAN(CHOOSE(SEQUENCE(2),1,1/0))", &c),
            Variant::Error(ExcelError::DivZero)
        );
    }

    #[test]
    fn test_mode_mult() {
        let c = HashMap::new();
        assert_eq!(
            calc("=MODE.MULT(1,2,2,3,3)", &c),
            Variant::Array(vec![Variant::Integer(2), Variant::Integer(3)])
        );
        assert_eq!(calc("=MODE.SNGL(1,2,2,3)", &c), Variant::Integer(2));
    }

    #[test]
    fn test_product() {
        let c = HashMap::new();
        assert_eq!(calc("=PRODUCT(2,3,4)", &c), Variant::Integer(24));
        assert_eq!(calc("=PRODUCT(\"2\",TRUE)", &c), Variant::Integer(2));
        assert_eq!(
            calc("=PRODUCT(2,1/0)", &c),
            Variant::Error(ExcelError::DivZero)
        );
    }

    #[test]
    fn test_row() {
        let c = HashMap::new();
        assert_eq!(calc("=ROW(A5)", &c), Variant::Integer(5));
        assert_eq!(calc("=ROW()", &c), Variant::Integer(1));
        assert_eq!(calc("=ROW(B3:C7)", &c), Variant::Integer(3));
        assert_eq!(calc("=ROWS(B3:C7)", &c), Variant::Integer(5));
        assert_eq!(calc("=ROWS(B3)", &c), Variant::Integer(1));
        assert_eq!(calc("=ROWS(SEQUENCE(3))", &c), Variant::Integer(3));
        assert_eq!(calc("=ROWS(SEQUENCE(2,3))", &c), Variant::Integer(2));
        assert_eq!(calc("=COLUMNS(SEQUENCE(2,3))", &c), Variant::Integer(3));
        assert_eq!(calc("=ROWS(TRANSPOSE(B3:C7))", &c), Variant::Integer(2));
    }

    #[test]
    fn test_date_serial() {
        let c = HashMap::new();
        // Jan 1 1900 = 1
        assert_eq!(calc("=DATE(1900,1,1)", &c), Variant::Date(1));
        // Jan 1 2000 = 36526
        assert_eq!(calc("=DATE(2000,1,1)", &c), Variant::Date(36526));
        // Excel normalizes month/day overflow and maps a two-digit year.
        assert_eq!(
            calc("=DATE(2020,0,1)", &c),
            Variant::Date(date_to_serial(2019, 12, 1))
        );
        assert_eq!(
            calc("=DATE(2020,13,1)", &c),
            Variant::Date(date_to_serial(2021, 1, 1))
        );
        assert_eq!(
            calc("=DATE(2020,1,0)", &c),
            Variant::Date(date_to_serial(2019, 12, 31))
        );
        assert_eq!(
            calc("=DATE(20,1,1)", &c),
            Variant::Date(date_to_serial(1920, 1, 1))
        );
        assert_eq!(
            calc("=DATE(2020.9,1.9,1.9)", &c),
            Variant::Date(date_to_serial(2020, 1, 1))
        );
    }

    #[test]
    fn test_eomonth() {
        let c = HashMap::new();
        // Jan 1 2000 serial=36526; EOMONTH(36526,0) = Jan 31 2000 = 36556
        assert_eq!(calc("=EOMONTH(DATE(2000,1,1),0)", &c), Variant::Date(36556));
        // EOMONTH(DATE(2000,1,1),1) = Feb 29 2000 (leap) = 36585
        assert_eq!(calc("=EOMONTH(DATE(2000,1,1),1)", &c), Variant::Date(36585));
    }

    #[test]
    fn test_networkdays() {
        let c = HashMap::new();
        // Mon Jan 3 2000 (serial 36528) to Fri Jan 7 2000 (serial 36532) = 5 workdays
        assert_eq!(calc("=NETWORKDAYS(36528,36532)", &c), Variant::Integer(5));
        // Mon Jan 3 to Mon Jan 10 = 6 workdays
        assert_eq!(calc("=NETWORKDAYS(36528,36535)", &c), Variant::Integer(6));
    }

    #[test]
    fn test_rank() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(30)),
            ((3, 1), Variant::Integer(20)),
        ]);
        assert_eq!(calc("=RANK(20,A1:A3,0)", &c), Variant::Integer(2));
        assert_eq!(calc("=RANK(20,A1:A3,1)", &c), Variant::Integer(2));
        assert_eq!(calc("=RANK(30,A1:A3,0)", &c), Variant::Integer(1));
    }

    #[test]
    fn test_ifs() {
        let c = HashMap::new();
        assert_eq!(
            calc("=IFS(FALSE,\"a\",TRUE,\"b\")", &c),
            Variant::Str("b".into())
        );
        assert_eq!(calc("=IFS(TRUE,42,FALSE,99)", &c), Variant::Integer(42));
        assert_eq!(
            calc("=IFS(CHOOSE(SEQUENCE(2),TRUE,FALSE),\"Y\",TRUE,\"N\")", &c),
            Variant::Array(vec![Variant::Str("Y".into()), Variant::Str("N".into())])
        );
        assert_eq!(
            calc(
                "=IFS(CHOOSE(SEQUENCE(2),TRUE,FALSE),SEQUENCE(2),TRUE,0)",
                &c
            ),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(0)])
        );
    }

    #[test]
    fn test_xlookup() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Str("one".into())),
            ((2, 1), Variant::Integer(2)),
            ((2, 2), Variant::Str("two".into())),
            ((3, 1), Variant::Integer(3)),
            ((3, 2), Variant::Str("three".into())),
        ]);
        assert_eq!(
            calc("=XLOOKUP(2,A1:A3,B1:B3)", &c),
            Variant::Str("two".into())
        );
        assert_eq!(
            calc("=XLOOKUP(99,A1:A3,B1:B3,\"N/A\")", &c),
            Variant::Str("N/A".into())
        );
        assert_eq!(
            calc("=XLOOKUP(2,A1:A3,B1:B2,\"missing\")", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_xlookup_wildcard_and_binary_search_modes() {
        let c = cells_from(&[
            ((1, 1), Variant::Str("alpha".into())),
            ((1, 2), Variant::Str("A".into())),
            ((2, 1), Variant::Str("beta".into())),
            ((2, 2), Variant::Str("B".into())),
        ]);
        assert_eq!(
            calc("=XLOOKUP(\"a*\",A1:A2,B1:B2,\"missing\",2)", &c),
            Variant::Str("A".into())
        );
        let ascending = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Str("one".into())),
            ((2, 1), Variant::Integer(2)),
            ((2, 2), Variant::Str("two".into())),
            ((3, 1), Variant::Integer(3)),
            ((3, 2), Variant::Str("three".into())),
        ]);
        assert_eq!(
            calc("=XLOOKUP(2,A1:A3,B1:B3,\"missing\",0,2)", &ascending),
            Variant::Str("two".into())
        );
        assert_eq!(
            calc("=XLOOKUP(0,A1:A3,B1:B3,\"missing\",1,2)", &ascending),
            Variant::Str("one".into())
        );
        let descending = cells_from(&[
            ((1, 1), Variant::Integer(3)),
            ((1, 2), Variant::Str("three".into())),
            ((2, 1), Variant::Integer(2)),
            ((2, 2), Variant::Str("two".into())),
            ((3, 1), Variant::Integer(1)),
            ((3, 2), Variant::Str("one".into())),
        ]);
        assert_eq!(
            calc("=XLOOKUP(2,A1:A3,B1:B3,\"missing\",0,-2)", &descending),
            Variant::Str("two".into())
        );
        assert_eq!(
            calc("=XLOOKUP(4,A1:A3,B1:B3,\"missing\",-1,-2)", &descending),
            Variant::Str("three".into())
        );
        assert_eq!(
            calc("=XLOOKUP(2,A1:A3,B1:B3,\"missing\",0.5)", &ascending),
            Variant::Error(ExcelError::Value)
        );

        let unsorted_ascending = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Str("one".into())),
            ((2, 1), Variant::Integer(3)),
            ((2, 2), Variant::Str("three".into())),
            ((3, 1), Variant::Integer(2)),
            ((3, 2), Variant::Str("two".into())),
        ]);
        assert_eq!(
            calc(
                "=XLOOKUP(2,A1:A3,B1:B3,\"missing\",0,2)",
                &unsorted_ascending
            ),
            Variant::Error(ExcelError::NA)
        );
        assert_eq!(
            calc("=XMATCH(2,A1:A3,0,2)", &unsorted_ascending),
            Variant::Error(ExcelError::NA)
        );

        let unsorted_descending = cells_from(&[
            ((1, 1), Variant::Integer(3)),
            ((1, 2), Variant::Str("three".into())),
            ((2, 1), Variant::Integer(1)),
            ((2, 2), Variant::Str("one".into())),
            ((3, 1), Variant::Integer(2)),
            ((3, 2), Variant::Str("two".into())),
        ]);
        assert_eq!(
            calc(
                "=XLOOKUP(2,A1:A3,B1:B3,\"missing\",0,-2)",
                &unsorted_descending
            ),
            Variant::Error(ExcelError::NA)
        );
        assert_eq!(
            calc("=XMATCH(2,A1:A3,0,-2)", &unsorted_descending),
            Variant::Error(ExcelError::NA)
        );
    }

    #[test]
    fn test_subtotal() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(20)),
            ((3, 1), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=SUBTOTAL(9,A1:A3)", &c), Variant::Integer(60));
        assert_eq!(calc("=SUBTOTAL(1,A1:A3)", &c), Variant::Float(20.0));
        assert_eq!(calc("=SUBTOTAL(109,A1:A3)", &c), Variant::Integer(60));
        assert!(matches!(
            calc("=SUBTOTAL(7,A1:A3)", &c),
            Variant::Float(value) if (value - 10.0).abs() < 1e-12
        ));
        assert!(matches!(
            calc("=SUBTOTAL(108,A1:A3)", &c),
            Variant::Float(value) if (value - (200.0_f64 / 3.0).sqrt()).abs() < 1e-12
        ));
        assert!(matches!(
            calc("=SUBTOTAL(10,A1:A3)", &c),
            Variant::Float(value) if (value - 100.0).abs() < 1e-12
        ));
        assert!(matches!(
            calc("=SUBTOTAL(111,A1:A3)", &c),
            Variant::Float(value) if (value - (200.0 / 3.0)).abs() < 1e-12
        ));
    }

    #[test]
    fn test_subtotal_ignores_nested_aggregates_and_propagates_errors() {
        let mut nested = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(30)),
            ((3, 1), Variant::Integer(5)),
        ]);
        nested.insert(
            (2, 1),
            CellContent {
                formula: Some("=AGGREGATE(9,4,B1:B2)".into()),
                value: Variant::Integer(30),
            },
        );
        assert_eq!(calc("=SUBTOTAL(9,A1:A3)", &nested), Variant::Integer(15));
        assert_eq!(calc("=SUBTOTAL(3,A1:A3)", &nested), Variant::Integer(2));

        let errors = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Error(ExcelError::Value)),
            ((3, 1), Variant::Integer(5)),
        ]);
        assert_eq!(
            calc("=SUBTOTAL(9,A1:A3)", &errors),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_concat() {
        let c = HashMap::new();
        assert_eq!(
            calc("=CONCAT(\"Hello\",\" \",\"World\")", &c),
            Variant::Str("Hello World".into())
        );
    }

    // ── Error variants ────────────────────────────────────────────────────────

    #[test]
    fn test_div_zero_stored() {
        let c = HashMap::new();
        assert_eq!(calc("=1/0", &c), Variant::Error(ExcelError::DivZero));
    }

    #[test]
    fn test_iferror_with_error_variant() {
        let c = HashMap::new();
        assert_eq!(calc("=IFERROR(1/0,99)", &c), Variant::Integer(99));
        assert_eq!(calc("=IFERROR(10,99)", &c), Variant::Integer(10));
    }

    #[test]
    fn test_iserror_div_zero() {
        let c = HashMap::new();
        assert_eq!(calc("=ISERROR(1/0)", &c), Variant::Boolean(true));
        assert_eq!(calc("=ISERROR(1)", &c), Variant::Boolean(false));
    }

    #[test]
    fn test_isna_vlookup() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Str("one".into())),
        ]);
        assert_eq!(
            calc("=ISNA(VLOOKUP(99,A1:B1,2,FALSE))", &c),
            Variant::Boolean(true)
        );
        assert_eq!(
            calc("=ISNA(VLOOKUP(1,A1:B1,2,FALSE))", &c),
            Variant::Boolean(false)
        );
    }

    #[test]
    fn test_error_propagates_through_arithmetic() {
        let c = HashMap::new();
        // Error propagates: 1/0 + 5 → #DIV/0!
        assert_eq!(calc("=1/0+5", &c), Variant::Error(ExcelError::DivZero));
    }

    #[test]
    fn test_error_display() {
        use crate::vm::ExcelError;
        assert_eq!(ExcelError::DivZero.as_str(), "#DIV/0!");
        assert_eq!(ExcelError::NA.as_str(), "#N/A");
        assert_eq!(ExcelError::Value.as_str(), "#VALUE!");
    }

    // ── Date variant ─────────────────────────────────────────────────────────

    #[test]
    fn test_date_variant() {
        let c = HashMap::new();
        assert_eq!(calc("=DATE(2000,1,1)", &c), Variant::Date(36526));
    }

    #[test]
    fn test_date_display() {
        assert_eq!(Variant::Date(36526).to_string(), "2000-01-01");
    }

    #[test]
    fn test_text_date_format() {
        let c = HashMap::new();
        assert_eq!(
            calc("=TEXT(DATE(2000,6,15),\"YYYY/MM/DD\")", &c),
            Variant::Str("2000/06/15".into())
        );
        assert_eq!(
            calc("=TEXT(DATE(2000,6,15),\"MM-DD-YYYY\")", &c),
            Variant::Str("06-15-2000".into())
        );
    }

    #[test]
    fn test_year_month_day_on_date_variant() {
        let c = HashMap::new();
        assert_eq!(calc("=YEAR(DATE(2000,6,15))", &c), Variant::Integer(2000));
        assert_eq!(calc("=MONTH(DATE(2000,6,15))", &c), Variant::Integer(6));
        assert_eq!(calc("=DAY(DATE(2000,6,15))", &c), Variant::Integer(15));
    }

    #[test]
    fn test_match_na_when_not_found() {
        let c = cells_from(&[((1, 1), Variant::Integer(1)), ((2, 1), Variant::Integer(2))]);
        assert_eq!(
            calc("=MATCH(99,A1:A2,0)", &c),
            Variant::Error(ExcelError::NA)
        );
    }

    #[test]
    fn test_averageif() {
        let c = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((1, 2), Variant::Integer(10)),
            ((2, 1), Variant::Str("b".into())),
            ((2, 2), Variant::Integer(20)),
            ((3, 1), Variant::Str("a".into())),
            ((3, 2), Variant::Integer(30)),
        ]);
        assert_eq!(
            calc("=AVERAGEIF(A1:A3,\"a\",B1:B3)", &c),
            Variant::Float(20.0)
        );
        let with_error = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((1, 2), Variant::Error(ExcelError::Value)),
        ]);
        assert_eq!(
            calc("=AVERAGEIF(A1:A1,\"a\",B1:B1)", &with_error),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    // 3.14 here is an arbitrary decimal test value for TRUNC/MOD, not an approximation of π.
    #[allow(clippy::approx_constant)]
    fn test_int_trunc_mod() {
        let c = HashMap::new();
        assert_eq!(calc("=INT(3.9)", &c), Variant::Integer(3));
        assert_eq!(calc("=INT(-3.1)", &c), Variant::Integer(-4));
        assert_eq!(calc("=TRUNC(3.9)", &c), Variant::Integer(3));
        assert_eq!(calc("=TRUNC(-3.9)", &c), Variant::Integer(-3));
        assert_eq!(calc("=TRUNC(3.14159,2)", &c), Variant::Float(3.14));
        assert_eq!(calc("=MOD(10,3)", &c), Variant::Integer(1));
        assert_eq!(calc("=MOD(-10,3)", &c), Variant::Integer(2));
    }

    #[test]
    fn test_large_small() {
        let c2 = cells_from(&[
            ((1, 1), Variant::Integer(5)),
            ((2, 1), Variant::Integer(1)),
            ((3, 1), Variant::Integer(3)),
        ]);
        assert_eq!(calc("=LARGE(A1:A3,1)", &c2), Variant::Integer(5));
        assert_eq!(calc("=LARGE(A1:A3,2)", &c2), Variant::Integer(3));
        assert_eq!(calc("=SMALL(A1:A3,1)", &c2), Variant::Integer(1));
        assert_eq!(calc("=SMALL(A1:A3,2)", &c2), Variant::Integer(3));
    }

    #[test]
    fn test_sumproduct() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(2)),
            ((1, 2), Variant::Integer(3)),
            ((2, 1), Variant::Integer(4)),
            ((2, 2), Variant::Integer(5)),
        ]);
        // SUMPRODUCT(A1:A2, B1:B2) = 2*3 + 4*5 = 26
        assert_eq!(calc("=SUMPRODUCT(A1:A2,B1:B2)", &c), Variant::Integer(26));
    }

    #[test]
    fn test_percentile() {
        let c2 = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(3)),
            ((4, 1), Variant::Integer(4)),
            ((5, 1), Variant::Integer(5)),
        ]);
        assert_eq!(calc("=PERCENTILE(A1:A5,0.5)", &c2), Variant::Integer(3));
        let _ = c2;
    }

    #[test]
    fn test_maxifs_minifs() {
        let c = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((1, 2), Variant::Integer(10)),
            ((2, 1), Variant::Str("b".into())),
            ((2, 2), Variant::Integer(20)),
            ((3, 1), Variant::Str("a".into())),
            ((3, 2), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=MAXIFS(B1:B3,A1:A3,\"a\")", &c), Variant::Integer(30));
        assert_eq!(calc("=MINIFS(B1:B3,A1:A3,\"a\")", &c), Variant::Integer(10));
        let with_error = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((1, 2), Variant::Error(ExcelError::Ref)),
        ]);
        assert_eq!(
            calc("=MAXIFS(B1:B1,A1:A1,\"a\")", &with_error),
            Variant::Error(ExcelError::Ref)
        );
        assert_eq!(
            calc("=MINIFS(B1:B1,A1:A1,\"a\")", &with_error),
            Variant::Error(ExcelError::Ref)
        );
    }

    #[test]
    fn test_rand_randbetween() {
        let c = HashMap::new();
        if let Variant::Float(v) = calc("=RAND()", &c) {
            assert!((0.0..1.0).contains(&v));
        } else {
            panic!("RAND should return Float");
        }
        if let Variant::Integer(v) = calc("=RANDBETWEEN(1,10)", &c) {
            assert!((1..=10).contains(&v));
        } else {
            panic!("RANDBETWEEN should return Integer");
        }
    }

    #[test]
    fn test_string_upper_lower_proper_trim() {
        let c = HashMap::new();
        assert_eq!(calc("=UPPER(\"hello\")", &c), Variant::Str("HELLO".into()));
        assert_eq!(calc("=LOWER(\"HELLO\")", &c), Variant::Str("hello".into()));
        assert_eq!(
            calc("=PROPER(\"hello world\")", &c),
            Variant::Str("Hello World".into())
        );
        assert_eq!(
            calc("=TRIM(\"  hello   world  \")", &c),
            Variant::Str("hello world".into())
        );
    }

    #[test]
    fn test_unicode_functions_are_not_legacy_aliases() {
        let c = HashMap::new();
        assert_eq!(calc("=UNICHAR(128512)", &c), Variant::Str("😀".into()));
        assert_eq!(calc("=UNICODE(\"😀ok\")", &c), Variant::Integer(128512));
        assert_eq!(calc("=UNICHAR(0)", &c), Variant::Error(ExcelError::Value));
        assert_eq!(
            calc("=UNICHAR(55296)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=UNICODE(\"\")", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=UNICHAR(65.5)", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_substitute_replace() {
        let c = HashMap::new();
        assert_eq!(
            calc("=SUBSTITUTE(\"aabbaa\",\"a\",\"x\")", &c),
            Variant::Str("xxbbxx".into())
        );
        assert_eq!(
            calc("=SUBSTITUTE(\"aabbaa\",\"a\",\"x\",2)", &c),
            Variant::Str("axbbaa".into())
        );
        assert_eq!(
            calc("=REPLACE(\"Hello\",1,2,\"ZZ\")", &c),
            Variant::Str("ZZllo".into())
        );
    }

    #[test]
    fn test_find_search() {
        let c = HashMap::new();
        assert_eq!(calc("=FIND(\"lo\",\"Hello\")", &c), Variant::Integer(4));
        assert_eq!(calc("=SEARCH(\"LO\",\"Hello\")", &c), Variant::Integer(4));
        assert_eq!(calc("=SEARCH(\"h*o\",\"Hello\")", &c), Variant::Integer(1));
        assert_eq!(calc("=FINDB(\"a\",\"あa\")", &c), Variant::Integer(3));
        assert_eq!(calc("=SEARCHB(\"A\",\"あa\")", &c), Variant::Integer(3));
        assert_eq!(
            calc("=REPLACEB(\"あいう\",3,2,\"X\")", &c),
            Variant::Str("あXう".into())
        );
    }

    #[test]
    fn test_find_with_an_empty_search_string_does_not_panic() {
        // Regression: fuzz corpus found `FIND("","...")` panicked on
        // `.windows(0)` ("window size must be non-zero"). Real Excel
        // matches trivially at the start position -- same convention
        // VBA InStr already uses (src/vm/mod.rs).
        let c = HashMap::new();
        assert_eq!(calc("=FIND(\"\",\"abc\")", &c), Variant::Integer(1));
        assert_eq!(calc("=FIND(\"\",\"abc\",3)", &c), Variant::Integer(3));
    }

    #[test]
    fn test_find_with_a_start_beyond_the_haystack_errors_instead_of_panicking() {
        let c = HashMap::new();
        assert!(evaluate(&fparse("=FIND(\"a\",\"abc\",10)").unwrap(), &c).is_err());
        assert!(evaluate(&fparse("=FIND(\"\",\"abc\",10)").unwrap(), &c).is_err());
    }

    #[test]
    fn test_find_reproduces_the_exact_fuzz_crash_input() {
        // The exact bytes fuzz/fuzz_targets/fuzz_formula_eval.rs's corpus found:
        // "FIND(v0\t\t,D)" -- an unquoted bare name `v0` (unresolved, evaluates
        // to Variant::Empty -> to_str -> "") as the search string. Reproduced
        // with that fuzz target's own 5x5 (r*c) cell setup, not this module's
        // usual empty-map `c`, so this is a faithful reproduction, not just a
        // simplified equivalent. Verified (by temporarily reverting the fix)
        // that this test and the two above actually fail without it.
        let raw = std::str::from_utf8(&[70, 73, 78, 68, 40, 118, 48, 9, 9, 44, 68, 41]).unwrap();
        assert_eq!(raw, "FIND(v0\t\t,D)");
        let mut cells: HashMap<(u32, u32), CellContent> = HashMap::new();
        for r in 1u32..=5 {
            for col in 1u32..=5 {
                cells.insert(
                    (r, col),
                    CellContent {
                        formula: None,
                        value: Variant::Integer((r * col) as i64),
                    },
                );
            }
        }
        let expr = fparse(raw).unwrap();
        let _ = evaluate(&expr, &cells); // must not panic; result value not asserted
    }

    #[test]
    fn test_textjoin() {
        let c = HashMap::new();
        assert_eq!(
            calc("=TEXTJOIN(\"-\",TRUE,\"a\",\"\",\"b\")", &c),
            Variant::Str("a-b".into())
        );
        assert_eq!(
            calc("=TEXTJOIN(\"-\",FALSE,\"a\",\"\",\"b\")", &c),
            Variant::Str("a--b".into())
        );
    }

    #[test]
    fn test_char_code() {
        let c = HashMap::new();
        assert_eq!(calc("=CHAR(65)", &c), Variant::Str("A".into()));
        assert_eq!(calc("=CODE(\"A\")", &c), Variant::Integer(65));
        assert_eq!(calc("=UNICHAR(9786)", &c), Variant::Str("☺".into()));
        assert_eq!(calc("=UNICODE(\"☺\")", &c), Variant::Integer(9786));
    }

    #[test]
    fn test_exact() {
        let c = HashMap::new();
        assert_eq!(
            calc("=EXACT(\"Hello\",\"Hello\")", &c),
            Variant::Boolean(true)
        );
        assert_eq!(
            calc("=EXACT(\"Hello\",\"hello\")", &c),
            Variant::Boolean(false)
        );
    }

    #[test]
    // 3.14 is an arbitrary decimal test value for VALUE()'s string-to-number parsing, not π.
    #[allow(clippy::approx_constant)]
    fn test_value() {
        let c = HashMap::new();
        assert_eq!(calc("=VALUE(\"42\")", &c), Variant::Integer(42));
        assert_eq!(calc("=VALUE(\"3.14\")", &c), Variant::Float(3.14));
    }

    #[test]
    fn test_euroconvert() {
        let c = HashMap::new();
        assert_eq!(
            calc("=EUROCONVERT(1.2,\"DEM\",\"EUR\")", &c),
            Variant::Float(0.61)
        );
        assert!(matches!(
            calc("=EUROCONVERT(1,\"FRF\",\"DEM\",TRUE,3)", &c),
            Variant::Float(value) if (value - 0.29728616).abs() < 1e-12
        ));
        assert_eq!(
            calc("=EUROCONVERT(7,\"EUR\",\"EUR\")", &c),
            Variant::Integer(7)
        );
        assert_eq!(
            calc("=EUROCONVERT(1,\"EUR\",\"DEM\",FALSE,2)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=EUROCONVERT(1,\"JPY\",\"EUR\")", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_asc_jis() {
        let c = HashMap::new();
        // Full-width A (U+FF21) → half-width A
        assert_eq!(calc("=ASC(\"Ａ\")", &c), Variant::Str("A".into()));
        // Half-width A → full-width A
        assert_eq!(calc("=JIS(\"A\")", &c), Variant::Str("Ａ".into()));
        assert_eq!(calc("=DBCS(\"A\")", &c), Variant::Str("Ａ".into()));
    }

    #[test]
    fn test_year_month_day_weekday() {
        let c = HashMap::new();
        // DATE(2000,6,15) = June 15 2000
        assert_eq!(calc("=YEAR(DATE(2000,6,15))", &c), Variant::Integer(2000));
        assert_eq!(calc("=MONTH(DATE(2000,6,15))", &c), Variant::Integer(6));
        assert_eq!(calc("=DAY(DATE(2000,6,15))", &c), Variant::Integer(15));
        // June 15 2000 was a Thursday; WEEKDAY type=2: Mon=1,...Thu=4
        assert_eq!(calc("=WEEKDAY(DATE(2000,6,15),2)", &c), Variant::Integer(4));
        // The 11–17 forms select Monday through Sunday as the first day.
        // June 15 2000 was Thursday, so the results are 4, 3, 2, 1, 7, 6, 5.
        assert_eq!(
            calc("=WEEKDAY(DATE(2000,6,15),11)", &c),
            Variant::Integer(4)
        );
        assert_eq!(
            calc("=WEEKDAY(DATE(2000,6,15),12)", &c),
            Variant::Integer(3)
        );
        assert_eq!(
            calc("=WEEKDAY(DATE(2000,6,15),13)", &c),
            Variant::Integer(2)
        );
        assert_eq!(
            calc("=WEEKDAY(DATE(2000,6,15),14)", &c),
            Variant::Integer(1)
        );
        assert_eq!(
            calc("=WEEKDAY(DATE(2000,6,15),15)", &c),
            Variant::Integer(7)
        );
        assert_eq!(
            calc("=WEEKDAY(DATE(2000,6,15),16)", &c),
            Variant::Integer(6)
        );
        assert_eq!(
            calc("=WEEKDAY(DATE(2000,6,15),17)", &c),
            Variant::Integer(5)
        );
        assert!(evaluate(&fparse("=WEEKDAY(DATE(2000,6,15),2.5)").unwrap(), &c).is_err());
        assert!(evaluate(&fparse("=WEEKNUM(DATE(2000,6,15),2.5)").unwrap(), &c).is_err());
    }

    #[test]
    fn test_days_edate() {
        let c = HashMap::new();
        assert_eq!(
            calc("=DAYS(DATE(2000,2,1),DATE(2000,1,1))", &c),
            Variant::Integer(31)
        );
        // EDATE(Jan 31 2000, 1) = Feb 29 2000 (leap year, clamped)
        assert_eq!(
            calc("=EDATE(DATE(2000,1,31),1)", &c),
            Variant::Date(date_to_serial(2000, 2, 29))
        );
    }

    #[test]
    fn test_datedif() {
        let c = HashMap::new();
        let s = calc("=DATEDIF(DATE(2000,1,1),DATE(2001,6,15),\"Y\")", &c);
        assert_eq!(s, Variant::Integer(1));
        let m = calc("=DATEDIF(DATE(2000,1,1),DATE(2000,3,1),\"M\")", &c);
        assert_eq!(m, Variant::Integer(2));
        let d = calc("=DATEDIF(DATE(2000,1,1),DATE(2000,1,10),\"D\")", &c);
        assert_eq!(d, Variant::Integer(9));
    }

    #[test]
    fn test_datevalue_timevalue() {
        let c = HashMap::new();
        assert_eq!(calc("=DATEVALUE(\"2000/01/01\")", &c), Variant::Date(36526));
        assert_eq!(calc("=DATEVALUE(\"2000-01-01\")", &c), Variant::Date(36526));
        assert_eq!(
            calc("=DATEVALUE(\"2000-01-01T12:30:00\")", &c),
            Variant::Date(36526)
        );
        assert_eq!(calc("=DATEVALUE(\"2000.01.01\")", &c), Variant::Date(36526));
        assert_eq!(
            calc("=DATEVALUE(\"January 1, 2000\")", &c),
            Variant::Date(36526)
        );
        assert_eq!(calc("=DATEVALUE(\"1 Jan 2000\")", &c), Variant::Date(36526));
        assert_eq!(
            calc("=DATEVALUE(\"2020-02-31\")", &c),
            Variant::Date(date_to_serial(2020, 3, 2))
        );
        if let Variant::Float(v) = calc("=TIMEVALUE(\"12:00:00\")", &c) {
            assert!((v - 0.5).abs() < 1e-9);
        } else {
            panic!("TIMEVALUE should return Float");
        }
        if let Variant::Float(v) = calc("=TIMEVALUE(\"12:30 PM\")", &c) {
            assert!((v - 0.5208333333333334).abs() < 1e-9);
        } else {
            panic!("TIMEVALUE should accept PM notation");
        }
        if let Variant::Float(v) = calc("=TIMEVALUE(\"12:30 AM\")", &c) {
            assert!((v - 0.020833333333333332).abs() < 1e-9);
        } else {
            panic!("TIMEVALUE should accept AM notation");
        }
        let invalid = evaluate(&fparse("=TIMEVALUE(\"24:00\")").unwrap(), &c);
        assert!(invalid.is_err());
    }

    #[test]
    fn test_time_hour_minute_second() {
        let c = HashMap::new();
        if let Variant::Float(v) = calc("=TIME(12,30,0)", &c) {
            assert!((v - 0.520833).abs() < 1e-5);
        }
        assert_eq!(calc("=HOUR(TIME(12,30,45))", &c), Variant::Integer(12));
        assert_eq!(calc("=MINUTE(TIME(12,30,45))", &c), Variant::Integer(30));
        assert_eq!(calc("=SECOND(TIME(12,30,45))", &c), Variant::Integer(45));
    }

    #[test]
    fn test_networkdays_intl() {
        let c = HashMap::new();
        // Mon Jan 3 2000 to Fri Jan 7 2000, default weekend (Sat+Sun) = 5
        assert_eq!(
            calc("=NETWORKDAYS.INTL(36528,36532,1)", &c),
            Variant::Integer(5)
        );
        // same range, weekend = Mon only (code 12): 4 days (Tue-Fri)
        assert_eq!(
            calc("=NETWORKDAYS.INTL(36528,36532,12)", &c),
            Variant::Integer(4)
        );
        assert_eq!(
            calc("=NETWORKDAYS.INTL(36528,36532,\"0000011\")", &c),
            Variant::Integer(5)
        );
    }

    #[test]
    fn test_workday_intl() {
        let c = HashMap::new();
        // 5 workdays after Mon Jan 3 2000 (serial 36528) = Fri Jan 7 2000 (36532)? no, it's Mon Jan 10 (36535)
        // Actually: Jan 3+1=Tue4, +2=Wed5, +3=Thu6, +4=Fri7, +5=Mon10 = 36535
        assert_eq!(calc("=WORKDAY.INTL(36528,5,1)", &c), Variant::Date(36535));
        assert_eq!(
            calc("=WORKDAY.INTL(36528,5,\"0000011\")", &c),
            Variant::Date(36535)
        );
    }

    #[test]
    fn test_switch() {
        let c = HashMap::new();
        assert_eq!(
            calc("=SWITCH(2,1,\"one\",2,\"two\",3,\"three\")", &c),
            Variant::Str("two".into())
        );
        assert_eq!(
            calc("=SWITCH(99,1,\"one\",\"default\")", &c),
            Variant::Str("default".into())
        );
        assert_eq!(
            calc("=SWITCH(CHOOSE(SEQUENCE(2),1,2),1,\"one\",2,\"two\")", &c),
            Variant::Array(vec![Variant::Str("one".into()), Variant::Str("two".into())])
        );
        assert_eq!(
            calc("=SWITCH(CHOOSE(SEQUENCE(2),1,3),1,\"one\",\"other\")", &c),
            Variant::Array(vec![
                Variant::Str("one".into()),
                Variant::Str("other".into())
            ])
        );
        assert_eq!(
            calc("=SWITCH(CHOOSE(SEQUENCE(2),1,3),1,\"one\")", &c),
            Variant::Array(vec![
                Variant::Str("one".into()),
                Variant::Error(ExcelError::NA)
            ])
        );
    }

    #[test]
    fn test_if_broadcasts_array_conditions_and_branches() {
        let c = HashMap::new();
        assert_eq!(
            calc("=IF(CHOOSE(SEQUENCE(2),TRUE,FALSE),\"Y\",\"N\")", &c),
            Variant::Array(vec![Variant::Str("Y".into()), Variant::Str("N".into())])
        );
        assert_eq!(
            calc("=IF(CHOOSE(SEQUENCE(2),TRUE,FALSE),SEQUENCE(2),0)", &c),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(0)])
        );
        assert_eq!(
            calc("=IF(CHOOSE(SEQUENCE(3),TRUE,FALSE),1,2)", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_xor() {
        let c = HashMap::new();
        assert_eq!(calc("=XOR(TRUE,FALSE)", &c), Variant::Boolean(true));
        assert_eq!(calc("=XOR(TRUE,TRUE)", &c), Variant::Boolean(false));
        assert_eq!(calc("=XOR(TRUE,TRUE,TRUE)", &c), Variant::Boolean(true));
        assert_eq!(
            calc("=XOR(TRUE,1/0)", &c),
            Variant::Error(ExcelError::DivZero)
        );
    }

    #[test]
    fn test_choose_column() {
        let c = HashMap::new();
        assert_eq!(
            calc("=CHOOSE(2,\"a\",\"b\",\"c\")", &c),
            Variant::Str("b".into())
        );
        let range = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(20)),
        ]);
        assert_eq!(
            calc("=CHOOSE(1,A1:A2,\"unused\")", &range),
            Variant::Array(vec![Variant::Integer(10), Variant::Integer(20)])
        );
        assert_eq!(
            calc("=CHOOSE(SEQUENCE(2),10,20)", &c),
            Variant::Array(vec![Variant::Integer(10), Variant::Integer(20)])
        );
        assert_eq!(
            calc("=CHOOSE(SEQUENCE(3),10,20)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(calc("=COLUMN(C1)", &c), Variant::Integer(3));
        assert_eq!(calc("=COLUMN()", &c), Variant::Integer(1));
        assert_eq!(calc("=COLUMNS(B3:C7)", &c), Variant::Integer(2));
        assert_eq!(calc("=COLUMNS(B3)", &c), Variant::Integer(1));
        assert_eq!(calc("=COLUMNS(SEQUENCE(2,3))", &c), Variant::Integer(3));
        assert_eq!(calc("=COLUMNS(TRANSPOSE(B3:C7))", &c), Variant::Integer(5));
    }

    #[test]
    fn test_lookup_xmatch() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(3)),
            ((1, 2), Variant::Str("one".into())),
            ((2, 2), Variant::Str("two".into())),
            ((3, 2), Variant::Str("three".into())),
        ]);
        assert_eq!(
            calc("=LOOKUP(2,A1:A3,B1:B3)", &c),
            Variant::Str("two".into())
        );
        assert_eq!(
            calc("=LOOKUP(2,A1:A3,B1:B2)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=LOOKUP(2,A1:A3,B1:D1)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(calc("=XMATCH(2,A1:A3,0)", &c), Variant::Integer(2));
        assert_eq!(calc("=XMATCH(2,A1:A3)", &c), Variant::Integer(2));
        assert_eq!(calc("=XMATCH(\"t*\",B1:B3,2)", &c), Variant::Integer(2));
        assert_eq!(calc("=XMATCH(2,A1:A3,0,2)", &c), Variant::Integer(2));
        assert_eq!(calc("=XMATCH(4,A1:A3,-1,2)", &c), Variant::Integer(3));
        assert_eq!(
            calc("=XMATCH(2,A1:A3,0.5)", &c),
            Variant::Error(ExcelError::Value)
        );

        let unsorted = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(3)),
            ((3, 1), Variant::Integer(2)),
        ]);
        assert_eq!(
            calc("=LOOKUP(2,A1:A3,B1:B3)", &unsorted),
            Variant::Error(ExcelError::NA)
        );

        let errored = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Error(ExcelError::DivZero)),
        ]);
        assert_eq!(
            calc("=LOOKUP(2,A1:A2)", &errored),
            Variant::Error(ExcelError::Value)
        );

        let matrix = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Integer(2)),
            ((2, 1), Variant::Integer(3)),
            ((2, 2), Variant::Integer(4)),
        ]);
        assert_eq!(
            calc("=XMATCH(2,A1:B2,0)", &matrix),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_is_functions() {
        let c = HashMap::new();
        assert_eq!(calc("=ISBLANK(\"\")", &c), Variant::Boolean(false));
        assert_eq!(calc("=ISERROR(1/0)", &c), Variant::Boolean(true));
        assert_eq!(calc("=ISERROR(1)", &c), Variant::Boolean(false));
        assert_eq!(calc("=ISNUMBER(42)", &c), Variant::Boolean(true));
        assert_eq!(calc("=ISNUMBER(\"x\")", &c), Variant::Boolean(false));
        assert_eq!(calc("=ISTEXT(\"a\")", &c), Variant::Boolean(true));
        assert_eq!(calc("=ISLOGICAL(TRUE)", &c), Variant::Boolean(true));
        assert_eq!(calc("=ISNONTEXT(42)", &c), Variant::Boolean(true));
        assert_eq!(calc("=ISNA(1)", &c), Variant::Boolean(false));
        assert_eq!(calc("=AREAS(A1)", &c), Variant::Integer(1));
        assert_eq!(calc("=AREAS(A1:B2)", &c), Variant::Integer(1));
        assert_eq!(calc("=AREAS(1+2)", &c), Variant::Error(ExcelError::Value));
        assert_eq!(
            calc("=INFO(\"release\")", &c),
            Variant::Str(env!("CARGO_PKG_VERSION").into())
        );
        assert_eq!(
            calc("=INFO(\"recalc\")", &c),
            Variant::Str("Automatic".into())
        );
        assert_eq!(
            calc("=INFO(\"memused\")", &c),
            Variant::Error(ExcelError::NA)
        );
    }

    #[test]
    fn test_aggregate() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(20)),
            ((3, 1), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=AGGREGATE(9,0,A1:A3)", &c), Variant::Integer(60));
        assert_eq!(calc("=AGGREGATE(1,0,A1:A3)", &c), Variant::Float(20.0));
        assert_eq!(calc("=AGGREGATE(4,0,A1:A3)", &c), Variant::Integer(30));
        assert_eq!(calc("=AGGREGATE(12,0,A1:A3)", &c), Variant::Integer(20));
        assert_eq!(calc("=AGGREGATE(14,0,A1:A3,2)", &c), Variant::Integer(20));
        assert_eq!(calc("=AGGREGATE(15,0,A1:A3,2)", &c), Variant::Integer(20));
        assert!(matches!(
            calc("=AGGREGATE(7,0,A1:A3)", &c),
            Variant::Float(value) if (value - 10.0).abs() < 1e-12
        ));
        assert!(matches!(
            calc("=AGGREGATE(8,0,A1:A3)", &c),
            Variant::Float(value) if (value - (200.0_f64 / 3.0).sqrt()).abs() < 1e-12
        ));
        assert_eq!(calc("=AGGREGATE(13,0,A1:A3)", &c), Variant::Integer(10));
        assert_eq!(calc("=AGGREGATE(17,0,A1:A3,2)", &c), Variant::Integer(20));
        assert_eq!(calc("=AGGREGATE(18,0,A1:A3,0.5)", &c), Variant::Integer(20));
        assert_eq!(calc("=AGGREGATE(20,0,A1:A3,20)", &c), Variant::Float(0.5));
        assert_eq!(calc("=AGGREGATE(21,0,A1:A3,20)", &c), Variant::Float(0.5));
    }

    #[test]
    fn test_aggregate_options_filter_nested_and_errors() {
        let mut nested = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(30)),
            ((3, 1), Variant::Integer(5)),
        ]);
        nested.insert(
            (2, 1),
            CellContent {
                formula: Some("=SUBTOTAL(9,B1:B2)".into()),
                value: Variant::Integer(30),
            },
        );
        // Option 0 ignores a nested SUBTOTAL; option 4 includes it.
        assert_eq!(calc("=AGGREGATE(9,0,A1:A3)", &nested), Variant::Integer(15));
        assert_eq!(calc("=AGGREGATE(9,4,A1:A3)", &nested), Variant::Integer(45));

        let errors = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Error(ExcelError::DivZero)),
            ((3, 1), Variant::Integer(5)),
        ]);
        assert_eq!(
            calc("=AGGREGATE(9,0,A1:A3)", &errors),
            Variant::Error(ExcelError::DivZero)
        );
        assert_eq!(calc("=AGGREGATE(9,6,A1:A3)", &errors), Variant::Integer(15));
        assert_eq!(
            calc("=AGGREGATE(9,1.5,A1:A3)", &errors),
            Variant::Error(ExcelError::Value)
        );
    }

    // ── Phase 10: numerical functions ─────────────────────────────────────────

    fn approx(v: Variant, expected: f64) {
        let f = match v {
            Variant::Float(f) => f,
            Variant::Integer(n) => n as f64,
            _ => panic!("not numeric"),
        };
        assert!(
            (f - expected).abs() < 1e-9,
            "expected {}, got {}",
            expected,
            f
        );
    }

    #[test]
    fn test_stdev_var() {
        let c = HashMap::new();
        // [2,4,4,4,5,5,7,9]: mean=5, sum_sq_dev=32
        // sample stdev = sqrt(32/7) ≈ 2.138
        approx(calc("=STDEV(2,4,4,4,5,5,7,9)", &c), (32.0f64 / 7.0).sqrt());
        approx(
            calc("=STDEV.S(2,4,4,4,5,5,7,9)", &c),
            (32.0f64 / 7.0).sqrt(),
        );
        // population stdev = sqrt(32/8) = 2.0
        approx(calc("=STDEVP(2,4,4,4,5,5,7,9)", &c), 2.0);
        approx(calc("=STDEV.P(2,4,4,4,5,5,7,9)", &c), 2.0);
        // VAR: sample=32/7, population=32/8=4
        approx(calc("=VAR(2,4,4,4,5,5,7,9)", &c), 32.0 / 7.0);
        approx(calc("=VARP(2,4,4,4,5,5,7,9)", &c), 4.0);
        approx(calc("=STDEV.S(SEQUENCE(3))", &c), 1.0);
        approx(calc("=VAR.S(SEQUENCE(3))", &c), 1.0);
        approx(calc("=STDEV.P(SEQUENCE(3))", &c), (2.0_f64 / 3.0).sqrt());
    }

    #[test]
    fn test_floor_ceiling_mround() {
        let c = HashMap::new();
        assert_eq!(calc("=FLOOR(3.7,1)", &c), Variant::Integer(3));
        approx(calc("=FLOOR(3.7,0.5)", &c), 3.5);
        assert_eq!(calc("=CEILING(3.2,1)", &c), Variant::Integer(4));
        approx(calc("=CEILING(3.2,0.5)", &c), 3.5);
        assert_eq!(calc("=MROUND(10,3)", &c), Variant::Integer(9));
        assert_eq!(calc("=MROUND(11,3)", &c), Variant::Integer(12));
        approx(calc("=MROUND(1.4,0.5)", &c), 1.5);
    }

    #[test]
    // 3.14 is an arbitrary decimal test value for ABS(), not an approximation of π.
    #[allow(clippy::approx_constant)]
    fn test_math_functions() {
        let c = HashMap::new();
        assert_eq!(calc("=ABS(-5)", &c), Variant::Integer(5));
        approx(calc("=ABS(3.14)", &c), 3.14);
        assert_eq!(calc("=SQRT(9)", &c), Variant::Integer(3));
        approx(calc("=SQRT(2)", &c), 2f64.sqrt());
        approx(calc("=SQRTPI(4)", &c), (4.0 * std::f64::consts::PI).sqrt());
        assert_eq!(calc("=SQRTPI(-1)", &c), Variant::Error(ExcelError::Num));
        assert_eq!(calc("=POWER(2,10)", &c), Variant::Integer(1024));
        approx(calc("=EXP(0)", &c), 1.0);
        approx(calc("=EXP(1)", &c), std::f64::consts::E);
        approx(calc("=LOG(100)", &c), 2.0);
        approx(calc("=LOG(8,2)", &c), 3.0);
        approx(calc("=LOG10(1000)", &c), 3.0);
        approx(calc("=LN(1)", &c), 0.0);
    }

    #[test]
    fn test_trig_functions() {
        let c = HashMap::new();
        let pi = std::f64::consts::PI;
        approx(calc("=PI()", &c), pi);
        approx(calc("=SIN(PI()/2)", &c), 1.0);
        approx(calc("=COS(0)", &c), 1.0);
        approx(calc("=TAN(0)", &c), 0.0);
        approx(calc("=DEGREES(PI())", &c), 180.0);
        approx(calc("=RADIANS(180)", &c), pi);
        approx(calc("=ATAN2(1,1)", &c), pi / 4.0);
        approx(calc("=ASIN(1)", &c), pi / 2.0);
        approx(calc("=ACOS(1)", &c), 0.0);
        approx(calc("=ATAN(1)", &c), pi / 4.0);
    }

    #[test]
    fn test_countblank() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Empty),
            ((3, 1), Variant::Str("".into())),
        ]);
        assert_eq!(calc("=COUNTBLANK(A1:A3)", &c), Variant::Integer(2));
    }

    #[test]
    fn test_address() {
        let c = HashMap::new();
        assert_eq!(calc("=ADDRESS(1,1)", &c), Variant::Str("$A$1".into()));
        assert_eq!(calc("=ADDRESS(2,3,4)", &c), Variant::Str("C2".into()));
        assert_eq!(calc("=ADDRESS(1,27)", &c), Variant::Str("$AA$1".into()));
    }

    #[test]
    fn test_indirect() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(42),
            },
        );
        c.insert(
            (3, 2),
            CellContent {
                formula: None,
                value: Variant::Str("hello".into()),
            },
        );
        // INDIRECT("A1") → value at A1
        assert_eq!(calc("=INDIRECT(\"A1\")", &c), Variant::Integer(42));
        // INDIRECT("B3") → value at B3
        assert_eq!(calc("=INDIRECT(\"B3\")", &c), Variant::Str("hello".into()));
        // INDIRECT of empty cell → Empty
        assert_eq!(calc("=INDIRECT(\"C5\")", &c), Variant::Empty);
        // INDIRECT with range reference → bounded row-major array
        assert_eq!(
            calc("=INDIRECT(\"A1:B3\")", &c),
            Variant::Array(vec![
                Variant::Integer(42),
                Variant::Empty,
                Variant::Empty,
                Variant::Empty,
                Variant::Empty,
                Variant::Str("hello".into()),
            ])
        );
        assert_eq!(calc("=SUM(INDIRECT(\"A1:A1\"))", &c), Variant::Integer(42));
        assert_eq!(calc("=INDIRECT(\"R1C1\",FALSE)", &c), Variant::Integer(42));
    }

    #[test]
    fn test_spill_sort() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(3),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        assert_eq!(
            calc("=SORT(A1:A3)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3)
            ])
        );
        assert_eq!(
            calc("=SORT(A1:A3,1,-1)", &c),
            Variant::Array(vec![
                Variant::Integer(3),
                Variant::Integer(2),
                Variant::Integer(1)
            ])
        );
        c.extend([
            (
                (1, 2),
                CellContent {
                    formula: None,
                    value: Variant::Integer(2),
                },
            ),
            (
                (2, 2),
                CellContent {
                    formula: None,
                    value: Variant::Integer(1),
                },
            ),
            (
                (3, 2),
                CellContent {
                    formula: None,
                    value: Variant::Integer(1),
                },
            ),
            (
                (1, 3),
                CellContent {
                    formula: None,
                    value: Variant::Integer(20),
                },
            ),
            (
                (2, 3),
                CellContent {
                    formula: None,
                    value: Variant::Integer(30),
                },
            ),
            (
                (3, 3),
                CellContent {
                    formula: None,
                    value: Variant::Integer(10),
                },
            ),
        ]);
        assert_eq!(
            calc("=SORTBY(A1:A3,B1:B3,1,C1:C3,-1)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
            ])
        );
    }

    #[test]
    fn test_spill_unique() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        c.insert(
            (4, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(3),
            },
        );
        assert_eq!(
            calc("=UNIQUE(A1:A4)", &c),
            Variant::Array(vec![
                Variant::Integer(2),
                Variant::Integer(1),
                Variant::Integer(3)
            ])
        );
    }

    #[test]
    fn test_spill_sequence() {
        let c = HashMap::new();
        assert_eq!(
            calc("=SEQUENCE(4)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Integer(4)
            ])
        );
        assert_eq!(
            calc("=SEQUENCE(3,1,5,2)", &c),
            Variant::Array(vec![
                Variant::Integer(5),
                Variant::Integer(7),
                Variant::Integer(9)
            ])
        );
        assert_eq!(calc("=SINGLE(SEQUENCE(4))", &c), Variant::Integer(1));
        assert_eq!(calc("=SINGLE(42)", &c), Variant::Integer(42));
        assert!(evaluate(&fparse("=SEQUENCE(2.5)").unwrap(), &c).is_err());
        assert!(evaluate(&fparse("=SEQUENCE(0)").unwrap(), &c).is_err());
    }

    #[test]
    fn test_spill_filter_range_condition() {
        // FILTER(A1:A4, B1:B4) — keep rows where B is truthy
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(10),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(20),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(30),
            },
        );
        c.insert(
            (4, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(40),
            },
        );
        c.insert(
            (1, 2),
            CellContent {
                formula: None,
                value: Variant::Boolean(true),
            },
        );
        c.insert(
            (2, 2),
            CellContent {
                formula: None,
                value: Variant::Boolean(false),
            },
        );
        c.insert(
            (3, 2),
            CellContent {
                formula: None,
                value: Variant::Boolean(true),
            },
        );
        c.insert(
            (4, 2),
            CellContent {
                formula: None,
                value: Variant::Boolean(false),
            },
        );
        assert_eq!(
            calc("=FILTER(A1:A4, B1:B4)", &c),
            Variant::Array(vec![Variant::Integer(10), Variant::Integer(30)])
        );
    }

    #[test]
    fn test_spill_filter_inline_comparison() {
        // FILTER(A1:A4, A1:A4>15)
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(10),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(20),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(30),
            },
        );
        c.insert(
            (4, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(5),
            },
        );
        assert_eq!(
            calc("=FILTER(A1:A4, A1:A4>15)", &c),
            Variant::Array(vec![Variant::Integer(20), Variant::Integer(30)])
        );
        assert_eq!(
            calc("=FILTER(A1:A4, (A1:A4>15)*(A1:A4<30))", &c),
            Variant::Integer(20)
        );
        assert_eq!(
            calc("=FILTER(A1:A4, (A1:A4<10)+(A1:A4>25))", &c),
            Variant::Array(vec![Variant::Integer(30), Variant::Integer(5)])
        );
    }

    #[test]
    fn test_spill_transpose() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        c.insert(
            (1, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        c.insert(
            (1, 3),
            CellContent {
                formula: None,
                value: Variant::Integer(3),
            },
        );
        // TRANSPOSE(A1:C1) — 1 row × 3 cols → 3 rows × 1 col (flat = same values)
        assert_eq!(
            calc("=TRANSPOSE(A1:C1)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3)
            ])
        );
    }

    #[test]
    fn test_let_basic() {
        let c = HashMap::new();
        // LET(x, 5, x+1) → 6
        assert_eq!(calc("=LET(x, 5, x+1)", &c), Variant::Integer(6));
        // LET(x, 3, y, 4, x*y) → 12
        assert_eq!(calc("=LET(x, 3, y, 4, x*y)", &c), Variant::Integer(12));
        // LET with string
        assert_eq!(calc("=LET(s, \"hello\", LEN(s))", &c), Variant::Integer(5));
    }

    #[test]
    fn test_isomitted_tracks_missing_lambda_bindings() {
        let c = HashMap::new();
        let lambda = fparse("=LAMBDA(value,ISOMITTED(value))").unwrap();
        assert_eq!(
            call_lambda(&lambda, vec![], &c).unwrap(),
            Variant::Boolean(true)
        );
        assert_eq!(
            call_lambda(&lambda, vec![Variant::Empty], &c).unwrap(),
            Variant::Boolean(false)
        );
        assert_eq!(
            calc("=LAMBDA(value,ISOMITTED(value))()", &c),
            Variant::Boolean(true)
        );
        assert_eq!(
            calc("=LAMBDA(value,ISOMITTED(value))(\"\")", &c),
            Variant::Boolean(false)
        );
        assert_eq!(calc("=LAMBDA(value,value+1)(2)", &c), Variant::Integer(3));
    }

    #[test]
    fn test_makearray_lambda() {
        let c = HashMap::new();
        assert_eq!(
            calc("=MAKEARRAY(2,3,LAMBDA(r,c,r*10+c))", &c),
            Variant::Array(vec![
                Variant::Integer(11),
                Variant::Integer(12),
                Variant::Integer(13),
                Variant::Integer(21),
                Variant::Integer(22),
                Variant::Integer(23),
            ])
        );
        assert!(evaluate(&fparse("=MAKEARRAY(0,2,LAMBDA(r,c,r+c))").unwrap(), &c).is_err());
    }

    #[test]
    fn test_map_lambda() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(3),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(4),
            },
        );
        // MAP(A1:A3, LAMBDA(x, x*2)) → [4, 6, 8]
        assert_eq!(
            calc("=MAP(A1:A3, LAMBDA(x, x*2))", &c),
            Variant::Array(vec![
                Variant::Integer(4),
                Variant::Integer(6),
                Variant::Integer(8)
            ])
        );
    }

    #[test]
    fn test_reduce_lambda() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(3),
            },
        );
        c.insert(
            (4, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(4),
            },
        );
        // REDUCE(0, A1:A4, LAMBDA(acc, x, acc+x)) → 10
        assert_eq!(
            calc("=REDUCE(0, A1:A4, LAMBDA(acc, x, acc+x))", &c),
            Variant::Integer(10)
        );
        // REDUCE(1, A1:A4, LAMBDA(acc, x, acc*x)) → 24
        assert_eq!(
            calc("=REDUCE(1, A1:A4, LAMBDA(acc, x, acc*x))", &c),
            Variant::Integer(24)
        );
    }

    #[test]
    fn test_scan_lambda() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(3),
            },
        );
        // SCAN(0, A1:A3, LAMBDA(acc, x, acc+x)) → [1, 3, 6] (cumulative sum)
        assert_eq!(
            calc("=SCAN(0, A1:A3, LAMBDA(acc, x, acc+x))", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(3),
                Variant::Integer(6)
            ])
        );
    }

    #[test]
    fn test_let_with_map() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(10),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(20),
            },
        );
        // LET(factor, 3, MAP(A1:A2, LAMBDA(x, x*factor))) → [30, 60]
        assert_eq!(
            calc("=LET(factor, 3, MAP(A1:A2, LAMBDA(x, x*factor)))", &c),
            Variant::Array(vec![Variant::Integer(30), Variant::Integer(60)])
        );
    }

    #[test]
    fn test_offset_negative_ref_error() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(10),
            },
        );
        // OFFSET(A1, -1, 0) goes to row 0 → #REF!
        assert_eq!(
            calc("=OFFSET(A1,-1,0)", &c),
            Variant::Error(ExcelError::Ref)
        );
        // OFFSET(A1, 0, -1) goes to col 0 → #REF!
        assert_eq!(
            calc("=OFFSET(A1,0,-1)", &c),
            Variant::Error(ExcelError::Ref)
        );
    }

    #[test]
    fn test_vlookup_col_index_bounds() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        c.insert(
            (1, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        c.insert(
            (1, 3),
            CellContent {
                formula: None,
                value: Variant::Integer(3),
            },
        );
        // col_index=0 → #VALUE!
        assert_eq!(
            calc("=VLOOKUP(1,A1:C1,0,TRUE)", &c),
            Variant::Error(ExcelError::Value)
        );
        // col_index > width (3) → #REF!
        assert_eq!(
            calc("=VLOOKUP(1,A1:C1,5,TRUE)", &c),
            Variant::Error(ExcelError::Ref)
        );
        // col_index within range → normal
        assert_eq!(calc("=VLOOKUP(1,A1:C1,2,TRUE)", &c), Variant::Integer(2));
        assert_eq!(
            calc("=VLOOKUP(1,A1:C1,1.5,TRUE)", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_hlookup_row_index_bounds() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Integer(2)),
            ((2, 1), Variant::Str("one".into())),
            ((2, 2), Variant::Str("two".into())),
        ]);
        assert_eq!(
            calc("=HLOOKUP(2,A1:B2,2,FALSE)", &c),
            Variant::Str("two".into())
        );
        assert_eq!(
            calc("=HLOOKUP(1,A1:B2,0,FALSE)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=HLOOKUP(1,A1:B2,3,FALSE)", &c),
            Variant::Error(ExcelError::Ref)
        );
        assert_eq!(
            calc("=HLOOKUP(1,A1:B2,1.5,FALSE)", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_approximate_lookup_requires_sorted_keys() {
        let vlookup = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Str("one".into())),
            ((2, 1), Variant::Integer(3)),
            ((2, 2), Variant::Str("three".into())),
            ((3, 1), Variant::Integer(2)),
            ((3, 2), Variant::Str("two".into())),
        ]);
        assert_eq!(
            calc("=VLOOKUP(2,A1:B3,2,TRUE)", &vlookup),
            Variant::Error(ExcelError::NA)
        );

        let hlookup = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Integer(3)),
            ((1, 3), Variant::Integer(2)),
            ((2, 1), Variant::Str("one".into())),
            ((2, 2), Variant::Str("three".into())),
            ((2, 3), Variant::Str("two".into())),
        ]);
        assert_eq!(
            calc("=HLOOKUP(2,A1:C2,2,TRUE)", &hlookup),
            Variant::Error(ExcelError::NA)
        );
    }

    #[test]
    fn test_wildcard_match_many_stars() {
        let c = HashMap::new();
        // Many-star pattern must complete without stack overflow or timeout
        assert_eq!(
            calc("=COUNTIF(A1:A1,\"*****hello\")", &c),
            Variant::Integer(0)
        );
        // Correct wildcard behavior
        let mut c2 = HashMap::new();
        c2.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Str("hello world".into()),
            },
        );
        assert_eq!(calc("=COUNTIF(A1:A1,\"*world\")", &c2), Variant::Integer(1));
        assert_eq!(calc("=COUNTIF(A1:A1,\"*xyz*\")", &c2), Variant::Integer(0));
    }

    #[test]
    fn test_offset() {
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(10),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(20),
            },
        );
        c.insert(
            (1, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(30),
            },
        );
        c.insert(
            (2, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(40),
            },
        );
        c.insert(
            (3, 3),
            CellContent {
                formula: None,
                value: Variant::Str("far".into()),
            },
        );
        // OFFSET(A1, 1, 0) → A2 = 20
        assert_eq!(calc("=OFFSET(A1, 1, 0)", &c), Variant::Integer(20));
        // OFFSET(A1, 0, 1) → B1 = 30
        assert_eq!(calc("=OFFSET(A1, 0, 1)", &c), Variant::Integer(30));
        // OFFSET(A1, 2, 2) → C3 = "far"
        assert_eq!(calc("=OFFSET(A1, 2, 2)", &c), Variant::Str("far".into()));
        // OFFSET with negative offset: A2 → 1 row up = A1 = 10
        assert_eq!(calc("=OFFSET(A2, -1, 0)", &c), Variant::Integer(10));
        assert_eq!(
            calc("=OFFSET(A1,0,0,2,2)", &c),
            Variant::Array(vec![
                Variant::Integer(10),
                Variant::Integer(30),
                Variant::Integer(20),
                Variant::Integer(40),
            ])
        );
        assert_eq!(calc("=SUM(OFFSET(A1,0,0,2,2))", &c), Variant::Integer(100));
        assert_eq!(calc("=SUM(OFFSET(A1:B2,0,0))", &c), Variant::Integer(100));
    }

    #[test]
    fn test_workday() {
        let c = HashMap::new();
        // DATE(2024,1,1) = Monday. 5 workdays later = Friday Jan 5
        // DATE(2024,1,1) serial: we use calc which returns Variant::Date; compare via float
        let start = calc("=DATE(2024,1,1)", &c);
        let result = calc("=WORKDAY(DATE(2024,1,1),5)", &c);
        let fri = calc("=DATE(2024,1,8)", &c); // Mon+5 workdays = Mon Jan 8
        assert_eq!(result, fri);
        // -1 workday from Monday = previous Friday
        let prev_fri = calc("=DATE(2023,12,29)", &c);
        assert_eq!(calc("=WORKDAY(DATE(2024,1,1),-1)", &c), prev_fri);
        // 0 days = same day
        assert_eq!(calc("=WORKDAY(DATE(2024,1,1),0)", &c), start);
    }

    #[test]
    fn test_pmt() {
        let c = HashMap::new();
        // PMT(5%/12, 60 months, $10,000) ≈ -$188.71
        let result = match calc("=PMT(0.05/12,60,10000)", &c) {
            Variant::Float(f) => f,
            Variant::Integer(i) => i as f64,
            other => panic!("expected float, got {:?}", other),
        };
        assert!(
            (result + 188.71).abs() < 0.01,
            "PMT ≈ -188.71, got {}",
            result
        );
        // rate=0: PMT = -pv/nper
        assert_eq!(calc("=PMT(0,10,1000)", &c), Variant::Float(-100.0));
    }

    #[test]
    fn test_ispmt() {
        let c = HashMap::new();
        match calc("=ISPMT(0.1,1,3,800000)", &c) {
            Variant::Float(value) => assert!((value + 53333.3333333333).abs() < 1e-8),
            other => panic!("ISPMT: {:?}", other),
        }
        assert_eq!(
            calc("=ISPMT(0.1,0,3,800000)", &c),
            Variant::Error(ExcelError::Num)
        );
    }

    #[test]
    fn test_textsplit() {
        let c = HashMap::new();
        assert_eq!(
            calc("=TEXTSPLIT(\"a,b,c\",\",\")", &c),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("b".into()),
                Variant::Str("c".into())
            ])
        );
        // ignore_empty
        assert_eq!(
            calc("=TEXTSPLIT(\"a,,b\",\",\",\"\",TRUE)", &c),
            Variant::Array(vec![Variant::Str("a".into()), Variant::Str("b".into())])
        );
        // single result → scalar
        assert_eq!(
            calc("=TEXTSPLIT(\"hello\",\",\")", &c),
            Variant::Str("hello".into())
        );
        assert_eq!(
            calc("=TEXTSPLIT(\"a,b;c\",\",\",\";\",FALSE,0,\"_\")", &c),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("b".into()),
                Variant::Str("c".into()),
                Variant::Str("_".into())
            ])
        );
        assert_eq!(
            calc("=TEXTSPLIT(\"a,bXc,d\",\",\",\"x\",FALSE,1,\"_\")", &c),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("b".into()),
                Variant::Str("c".into()),
                Variant::Str("d".into())
            ])
        );
        // Array delimiters are accepted from a dynamic-array expression.
        assert_eq!(
            calc("=TEXTSPLIT(\"a,b;c\",VSTACK(\",\",\";\"),\"\",FALSE)", &c),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("b".into()),
                Variant::Str("c".into())
            ])
        );
        // An omitted/empty column delimiter produces a single column.
        assert_eq!(
            calc("=TEXTSPLIT(\"a;b\",\"\",\";\")", &c),
            Variant::Array(vec![Variant::Str("a".into()), Variant::Str("b".into())])
        );
        assert_eq!(
            calc(
                "=TEXTSPLIT(\"a,b;c,d|e,f\",\",\",VSTACK(\";\",\"|\"),FALSE)",
                &c
            ),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("b".into()),
                Variant::Str("c".into()),
                Variant::Str("d".into()),
                Variant::Str("e".into()),
                Variant::Str("f".into())
            ])
        );
    }

    #[test]
    fn test_filterxml() {
        let c = HashMap::new();
        assert_eq!(
            calc(
                "=FILTERXML(\"<root><item id=\"\"a\"\">A</item><item id=\"\"b\"\">B</item></root>\",\"//item\")",
                &c
            ),
            Variant::Array(vec![Variant::Str("A".into()), Variant::Str("B".into())])
        );
        assert_eq!(
            calc(
                "=FILTERXML(\"<root><item id=\"\"a\"\">A</item><item id=\"\"b\"\">B</item></root>\",\"//item/@id\")",
                &c
            ),
            Variant::Array(vec![Variant::Str("a".into()), Variant::Str("b".into())])
        );
        assert_eq!(
            calc(
                "=FILTERXML(\"<root><item>A</item></root>\",\"/root/item\")",
                &c
            ),
            Variant::Str("A".into())
        );
        assert_eq!(
            calc("=FILTERXML(\"<root>\",\"//item\")", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=FILTERXML(\"<root/>\",\"//item\")", &c),
            Variant::Error(ExcelError::NA)
        );
    }

    #[test]
    fn test_textbefore_textafter() {
        let c = HashMap::new();
        assert_eq!(
            calc("=TEXTBEFORE(\"hello world\",\" \")", &c),
            Variant::Str("hello".into())
        );
        assert_eq!(
            calc("=TEXTAFTER(\"hello world\",\" \")", &c),
            Variant::Str("world".into())
        );
        // instance_num = 2
        assert_eq!(
            calc("=TEXTBEFORE(\"a-b-c\",\"-\",2)", &c),
            Variant::Str("a-b".into())
        );
        assert_eq!(
            calc("=TEXTAFTER(\"a-b-c\",\"-\",2)", &c),
            Variant::Str("c".into())
        );
        // negative instance_num (from end)
        assert_eq!(
            calc("=TEXTBEFORE(\"a-b-c\",\"-\",-1)", &c),
            Variant::Str("a-b".into())
        );
        // not found → #N/A
        assert_eq!(
            calc("=TEXTBEFORE(\"hello\",\",\")", &c),
            Variant::Error(ExcelError::NA)
        );
        assert_eq!(
            calc("=TEXTBEFORE(\"Socrates\",\" \",1,0,1)", &c),
            Variant::Str("Socrates".into())
        );
        assert_eq!(
            calc("=TEXTAFTER(\"Socrates\",\" \",1,0,1)", &c),
            Variant::Str(String::new())
        );
        assert_eq!(
            calc("=TEXTBEFORE(\"hello\",\"\",-1)", &c),
            Variant::Str("hello".into())
        );
        assert_eq!(
            calc("=TEXTAFTER(\"hello\",\"\")", &c),
            Variant::Str("hello".into())
        );
        assert_eq!(
            calc("=TEXTBEFORE(\"hello\",\",\",1,0,0,\"missing\")", &c),
            Variant::Str("missing".into())
        );
    }

    #[test]
    fn test_valuetotext() {
        let c = HashMap::new();
        assert_eq!(calc("=VALUETOTEXT(42)", &c), Variant::Str("42".into()));
        assert_eq!(calc("=VALUETOTEXT(3.14)", &c), Variant::Str("3.14".into()));
        assert_eq!(calc("=VALUETOTEXT(TRUE)", &c), Variant::Str("TRUE".into()));
        assert_eq!(calc("=VALUETOTEXT(\"hi\")", &c), Variant::Str("hi".into()));
        assert_eq!(
            calc("=VALUETOTEXT(\"hi\",1)", &c),
            Variant::Str("\"hi\"".into())
        );
    }

    #[test]
    fn test_take_drop() {
        let c = HashMap::new();
        // TAKE first 3 of [1,2,3,4,5]
        assert_eq!(
            calc("=TAKE(SEQUENCE(5),3)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3)
            ])
        );
        // TAKE last 2
        assert_eq!(
            calc("=TAKE(SEQUENCE(5),-2)", &c),
            Variant::Array(vec![Variant::Integer(4), Variant::Integer(5)])
        );
        // DROP first 2
        assert_eq!(
            calc("=DROP(SEQUENCE(5),2)", &c),
            Variant::Array(vec![
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Integer(5)
            ])
        );
        // DROP last 3
        assert_eq!(
            calc("=DROP(SEQUENCE(5),-3)", &c),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(2)])
        );
        // TAKE more than available → all
        assert_eq!(
            calc("=TAKE(SEQUENCE(3),10)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3)
            ])
        );
        assert!(evaluate(&fparse("=TAKE(SEQUENCE(3),1.5)").unwrap(), &c).is_err());
        assert!(evaluate(&fparse("=DROP(SEQUENCE(3),0)").unwrap(), &c).is_err());
    }

    #[test]
    fn test_aggregate_functions_flatten_formula_arrays() {
        let mut cells = HashMap::new();
        cells.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        cells.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        cells.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(3),
            },
        );
        assert_eq!(calc("=SUM(TRANSPOSE(A1:A3))", &cells), Variant::Integer(6));
        assert_eq!(calc("=SUM(SEQUENCE(3))", &cells), Variant::Integer(6));
    }

    #[test]
    fn test_vstack_hstack() {
        let c = HashMap::new();
        // VSTACK concatenates arrays
        assert_eq!(
            calc("=VSTACK(SEQUENCE(3),SEQUENCE(2))", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Integer(1),
                Variant::Integer(2),
            ])
        );
        // HSTACK preserves the row shape of its 2-row inputs.
        assert_eq!(
            calc("=HSTACK(SEQUENCE(2),SEQUENCE(2))", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(2),
            ])
        );
        assert_eq!(
            calc("=HSTACK(SEQUENCE(2),SEQUENCE(3))", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(2),
                Variant::Error(ExcelError::NA),
                Variant::Integer(3),
            ])
        );
        assert_eq!(
            calc("=VSTACK(SEQUENCE(1,2),SEQUENCE(2,1))", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(1),
                Variant::Error(ExcelError::NA),
                Variant::Integer(2),
                Variant::Error(ExcelError::NA),
            ])
        );
        assert_eq!(
            calc("=EXPAND(SEQUENCE(2,2),3,4,0)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(0),
                Variant::Integer(0),
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Integer(0),
                Variant::Integer(0),
                Variant::Integer(0),
                Variant::Integer(0),
                Variant::Integer(0),
                Variant::Integer(0),
            ])
        );
        assert!(matches!(
            calc("=EXPAND(SEQUENCE(2),1,1,0)", &c),
            Variant::Error(ExcelError::Num)
        ));
    }

    #[test]
    fn test_trimrange() {
        let c = cells_from(&[((2, 2), Variant::Integer(1)), ((3, 3), Variant::Integer(2))]);
        assert_eq!(
            calc("=TRIMRANGE(A1:C4)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Empty,
                Variant::Empty,
                Variant::Integer(2),
            ])
        );
        assert_eq!(
            calc("=TRIMRANGE(A1:C4,0,0)", &c),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Empty,
                Variant::Empty,
                Variant::Empty,
                Variant::Integer(1),
                Variant::Empty,
                Variant::Empty,
                Variant::Empty,
                Variant::Integer(2),
                Variant::Empty,
                Variant::Empty,
                Variant::Empty,
            ])
        );
        assert!(evaluate(&fparse("=TRIMRANGE(A1:C4,4,3)").unwrap(), &c).is_err());
    }

    #[test]
    fn test_groupby() {
        let c = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((2, 1), Variant::Str("b".into())),
            ((3, 1), Variant::Str("a".into())),
            ((4, 1), Variant::Str("b".into())),
            ((1, 2), Variant::Integer(1)),
            ((2, 2), Variant::Integer(2)),
            ((3, 2), Variant::Integer(3)),
            ((4, 2), Variant::Integer(4)),
        ]);
        assert_eq!(
            calc("=GROUPBY(A1:A4,B1:B4,\"SUM\")", &c),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Integer(4),
                Variant::Str("b".into()),
                Variant::Integer(6),
            ])
        );
        assert_eq!(
            calc("=GROUPBY(A1:A4,B1:B4,\"COUNT\")", &c),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Integer(2),
                Variant::Str("b".into()),
                Variant::Integer(2),
            ])
        );
        assert_eq!(
            calc("=GROUPBY(A1:A4,B1:B4,LAMBDA(values,SUM(values)))", &c),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Integer(4),
                Variant::Str("b".into()),
                Variant::Integer(6),
            ])
        );
        assert_eq!(
            calc(
                "=GROUPBY(A1:A4,B1:B4,HSTACK(LAMBDA(values,SUM(values)),LAMBDA(values,MAX(values))))",
                &c
            ),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Integer(4),
                Variant::Integer(3),
                Variant::Str("b".into()),
                Variant::Integer(6),
                Variant::Integer(4),
            ])
        );
        assert_eq!(
            calc(
                "=GROUPBY(A1:A4,B1:B4,LAMBDA(values,total,SUM(values)/SUM(total)))",
                &c
            ),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Float(0.4),
                Variant::Str("b".into()),
                Variant::Float(0.6),
            ])
        );
        assert_eq!(
            calc("=GROUPBY(A1:A4,B1:B4,HSTACK(SUM,AVERAGE))", &c),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Integer(4),
                Variant::Float(2.0),
                Variant::Str("b".into()),
                Variant::Integer(6),
                Variant::Float(3.0),
            ])
        );
        let mut c_multi = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((2, 1), Variant::Str("a".into())),
            ((3, 1), Variant::Str("b".into())),
            ((4, 1), Variant::Str("b".into())),
            ((1, 2), Variant::Str("x".into())),
            ((2, 2), Variant::Str("y".into())),
            ((3, 2), Variant::Str("x".into())),
            ((4, 2), Variant::Str("y".into())),
            ((1, 3), Variant::Integer(1)),
            ((2, 3), Variant::Integer(2)),
            ((3, 3), Variant::Integer(3)),
            ((4, 3), Variant::Integer(4)),
        ]);
        for (row, value) in [(1, 10), (2, 20), (3, 30), (4, 40)] {
            c_multi.insert(
                (row, 4),
                CellContent {
                    formula: None,
                    value: Variant::Integer(value),
                },
            );
        }
        assert_eq!(
            calc("=GROUPBY(A1:B4,C1:C4,\"SUM\")", &c_multi),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("x".into()),
                Variant::Integer(1),
                Variant::Str("a".into()),
                Variant::Str("y".into()),
                Variant::Integer(2),
                Variant::Str("b".into()),
                Variant::Str("x".into()),
                Variant::Integer(3),
                Variant::Str("b".into()),
                Variant::Str("y".into()),
                Variant::Integer(4),
            ])
        );
        assert_eq!(
            calc("=GROUPBY(A1:B4,C1:D4,\"SUM\")", &c_multi),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("x".into()),
                Variant::Integer(1),
                Variant::Integer(10),
                Variant::Str("a".into()),
                Variant::Str("y".into()),
                Variant::Integer(2),
                Variant::Integer(20),
                Variant::Str("b".into()),
                Variant::Str("x".into()),
                Variant::Integer(3),
                Variant::Integer(30),
                Variant::Str("b".into()),
                Variant::Str("y".into()),
                Variant::Integer(4),
                Variant::Integer(40),
            ])
        );
        assert_eq!(
            calc("=GROUPBY(A1:B4,C1:D4,\"SUM\",0,0,VSTACK(1,-1))", &c_multi),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("y".into()),
                Variant::Integer(2),
                Variant::Integer(20),
                Variant::Str("a".into()),
                Variant::Str("x".into()),
                Variant::Integer(1),
                Variant::Integer(10),
                Variant::Str("b".into()),
                Variant::Str("y".into()),
                Variant::Integer(4),
                Variant::Integer(40),
                Variant::Str("b".into()),
                Variant::Str("x".into()),
                Variant::Integer(3),
                Variant::Integer(30),
            ])
        );
        assert_eq!(
            calc("=GROUPBY(A1:B4,C1:D4,\"SUM\",0,2)", &c_multi),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("x".into()),
                Variant::Integer(1),
                Variant::Integer(10),
                Variant::Str("a".into()),
                Variant::Str("y".into()),
                Variant::Integer(2),
                Variant::Integer(20),
                Variant::Str("a".into()),
                Variant::Str("Subtotal".into()),
                Variant::Integer(3),
                Variant::Integer(30),
                Variant::Str("b".into()),
                Variant::Str("x".into()),
                Variant::Integer(3),
                Variant::Integer(30),
                Variant::Str("b".into()),
                Variant::Str("y".into()),
                Variant::Integer(4),
                Variant::Integer(40),
                Variant::Str("b".into()),
                Variant::Str("Subtotal".into()),
                Variant::Integer(7),
                Variant::Integer(70),
                Variant::Str("Total".into()),
                Variant::Empty,
                Variant::Integer(10),
                Variant::Integer(100),
            ])
        );
        assert_eq!(
            calc("=GROUPBY(A1:B4,C1:D4,\"SUM\",0,-2)", &c_multi),
            Variant::Array(vec![
                Variant::Str("Total".into()),
                Variant::Empty,
                Variant::Integer(10),
                Variant::Integer(100),
                Variant::Str("a".into()),
                Variant::Str("Subtotal".into()),
                Variant::Integer(3),
                Variant::Integer(30),
                Variant::Str("a".into()),
                Variant::Str("x".into()),
                Variant::Integer(1),
                Variant::Integer(10),
                Variant::Str("a".into()),
                Variant::Str("y".into()),
                Variant::Integer(2),
                Variant::Integer(20),
                Variant::Str("b".into()),
                Variant::Str("Subtotal".into()),
                Variant::Integer(7),
                Variant::Integer(70),
                Variant::Str("b".into()),
                Variant::Str("x".into()),
                Variant::Integer(3),
                Variant::Integer(30),
                Variant::Str("b".into()),
                Variant::Str("y".into()),
                Variant::Integer(4),
                Variant::Integer(40),
            ])
        );
        let interleaved = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((2, 1), Variant::Str("b".into())),
            ((3, 1), Variant::Str("a".into())),
            ((4, 1), Variant::Str("b".into())),
            ((1, 2), Variant::Str("x".into())),
            ((2, 2), Variant::Str("x".into())),
            ((3, 2), Variant::Str("y".into())),
            ((4, 2), Variant::Str("y".into())),
            ((1, 3), Variant::Integer(1)),
            ((2, 3), Variant::Integer(2)),
            ((3, 3), Variant::Integer(3)),
            ((4, 3), Variant::Integer(4)),
        ]);
        assert_eq!(
            calc("=GROUPBY(A1:B4,C1:C4,\"SUM\",0,2)", &interleaved),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("x".into()),
                Variant::Integer(1),
                Variant::Str("b".into()),
                Variant::Str("x".into()),
                Variant::Integer(2),
                Variant::Str("a".into()),
                Variant::Str("y".into()),
                Variant::Integer(3),
                Variant::Str("a".into()),
                Variant::Str("Subtotal".into()),
                Variant::Integer(4),
                Variant::Str("b".into()),
                Variant::Str("y".into()),
                Variant::Integer(4),
                Variant::Str("b".into()),
                Variant::Str("Subtotal".into()),
                Variant::Integer(6),
                Variant::Str("Total".into()),
                Variant::Empty,
                Variant::Integer(10),
            ])
        );
        let c_headers = cells_from(&[
            ((1, 1), Variant::Str("Category".into())),
            ((2, 1), Variant::Str("a".into())),
            ((3, 1), Variant::Str("b".into())),
            ((4, 1), Variant::Str("a".into())),
            ((1, 2), Variant::Str("Amount".into())),
            ((2, 2), Variant::Integer(1)),
            ((3, 2), Variant::Integer(2)),
            ((4, 2), Variant::Integer(3)),
        ]);
        assert_eq!(
            calc("=GROUPBY(A1:A4,B1:B4,SUM,3,1)", &c_headers),
            Variant::Array(vec![
                Variant::Str("Category".into()),
                Variant::Str("Amount".into()),
                Variant::Str("a".into()),
                Variant::Integer(4),
                Variant::Str("b".into()),
                Variant::Integer(2),
                Variant::Str("Total".into()),
                Variant::Integer(6),
            ])
        );
        assert_eq!(
            calc("=GROUPBY(A1:A4,B1:B4,SUM,2,0)", &c),
            Variant::Array(vec![
                Variant::Str("Field1".into()),
                Variant::Str("Values".into()),
                Variant::Str("a".into()),
                Variant::Integer(4),
                Variant::Str("b".into()),
                Variant::Integer(6),
            ])
        );
    }

    #[test]
    fn test_pivotby() {
        let c = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((2, 1), Variant::Str("a".into())),
            ((3, 1), Variant::Str("b".into())),
            ((4, 1), Variant::Str("b".into())),
            ((1, 2), Variant::Str("x".into())),
            ((2, 2), Variant::Str("y".into())),
            ((3, 2), Variant::Str("x".into())),
            ((4, 2), Variant::Str("y".into())),
            ((1, 3), Variant::Integer(1)),
            ((2, 3), Variant::Integer(2)),
            ((3, 3), Variant::Integer(3)),
            ((4, 3), Variant::Integer(4)),
            ((1, 4), Variant::Integer(10)),
            ((2, 4), Variant::Integer(20)),
            ((3, 4), Variant::Integer(30)),
            ((4, 4), Variant::Integer(40)),
            ((1, 5), Variant::Boolean(true)),
            ((2, 5), Variant::Boolean(false)),
            ((3, 5), Variant::Boolean(true)),
            ((4, 5), Variant::Boolean(false)),
        ]);
        assert_eq!(
            calc("=PIVOTBY(A1:A4,B1:B4,C1:C4,\"SUM\")", &c),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("y".into()),
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Integer(4),
            ])
        );
        let c_with_headers = cells_from(&[
            ((1, 1), Variant::Str("Category".into())),
            ((2, 1), Variant::Str("a".into())),
            ((3, 1), Variant::Str("a".into())),
            ((4, 1), Variant::Str("b".into())),
            ((5, 1), Variant::Str("b".into())),
            ((1, 2), Variant::Str("Year".into())),
            ((2, 2), Variant::Str("x".into())),
            ((3, 2), Variant::Str("y".into())),
            ((4, 2), Variant::Str("x".into())),
            ((5, 2), Variant::Str("y".into())),
            ((1, 3), Variant::Str("Amount".into())),
            ((2, 3), Variant::Integer(1)),
            ((3, 3), Variant::Integer(2)),
            ((4, 3), Variant::Integer(3)),
            ((5, 3), Variant::Integer(4)),
        ]);
        assert_eq!(
            calc("=PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,1)", &c_with_headers),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Integer(4),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:A5,B1:B5,C1:C5,SUM)", &c_with_headers),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Integer(4),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:A5,B1:B5,C1:C5,SUM,3)", &c_with_headers),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("y".into()),
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Integer(4),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:A4,B1:B4,C1:D4,\"SUM\")", &c),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("x".into()),
                Variant::Str("y".into()),
                Variant::Str("y".into()),
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(10),
                Variant::Integer(2),
                Variant::Integer(20),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Integer(30),
                Variant::Integer(4),
                Variant::Integer(40),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:A4,B1:B4,C1:C4,\"SUM\",2,0,1,0,1,E1:E4)", &c),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Str("b".into()),
                Variant::Integer(3),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:A4,B1:B4,C1:C4,\"SUM\",2,1,1,1,1)", &c),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("y".into()),
                Variant::Str("Total".into()),
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Integer(7),
                Variant::Str("Total".into()),
                Variant::Integer(4),
                Variant::Integer(6),
                Variant::Integer(10),
            ])
        );
        let composite = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((2, 1), Variant::Str("a".into())),
            ((3, 1), Variant::Str("b".into())),
            ((4, 1), Variant::Str("b".into())),
            ((1, 2), Variant::Str("x".into())),
            ((2, 2), Variant::Str("y".into())),
            ((3, 2), Variant::Str("x".into())),
            ((4, 2), Variant::Str("y".into())),
            ((1, 3), Variant::Str("x".into())),
            ((2, 3), Variant::Str("y".into())),
            ((3, 3), Variant::Str("x".into())),
            ((4, 3), Variant::Str("y".into())),
            ((1, 4), Variant::Integer(10)),
            ((2, 4), Variant::Integer(20)),
            ((3, 4), Variant::Integer(30)),
            ((4, 4), Variant::Integer(40)),
        ]);
        assert_eq!(
            calc("=PIVOTBY(A1:B4,C1:C4,D1:D4,\"SUM\")", &composite),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("y".into()),
                Variant::Str("a".into()),
                Variant::Str("x".into()),
                Variant::Integer(10),
                Variant::Integer(0),
                Variant::Str("a".into()),
                Variant::Str("y".into()),
                Variant::Integer(0),
                Variant::Integer(20),
                Variant::Str("b".into()),
                Variant::Str("x".into()),
                Variant::Integer(30),
                Variant::Integer(0),
                Variant::Str("b".into()),
                Variant::Str("y".into()),
                Variant::Integer(0),
                Variant::Integer(40),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:A4,B1:B4,C1:C4,\"SUM\",0)", &c),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Integer(4),
            ])
        );
        assert_eq!(
            calc(
                "=PIVOTBY(A1:A4,B1:B4,C1:C4,PERCENTOF,0,1,1,0,1,SEQUENCE(4,1),2)",
                &c
            ),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Float(0.1),
                Variant::Float(0.2),
                Variant::Float(0.3),
                Variant::Str("b".into()),
                Variant::Float(0.3),
                Variant::Float(0.4),
                Variant::Float(0.7),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:A4,B1:B4,C1:C4,\"SUM\",2,0,-2,0,1)", &c),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("y".into()),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(2),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:A4,B1:B4,C1:C4,\"SUM\",2,0,1,0,-2)", &c),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Str("y".into()),
                Variant::Str("x".into()),
                Variant::Str("a".into()),
                Variant::Integer(2),
                Variant::Integer(1),
                Variant::Str("b".into()),
                Variant::Integer(4),
                Variant::Integer(3),
            ])
        );
        assert_eq!(
            calc(
                "=PIVOTBY(A1:B4,C1:C4,D1:D4,\"SUM\",2,0,HSTACK(-2,1),0)",
                &composite
            ),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("y".into()),
                Variant::Str("a".into()),
                Variant::Str("y".into()),
                Variant::Integer(0),
                Variant::Integer(20),
                Variant::Str("b".into()),
                Variant::Str("y".into()),
                Variant::Integer(0),
                Variant::Integer(40),
                Variant::Str("a".into()),
                Variant::Str("x".into()),
                Variant::Integer(10),
                Variant::Integer(0),
                Variant::Str("b".into()),
                Variant::Str("x".into()),
                Variant::Integer(30),
                Variant::Integer(0),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:B4,C1:C4,D1:D4,\"SUM\",0,2,1,0)", &composite),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Str("x".into()),
                Variant::Integer(10),
                Variant::Integer(0),
                Variant::Integer(10),
                Variant::Str("a".into()),
                Variant::Str("y".into()),
                Variant::Integer(0),
                Variant::Integer(20),
                Variant::Integer(20),
                Variant::Str("a".into()),
                Variant::Str("Subtotal".into()),
                Variant::Integer(10),
                Variant::Integer(20),
                Variant::Integer(30),
                Variant::Str("b".into()),
                Variant::Str("x".into()),
                Variant::Integer(30),
                Variant::Integer(0),
                Variant::Integer(30),
                Variant::Str("b".into()),
                Variant::Str("y".into()),
                Variant::Integer(0),
                Variant::Integer(40),
                Variant::Integer(40),
                Variant::Str("b".into()),
                Variant::Str("Subtotal".into()),
                Variant::Integer(30),
                Variant::Integer(40),
                Variant::Integer(70),
            ])
        );
        let col_composite = cells_from(&[
            ((1, 1), Variant::Str("a".into())),
            ((2, 1), Variant::Str("a".into())),
            ((3, 1), Variant::Str("a".into())),
            ((4, 1), Variant::Str("a".into())),
            ((1, 2), Variant::Str("p".into())),
            ((2, 2), Variant::Str("p".into())),
            ((3, 2), Variant::Str("q".into())),
            ((4, 2), Variant::Str("q".into())),
            ((1, 3), Variant::Str("x".into())),
            ((2, 3), Variant::Str("y".into())),
            ((3, 3), Variant::Str("x".into())),
            ((4, 3), Variant::Str("y".into())),
            ((1, 4), Variant::Integer(1)),
            ((2, 4), Variant::Integer(2)),
            ((3, 4), Variant::Integer(3)),
            ((4, 4), Variant::Integer(4)),
        ]);
        assert_eq!(
            calc(
                "=PIVOTBY(A1:A4,B1:C4,D1:D4,\"SUM\",0,0,1,2)",
                &col_composite
            ),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Integer(7),
                Variant::Empty,
                Variant::Empty,
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Integer(7),
                Variant::Integer(10),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:A4,B1:B4,C1:C4,HSTACK(SUM,AVERAGE))", &c),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("x".into()),
                Variant::Str("y".into()),
                Variant::Str("y".into()),
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Float(1.0),
                Variant::Integer(2),
                Variant::Float(2.0),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Float(3.0),
                Variant::Integer(4),
                Variant::Float(4.0),
            ])
        );
        assert_eq!(
            calc(
                "=PIVOTBY(A1:A4,B1:B4,C1:C4,HSTACK(LAMBDA(x,SUM(x)),LAMBDA(x,MAX(x))))",
                &c
            ),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("x".into()),
                Variant::Str("y".into()),
                Variant::Str("y".into()),
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(2),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Integer(4),
            ])
        );
        assert_eq!(
            calc(
                "=PIVOTBY(A1:A4,B1:B4,C1:C4,LAMBDA(subset,total,SUM(subset)/SUM(total)),0,1,1,0,1,SEQUENCE(4,1),2)",
                &c
            ),
            Variant::Array(vec![
                Variant::Str("a".into()),
                Variant::Float(0.1),
                Variant::Float(0.2),
                Variant::Float(0.3),
                Variant::Str("b".into()),
                Variant::Float(0.3),
                Variant::Float(0.4),
                Variant::Float(0.7),
            ])
        );
        assert_eq!(
            calc("=PIVOTBY(A1:A4,B1:B4,C1:C4,VSTACK(SUM,AVERAGE))", &c),
            Variant::Array(vec![
                Variant::Empty,
                Variant::Str("x".into()),
                Variant::Str("y".into()),
                Variant::Str("a".into()),
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Str("a".into()),
                Variant::Float(1.0),
                Variant::Float(2.0),
                Variant::Str("b".into()),
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Str("b".into()),
                Variant::Float(3.0),
                Variant::Float(4.0),
            ])
        );
    }

    #[test]
    fn test_choosecols_chooserows() {
        let c = HashMap::new();
        // CHOOSECOLS(SEQUENCE(5), 1, 3, 5) → [1, 3, 5]
        assert_eq!(
            calc("=CHOOSECOLS(SEQUENCE(5),1,3,5)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(3),
                Variant::Integer(5)
            ])
        );
        // Negative index: -1 = last element
        assert_eq!(calc("=CHOOSECOLS(SEQUENCE(5),-1)", &c), Variant::Integer(5));
        // CHOOSEROWS (same logic)
        assert_eq!(
            calc("=CHOOSEROWS(SEQUENCE(4),2,4)", &c),
            Variant::Array(vec![Variant::Integer(2), Variant::Integer(4)])
        );
        // Out-of-bounds → #VALUE!
        assert_eq!(
            calc("=CHOOSECOLS(SEQUENCE(3),5)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=CHOOSECOLS(SEQUENCE(2,3),3,1)", &c),
            Variant::Array(vec![
                Variant::Integer(3),
                Variant::Integer(1),
                Variant::Integer(6),
                Variant::Integer(4),
            ])
        );
        assert_eq!(
            calc("=CHOOSEROWS(SEQUENCE(2,3),2)", &c),
            Variant::Array(vec![
                Variant::Integer(4),
                Variant::Integer(5),
                Variant::Integer(6),
            ])
        );
        assert_eq!(
            calc("=CHOOSECOLS(SEQUENCE(2,3),1.5)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=CHOOSEROWS(SEQUENCE(3),1.5)", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_combin() {
        let c = HashMap::new();
        assert_eq!(calc("=COMBIN(4,2)", &c), Variant::Integer(6));
        assert_eq!(calc("=COMBIN(10,3)", &c), Variant::Integer(120));
        assert_eq!(calc("=COMBIN(0,0)", &c), Variant::Integer(1));
        assert_eq!(calc("=COMBIN(5,0)", &c), Variant::Integer(1));
        assert_eq!(calc("=COMBIN(5,5)", &c), Variant::Integer(1));
        // Error cases
        assert_eq!(calc("=COMBIN(3,5)", &c), Variant::Error(ExcelError::Num));
        assert_eq!(calc("=COMBIN(-1,0)", &c), Variant::Error(ExcelError::Num));
        assert_eq!(calc("=COMBIN(5,-1)", &c), Variant::Error(ExcelError::Num));
    }

    #[test]
    fn test_dget() {
        // Database (A1:B4):
        //   A1=Name  B1=Score
        //   A2=Alice B2=90
        //   A3=Bob   B3=75
        //   A4=Carol B4=85
        // Criteria (D1:D2):
        //   D1=Name  D2=Alice
        let mut c = HashMap::new();
        // headers
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Str("Name".into()),
            },
        );
        c.insert(
            (1, 2),
            CellContent {
                formula: None,
                value: Variant::Str("Score".into()),
            },
        );
        // data rows
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Str("Alice".into()),
            },
        );
        c.insert(
            (2, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(90),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Str("Bob".into()),
            },
        );
        c.insert(
            (3, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(75),
            },
        );
        c.insert(
            (4, 1),
            CellContent {
                formula: None,
                value: Variant::Str("Carol".into()),
            },
        );
        c.insert(
            (4, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(85),
            },
        );
        // criteria: Name = Alice
        c.insert(
            (1, 4),
            CellContent {
                formula: None,
                value: Variant::Str("Name".into()),
            },
        );
        c.insert(
            (2, 4),
            CellContent {
                formula: None,
                value: Variant::Str("Alice".into()),
            },
        );

        // DGET by field name
        assert_eq!(
            calc("=DGET(A1:B4,\"Score\",D1:D2)", &c),
            Variant::Integer(90)
        );
        // DGET by field index
        assert_eq!(calc("=DGET(A1:B4,2,D1:D2)", &c), Variant::Integer(90));

        // No match → #VALUE!
        c.insert(
            (2, 4),
            CellContent {
                formula: None,
                value: Variant::Str("Zara".into()),
            },
        );
        assert_eq!(
            calc("=DGET(A1:B4,\"Score\",D1:D2)", &c),
            Variant::Error(ExcelError::Value)
        );

        // Multiple matches → #NUM!
        c.insert(
            (2, 4),
            CellContent {
                formula: None,
                value: Variant::Str("Alice".into()),
            },
        );
        c.insert(
            (3, 4),
            CellContent {
                formula: None,
                value: Variant::Str("Bob".into()),
            },
        );
        // criteria D1:D3 with two rows: Alice OR Bob
        c.insert(
            (1, 4),
            CellContent {
                formula: None,
                value: Variant::Str("Name".into()),
            },
        );
        assert_eq!(
            calc("=DGET(A1:B4,\"Score\",D1:D3)", &c),
            Variant::Error(ExcelError::Num)
        );
    }

    #[test]
    fn test_dsum_daverage_dcount_dmax_dmin() {
        // Database (A1:B5):
        //   A1=Name  B1=Score
        //   A2=Alice B2=90
        //   A3=Bob   B3=70
        //   A4=Alice B4=80
        //   A5=Carol B5=60
        // Criteria (D1:D2): Name = Alice
        let mut c = HashMap::new();
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Str("Name".into()),
            },
        );
        c.insert(
            (1, 2),
            CellContent {
                formula: None,
                value: Variant::Str("Score".into()),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Str("Alice".into()),
            },
        );
        c.insert(
            (2, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(90),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Str("Bob".into()),
            },
        );
        c.insert(
            (3, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(70),
            },
        );
        c.insert(
            (4, 1),
            CellContent {
                formula: None,
                value: Variant::Str("Alice".into()),
            },
        );
        c.insert(
            (4, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(80),
            },
        );
        c.insert(
            (5, 1),
            CellContent {
                formula: None,
                value: Variant::Str("Carol".into()),
            },
        );
        c.insert(
            (5, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(60),
            },
        );
        c.insert(
            (1, 4),
            CellContent {
                formula: None,
                value: Variant::Str("Name".into()),
            },
        );
        c.insert(
            (2, 4),
            CellContent {
                formula: None,
                value: Variant::Str("Alice".into()),
            },
        );

        // DSUM: 90 + 80 = 170
        assert_eq!(
            calc("=DSUM(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Integer(170)
        );
        assert_eq!(calc("=DSUM(A1:B5,2,D1:D2)", &c), Variant::Integer(170));
        // DAVERAGE: (90 + 80) / 2 = 85.0
        assert_eq!(
            calc("=DAVERAGE(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Float(85.0)
        );
        // DCOUNT: 2 numeric values
        assert_eq!(
            calc("=DCOUNT(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Integer(2)
        );
        // DCOUNTA: 2 non-empty values
        assert_eq!(
            calc("=DCOUNTA(A1:B5,\"Name\",D1:D2)", &c),
            Variant::Integer(2)
        );
        // DMAX: 90
        assert_eq!(
            calc("=DMAX(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Integer(90)
        );
        // DMIN: 80
        assert_eq!(
            calc("=DMIN(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Integer(80)
        );
        assert_eq!(
            calc("=DPRODUCT(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Integer(7200)
        );
        match calc("=DSTDEV(A1:B5,\"Score\",D1:D2)", &c) {
            Variant::Float(value) => assert!((value - 7.0710678119).abs() < 1e-9),
            other => panic!("DSTDEV unexpected: {:?}", other),
        }
        match calc("=DSTDEVP(A1:B5,\"Score\",D1:D2)", &c) {
            Variant::Float(value) => assert!((value - 5.0).abs() < 1e-9),
            other => panic!("DSTDEVP unexpected: {:?}", other),
        }
        assert_eq!(
            calc("=DVAR(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Float(50.0)
        );
        assert_eq!(
            calc("=DVARP(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Float(25.0)
        );

        // A criteria range containing only its header selects every record.
        assert_eq!(
            calc("=DSUM(A1:B5,\"Score\",D1:D1)", &c),
            Variant::Integer(300)
        );
        assert_eq!(
            calc("=DCOUNT(A1:B5,\"Score\",D1:D1)", &c),
            Variant::Integer(4)
        );
        assert_eq!(
            calc("=DGET(A1:B5,\"Score\",D1:D1)", &c),
            Variant::Error(ExcelError::Num)
        );

        // No matches → DSUM=0, DAVERAGE=#DIV/0!, DCOUNT=0, DMAX=0, DMIN=0
        c.insert(
            (2, 4),
            CellContent {
                formula: None,
                value: Variant::Str("Zara".into()),
            },
        );
        assert_eq!(
            calc("=DSUM(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Integer(0)
        );
        assert_eq!(
            calc("=DAVERAGE(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Error(ExcelError::DivZero)
        );
        assert_eq!(
            calc("=DCOUNT(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Integer(0)
        );
        assert_eq!(
            calc("=DMAX(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Integer(0)
        );
        assert_eq!(
            calc("=DMIN(A1:B5,\"Score\",D1:D2)", &c),
            Variant::Integer(0)
        );
    }

    #[test]
    fn test_statistical_summary_functions() {
        let mut c = HashMap::new();
        assert_eq!(calc("=SUM(\"2\",TRUE)", &c), Variant::Integer(3));
        assert_eq!(calc("=AVERAGE(\"2\",TRUE)", &c), Variant::Float(1.5));
        assert_eq!(calc("=AVERAGE(3,\"2\")", &c), Variant::Float(2.5));
        c.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        c.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Str("2".into()),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Boolean(true),
            },
        );
        assert_eq!(calc("=SUM(A1:A3)", &c), Variant::Integer(1));
        assert_eq!(calc("=AVERAGE(A1:A3)", &c), Variant::Float(1.0));
        assert_eq!(calc("=MODE.SNGL(1,2,2,3)", &c), Variant::Integer(2));
        match calc("=VARA(1,2,TRUE,\"x\")", &c) {
            Variant::Float(value) => assert!((value - 2.0 / 3.0).abs() < 1e-12),
            other => panic!("VARA: {:?}", other),
        }
        match calc("=VARPA(1,2,TRUE,\"x\")", &c) {
            Variant::Float(value) => assert!((value - 0.5).abs() < 1e-12),
            other => panic!("VARPA: {:?}", other),
        }
        match calc("=STDEVA(1,2,TRUE,\"x\")", &c) {
            Variant::Float(value) => assert!((value - (2.0_f64 / 3.0).sqrt()).abs() < 1e-12),
            other => panic!("STDEVA: {:?}", other),
        }
        match calc("=STDEVPA(1,2,TRUE,\"x\")", &c) {
            Variant::Float(value) => assert!((value - 0.5_f64.sqrt()).abs() < 1e-12),
            other => panic!("STDEVPA: {:?}", other),
        }
        for (row, value) in [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 100.0, 200.0]
            .into_iter()
            .enumerate()
        {
            c.insert(
                (row as u32 + 1, 1),
                CellContent {
                    formula: None,
                    value: Variant::Float(value),
                },
            );
        }
        assert_eq!(calc("=RANK.EQ(2,A1:A3,0)", &c), Variant::Integer(2));
        assert_eq!(calc("=TRIMMEAN(A1:A8,0.5)", &c), Variant::Float(4.5));
        c.insert(
            (1, 2),
            CellContent {
                formula: None,
                value: Variant::Float(2.0),
            },
        );
        c.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Float(2.0),
            },
        );
        assert_eq!(calc("=RANK.AVG(2,A1:A3,0)", &c), Variant::Float(1.5));
        for (row, value) in [3.0, 5.0, 100.0].into_iter().enumerate() {
            c.insert(
                (row as u32 + 1, 2),
                CellContent {
                    formula: None,
                    value: Variant::Float(value),
                },
            );
        }
        assert_eq!(
            calc("=FREQUENCY(A1:A8,B1:B3)", &c),
            Variant::Array(vec![
                Variant::Integer(3),
                Variant::Integer(2),
                Variant::Integer(2),
                Variant::Integer(1),
            ])
        );
        match calc("=SKEW(1,2,3,4,5)", &c) {
            Variant::Float(value) => assert!(value.abs() < 1e-12),
            other => panic!("SKEW unexpected: {:?}", other),
        }
        match calc("=SKEW.P(1,2,3,4,5)", &c) {
            Variant::Float(value) => assert!(value.abs() < 1e-12),
            other => panic!("SKEW.P unexpected: {:?}", other),
        }
        match calc("=KURT(1,2,3,4,5)", &c) {
            Variant::Float(value) => assert!((value + 1.2).abs() < 1e-12),
            other => panic!("KURT unexpected: {:?}", other),
        }
    }

    #[test]
    fn test_financial_functions() {
        let c = HashMap::new();

        // FV: invest 1000 at 5%/yr for 3 years, no recurring payment → 1157.625
        match calc("=FV(0.05,3,0,-1000)", &c) {
            Variant::Float(f) => assert!((f - 1157.625).abs() < 0.01),
            other => panic!("FV unexpected: {:?}", other),
        }

        // PV: what is the PV of FV=1157.625 in 3 years at 5%?
        match calc("=PV(0.05,3,0,1157.625)", &c) {
            Variant::Float(f) => assert!((f + 1000.0).abs() < 0.01),
            other => panic!("PV unexpected: {:?}", other),
        }

        // PMT round-trip: 5% for 3 years, pv=1000
        let pmt_val = match calc("=PMT(0.05,3,-1000)", &c) {
            Variant::Float(f) => f,
            other => panic!("PMT unexpected: {:?}", other),
        };
        assert!((pmt_val - 367.2088).abs() < 0.01);

        // NPER: how many periods to pay off pv=1000 at 5% with pmt=-367.21?
        match calc("=NPER(0.05,-367.2088,1000)", &c) {
            Variant::Float(f) => assert!((f - 3.0).abs() < 0.01),
            other => panic!("NPER unexpected: {:?}", other),
        }

        // RATE: 3 periods, pmt=-367.21, pv=1000 → should converge to ~5%
        match calc("=RATE(3,-367.2088,1000)", &c) {
            Variant::Float(f) => assert!((f - 0.05).abs() < 1e-5),
            other => panic!("RATE unexpected: {:?}", other),
        }

        // IPMT: interest portion of period 1 payment on 1000 at 5% for 3 periods
        // period 1 interest = 1000 * 0.05 = 50
        match calc("=IPMT(0.05,1,3,1000)", &c) {
            Variant::Float(f) => assert!((f + 50.0).abs() < 0.01),
            other => panic!("IPMT unexpected: {:?}", other),
        }

        // PPMT: principal portion = PMT - IPMT = 367.21 - 50 = 317.21
        match calc("=PPMT(0.05,1,3,1000)", &c) {
            Variant::Float(f) => assert!((f + 317.21).abs() < 0.01),
            other => panic!("PPMT unexpected: {:?}", other),
        }

        // NPV: cash flows [50, 60, 70] at 10%
        // Excel NPV discounts each arg: 50/1.1 + 60/1.21 + 70/1.331 ≈ 147.63
        let c_npv = cells_from(&[
            ((1, 1), Variant::Float(50.0)),
            ((2, 1), Variant::Float(60.0)),
            ((3, 1), Variant::Float(70.0)),
        ]);
        match calc("=NPV(0.1,A1:A3)", &c_npv) {
            Variant::Float(f) => assert!((f - 147.63).abs() < 0.01),
            other => panic!("NPV unexpected: {:?}", other),
        }

        // IRR: cash flows [-100, 40, 60, 50] → find r where -100 + 40/(1+r) + 60/(1+r)^2 + 50/(1+r)^3 = 0
        // Verify: NPV at found rate ≈ 0
        let c_irr = cells_from(&[
            ((1, 1), Variant::Float(-100.0)),
            ((2, 1), Variant::Float(40.0)),
            ((3, 1), Variant::Float(60.0)),
            ((4, 1), Variant::Float(50.0)),
        ]);
        let irr_rate = match calc("=IRR(A1:A4)", &c_irr) {
            Variant::Float(f) => f,
            other => panic!("IRR unexpected: {:?}", other),
        };
        // Verify NPV ≈ 0 at that rate
        let npv_check = -100.0
            + 40.0 / (1.0 + irr_rate)
            + 60.0 / (1.0 + irr_rate).powi(2)
            + 50.0 / (1.0 + irr_rate).powi(3);
        assert!(
            npv_check.abs() < 0.001,
            "IRR: NPV at found rate should be ~0, got {}",
            npv_check
        );

        // MIRR: [-120, 50, 60, 70] at finance_rate=10%, reinvest_rate=12%
        // pv_neg = -120; fv_pos = 50*1.12^2 + 60*1.12 + 70 = 199.92
        // MIRR = (199.92/120)^(1/3) - 1 ≈ 18.58%
        let c_mirr = cells_from(&[
            ((1, 1), Variant::Float(-120.0)),
            ((2, 1), Variant::Float(50.0)),
            ((3, 1), Variant::Float(60.0)),
            ((4, 1), Variant::Float(70.0)),
        ]);
        match calc("=MIRR(A1:A4,0.1,0.12)", &c_mirr) {
            Variant::Float(f) => assert!((f - 0.1858).abs() < 0.001),
            other => panic!("MIRR unexpected: {:?}", other),
        }

        // XNPV: values=[-100, 50, 60] at dates=[0, 180, 365] (serial days from base)
        // Using small day numbers to simplify: date0=1, date1=181, date2=366
        let c_xnpv = cells_from(&[
            ((1, 1), Variant::Float(-100.0)),
            ((1, 2), Variant::Integer(1)),
            ((2, 1), Variant::Float(50.0)),
            ((2, 2), Variant::Integer(181)),
            ((3, 1), Variant::Float(60.0)),
            ((3, 2), Variant::Integer(366)),
        ]);
        match calc("=XNPV(0.1,A1:A3,B1:B3)", &c_xnpv) {
            Variant::Float(f) => {
                let expected =
                    -100.0 + 50.0 / 1.1f64.powf(180.0 / 365.0) + 60.0 / 1.1f64.powf(365.0 / 365.0);
                assert!(
                    (f - expected).abs() < 0.01,
                    "XNPV: got {}, expected {}",
                    f,
                    expected
                );
            }
            other => panic!("XNPV unexpected: {:?}", other),
        }

        // XIRR: same cash flows, find rate where XNPV=0
        match calc("=XIRR(A1:A3,B1:B3)", &c_xnpv) {
            Variant::Float(f) => assert!(f > 0.0 && f < 2.0, "XIRR out of range: {}", f),
            other => panic!("XIRR unexpected: {:?}", other),
        }

        let c_schedule =
            cells_from(&[((1, 1), Variant::Float(0.1)), ((2, 1), Variant::Float(0.2))]);
        match calc("=FVSCHEDULE(100,A1:A2)", &c_schedule) {
            Variant::Float(value) => assert!((value - 132.0).abs() < 1e-12),
            other => panic!("FVSCHEDULE unexpected: {:?}", other),
        }
        match calc("=DOLLARDE(1.02,16)", &c) {
            Variant::Float(value) => assert!((value - 1.125).abs() < 1e-12),
            other => panic!("DOLLARDE unexpected: {:?}", other),
        }
        match calc("=DOLLARFR(1.125,16)", &c) {
            Variant::Float(value) => assert!((value - 1.02).abs() < 1e-12),
            other => panic!("DOLLARFR unexpected: {:?}", other),
        }
        let cum_interest = match calc("=CUMIPMT(0.05,3,1000,1,2,0)", &c) {
            Variant::Float(value) => value,
            other => panic!("CUMIPMT unexpected: {:?}", other),
        };
        let cum_principal = match calc("=CUMPRINC(0.05,3,1000,1,2,0)", &c) {
            Variant::Float(value) => value,
            other => panic!("CUMPRINC unexpected: {:?}", other),
        };
        let pmt = match calc("=PMT(0.05,3,1000)", &c) {
            Variant::Float(value) => value,
            other => panic!("PMT unexpected: {:?}", other),
        };
        assert!((cum_interest + cum_principal - 2.0 * pmt).abs() < 1e-9);

        assert!(
            matches!(calc("=PDURATION(0.1,100,200)", &c), Variant::Float(value) if (value - 7.272540897341714).abs() < 1e-9)
        );
        assert!(
            matches!(calc("=PRICEDISC(1,181,0.1,100,2)", &c), Variant::Float(value) if (value - 95.0).abs() < 1e-9)
        );
        assert!(
            matches!(calc("=DISC(1,181,95,100,2)", &c), Variant::Float(value) if (value - 0.1).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=RECEIVED(1,181,95,0.1,2)", &c), Variant::Float(value) if (value - 100.0).abs() < 1e-9)
        );
        assert!(
            matches!(calc("=YIELDDISC(1,181,95,100,2)", &c), Variant::Float(value) if (value - (100.0 / 95.0 - 1.0) / 0.5).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=ACCRINTM(1,181,0.1,1000,2)", &c), Variant::Float(value) if (value - 50.0).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=ACCRINT(DATE(2020,1,15),DATE(2020,7,15),DATE(2020,3,1),0.08,100,2,0,1)", &c), Variant::Float(value) if (value - (4.0 * 46.0 / 180.0)).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=TBILLPRICE(1,181,0.1)", &c), Variant::Float(value) if (value - 95.0).abs() < 1e-9)
        );
        assert!(
            matches!(calc("=TBILLYIELD(1,181,95)", &c), Variant::Float(value) if (value - (10.0 / 95.0)).abs() < 1e-9)
        );
        assert!(
            matches!(calc("=TBILLEQ(1,181,0.1)", &c), Variant::Float(value) if (value - (36.5 / 342.0)).abs() < 1e-12)
        );
        assert_eq!(
            calc("=COUPDAYBS(DATE(2020,3,1),DATE(2021,1,15),2,0)", &c),
            Variant::Float(46.0)
        );
        assert_eq!(
            calc("=COUPDAYS(DATE(2020,3,1),DATE(2021,1,15),2,0)", &c),
            Variant::Float(180.0)
        );
        assert_eq!(
            calc("=COUPDAYSNC(DATE(2020,3,1),DATE(2021,1,15),2,0)", &c),
            Variant::Float(134.0)
        );
        assert_eq!(
            calc("=COUPNCD(DATE(2020,3,1),DATE(2021,1,15),2,0)", &c),
            Variant::Date(date_to_serial(2020, 7, 15))
        );
        assert_eq!(
            calc("=COUPPCD(DATE(2020,3,1),DATE(2021,1,15),2,0)", &c),
            Variant::Date(date_to_serial(2020, 1, 15))
        );
        assert_eq!(
            calc("=COUPNUM(DATE(2020,3,1),DATE(2021,1,15),2,0)", &c),
            Variant::Integer(2)
        );
        assert!(
            matches!(calc("=INTRATE(1,181,95,100,2)", &c), Variant::Float(value) if (value - (100.0 / 95.0 - 1.0) / 0.5).abs() < 1e-12)
        );
        let price_mat = 108.0 / 1.05;
        assert!(
            matches!(calc("=PRICEMAT(DATE(2020,7,1),DATE(2021,1,1),DATE(2020,1,1),0.08,0.1,0)", &c), Variant::Float(value) if (value - price_mat).abs() < 1e-12)
        );
        assert!(
            matches!(calc(&format!("=YIELDMAT(DATE(2020,7,1),DATE(2021,1,1),DATE(2020,1,1),0.08,{price_mat},0)"), &c), Variant::Float(value) if (value - 0.1).abs() < 1e-12)
        );
        let duration = match calc("=DURATION(DATE(2020,3,1),DATE(2021,1,15),0.08,0.1,2,0)", &c) {
            Variant::Float(value) => value,
            other => panic!("DURATION unexpected: {:?}", other),
        };
        let mduration = match calc(
            "=MDURATION(DATE(2020,3,1),DATE(2021,1,15),0.08,0.1,2,0)",
            &c,
        ) {
            Variant::Float(value) => value,
            other => panic!("MDURATION unexpected: {:?}", other),
        };
        assert!(duration.is_finite() && duration > 0.0);
        assert!((mduration - duration / 1.05).abs() < 1e-12);
        let price = match calc(
            "=PRICE(DATE(2020,3,1),DATE(2021,1,15),0.08,0.1,100,2,0)",
            &c,
        ) {
            Variant::Float(value) => value,
            other => panic!("PRICE unexpected: {:?}", other),
        };
        assert!(price.is_finite() && price > 0.0);
        assert!(matches!(
            calc(
                &format!(
                    "=YIELD(DATE(2020,3,1),DATE(2021,1,15),0.08,{price},100,2,0)"
                ),
                &c
            ),
            Variant::Float(value) if (value - 0.1).abs() < 1e-10
        ));
        let odd_first_price = match calc(
            "=ODDFPRICE(DATE(2024,2,1),DATE(2025,1,1),DATE(2024,1,1),DATE(2024,7,1),0.08,0.1,100,2,0)",
            &c,
        ) {
            Variant::Float(value) => value,
            other => panic!("ODDFPRICE unexpected: {:?}", other),
        };
        assert!(odd_first_price.is_finite() && odd_first_price > 0.0);
        assert!(matches!(
            calc(
                &format!(
                    "=ODDFYIELD(DATE(2024,2,1),DATE(2025,1,1),DATE(2024,1,1),DATE(2024,7,1),0.08,{odd_first_price},100,2,0)"
                ),
                &c
            ),
            Variant::Float(value) if (value - 0.1).abs() < 1e-9
        ));
        let odd_last_price = match calc(
            "=ODDLPRICE(DATE(2024,2,1),DATE(2025,3,1),DATE(2024,7,1),0.08,0.1,100,2,0)",
            &c,
        ) {
            Variant::Float(value) => value,
            other => panic!("ODDLPRICE unexpected: {:?}", other),
        };
        assert!(odd_last_price.is_finite() && odd_last_price > 0.0);
        assert!(matches!(
            calc(
                &format!(
                    "=ODDLYIELD(DATE(2024,2,1),DATE(2025,3,1),DATE(2024,7,1),0.08,{odd_last_price},100,2,0)"
                ),
                &c
            ),
            Variant::Float(value) if (value - 0.1).abs() < 1e-9
        ));
        assert!(
            matches!(calc("=AMORLINC(1000,DATE(2020,1,1),DATE(2020,7,1),0,0,0.2,0)", &c), Variant::Float(value) if (value - 100.0).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=AMORLINC(1000,DATE(2020,1,1),DATE(2020,7,1),0,1,0.2,0)", &c), Variant::Float(value) if (value - 200.0).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=AMORLINC(1000,DATE(2020,1,1),DATE(2020,7,1),0,5,0.2,0)", &c), Variant::Float(value) if (value - 100.0).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=AMORDEGRC(1000,DATE(2020,1,1),DATE(2020,7,1),0,0,0.2,0)", &c), Variant::Float(value) if (value - 100.0).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=AMORDEGRC(1000,DATE(2020,1,1),DATE(2020,7,1),0,1,0.2,0)", &c), Variant::Float(value) if (value - 360.0).abs() < 1e-12)
        );
    }

    #[test]
    fn test_weeknum_isoweeknum() {
        let c = HashMap::new();
        // ISOWEEKNUM
        // Jan 4, 2016 (Mon) = ISO week 1 of 2016
        assert_eq!(calc("=ISOWEEKNUM(DATE(2016,1,4))", &c), Variant::Integer(1));
        // Jan 1, 2016 (Fri) = ISO week 53 of 2015
        assert_eq!(
            calc("=ISOWEEKNUM(DATE(2016,1,1))", &c),
            Variant::Integer(53)
        );
        // Dec 31, 2018 (Mon) = ISO week 1 of 2019
        assert_eq!(
            calc("=ISOWEEKNUM(DATE(2018,12,31))", &c),
            Variant::Integer(1)
        );
        // Dec 28, 2015 (Mon) = ISO week 53 of 2015
        assert_eq!(
            calc("=ISOWEEKNUM(DATE(2015,12,28))", &c),
            Variant::Integer(53)
        );
        // WEEKNUM type 1 (Sun-start); Jan 1, 2023 = Sunday
        assert_eq!(calc("=WEEKNUM(DATE(2023,1,1))", &c), Variant::Integer(1));
        assert_eq!(calc("=WEEKNUM(DATE(2023,1,7))", &c), Variant::Integer(1)); // Sat still w1
        assert_eq!(calc("=WEEKNUM(DATE(2023,1,8))", &c), Variant::Integer(2)); // next Sunday
        // WEEKNUM type 2 (Mon-start)
        assert_eq!(calc("=WEEKNUM(DATE(2023,1,1),2)", &c), Variant::Integer(1)); // Sun in w1
        assert_eq!(calc("=WEEKNUM(DATE(2023,1,2),2)", &c), Variant::Integer(2)); // Mon starts w2
        // WEEKNUM type 21 = ISOWEEKNUM
        assert_eq!(
            calc("=WEEKNUM(DATE(2016,1,1),21)", &c),
            Variant::Integer(53)
        );
        assert_eq!(calc("=WEEKNUM(DATE(2016,1,4),21)", &c), Variant::Integer(1));
    }

    #[test]
    fn test_rept_numbervalue() {
        let c = HashMap::new();
        // REPT
        assert_eq!(calc("=REPT(\"ab\",3)", &c), Variant::Str("ababab".into()));
        assert_eq!(calc("=REPT(\"x\",0)", &c), Variant::Str("".into()));
        assert_eq!(
            calc("=REPT(\"x\",-1)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(calc("=REPT(\"hi\",1.9)", &c), Variant::Str("hi".into())); // truncate

        // NUMBERVALUE — default separators
        assert_eq!(
            calc("=NUMBERVALUE(\"1234.56\")", &c),
            Variant::Float(1234.56)
        );
        match calc("=NUMBERVALUE(\"1,234.56\")", &c) {
            Variant::Float(f) => assert!((f - 1234.56).abs() < 1e-9),
            other => panic!("NUMBERVALUE unexpected: {:?}", other),
        }
        assert_eq!(calc("=NUMBERVALUE(\"\")", &c), Variant::Integer(0));
        assert_eq!(
            calc("=NUMBERVALUE(\"abc\")", &c),
            Variant::Error(ExcelError::Value)
        );
        // European locale: decimal=",", group="."
        match calc("=NUMBERVALUE(\"1.234,56\",\",\",\".\")", &c) {
            Variant::Float(f) => assert!((f - 1234.56).abs() < 1e-9),
            other => panic!("NUMBERVALUE EU unexpected: {:?}", other),
        }
        // Whole number
        assert_eq!(calc("=NUMBERVALUE(\"42\")", &c), Variant::Integer(42));
        // Invalid: same separator
        assert_eq!(
            calc("=NUMBERVALUE(\"1.2\",\".\",\".\")", &c),
            Variant::Error(ExcelError::Value)
        );
    }

    #[test]
    fn test_lookup_info_functions() {
        let c = HashMap::new();

        // NA
        assert_eq!(calc("=NA()", &c), Variant::Error(ExcelError::NA));

        // N
        assert_eq!(calc("=N(42)", &c), Variant::Integer(42));
        assert_eq!(calc("=N(TRUE)", &c), Variant::Integer(1));
        assert_eq!(calc("=N(FALSE)", &c), Variant::Integer(0));
        assert_eq!(calc("=N(\"abc\")", &c), Variant::Integer(0));

        // TYPE
        assert_eq!(calc("=TYPE(1)", &c), Variant::Integer(1));
        assert_eq!(calc("=TYPE(1.5)", &c), Variant::Integer(1));
        assert_eq!(calc("=TYPE(\"text\")", &c), Variant::Integer(2));
        assert_eq!(calc("=TYPE(TRUE)", &c), Variant::Integer(4));

        // ERROR.TYPE
        assert_eq!(calc("=ERROR.TYPE(1/0)", &c), Variant::Integer(2)); // #DIV/0! = 2
        assert_eq!(calc("=ERROR.TYPE(NA())", &c), Variant::Integer(7)); // #N/A = 7
        assert_eq!(calc("=ERROR.TYPE(42)", &c), Variant::Error(ExcelError::NA)); // not an error

        // External-service worksheet functions are recognized but remain
        // headless-safe and fail closed without evaluating external inputs.
        assert_eq!(
            calc("=WEBSERVICE(\"https://example.invalid\")", &c),
            Variant::Error(ExcelError::NA)
        );
        assert_eq!(
            calc("=CUBESETCOUNT(\"local-cube\")", &c),
            Variant::Error(ExcelError::NA)
        );
        assert_eq!(
            calc("=TRANSLATE(\"hello\",\"en\",\"EN\")", &c),
            Variant::Str("hello".into())
        );
        assert_eq!(
            calc("=TRANSLATE(\"hello\",\"en\",\"ja\")", &c),
            Variant::Error(ExcelError::NA)
        );

        // FORMULATEXT
        let mut cf = cells_from(&[((1, 1), Variant::Integer(10))]);
        cf.insert(
            (2, 1),
            CellContent {
                formula: Some("=A1*2".into()),
                value: Variant::Integer(20),
            },
        );
        assert_eq!(
            calc("=FORMULATEXT(A1)", &cf),
            Variant::Error(ExcelError::NA)
        ); // no formula
        assert_eq!(calc("=FORMULATEXT(A2)", &cf), Variant::Str("=A1*2".into()));

        // CELL
        let cv = cells_from(&[
            ((3, 2), Variant::Str("hello".into())),
            ((4, 2), Variant::Integer(99)),
        ]);
        assert_eq!(calc("=CELL(\"row\",B3)", &cv), Variant::Integer(3));
        assert_eq!(calc("=CELL(\"col\",B3)", &cv), Variant::Integer(2));
        assert_eq!(
            calc("=CELL(\"address\",B3)", &cv),
            Variant::Str("$B$3".into())
        );
        assert_eq!(calc("=CELL(\"type\",B3)", &cv), Variant::Str("l".into())); // Str = label
        assert_eq!(calc("=CELL(\"type\",B4)", &cv), Variant::Str("v".into())); // number = value
        assert_eq!(calc("=CELL(\"contents\",B4)", &cv), Variant::Integer(99));
        assert_eq!(calc("=CELL(\"filename\",B3)", &cv), Variant::Str("".into()));
        assert_eq!(calc("=CELL(\"width\",B3)", &cv), Variant::Integer(8));

        // col_to_letter helper: col 1=A, 26=Z, 27=AA
        assert_eq!(col_to_letter(1), "A");
        assert_eq!(col_to_letter(26), "Z");
        assert_eq!(col_to_letter(27), "AA");
        assert_eq!(col_to_letter(28), "AB");
    }

    #[test]
    fn test_bug_fixes() {
        let c = HashMap::new();
        // Fix 1: partial_cmp NaN safety — sort_by now uses unwrap_or(Equal)
        let c_nums = cells_from(&[
            ((1, 1), Variant::Integer(3)),
            ((2, 1), Variant::Integer(1)),
            ((3, 1), Variant::Integer(2)),
        ]);
        assert_eq!(calc("=MEDIAN(A1:A3)", &c_nums), Variant::Integer(2));
        assert_eq!(calc("=LARGE(A1:A3,1)", &c_nums), Variant::Integer(3));
        assert_eq!(calc("=SMALL(A1:A3,1)", &c_nums), Variant::Integer(1));

        // Fix 2: LCM overflow → #NUM!
        let large = cells_from(&[
            ((1, 1), Variant::Integer(4611686018427387903i64)), // 2^62-1
            ((2, 1), Variant::Integer(5)),
        ]);
        assert_eq!(calc("=LCM(A1,A2)", &large), Variant::Error(ExcelError::Num));
        // Normal LCM still works
        assert_eq!(calc("=LCM(4,6)", &c), Variant::Integer(12));

        // Fix 3: Huge range → error (size limit)
        // Not easily testable without a giant formula, but confirmed via code review

        // Fix 4: IPMT/PPMT with fractional per → #VALUE!
        assert_eq!(
            calc("=IPMT(0.05,1.9,10,1000)", &c),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=PPMT(0.05,1.9,10,1000)", &c),
            Variant::Error(ExcelError::Value)
        );
        // Integer per still works
        match calc("=IPMT(0.05,1,3,1000)", &c) {
            Variant::Float(f) => assert!((f + 50.0).abs() < 0.01),
            other => panic!("IPMT unexpected: {:?}", other),
        }

        // Fix 5: "<>text" criteria now correctly does string inequality
        let cs = cells_from(&[
            ((1, 1), Variant::Str("apple".into())),
            ((2, 1), Variant::Str("banana".into())),
            ((3, 1), Variant::Str("apple".into())),
        ]);
        // COUNTIF with "<>apple" should count non-apple cells → 1
        assert_eq!(
            calc("=COUNTIF(A1:A3,\"<>apple\")", &cs),
            Variant::Integer(1)
        );
        // COUNTIF with "<>5" (numeric NE) should work for numeric cells
        let cn = cells_from(&[
            ((1, 1), Variant::Integer(3)),
            ((2, 1), Variant::Integer(5)),
            ((3, 1), Variant::Integer(7)),
        ]);
        assert_eq!(calc("=COUNTIF(A1:A3,\"<>5\")", &cn), Variant::Integer(2));
    }

    #[test]
    fn test_math_combinatorics() {
        let c = HashMap::new();
        // FACT
        assert_eq!(calc("=FACT(0)", &c), Variant::Integer(1));
        assert_eq!(calc("=FACT(5)", &c), Variant::Integer(120));
        assert_eq!(calc("=FACT(10)", &c), Variant::Integer(3628800));
        assert_eq!(calc("=FACT(-1)", &c), Variant::Error(ExcelError::Num));
        // PERMUT
        assert_eq!(calc("=PERMUT(5,2)", &c), Variant::Integer(20));
        assert_eq!(calc("=PERMUT(5,0)", &c), Variant::Integer(1));
        assert_eq!(calc("=PERMUTATIONA(3,2)", &c), Variant::Integer(9));
        // GCD
        assert_eq!(calc("=GCD(12,8)", &c), Variant::Integer(4));
        assert_eq!(calc("=GCD(0,5)", &c), Variant::Integer(5));
        // LCM
        assert_eq!(calc("=LCM(4,6)", &c), Variant::Integer(12));
        assert_eq!(calc("=LCM(3,0)", &c), Variant::Integer(0));
        // QUOTIENT
        assert_eq!(calc("=QUOTIENT(10,3)", &c), Variant::Integer(3));
        assert_eq!(calc("=QUOTIENT(-10,3)", &c), Variant::Integer(-3));
        assert_eq!(
            calc("=QUOTIENT(10,0)", &c),
            Variant::Error(ExcelError::DivZero)
        );
        // SIGN
        assert_eq!(calc("=SIGN(5)", &c), Variant::Integer(1));
        assert_eq!(calc("=SIGN(-3)", &c), Variant::Integer(-1));
        assert_eq!(calc("=SIGN(0)", &c), Variant::Integer(0));
        assert_eq!(calc("=SLN(1000,100,5)", &c), Variant::Float(180.0));
        assert_eq!(calc("=SYD(1000,100,5,1)", &c), Variant::Float(300.0));
        assert_eq!(calc("=DDB(1000,100,5,1)", &c), Variant::Float(400.0));
        assert_eq!(
            calc("=VDB(1000,100,5,0,1,2,TRUE)", &c),
            Variant::Float(400.0)
        );
        assert_eq!(
            calc("=VDB(1000,100,5,0,1,2,FALSE)", &c),
            Variant::Float(400.0)
        );
        assert_eq!(
            calc("=VDB(1000,100,5,2,1)", &c),
            Variant::Error(ExcelError::Num)
        );
        assert!(
            matches!(calc("=DB(1000,100,5,1)", &c), Variant::Float(v) if (v - 369.0).abs() < 1.0)
        );
        match calc("=EFFECT(0.1,4)", &c) {
            Variant::Float(v) => assert!((v - 0.10381289).abs() < 1e-7),
            other => panic!("EFFECT: {:?}", other),
        }
        match calc("=NOMINAL(0.10381289,4)", &c) {
            Variant::Float(v) => assert!((v - 0.1).abs() < 1e-7),
            other => panic!("NOMINAL: {:?}", other),
        }
        match calc("=RRI(5,100,150)", &c) {
            Variant::Float(v) => assert!((v - 0.08447177).abs() < 1e-7),
            other => panic!("RRI: {:?}", other),
        }
        assert_eq!(
            calc("=SLN(1000,100,0)", &c),
            Variant::Error(ExcelError::DivZero)
        );
    }

    #[test]
    fn test_regression_array_functions() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(2)),
            ((2, 1), Variant::Integer(4)),
            ((3, 1), Variant::Integer(6)),
            ((1, 2), Variant::Integer(1)),
            ((2, 2), Variant::Integer(2)),
            ((3, 2), Variant::Integer(3)),
            ((1, 3), Variant::Integer(4)),
            ((2, 3), Variant::Integer(5)),
            ((1, 4), Variant::Integer(2)),
            ((2, 4), Variant::Integer(4)),
            ((3, 4), Variant::Integer(8)),
            ((1, 5), Variant::Integer(1)),
            ((2, 5), Variant::Integer(2)),
            ((3, 5), Variant::Integer(3)),
            ((1, 6), Variant::Integer(4)),
        ]);
        assert_eq!(
            calc("=TREND(A1:A3,B1:B3,C1:C2)", &c),
            Variant::Array(vec![Variant::Integer(8), Variant::Integer(10)])
        );
        assert_eq!(
            calc("=LINEST(A1:A3,B1:B3)", &c),
            Variant::Array(vec![Variant::Integer(2), Variant::Integer(0)])
        );
        match calc("=LINEST(A1:A3,B1:B3,TRUE,TRUE)", &c) {
            Variant::Array(values) => {
                assert_eq!(values.len(), 10);
                assert_eq!(values[0], Variant::Integer(2));
                assert_eq!(values[1], Variant::Integer(0));
                assert_eq!(values[7], Variant::Integer(1));
            }
            other => panic!("LINEST stats unexpected: {:?}", other),
        }
        let multivariate = cells_from(&[
            ((1, 7), Variant::Integer(10)),
            ((2, 7), Variant::Integer(12)),
            ((3, 7), Variant::Integer(13)),
            ((4, 7), Variant::Integer(15)),
            ((5, 7), Variant::Integer(14)),
            ((6, 7), Variant::Integer(16)),
            ((1, 8), Variant::Integer(1)),
            ((2, 8), Variant::Integer(2)),
            ((3, 8), Variant::Integer(1)),
            ((4, 8), Variant::Integer(2)),
            ((5, 8), Variant::Integer(3)),
            ((6, 8), Variant::Integer(1)),
            ((1, 9), Variant::Integer(1)),
            ((2, 9), Variant::Integer(1)),
            ((3, 9), Variant::Integer(2)),
            ((4, 9), Variant::Integer(2)),
            ((5, 9), Variant::Integer(1)),
            ((6, 9), Variant::Integer(3)),
        ]);
        match calc("=LINEST(G1:G6,H1:I6)", &multivariate) {
            Variant::Array(values) => {
                assert_eq!(values.len(), 3);
                for (value, expected) in values.iter().zip([3.0, 2.0, 5.0]) {
                    assert!((as_f64(value).unwrap() - expected).abs() < 1e-9);
                }
            }
            other => panic!("multivariate LINEST unexpected: {:?}", other),
        }
        match calc("=INDEX(LINEST(G1:G6,H1:I6),1,1)", &multivariate) {
            Variant::Integer(value) => assert_eq!(value, 3),
            Variant::Float(value) => assert!((value - 3.0).abs() < 1e-9),
            other => panic!("LINEST INDEX projection unexpected: {:?}", other),
        }
        match calc("=LINEST(G1:G6,H1:I6,TRUE,TRUE)", &multivariate) {
            Variant::Array(values) => {
                assert_eq!(values.len(), 15);
                for (value, expected) in values.iter().take(3).zip([3.0, 2.0, 5.0]) {
                    assert!((as_f64(value).unwrap() - expected).abs() < 1e-9);
                }
                assert!((as_f64(&values[6]).unwrap() - 1.0).abs() < 1e-9);
            }
            other => panic!("multivariate LINEST stats unexpected: {:?}", other),
        }
        match calc("=INDEX(LINEST(G1:G6,H1:I6,TRUE,TRUE),3,1)", &multivariate) {
            Variant::Integer(value) => assert_eq!(value, 1),
            Variant::Float(value) => assert!((value - 1.0).abs() < 1e-9),
            other => panic!("LINEST stats INDEX projection unexpected: {:?}", other),
        }
        match calc("=LOGEST(G1:G6,H1:I6)", &multivariate) {
            Variant::Array(values) => {
                assert_eq!(values.len(), 3);
                assert!(
                    values
                        .iter()
                        .all(|value| as_f64(value).is_some_and(f64::is_finite))
                );
            }
            other => panic!("multivariate LOGEST unexpected: {:?}", other),
        }
        match calc("=TREND(G1:G6,H1:I6,H1:I2)", &multivariate) {
            Variant::Array(values) => {
                assert_eq!(values.len(), 2);
                assert!((as_f64(&values[0]).unwrap() - 10.0).abs() < 1e-9);
                assert!((as_f64(&values[1]).unwrap() - 12.0).abs() < 1e-9);
            }
            other => panic!("multivariate TREND unexpected: {:?}", other),
        }
        let growth_multi = cells_from(&[
            ((1, 1), Variant::Float(1.3_f64.exp())),
            ((2, 1), Variant::Float(1.4_f64.exp())),
            ((3, 1), Variant::Float(1.5_f64.exp())),
            ((4, 1), Variant::Float(1.6_f64.exp())),
            ((5, 1), Variant::Float(1.5_f64.exp())),
            ((6, 1), Variant::Float(1.7_f64.exp())),
            ((1, 2), Variant::Integer(1)),
            ((2, 2), Variant::Integer(2)),
            ((3, 2), Variant::Integer(1)),
            ((4, 2), Variant::Integer(2)),
            ((5, 2), Variant::Integer(3)),
            ((6, 2), Variant::Integer(1)),
            ((1, 3), Variant::Integer(1)),
            ((2, 3), Variant::Integer(1)),
            ((3, 3), Variant::Integer(2)),
            ((4, 3), Variant::Integer(2)),
            ((5, 3), Variant::Integer(1)),
            ((6, 3), Variant::Integer(3)),
        ]);
        match calc("=GROWTH(A1:A6,B1:C6,B1:C2)", &growth_multi) {
            Variant::Array(values) => {
                assert_eq!(values.len(), 2);
                assert!((as_f64(&values[0]).unwrap() - 1.3_f64.exp()).abs() < 1e-9);
                assert!((as_f64(&values[1]).unwrap() - 1.4_f64.exp()).abs() < 1e-9);
            }
            other => panic!("multivariate GROWTH unexpected: {:?}", other),
        }
        match calc("=GROWTH(D1:D3,E1:E3,F1:F1)", &c) {
            Variant::Array(values) => match values.as_slice() {
                [Variant::Float(value)] => assert!((*value - 16.0).abs() < 1e-9),
                other => panic!("GROWTH values unexpected: {:?}", other),
            },
            other => panic!("GROWTH result unexpected: {:?}", other),
        }
        match calc("=LOGEST(D1:D3,E1:E3)", &c) {
            Variant::Array(values) => match values.as_slice() {
                [Variant::Float(base), Variant::Float(factor)] => {
                    assert!((*base - 2.0).abs() < 1e-9);
                    assert!((*factor - 1.0).abs() < 1e-9);
                }
                other => panic!("LOGEST values unexpected: {:?}", other),
            },
            other => panic!("LOGEST result unexpected: {:?}", other),
        }
    }

    #[test]
    fn test_prob_function() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(3)),
            ((4, 1), Variant::Integer(4)),
            ((1, 2), Variant::Float(0.1)),
            ((2, 2), Variant::Float(0.2)),
            ((3, 2), Variant::Float(0.3)),
            ((4, 2), Variant::Float(0.4)),
        ]);
        assert_eq!(calc("=PROB(A1:A4,B1:B4,2,3)", &c), Variant::Float(0.5));
        assert_eq!(calc("=PROB(A1:A4,B1:B4,2)", &c), Variant::Float(0.2));
        assert!(evaluate(&fparse("=PROB(A1:A4,B1:B4,4,2)").unwrap(), &c).is_err());
    }

    #[test]
    fn test_statistical_extended() {
        // CORREL
        let c = cells_from(&[
            ((1, 1), Variant::Float(1.0)),
            ((1, 2), Variant::Float(2.0)),
            ((2, 1), Variant::Float(2.0)),
            ((2, 2), Variant::Float(4.0)),
            ((3, 1), Variant::Float(3.0)),
            ((3, 2), Variant::Float(6.0)),
        ]);
        match calc("=CORREL(A1:A3,B1:B3)", &c) {
            Variant::Float(f) => assert!((f - 1.0).abs() < 1e-9),
            other => panic!("CORREL: {:?}", other),
        }
        match calc("=PEARSON(A1:A3,B1:B3)", &c) {
            Variant::Float(f) => assert!((f - 1.0).abs() < 1e-9),
            other => panic!("PEARSON: {:?}", other),
        }
        match calc("=CORREL(SEQUENCE(3),SEQUENCE(3))", &c) {
            Variant::Float(f) => assert!((f - 1.0).abs() < 1e-9),
            other => panic!("CORREL dynamic array: {:?}", other),
        }
        match calc("=COVARIANCE.P(SEQUENCE(3),SEQUENCE(3))", &c) {
            Variant::Float(f) => assert!((f - (2.0 / 3.0)).abs() < 1e-9),
            other => panic!("COVARIANCE.P dynamic array: {:?}", other),
        }
        match calc("=SLOPE(B1:B3,A1:A3)", &c) {
            Variant::Float(f) => assert!((f - 2.0).abs() < 1e-9),
            other => panic!("SLOPE: {:?}", other),
        }
        match calc("=INTERCEPT(B1:B3,A1:A3)", &c) {
            Variant::Float(f) => assert!(f.abs() < 1e-9),
            other => panic!("INTERCEPT: {:?}", other),
        }
        match calc("=RSQ(B1:B3,A1:A3)", &c) {
            Variant::Float(f) => assert!((f - 1.0).abs() < 1e-9),
            other => panic!("RSQ: {:?}", other),
        }
        match calc("=FORECAST.LINEAR(4,B1:B3,A1:A3)", &c) {
            Variant::Float(f) => assert!((f - 8.0).abs() < 1e-9),
            other => panic!("FORECAST.LINEAR: {:?}", other),
        }
        let ets = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(3)),
            ((4, 1), Variant::Integer(4)),
            ((1, 2), Variant::Integer(10)),
            ((2, 2), Variant::Integer(20)),
            ((3, 2), Variant::Integer(10)),
            ((4, 2), Variant::Integer(20)),
        ]);
        assert_eq!(
            calc("=FORECAST.ETS.SEASONALITY(B1:B4,A1:A4)", &ets),
            Variant::Integer(2)
        );
        assert_eq!(
            calc("=FORECAST.ETS(5,B1:B4,A1:A4)", &ets),
            Variant::Float(10.0)
        );
        assert_eq!(
            calc("=FORECAST.ETS(5,B1:B4,A1:A4,0)", &ets),
            Variant::Float(20.0)
        );
        match calc("=FORECAST.ETS.CONFINT(5,B1:B4,A1:A4)", &ets) {
            Variant::Float(value) => assert!(value.is_finite() && value >= 0.0),
            other => panic!("FORECAST.ETS.CONFINT: {:?}", other),
        }
        assert_eq!(
            calc("=FORECAST.ETS.STAT(B1:B4,A1:A4,8)", &ets),
            Variant::Float(1.0)
        );
        match calc("=FORECAST.ETS.STAT(B1:B4,A1:A4,7,0)", &ets) {
            Variant::Float(value) => assert!(value.is_finite() && value > 0.0),
            other => panic!("FORECAST.ETS.STAT: {:?}", other),
        }
        match calc("=STEYX(B1:B3,A1:A3)", &c) {
            Variant::Float(f) => assert!(f.abs() < 1e-9),
            other => panic!("STEYX: {:?}", other),
        }
        // COVARIANCE.P
        let c2 = cells_from(&[
            ((1, 1), Variant::Float(2.0)),
            ((1, 2), Variant::Float(4.0)),
            ((2, 1), Variant::Float(4.0)),
            ((2, 2), Variant::Float(6.0)),
        ]);
        // cov_p = ((2-3)*(4-5) + (4-3)*(6-5)) / 2 = 1.0
        match calc("=COVARIANCE.P(A1:A2,B1:B2)", &c2) {
            Variant::Float(f) => assert!((f - 1.0).abs() < 1e-9),
            other => panic!("COVARIANCE.P: {:?}", other),
        }
        // NORM.DIST (cumulative)
        let cn = HashMap::new();
        match calc("=NORM.DIST(0,0,1,TRUE)", &cn) {
            Variant::Float(f) => assert!((f - 0.5).abs() < 1e-6),
            other => panic!("NORM.DIST CDF: {:?}", other),
        }
        match calc("=NORM.DIST(1.96,0,1,TRUE)", &cn) {
            Variant::Float(f) => assert!((f - 0.975).abs() < 0.001),
            other => panic!("NORM.DIST 1.96: {:?}", other),
        }
        // NORM.DIST (PDF)
        match calc("=NORM.DIST(0,0,1,FALSE)", &cn) {
            Variant::Float(f) => assert!((f - 0.39894).abs() < 1e-4),
            other => panic!("NORM.DIST PDF: {:?}", other),
        }
        // NORM.INV
        match calc("=NORM.INV(0.975,0,1)", &cn) {
            Variant::Float(f) => assert!((f - 1.96).abs() < 0.001),
            other => panic!("NORM.INV: {:?}", other),
        }
        // T.DIST (cumulative) — t=0 should give 0.5
        match calc("=T.DIST(0,10,TRUE)", &cn) {
            Variant::Float(f) => assert!((f - 0.5).abs() < 1e-6),
            other => panic!("T.DIST(0): {:?}", other),
        }
        // T.DIST: t(10) 97.5th percentile ≈ 2.228
        match calc("=T.DIST(2.228,10,TRUE)", &cn) {
            Variant::Float(f) => assert!((f - 0.975).abs() < 0.001),
            other => panic!("T.DIST df=10: {:?}", other),
        }
        match calc("=T.DIST.2T(2.228,10)", &cn) {
            Variant::Float(f) => assert!((f - 0.05).abs() < 0.001),
            other => panic!("T.DIST.2T: {:?}", other),
        }
        match calc("=T.DIST.RT(2.228,10)", &cn) {
            Variant::Float(f) => assert!((f - 0.025).abs() < 0.001),
            other => panic!("T.DIST.RT: {:?}", other),
        }
        match calc("=T.INV(0.975,10)", &cn) {
            Variant::Float(f) => assert!((f - 2.228).abs() < 0.002),
            other => panic!("T.INV: {:?}", other),
        }
        match calc("=T.INV.2T(0.05,10)", &cn) {
            Variant::Float(f) => assert!((f - 2.228).abs() < 0.002),
            other => panic!("T.INV.2T: {:?}", other),
        }
        match calc("=TDIST(2.228,10,2)", &cn) {
            Variant::Float(f) => assert!((f - 0.05).abs() < 0.001),
            other => panic!("TDIST: {:?}", other),
        }
        match calc("=TINV(0.05,10)", &cn) {
            Variant::Float(f) => assert!((f - 2.228).abs() < 0.002),
            other => panic!("TINV: {:?}", other),
        }
        match calc("=GAMMA(5)", &cn) {
            Variant::Float(f) => assert!((f - 24.0).abs() < 1e-6),
            other => panic!("GAMMA: {:?}", other),
        }
        match calc("=GAMMALN(5)", &cn) {
            Variant::Float(f) => assert!((f - 24.0_f64.ln()).abs() < 1e-9),
            other => panic!("GAMMALN: {:?}", other),
        }
        assert_eq!(
            calc("=GAMMADIST(1,2,1,TRUE)", &cn),
            calc("=GAMMA.DIST(1,2,1,TRUE)", &cn)
        );
        assert_eq!(calc("=CRITBINOM(10,0.5,0.5)", &cn), Variant::Integer(5));
        assert_eq!(
            calc("=CHIINV(0.5,2)", &cn),
            calc("=CHISQ.INV.RT(0.5,2)", &cn)
        );
        assert_eq!(calc("=FINV(0.5,1,1)", &cn), calc("=F.INV.RT(0.5,1,1)", &cn));
        match calc("=BETA.DIST(0.5,2,2,TRUE)", &cn) {
            Variant::Float(f) => assert!((f - 0.5).abs() < 1e-9),
            other => panic!("BETA.DIST CDF: {:?}", other),
        }
        match calc("=BETA.DIST(0.5,2,2,FALSE)", &cn) {
            Variant::Float(f) => assert!((f - 1.5).abs() < 1e-6),
            other => panic!("BETA.DIST PDF: {:?}", other),
        }
        match calc("=BETA.INV(0.5,2,2)", &cn) {
            Variant::Float(f) => assert!((f - 0.5).abs() < 1e-8),
            other => panic!("BETA.INV: {:?}", other),
        }
        match calc("=CHISQ.DIST(2,2,TRUE)", &cn) {
            Variant::Float(f) => assert!((f - 0.6321205588).abs() < 1e-6),
            other => panic!("CHISQ.DIST CDF: {:?}", other),
        }
        match calc("=F.DIST(1,1,1,TRUE)", &cn) {
            Variant::Float(f) => assert!((f - 0.5).abs() < 1e-6),
            other => panic!("F.DIST CDF: {:?}", other),
        }
        let ftest_cells = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(4)),
            ((4, 1), Variant::Integer(8)),
        ]);
        match calc("=FTEST(A1:A4,A1:A4)", &ftest_cells) {
            Variant::Float(value) => assert!((value - 1.0).abs() < 1e-12),
            other => panic!("FTEST: {:?}", other),
        }
        match calc("=F.TEST(A1:A4,A1:A4)", &ftest_cells) {
            Variant::Float(value) => assert!((value - 1.0).abs() < 1e-12),
            other => panic!("F.TEST: {:?}", other),
        }
    }

    #[test]
    fn test_bahttext() {
        let c = HashMap::new();
        assert_eq!(
            calc("=BAHTTEXT(1234.56)", &c),
            Variant::Str("หนึ่งพันสองร้อยสามสิบสี่บาทห้าสิบหกสตางค์".into())
        );
        assert_eq!(
            calc("=BAHTTEXT(1000000)", &c),
            Variant::Str("หนึ่งล้านบาทถ้วน".into())
        );
        assert_eq!(
            calc("=BAHTTEXT(-0.25)", &c),
            Variant::Str("ลบศูนย์บาทยี่สิบห้าสตางค์".into())
        );
    }

    #[test]
    fn test_randarray() {
        let c = HashMap::new();
        // RANDARRAY(3) → 3 floats, all in [0, 1)
        let result = calc("=RANDARRAY(3)", &c);
        if let Variant::Array(arr) = result {
            assert_eq!(arr.len(), 3);
            for v in &arr {
                if let Variant::Float(f) = v {
                    assert!(*f >= 0.0 && *f < 1.0, "out of [0,1): {}", f);
                } else {
                    panic!("expected Float, got {:?}", v);
                }
            }
        } else {
            panic!("expected Array");
        }

        // RANDARRAY(2, 3) → 6 elements
        let result2 = calc("=RANDARRAY(2, 3)", &c);
        if let Variant::Array(arr) = result2 {
            assert_eq!(arr.len(), 6);
        } else {
            panic!("expected Array");
        }

        // RANDARRAY(5, 1, 1, 10, TRUE) → integers in [1, 10]
        let result3 = calc("=RANDARRAY(5, 1, 1, 10, TRUE)", &c);
        if let Variant::Array(arr) = result3 {
            assert_eq!(arr.len(), 5);
            for v in &arr {
                let n = match v {
                    Variant::Integer(i) => *i as f64,
                    Variant::Float(f) => *f,
                    other => panic!("expected numeric, got {:?}", other),
                };
                assert!((1.0..=10.0).contains(&n), "out of [1,10]: {}", n);
            }
        } else {
            panic!("expected Array");
        }
        assert!(evaluate(&fparse("=RANDARRAY(2.5)").unwrap(), &c).is_err());
        assert!(evaluate(&fparse("=RANDARRAY(0,2)").unwrap(), &c).is_err());
    }

    #[test]
    fn test_day_count_functions() {
        let cells = HashMap::new();
        assert_eq!(
            calc("=DAYS360(DATE(2020,1,1),DATE(2020,7,1))", &cells),
            Variant::Integer(180)
        );
        assert_eq!(
            calc("=DAYS360(DATE(2020,2,29),DATE(2020,3,31))", &cells),
            Variant::Integer(30)
        );
        assert_eq!(
            calc("=DAYS360(DATE(2020,2,29),DATE(2020,3,31),TRUE)", &cells),
            Variant::Integer(31)
        );
        match calc("=YEARFRAC(DATE(2020,1,1),DATE(2020,12,31),1)", &cells) {
            Variant::Float(f) => assert!((f - 365.0 / 366.0).abs() < 1e-12),
            other => panic!("YEARFRAC actual/actual: {:?}", other),
        }
        match calc("=YEARFRAC(DATE(2020,1,1),DATE(2021,1,1),3)", &cells) {
            Variant::Float(f) => assert!((f - 366.0 / 365.0).abs() < 1e-12),
            other => panic!("YEARFRAC actual/365: {:?}", other),
        }
        assert_eq!(
            calc("=YEARFRAC(DATE(2020,1,1),DATE(2020,1,2),9)", &cells),
            Variant::Error(ExcelError::Num)
        );
    }

    #[test]
    fn test_matrix_functions() {
        let cells = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((1, 2), Variant::Integer(2)),
            ((2, 1), Variant::Integer(3)),
            ((2, 2), Variant::Integer(4)),
        ]);
        assert_eq!(
            calc("=MUNIT(2)", &cells),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(0),
                Variant::Integer(0),
                Variant::Integer(1),
            ])
        );
        assert_eq!(calc("=MDETERM(A1:B2)", &cells), Variant::Integer(-2));
        assert_eq!(
            calc("=MMULT(A1:B2,MUNIT(2))", &cells),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Integer(4),
            ])
        );
        match calc("=MINVERSE(A1:B2)", &cells) {
            Variant::Array(values) => {
                assert!((to_float(&values[0]).unwrap() + 2.0).abs() < 1e-9);
                assert!((to_float(&values[1]).unwrap() - 1.0).abs() < 1e-9);
                assert!((to_float(&values[2]).unwrap() - 1.5).abs() < 1e-9);
                assert!((to_float(&values[3]).unwrap() + 0.5).abs() < 1e-9);
            }
            other => panic!("MINVERSE: {:?}", other),
        }
        assert_eq!(calc("=MINVERSE(MUNIT(1))", &cells), Variant::Integer(1));
    }

    #[test]
    fn test_text_and_base_conversion_functions() {
        let cells = HashMap::new();
        assert_eq!(
            calc("=CLEAN(\"a\"&CHAR(10)&\"b\")", &cells),
            Variant::Str("ab".into())
        );
        assert_eq!(calc("=T(\"text\")", &cells), Variant::Str("text".into()));
        assert_eq!(calc("=T(42)", &cells), Variant::Str(String::new()));
        assert_eq!(
            calc("=FIXED(1234.567,2,FALSE)", &cells),
            Variant::Str("1,234.57".into())
        );
        assert_eq!(
            calc("=FIXED(1234.567,0,TRUE)", &cells),
            Variant::Str("1235".into())
        );
        assert_eq!(
            calc("=DOLLAR(-1234.5,2)", &cells),
            Variant::Str("-$1,234.50".into())
        );
        assert_eq!(calc("=BASE(255,16,4)", &cells), Variant::Str("00FF".into()));
        assert_eq!(calc("=DECIMAL(\"FF\",16)", &cells), Variant::Integer(255));
        assert_eq!(calc("=BASE(10,1)", &cells), Variant::Error(ExcelError::Num));
    }

    #[test]
    fn test_tocol_torow() {
        let c = HashMap::new();
        // TOCOL on a literal range of constants (flat sequence)
        assert_eq!(calc("=TOCOL(1)", &c), Variant::Integer(1));
        // with ignore=0 (none): empties pass through
        let mut c2 = HashMap::new();
        c2.insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(10),
            },
        );
        c2.insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Empty,
            },
        );
        c2.insert(
            (3, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(30),
            },
        );
        // ignore=1 → skip blanks
        assert_eq!(
            calc("=TOCOL(A1:A3, 1)", &c2),
            Variant::Array(vec![Variant::Integer(10), Variant::Integer(30)])
        );
        // TOROW same semantics
        assert_eq!(
            calc("=TOROW(A1:A3, 1)", &c2),
            Variant::Array(vec![Variant::Integer(10), Variant::Integer(30)])
        );
        let mut matrix = HashMap::new();
        for (row, values) in [(1, [1, 2, 3]), (2, [4, 5, 6])] {
            for (col, value) in values.into_iter().enumerate() {
                matrix.insert(
                    (row, col as u32 + 1),
                    CellContent {
                        formula: None,
                        value: Variant::Integer(value),
                    },
                );
            }
        }
        assert_eq!(
            calc("=TOCOL(A1:C2,0,TRUE)", &matrix),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(4),
                Variant::Integer(2),
                Variant::Integer(5),
                Variant::Integer(3),
                Variant::Integer(6)
            ])
        );
        assert_eq!(
            calc("=TOROW(A1:C2,0,TRUE)", &matrix),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(4),
                Variant::Integer(2),
                Variant::Integer(5),
                Variant::Integer(3),
                Variant::Integer(6)
            ])
        );
        assert_eq!(
            calc("=TOCOL(VSTACK(SEQUENCE(1,3),SEQUENCE(1,3,4)),0,TRUE)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(4),
                Variant::Integer(2),
                Variant::Integer(5),
                Variant::Integer(3),
                Variant::Integer(6)
            ])
        );
        assert_eq!(
            calc("=TOROW(HSTACK(SEQUENCE(2,1),SEQUENCE(2,1,3)),0,TRUE)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Integer(4)
            ])
        );
    }

    #[test]
    fn test_wraprows() {
        let c = HashMap::new();
        // SEQUENCE(1,6,1,1) → [1,2,3,4,5,6]
        // WRAPROWS([1,2,3,4,5,6], 3) → [[1,2,3],[4,5,6]] → flat: [1,2,3,4,5,6]
        assert_eq!(
            calc("=WRAPROWS(SEQUENCE(6),3)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Integer(5),
                Variant::Integer(6),
            ])
        );
        // WRAPROWS with padding: [1,2,3,4,5] wrap=3 → [[1,2,3],[4,5,0]]
        let mut c2 = HashMap::new();
        for i in 1u32..=5 {
            c2.insert(
                (i, 1),
                CellContent {
                    formula: None,
                    value: Variant::Integer(i as i64),
                },
            );
        }
        assert_eq!(
            calc("=WRAPROWS(A1:A5, 3, 0)", &c2),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Integer(5),
                Variant::Integer(0),
            ])
        );
    }

    #[test]
    fn test_wrapcols() {
        let c = HashMap::new();
        // WRAPCOLS(SEQUENCE(6), 2):
        //   vals=[1,2,3,4,5,6], wrap_count=2, n_cols=3
        //   col0=[1,2], col1=[3,4], col2=[5,6]
        //   row-major 2D (2 rows × 3 cols): [1,3,5, 2,4,6]
        assert_eq!(
            calc("=WRAPCOLS(SEQUENCE(6),2)", &c),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(3),
                Variant::Integer(5),
                Variant::Integer(2),
                Variant::Integer(4),
                Variant::Integer(6),
            ])
        );
        // WRAPCOLS([1,2,3,4,5], 2, 0): n_cols=3, last col=[5,0]
        // row-major: [1,3,5, 2,4,0]
        let mut c2 = HashMap::new();
        for i in 1u32..=5 {
            c2.insert(
                (i, 1),
                CellContent {
                    formula: None,
                    value: Variant::Integer(i as i64),
                },
            );
        }
        assert_eq!(
            calc("=WRAPCOLS(A1:A5, 2, 0)", &c2),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(3),
                Variant::Integer(5),
                Variant::Integer(2),
                Variant::Integer(4),
                Variant::Integer(0),
            ])
        );
    }

    #[test]
    fn test_unique_and_sort_two_dimensional_arrays() {
        let mut cells = HashMap::new();
        for (row, values) in [(1, [2, 20]), (2, [1, 10]), (3, [2, 20])] {
            for (col, value) in values.into_iter().enumerate() {
                cells.insert(
                    (row, (col + 1) as u32),
                    CellContent {
                        formula: None,
                        value: Variant::Integer(value),
                    },
                );
            }
        }
        assert_eq!(
            calc("=UNIQUE(A1:B3)", &cells),
            Variant::Array(vec![
                Variant::Integer(2),
                Variant::Integer(20),
                Variant::Integer(1),
                Variant::Integer(10),
            ])
        );
        assert_eq!(
            calc("=UNIQUE(A1:B3,1)", &cells),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(10)])
        );
        assert_eq!(
            calc("=SORT(A1:B3,1,-1)", &cells),
            Variant::Array(vec![
                Variant::Integer(2),
                Variant::Integer(20),
                Variant::Integer(2),
                Variant::Integer(20),
                Variant::Integer(1),
                Variant::Integer(10),
            ])
        );
        assert_eq!(
            calc("=SORT(A1:B3,1.5,1)", &cells),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=SORT(A1:B3,1,2)", &cells),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=SORT(SEQUENCE(3),1,2)", &cells),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=SORTBY(SEQUENCE(3),SEQUENCE(3),2)", &cells),
            Variant::Error(ExcelError::Value)
        );
        assert_eq!(
            calc("=SORT(A1:B3,1,-1,1)", &cells),
            Variant::Array(vec![
                Variant::Integer(20),
                Variant::Integer(2),
                Variant::Integer(10),
                Variant::Integer(1),
                Variant::Integer(20),
                Variant::Integer(2),
            ])
        );
        assert_eq!(
            calc("=SORTBY(A1:B3,C1:C3,-1)", &{
                let mut sorted_cells = cells.clone();
                for (row, value) in [(1, 5), (2, 9), (3, 7)] {
                    sorted_cells.insert(
                        (row, 3),
                        CellContent {
                            formula: None,
                            value: Variant::Integer(value),
                        },
                    );
                }
                sorted_cells
            }),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(10),
                Variant::Integer(2),
                Variant::Integer(20),
                Variant::Integer(2),
                Variant::Integer(20),
            ])
        );
        assert_eq!(
            calc("=SORTBY(A1:B3,C1:C3,-1,D1:D3,1)", &{
                let mut sorted_cells = cells.clone();
                for (row, value) in [(1, 2), (2, 2), (3, 1)] {
                    sorted_cells.insert(
                        (row, 3),
                        CellContent {
                            formula: None,
                            value: Variant::Integer(value),
                        },
                    );
                }
                for (row, value) in [(1, 30), (2, 10), (3, 20)] {
                    sorted_cells.insert(
                        (row, 4),
                        CellContent {
                            formula: None,
                            value: Variant::Integer(value),
                        },
                    );
                }
                sorted_cells
            }),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(10),
                Variant::Integer(2),
                Variant::Integer(20),
                Variant::Integer(2),
                Variant::Integer(20),
            ])
        );
        assert_eq!(
            calc("=SORTBY(A1:B3,C1:C3,2)", &{
                let mut invalid_order_cells = cells.clone();
                for (row, value) in [(1, 5), (2, 9), (3, 7)] {
                    invalid_order_cells.insert(
                        (row, 3),
                        CellContent {
                            formula: None,
                            value: Variant::Integer(value),
                        },
                    );
                }
                invalid_order_cells
            }),
            Variant::Error(ExcelError::Value)
        );
        let mut filter_cells = cells.clone();
        for (row, value) in [(1, false), (2, true), (3, false)] {
            filter_cells.insert(
                (row, 3),
                CellContent {
                    formula: None,
                    value: Variant::Boolean(value),
                },
            );
        }
        assert_eq!(
            calc("=FILTER(A1:B3,C1:C3)", &filter_cells),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(10)])
        );
        assert_eq!(
            calc("=FILTER(SEQUENCE(2,3),SEQUENCE(1,3))", &cells),
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Integer(4),
                Variant::Integer(5),
                Variant::Integer(6),
            ])
        );
        assert_eq!(
            calc("=FILTER(SEQUENCE(2,3),SEQUENCE(1,3)>1)", &cells),
            Variant::Array(vec![
                Variant::Integer(2),
                Variant::Integer(3),
                Variant::Integer(5),
                Variant::Integer(6),
            ])
        );
        assert_eq!(
            calc("=FILTER(SEQUENCE(2,3),SEQUENCE(2,1)>1)", &cells),
            Variant::Array(vec![
                Variant::Integer(4),
                Variant::Integer(5),
                Variant::Integer(6)
            ])
        );
        assert!(
            evaluate(
                &fparse("=FILTER(SEQUENCE(2,3),SEQUENCE(2,3))").unwrap(),
                &cells
            )
            .is_err()
        );
    }

    #[test]
    fn test_extended_statistical_functions() {
        let cells = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(4)),
            ((4, 1), Variant::Integer(8)),
        ]);
        assert_eq!(calc("=SUMSQ(A1:A4)", &cells), Variant::Integer(85));
        assert!(
            matches!(calc("=GEOMEAN(A1:A4)", &cells), Variant::Float(v) if (v - 2.8284271247461903).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=HARMEAN(A1:A4)", &cells), Variant::Float(v) if (v - 2.1333333333333333).abs() < 1e-12)
        );
        assert_eq!(calc("=DEVSQ(A1:A4)", &cells), Variant::Float(28.75));
        assert_eq!(calc("=AVEDEV(A1:A4)", &cells), Variant::Float(2.25));
        assert_eq!(
            calc("=PERCENTILE.EXC(A1:A4,0.5)", &cells),
            Variant::Integer(3)
        );
        assert_eq!(calc("=QUARTILE.EXC(A1:A4,1)", &cells), Variant::Float(1.25));
        assert_eq!(
            calc("=PERCENTRANK.EXC(A1:A4,4)", &cells),
            Variant::Float(0.6)
        );
        assert!(evaluate(&fparse("=GEOMEAN(A1:A4)").unwrap(), &HashMap::new()).is_err());
    }

    #[test]
    fn test_estimation_and_hypothesis_functions() {
        let cells = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(3)),
        ]);
        let two_samples = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Integer(2)),
            ((3, 1), Variant::Integer(3)),
            ((1, 2), Variant::Integer(1)),
            ((2, 2), Variant::Integer(2)),
            ((3, 2), Variant::Integer(4)),
        ]);
        assert!(
            matches!(calc("=FISHER(0.5)", &cells), Variant::Float(v) if (v - 0.5493061443340549).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=FISHERINV(0.5493061443340549)", &cells), Variant::Float(v) if (v - 0.5).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=STANDARDIZE(10,4,3)", &cells), Variant::Float(v) if (v - 2.0).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=Z.TEST(A1:A3,2)", &cells), Variant::Float(v) if (v - 0.5).abs() < 1e-8)
        );
        assert!(
            matches!(calc("=CONFIDENCE.NORM(0.05,1,100)", &cells), Variant::Float(v) if (v - 0.1959963988186964).abs() < 1e-6)
        );
        let confidence_t = calc("=CONFIDENCE.T(0.05,1,10)", &cells);
        assert!(matches!(confidence_t, Variant::Float(v) if (v - 0.7154368582207706).abs() < 1e-4));
        for test_type in 1..=3 {
            let formula = format!("=TTEST(A1:A3,B1:B3,2,{test_type})");
            assert!(matches!(calc(&formula, &two_samples), Variant::Float(v) if v.is_finite()));
        }
        assert!(matches!(
            calc("=T.TEST(A1:A3,B1:B3,2,2)", &two_samples),
            Variant::Float(v) if v.is_finite()
        ));
        let chi_cells = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(20)),
            ((1, 2), Variant::Integer(20)),
            ((2, 2), Variant::Integer(40)),
            ((1, 3), Variant::Integer(12)),
            ((2, 3), Variant::Integer(18)),
            ((1, 4), Variant::Integer(18)),
            ((2, 4), Variant::Integer(42)),
        ]);
        assert!(matches!(
            calc("=CHITEST(A1:B2,C1:D2)", &chi_cells),
            Variant::Float(value) if (0.0..=1.0).contains(&value)
        ));
        assert!(matches!(
            calc("=CHISQ.TEST(A1:B2,C1:D2)", &chi_cells),
            Variant::Float(value) if (0.0..=1.0).contains(&value)
        ));
        assert_eq!(
            calc("=CHITEST(A1:A3,B1:B2)", &chi_cells),
            Variant::Error(ExcelError::NA)
        );
        assert_eq!(calc("=FISHER(1)", &cells), Variant::Error(ExcelError::Num));
        assert_eq!(
            calc("=STANDARDIZE(1,1,0)", &cells),
            Variant::Error(ExcelError::Num)
        );
    }

    #[test]
    fn test_distribution_functions() {
        let cells = HashMap::new();
        assert!(
            matches!(calc("=NORM.S.DIST(0,TRUE)", &cells), Variant::Float(v) if (v - 0.5).abs() < 1e-9)
        );
        assert!(
            matches!(calc("=NORM.S.DIST(0,FALSE)", &cells), Variant::Float(v) if (v - 0.39894228).abs() < 1e-7)
        );
        assert!(matches!(calc("=NORM.S.INV(0.5)", &cells), Variant::Float(v) if v.abs() < 1e-6));
        assert!(
            matches!(calc("=PHI(0)", &cells), Variant::Float(v) if (v - 0.39894228).abs() < 1e-7)
        );
        assert!(matches!(calc("=GAUSS(0)", &cells), Variant::Float(v) if v.abs() < 1e-9));
        assert!(
            matches!(calc("=BINOM.DIST(2,4,0.5,FALSE)", &cells), Variant::Float(v) if (v - 0.375).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=BINOM.DIST(2,4,0.5,TRUE)", &cells), Variant::Float(v) if (v - 0.6875).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=POISSON.DIST(0,2,FALSE)", &cells), Variant::Float(v) if (v - 0.13533528).abs() < 1e-7)
        );
        assert!(
            matches!(calc("=POISSON.DIST(2,2,TRUE)", &cells), Variant::Float(v) if (v - 0.6766764).abs() < 1e-6)
        );
        assert_eq!(
            calc("=BINOM.DIST(5,4,0.5,FALSE)", &cells),
            Variant::Error(ExcelError::Num)
        );
        assert_eq!(
            calc("=POISSON.DIST(1,0,TRUE)", &cells),
            Variant::Error(ExcelError::Num)
        );
        assert!(
            matches!(calc("=GAMMA.DIST(3,2,3,TRUE)", &cells), Variant::Float(v) if (v - 0.2642411176571153).abs() < 1e-6)
        );
        assert!(
            matches!(calc("=GAMMA.DIST(3,2,3,FALSE)", &cells), Variant::Float(v) if (v - 0.1226264803904808).abs() < 1e-6)
        );
        assert!(
            matches!(calc("=GAMMA.INV(0.2642411176571153,2,3)", &cells), Variant::Float(v) if (v - 3.0).abs() < 1e-5)
        );
        assert!(
            matches!(calc("=CHISQ.DIST.RT(0,2)", &cells), Variant::Float(v) if (v - 1.0).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=F.DIST.RT(0,1,1)", &cells), Variant::Float(v) if (v - 1.0).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=F.INV(0.5,1,1)", &cells), Variant::Float(v) if (v - 1.0).abs() < 1e-5)
        );
        assert!(
            matches!(calc("=WEIBULL.DIST(1,2,1,TRUE)", &cells), Variant::Float(v) if (v - (1.0 - (-1.0f64).exp())).abs() < 1e-8)
        );
        assert!(
            matches!(calc("=EXPON.DIST(1,1,TRUE)", &cells), Variant::Float(v) if (v - (1.0 - (-1.0f64).exp())).abs() < 1e-8)
        );
        assert!(
            matches!(calc("=LOGNORM.DIST(1,0,1,TRUE)", &cells), Variant::Float(v) if (v - 0.5).abs() < 1e-7)
        );
        assert!(
            matches!(calc("=LOGNORM.INV(0.5,0,1)", &cells), Variant::Float(v) if (v - 1.0).abs() < 1e-7)
        );
        assert!(
            matches!(calc("=BINOM.DIST.RANGE(10,0.5,3,5)", &cells), Variant::Float(v) if (v - (582.0 / 1024.0)).abs() < 1e-12)
        );
        assert_eq!(
            calc("=BINOM.INV(4,0.5,0.6875)", &cells),
            Variant::Integer(2)
        );
        assert_eq!(calc("=BINOM.INV(4,0.5,0)", &cells), Variant::Integer(0));
        assert!(matches!(
            calc("=F.DIST.2T(1,1,1)", &cells),
            Variant::Float(v) if (v - 1.0).abs() < 1e-9
        ));
        assert!(
            matches!(calc("=NEGBINOM.DIST(3,2,0.5,FALSE)", &cells), Variant::Float(v) if (v - 0.125).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=NEGBINOM.DIST(3,2,0.5,TRUE)", &cells), Variant::Float(v) if (v - 0.8125).abs() < 1e-12)
        );
        assert!(
            matches!(calc("=HYPGEOM.DIST(1,4,2,10,FALSE)", &cells), Variant::Float(v) if (v - (16.0 / 30.0)).abs() < 1e-12)
        );
    }
}
