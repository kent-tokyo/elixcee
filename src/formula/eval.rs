use std::cmp::Ordering;
use std::collections::HashMap;
use std::cell::RefCell;

use crate::vm::{CellContent, ExcelError, Variant};
use super::ast::{BinOpKind, FormulaExpr};

// ── LET/LAMBDA name-binding stack ────────────────────────────────────────────
// A stack of binding frames; each frame is pushed by LET or a lambda call.

thread_local! {
    static BINDINGS: RefCell<Vec<HashMap<String, Variant>>> = RefCell::new(vec![]);
}

fn push_bindings(frame: HashMap<String, Variant>) {
    BINDINGS.with(|b| b.borrow_mut().push(frame));
}

fn pop_bindings() {
    BINDINGS.with(|b| { b.borrow_mut().pop(); });
}

fn lookup_binding(name: &str) -> Option<Variant> {
    BINDINGS.with(|b| {
        let stack = b.borrow();
        for frame in stack.iter().rev() {
            if let Some(v) = frame.get(name) { return Some(v.clone()); }
        }
        None
    })
}

pub fn evaluate(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    match expr {
        FormulaExpr::Number(n) => Ok(as_integer_if_whole(*n)),
        FormulaExpr::Str(s)    => Ok(Variant::Str(s.clone())),
        FormulaExpr::Bool(b)   => Ok(Variant::Boolean(*b)),
        FormulaExpr::CellRef { col, row } => Ok(
            cells.get(&(*row, *col))
                .map(|c| c.value.clone())
                .unwrap_or(Variant::Empty),
        ),
        FormulaExpr::Range { .. } => Err("Range cannot be used as a scalar value".into()),
        FormulaExpr::UnaryMinus(inner) => match evaluate(inner, cells)? {
            Variant::Integer(n) => Ok(Variant::Integer(-n)),
            Variant::Float(f)   => Ok(Variant::Float(-f)),
            other => Err(format!("Unary minus on non-numeric value: {}", other)),
        },
        FormulaExpr::BinOp { op, lhs, rhs } => eval_binop(op, lhs, rhs, cells),
        FormulaExpr::FuncCall { name, args } => {
            // 0-arg "call" is a name reference (LET / LAMBDA parameter)
            if args.is_empty() {
                if let Some(v) = lookup_binding(name) { return Ok(v); }
            }
            eval_func(name, args, cells)
        }
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
    if let Variant::Error(_) = &l { return Ok(l); }
    if let Variant::Error(_) = &r { return Ok(r); }
    match op {
        BinOpKind::Add | BinOpKind::Sub | BinOpKind::Mul | BinOpKind::Div => {
            let lf = to_float(&l)?;
            let rf = to_float(&r)?;
            let result = match op {
                BinOpKind::Add => lf + rf,
                BinOpKind::Sub => lf - rf,
                BinOpKind::Mul => lf * rf,
                BinOpKind::Div => {
                    if rf == 0.0 { return Ok(Variant::Error(ExcelError::DivZero)); }
                    lf / rf
                }
                _ => unreachable!(),
            };
            Ok(as_integer_if_whole(result))
        }
        BinOpKind::Concat => Ok(Variant::Str(format!("{}{}", l, r))),
        BinOpKind::Eq  => Ok(Variant::Boolean(variant_eq(&l, &r))),
        BinOpKind::Ne  => Ok(Variant::Boolean(!variant_eq(&l, &r))),
        BinOpKind::Lt  => Ok(Variant::Boolean(variant_cmp(&l, &r)? == Ordering::Less)),
        BinOpKind::Le  => Ok(Variant::Boolean(variant_cmp(&l, &r)? != Ordering::Greater)),
        BinOpKind::Gt  => Ok(Variant::Boolean(variant_cmp(&l, &r)? == Ordering::Greater)),
        BinOpKind::Ge  => Ok(Variant::Boolean(variant_cmp(&l, &r)? != Ordering::Less)),
    }
}

// ── Type helpers ──────────────────────────────────────────────────────────────

fn to_float(v: &Variant) -> Result<f64, String> {
    match v {
        Variant::Integer(n) => Ok(*n as f64),
        Variant::Float(f)   => Ok(*f),
        Variant::Boolean(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Variant::Date(s)    => Ok(*s as f64),
        Variant::Error(e)   => Err(e.to_string()),
        Variant::Empty      => Ok(0.0),
        Variant::Str(s)     => s.parse::<f64>()
            .map_err(|_| format!("Cannot convert '{}' to a number", s)),
        Variant::Array(_)   => Err("Cannot convert array to number".into()),
        Variant::Record(_)  => Err("Cannot convert record to number".into()),
    }
}

fn to_str(v: &Variant) -> String {
    match v {
        Variant::Str(s)     => s.clone(),
        Variant::Integer(n) => n.to_string(),
        Variant::Float(f)   => f.to_string(),
        Variant::Boolean(b) => if *b { "TRUE".into() } else { "FALSE".into() },
        Variant::Date(s)    => serial_to_display(*s),
        Variant::Error(e)   => e.as_str().to_string(),
        Variant::Empty      => String::new(),
        Variant::Array(a)   => a.iter().map(|x| to_str(x)).collect::<Vec<_>>().join(", "),
        Variant::Record(_)  => "[Record]".into(),
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

fn variant_eq(a: &Variant, b: &Variant) -> bool {
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

fn variant_cmp(a: &Variant, b: &Variant) -> Result<Ordering, String> {
    let af = to_float(a)?;
    let bf = to_float(b)?;
    af.partial_cmp(&bf).ok_or_else(|| "Cannot compare NaN values".into())
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
    match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None }
}

/// Expand a single formula argument into a flat list of Variant values.
/// Range → all cells row-major. Scalar → single value.
fn collect_values(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Vec<Variant>, String> {
    match expr {
        FormulaExpr::Range { c1, r1, c2, r2 } => {
            let (rmin, rmax) = (r1.min(r2), r1.max(r2));
            let (cmin, cmax) = (c1.min(c2), c1.max(c2));
            let mut vals = vec![];
            for row in *rmin..=*rmax {
                for col in *cmin..=*cmax {
                    vals.push(cell_val(cells, row, col));
                }
            }
            Ok(vals)
        }
        other => Ok(vec![evaluate(other, cells)?]),
    }
}

fn collect_all(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Vec<Variant>, String> {
    let mut out = vec![];
    for a in args { out.extend(collect_values(a, cells)?); }
    Ok(out)
}

fn cell_val(cells: &HashMap<(u32, u32), CellContent>, row: u32, col: u32) -> Variant {
    cells.get(&(row, col)).map(|c| c.value.clone()).unwrap_or(Variant::Empty)
}

// ── Function dispatch ─────────────────────────────────────────────────────────

fn eval_func(
    name: &str,
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    match name {
        "SUM"         => func_sum(args, cells),
        "AVERAGE"     => func_average(args, cells),
        "MIN"         => func_min(args, cells),
        "MAX"         => func_max(args, cells),
        "COUNT"       => func_count(args, cells),
        "COUNTA"      => func_counta(args, cells),
        "IF"          => func_if(args, cells),
        "AND"         => func_and(args, cells),
        "OR"          => func_or(args, cells),
        "NOT"         => func_not(args, cells),
        "IFERROR"     => func_iferror(args, cells),
        "LEFT"        => func_left(args, cells),
        "RIGHT"       => func_right(args, cells),
        "MID"         => func_mid(args, cells),
        "LEN"         => func_len(args, cells),
        "LEFTB"       => func_leftb(args, cells),
        "RIGHTB"      => func_rightb(args, cells),
        "MIDB"        => func_midb(args, cells),
        "LENB"        => func_lenb(args, cells),
        "ROUND"       => func_round(args, cells),
        "ROUNDUP"     => func_roundup(args, cells),
        "ROUNDDOWN"   => func_rounddown(args, cells),
        "CONCATENATE" => func_concatenate(args, cells),
        "CONCAT"      => func_concatenate(args, cells),
        "TEXT"        => func_text(args, cells),
        "COUNTIF"     => func_countif(args, cells),
        "SUMIF"       => func_sumif(args, cells),
        "SUMIFS"      => func_sumifs(args, cells),
        "COUNTIFS"    => func_countifs(args, cells),
        "MEDIAN"      => func_median(args, cells),
        "MODE.MULT"   => func_mode_mult(args, cells),
        "PRODUCT"     => func_product(args, cells),
        "ROW"         => func_row(args, cells),
        "DATE"        => func_date(args, cells),
        "TODAY"       => func_today(args, cells),
        "NETWORKDAYS" => func_networkdays(args, cells),
        "WORKDAY"     => func_workday(args, cells),
        "RANK"        => func_rank(args, cells),
        "IFS"         => func_ifs(args, cells),
        "XLOOKUP"     => func_xlookup(args, cells),
        "EOMONTH"     => func_eomonth(args, cells),
        "SUBTOTAL"    => func_subtotal(args, cells),
        "AGGREGATE"   => func_aggregate(args, cells),
        // -- Numerical --
        "AVERAGEIF"   => func_averageif(args, cells),
        "AVERAGEIFS"  => func_averageifs(args, cells),
        "INT"         => func_int(args, cells),
        "LARGE"       => func_large(args, cells),
        "MAXIFS"      => func_maxifs(args, cells),
        "MINIFS"      => func_minifs(args, cells),
        "MOD"         => func_mod(args, cells),
        "PERCENTILE" | "PERCENTILE.INC" => func_percentile(args, cells),
        "PERCENTRANK" | "PERCENTRANK.INC" => func_percentrank(args, cells),
        "RAND"        => func_rand(args, cells),
        "RANDBETWEEN" => func_randbetween(args, cells),
        "SMALL"       => func_small(args, cells),
        "SUMPRODUCT"  => func_sumproduct(args, cells),
        "TRUNC"       => func_trunc(args, cells),
        // -- String --
        "ASC"         => func_asc(args, cells),
        "CHAR"        => func_char(args, cells),
        "CODE"        => func_code(args, cells),
        "EXACT"       => func_exact(args, cells),
        "FIND"        => func_find(args, cells),
        "JIS"         => func_jis(args, cells),
        "LOWER"       => func_lower(args, cells),
        "PROPER"      => func_proper(args, cells),
        "REPLACE"     => func_replace(args, cells),
        "SEARCH"      => func_search(args, cells),
        "SUBSTITUTE"  => func_substitute(args, cells),
        "TEXTJOIN"    => func_textjoin(args, cells),
        "TEXTSPLIT"   => func_textsplit(args, cells),
        "TEXTBEFORE"  => func_textbefore(args, cells),
        "TEXTAFTER"   => func_textafter(args, cells),
        "VALUETOTEXT" => func_valuetotext(args, cells),
        "TRIM"        => func_trim(args, cells),
        "UNICHAR"     => func_char(args, cells),
        "UNICODE"     => func_code(args, cells),
        "UPPER"       => func_upper(args, cells),
        "VALUE"       => func_value(args, cells),
        // -- Date/Time --
        "YEAR"        => func_year(args, cells),
        "MONTH"       => func_month(args, cells),
        "DAY"         => func_day(args, cells),
        "WEEKDAY"     => func_weekday(args, cells),
        "DAYS"        => func_days(args, cells),
        "EDATE"       => func_edate(args, cells),
        "DATEDIF"     => func_datedif(args, cells),
        "DATEVALUE"   => func_datevalue(args, cells),
        "NOW"         => func_now(args, cells),
        "TIME"        => func_time_fn(args, cells),
        "TIMEVALUE"   => func_timevalue(args, cells),
        "HOUR"        => func_hour(args, cells),
        "MINUTE"      => func_minute(args, cells),
        "SECOND"      => func_second(args, cells),
        "NETWORKDAYS.INTL" => func_networkdays_intl(args, cells),
        "WORKDAY.INTL"     => func_workday_intl(args, cells),
        // -- Logic --
        "SWITCH"      => func_switch(args, cells),
        "XOR"         => func_xor(args, cells),
        // -- Lookup --
        "CHOOSE"      => func_choose(args, cells),
        "COLUMN"      => func_column(args, cells),
        "LOOKUP"      => func_lookup(args, cells),
        "XMATCH"      => func_xmatch(args, cells),
        // -- Info --
        "ISBLANK"     => func_isblank(args, cells),
        "ISERROR"     => func_iserror(args, cells),
        "ISERR"       => func_iserror(args, cells),
        "ISNA"        => func_isna(args, cells),
        "ISNUMBER"    => func_isnumber(args, cells),
        "ISTEXT"      => func_istext(args, cells),
        "ISLOGICAL"   => func_islogical(args, cells),
        "ISNONTEXT"   => func_isnontext(args, cells),
        "VLOOKUP"     => func_vlookup(args, cells),
        "HLOOKUP"     => func_hlookup(args, cells),
        "INDEX"       => func_index(args, cells),
        "MATCH"       => func_match_fn(args, cells),
        // ── Statistics ───────────────────────────────────────────────────────
        "STDEV" | "STDEV.S" => func_stdev_s(args, cells),
        "STDEVP" | "STDEV.P" => func_stdev_p(args, cells),
        "VAR" | "VAR.S"  => func_var_s(args, cells),
        "VARP" | "VAR.P" => func_var_p(args, cells),
        // ── Rounding ─────────────────────────────────────────────────────────
        "FLOOR" | "FLOOR.MATH"     => func_floor(args, cells),
        "CEILING" | "CEILING.MATH" => func_ceiling(args, cells),
        "MROUND"                   => func_mround(args, cells),
        // ── Math ─────────────────────────────────────────────────────────────
        "ABS"   => func_abs(args, cells),
        "SQRT"  => func_sqrt(args, cells),
        "POWER" => func_power(args, cells),
        "EXP"   => func_exp(args, cells),
        "LOG"   => func_log(args, cells),
        "LOG10" => func_log10(args, cells),
        "LN"    => func_ln(args, cells),
        // ── Trigonometry ──────────────────────────────────────────────────────
        "PI"      => func_pi(args, cells),
        "SIN"     => func_trig1(args, cells, f64::sin),
        "COS"     => func_trig1(args, cells, f64::cos),
        "TAN"     => func_trig1(args, cells, f64::tan),
        "ASIN"    => func_trig1(args, cells, f64::asin),
        "ACOS"    => func_trig1(args, cells, f64::acos),
        "ATAN"    => func_trig1(args, cells, f64::atan),
        "ATAN2"   => func_atan2(args, cells),
        "DEGREES" => func_trig1(args, cells, f64::to_degrees),
        "RADIANS" => func_trig1(args, cells, f64::to_radians),
        // ── Info ─────────────────────────────────────────────────────────────
        "COUNTBLANK" => func_countblank(args, cells),
        "ADDRESS"    => func_address(args, cells),
        "INDIRECT"   => func_indirect(args, cells),
        "OFFSET"     => func_offset(args, cells),
        // ── Array / spill functions ───────────────────────────────────────────
        "FILTER"    => func_filter(args, cells),
        "UNIQUE"    => func_unique(args, cells),
        "SORT"      => func_sort(args, cells),
        "SORTBY"    => func_sortby(args, cells),
        "SEQUENCE"  => func_sequence(args, cells),
        "TRANSPOSE" => func_transpose(args, cells),
        "TOCOL"      => func_tocol(args, cells),
        "TOROW"      => func_torow(args, cells),
        "WRAPCOLS"   => func_wrapcols(args, cells),
        "WRAPROWS"   => func_wraprows(args, cells),
        "RANDARRAY"  => func_randarray(args, cells),
        "TAKE"       => func_take(args, cells),
        "DROP"       => func_drop(args, cells),
        "VSTACK"     => func_vstack(args, cells),
        "HSTACK"     => func_hstack(args, cells),
        "CHOOSECOLS" => func_choosecols(args, cells),
        "CHOOSEROWS" => func_chooserows(args, cells),
        // ── Math / Financial ─────────────────────────────────────────────────
        "COMBIN"    => func_combin(args, cells),
        "PMT"       => func_pmt(args, cells),
        "FV"        => func_fv(args, cells),
        "PV"        => func_pv(args, cells),
        "NPER"      => func_nper(args, cells),
        "RATE"      => func_rate(args, cells),
        "IPMT"      => func_ipmt(args, cells),
        "PPMT"      => func_ppmt(args, cells),
        "NPV"       => func_npv(args, cells),
        "IRR"       => func_irr(args, cells),
        "MIRR"      => func_mirr(args, cells),
        "XNPV"      => func_xnpv(args, cells),
        "XIRR"      => func_xirr(args, cells),
        // ── Database ─────────────────────────────────────────────────────────
        "DGET"      => func_dget(args, cells),
        "DSUM"      => func_dsum(args, cells),
        "DAVERAGE"  => func_daverage(args, cells),
        "DCOUNT"    => func_dcount(args, cells),
        "DCOUNTA"   => func_dcounta(args, cells),
        "DMAX"      => func_dmax(args, cells),
        "DMIN"      => func_dmin(args, cells),
        // ── LET / higher-order ───────────────────────────────────────────────
        "LET"       => func_let(args, cells),
        "LAMBDA"    => func_lambda(args, cells),
        "MAP"       => func_map(args, cells),
        "REDUCE"    => func_reduce(args, cells),
        "SCAN"      => func_scan(args, cells),
        "BYROW"     => func_byrow(args, cells),
        "BYCOL"     => func_bycol(args, cells),
        _ => Ok(Variant::Error(ExcelError::Name)),
    }
}

// ── Arithmetic ────────────────────────────────────────────────────────────────

fn func_sum(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let sum: f64 = collect_all(args, cells)?.iter()
        .filter_map(|v| if matches!(v, Variant::Str(_)) { None } else { to_float(v).ok() })
        .sum();
    Ok(as_integer_if_whole(sum))
}

fn func_average(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let nums: Vec<f64> = collect_all(args, cells)?.iter()
        .filter_map(as_f64)
        .collect();
    if nums.is_empty() { return Err("AVERAGE: no numeric values".into()); }
    Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
}

fn func_min(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let min = collect_all(args, cells)?.iter()
        .filter_map(as_f64)
        .reduce(f64::min);
    min.map(as_integer_if_whole).ok_or_else(|| "MIN: no numeric values".into())
}

fn func_max(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let max = collect_all(args, cells)?.iter()
        .filter_map(as_f64)
        .reduce(f64::max);
    max.map(as_integer_if_whole).ok_or_else(|| "MAX: no numeric values".into())
}

fn func_count(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    Ok(Variant::Integer(
        collect_all(args, cells)?.iter()
            .filter(|v| matches!(v, Variant::Integer(_) | Variant::Float(_)))
            .count() as i64,
    ))
}

fn func_counta(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    Ok(Variant::Integer(
        collect_all(args, cells)?.iter()
            .filter(|v| !matches!(v, Variant::Empty))
            .count() as i64,
    ))
}

// ── Logical ───────────────────────────────────────────────────────────────────

fn func_if(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("IF requires 2 or 3 arguments".into()); }
    if is_truthy(&evaluate(&args[0], cells)?) {
        evaluate(&args[1], cells)
    } else if args.len() == 3 {
        evaluate(&args[2], cells)
    } else {
        Ok(Variant::Boolean(false))
    }
}

fn func_and(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    for a in args { if !is_truthy(&evaluate(a, cells)?) { return Ok(Variant::Boolean(false)); } }
    Ok(Variant::Boolean(true))
}

fn func_or(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    for a in args { if is_truthy(&evaluate(a, cells)?) { return Ok(Variant::Boolean(true)); } }
    Ok(Variant::Boolean(false))
}

fn func_not(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("NOT requires 1 argument".into()); }
    Ok(Variant::Boolean(!is_truthy(&evaluate(&args[0], cells)?)))
}

fn func_iferror(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("IFERROR requires 2 arguments".into()); }
    match evaluate(&args[0], cells) {
        Ok(Variant::Error(_)) | Err(_) => evaluate(&args[1], cells),
        Ok(v) => Ok(v),
    }
}

// ── Text ──────────────────────────────────────────────────────────────────────

fn func_left(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 { return Err("LEFT requires 1 or 2 arguments".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    let n = if args.len() == 2 { to_float(&evaluate(&args[1], cells)?)? as usize } else { 1 };
    Ok(Variant::Str(s.chars().take(n).collect()))
}

fn func_right(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 { return Err("RIGHT requires 1 or 2 arguments".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    let n = if args.len() == 2 { to_float(&evaluate(&args[1], cells)?)? as usize } else { 1 };
    let chars: Vec<char> = s.chars().collect();
    Ok(Variant::Str(chars[chars.len().saturating_sub(n)..].iter().collect()))
}

fn func_mid(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 3 { return Err("MID requires 3 arguments".into()); }
    let s     = to_str(&evaluate(&args[0], cells)?);
    let start = (to_float(&evaluate(&args[1], cells)?)? as usize).saturating_sub(1);
    let len   = to_float(&evaluate(&args[2], cells)?)? as usize;
    Ok(Variant::Str(s.chars().skip(start).take(len).collect()))
}

fn func_len(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("LEN requires 1 argument".into()); }
    Ok(Variant::Integer(to_str(&evaluate(&args[0], cells)?).chars().count() as i64))
}

// DBCS byte width: ASCII = 1, everything else = 2 (matches Excel's LENB/LEFTB/etc.)
fn char_byte_width(c: char) -> usize {
    if (c as u32) <= 0x7F { 1 } else { 2 }
}

fn str_byte_len(s: &str) -> usize {
    s.chars().map(char_byte_width).sum()
}

fn func_lenb(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("LENB requires 1 argument".into()); }
    Ok(Variant::Integer(str_byte_len(&to_str(&evaluate(&args[0], cells)?)) as i64))
}

fn func_leftb(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 { return Err("LEFTB requires 1 or 2 arguments".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    let n = if args.len() == 2 { to_float(&evaluate(&args[1], cells)?)? as usize } else { 1 };
    let mut result = String::new();
    let mut bytes = 0;
    for c in s.chars() {
        let w = char_byte_width(c);
        if bytes + w > n { break; }
        bytes += w;
        result.push(c);
    }
    Ok(Variant::Str(result))
}

fn func_rightb(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 { return Err("RIGHTB requires 1 or 2 arguments".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    let n = if args.len() == 2 { to_float(&evaluate(&args[1], cells)?)? as usize } else { 1 };
    let chars: Vec<char> = s.chars().collect();
    let mut result = Vec::new();
    let mut bytes = 0;
    for &c in chars.iter().rev() {
        let w = char_byte_width(c);
        if bytes + w > n { break; }
        bytes += w;
        result.push(c);
    }
    result.reverse();
    Ok(Variant::Str(result.into_iter().collect()))
}

fn func_midb(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 3 { return Err("MIDB requires 3 arguments".into()); }
    let s         = to_str(&evaluate(&args[0], cells)?);
    let start_b   = (to_float(&evaluate(&args[1], cells)?)? as usize).saturating_sub(1);
    let num_bytes = to_float(&evaluate(&args[2], cells)?)? as usize;
    let mut result = String::new();
    let mut pos = 0usize;
    for c in s.chars() {
        let w = char_byte_width(c);
        if pos >= start_b {
            let taken = result.chars().map(char_byte_width).sum::<usize>();
            if taken + w > num_bytes { break; }
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

fn func_round(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("ROUND requires 2 arguments".into()); }
    let num    = to_float(&evaluate(&args[0], cells)?)?;
    let digits = to_float(&evaluate(&args[1], cells)?)? as i32;
    let mult   = round_multiplier(digits);
    Ok(as_integer_if_whole((num * mult).round() / mult))
}

fn func_roundup(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("ROUNDUP requires 2 arguments".into()); }
    let num    = to_float(&evaluate(&args[0], cells)?)?;
    let digits = to_float(&evaluate(&args[1], cells)?)? as i32;
    let mult   = round_multiplier(digits);
    let result = if num >= 0.0 { (num * mult).ceil() } else { (num * mult).floor() };
    Ok(as_integer_if_whole(result / mult))
}

fn func_rounddown(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("ROUNDDOWN requires 2 arguments".into()); }
    let num    = to_float(&evaluate(&args[0], cells)?)?;
    let digits = to_float(&evaluate(&args[1], cells)?)? as i32;
    let mult   = round_multiplier(digits);
    let result = if num >= 0.0 { (num * mult).floor() } else { (num * mult).ceil() };
    Ok(as_integer_if_whole(result / mult))
}

fn func_concatenate(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let mut s = String::new();
    for a in args { s.push_str(&to_str(&evaluate(a, cells)?)); }
    Ok(Variant::Str(s))
}

fn func_text(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("TEXT requires 2 arguments".into()); }
    let n   = to_float(&evaluate(&args[0], cells)?)?;
    let fmt = to_str(&evaluate(&args[1], cells)?);
    Ok(Variant::Str(apply_text_format(n, &fmt)))
}

fn apply_text_format(n: f64, fmt: &str) -> String {
    let fmt_upper = fmt.to_uppercase();
    // Date format detection: contains YYYY, YY, or DD (not percentage)
    if !fmt.ends_with('%') && (fmt_upper.contains("YYYY") || fmt_upper.contains("YY") || fmt_upper.contains("DD")) {
        let (y, m, d) = serial_to_ymd(n as i64);
        return fmt
            .replace("YYYY", &format!("{:04}", y))
            .replace("yyyy", &format!("{:04}", y))
            .replace("YY",   &format!("{:02}", y % 100))
            .replace("yy",   &format!("{:02}", y % 100))
            .replace("MM",   &format!("{:02}", m))
            .replace("mm",   &format!("{:02}", m))
            .replace("DD",   &format!("{:02}", d))
            .replace("dd",   &format!("{:02}", d));
    }
    if fmt.ends_with('%') {
        let dec = fmt.trim_end_matches('%').split('.').nth(1).map(|s| s.len()).unwrap_or(0);
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
        FormulaExpr::Range { c1, r1, c2, r2 } => Ok((*c1, *r1, *c2, *r2)),
        _ => Err(format!("{}: table argument must be a range", fname)),
    }
}

fn func_vlookup(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 4 { return Err("VLOOKUP requires 3 or 4 arguments".into()); }
    let key   = evaluate(&args[0], cells)?;
    let (c1, r1, c2, r2) = require_range(&args[1], "VLOOKUP")?;
    let col_n = to_float(&evaluate(&args[2], cells)?)? as u32;
    if col_n == 0 { return Ok(Variant::Error(ExcelError::Value)); }
    let exact = if args.len() == 4 { !is_truthy(&evaluate(&args[3], cells)?) } else { false };
    let ret_col = c1 + col_n - 1;
    if ret_col > c2 { return Ok(Variant::Error(ExcelError::Ref)); }

    if exact {
        for row in r1..=r2 {
            if variant_eq(&cell_val(cells, row, c1), &key) {
                return Ok(cell_val(cells, row, ret_col));
            }
        }
        Ok(Variant::Error(ExcelError::NA))
    } else {
        let mut best = None;
        for row in r1..=r2 {
            match variant_cmp(&cell_val(cells, row, c1), &key) {
                Ok(Ordering::Less) | Ok(Ordering::Equal) => best = Some(row),
                _ => break,
            }
        }
        Ok(best.map(|row| cell_val(cells, row, ret_col)).unwrap_or(Variant::Error(ExcelError::NA)))
    }
}

fn func_hlookup(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 4 { return Err("HLOOKUP requires 3 or 4 arguments".into()); }
    let key   = evaluate(&args[0], cells)?;
    let (c1, r1, c2, _r2) = require_range(&args[1], "HLOOKUP")?;
    let row_n = to_float(&evaluate(&args[2], cells)?)? as u32;
    let exact = if args.len() == 4 { !is_truthy(&evaluate(&args[3], cells)?) } else { false };
    let ret_row = r1 + row_n - 1;

    if exact {
        for col in c1..=c2 {
            if variant_eq(&cell_val(cells, r1, col), &key) {
                return Ok(cell_val(cells, ret_row, col));
            }
        }
        Ok(Variant::Error(ExcelError::NA))
    } else {
        let mut best = None;
        for col in c1..=c2 {
            match variant_cmp(&cell_val(cells, r1, col), &key) {
                Ok(Ordering::Less) | Ok(Ordering::Equal) => best = Some(col),
                _ => break,
            }
        }
        best.map(|col| cell_val(cells, ret_row, col))
            .map(Ok)
            .unwrap_or(Ok(Variant::Error(ExcelError::NA)))
    }
}

fn func_index(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("INDEX requires 2 or 3 arguments".into()); }
    let (c1, r1, _c2, _r2) = require_range(&args[0], "INDEX")?;
    let row_off = to_float(&evaluate(&args[1], cells)?)? as u32;
    let col_off = if args.len() == 3 { to_float(&evaluate(&args[2], cells)?)? as u32 } else { 1 };
    Ok(cell_val(cells, r1 + row_off - 1, c1 + col_off - 1))
}

fn func_match_fn(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("MATCH requires 2 or 3 arguments".into()); }
    let key   = evaluate(&args[0], cells)?;
    let vals  = collect_values(&args[1], cells)?;
    let mtype = if args.len() == 3 { to_float(&evaluate(&args[2], cells)?)? as i32 } else { 1 };

    match mtype {
        0 => Ok(vals.iter().position(|v| variant_eq(v, &key))
                 .map(|i| Variant::Integer((i + 1) as i64))
                 .unwrap_or(Variant::Error(ExcelError::NA))),
        1 => {
            let mut best = None;
            for (i, v) in vals.iter().enumerate() {
                if matches!(variant_cmp(v, &key), Ok(Ordering::Less) | Ok(Ordering::Equal)) {
                    best = Some(i);
                }
            }
            Ok(best.map(|i| Variant::Integer((i + 1) as i64))
                .unwrap_or(Variant::Error(ExcelError::NA)))
        }
        -1 => {
            let mut best = None;
            for (i, v) in vals.iter().enumerate() {
                if matches!(variant_cmp(v, &key), Ok(Ordering::Greater) | Ok(Ordering::Equal)) {
                    best = Some(i);
                }
            }
            Ok(best.map(|i| Variant::Integer((i + 1) as i64))
                .unwrap_or(Variant::Error(ExcelError::NA)))
        }
        t => Err(format!("MATCH: invalid match_type {}", t)),
    }
}

// ── Criteria matching (for COUNTIF / SUMIF / COUNTIFS / SUMIFS) ──────────────

fn matches_criteria(val: &Variant, criteria: &Variant) -> bool {
    let crit_str = match criteria {
        Variant::Str(s) => s.clone(),
        other => return variant_eq(val, other),
    };
    // Comparison operator prefix
    let (op, rest) = if let Some(r) = crit_str.strip_prefix(">=") { (">=", r) }
        else if let Some(r) = crit_str.strip_prefix("<=") { ("<=", r) }
        else if let Some(r) = crit_str.strip_prefix("<>") { ("<>", r) }
        else if let Some(r) = crit_str.strip_prefix('>') { (">", r) }
        else if let Some(r) = crit_str.strip_prefix('<') { ("<", r) }
        else { ("=", crit_str.as_str()) };

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
    let val_f = match to_float(val) { Ok(f) => f, Err(_) => return false };
    let crit_f = match rest.parse::<f64>() { Ok(f) => f, Err(_) => return false };
    match op {
        ">"  => val_f > crit_f,
        ">=" => val_f >= crit_f,
        "<"  => val_f < crit_f,
        "<=" => val_f <= crit_f,
        "<>" => (val_f - crit_f).abs() > f64::EPSILON,
        _    => false,
    }
}

fn wildcard_match(text: &str, pattern: &str) -> bool {
    // Iterative NFA simulation — O(|text| × |pattern|), no exponential recursion.
    let t: Vec<char> = text.chars().collect();
    let p: Vec<char> = pattern.chars().collect();
    let (n, m) = (t.len(), p.len());
    // dp[i][j] = can pattern[..j] match text[..i]
    let mut dp = vec![vec![false; m + 1]; n + 1];
    dp[0][0] = true;
    for j in 1..=m { if p[j - 1] == '*' { dp[0][j] = dp[0][j - 1]; } }
    for i in 1..=n {
        for j in 1..=m {
            dp[i][j] = match p[j - 1] {
                '*' => dp[i - 1][j] || dp[i][j - 1],
                '?' => dp[i - 1][j - 1],
                c   => dp[i - 1][j - 1] && t[i - 1] == c,
            };
        }
    }
    dp[n][m]
}

/// Like wildcard_match but pattern only needs to match a prefix of text (for SEARCH positioning).
fn wildcard_match_prefix(text: &[char], pattern: &[char]) -> bool {
    // Iterative NFA — avoids exponential recursion for patterns with many '*'.
    let (n, m) = (text.len(), pattern.len());
    let mut dp = vec![vec![false; m + 1]; n + 1];
    dp[0][0] = true;
    for j in 1..=m { if pattern[j - 1] == '*' { dp[0][j] = dp[0][j - 1]; } }
    for i in 1..=n {
        for j in 1..=m {
            dp[i][j] = match pattern[j - 1] {
                '*' => dp[i - 1][j] || dp[i][j - 1],
                '?' => dp[i - 1][j - 1],
                c   => dp[i - 1][j - 1] && text[i - 1] == c,
            };
        }
        // Pattern consumed at any text prefix → match found
        if dp[i][m] { return true; }
    }
    dp[0][m] // handles empty text with pattern that reduces to empty via '*'
}

// ── Statistical ───────────────────────────────────────────────────────────────

fn func_countif(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("COUNTIF requires 2 arguments".into()); }
    let vals = collect_values(&args[0], cells)?;
    let crit = evaluate(&args[1], cells)?;
    Ok(Variant::Integer(vals.iter().filter(|v| matches_criteria(v, &crit)).count() as i64))
}

fn func_sumif(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("SUMIF requires 2 or 3 arguments".into()); }
    let range_vals = collect_values(&args[0], cells)?;
    let crit = evaluate(&args[1], cells)?;
    let sum_vals = if args.len() == 3 { collect_values(&args[2], cells)? } else { range_vals.clone() };
    let total: f64 = range_vals.iter().zip(sum_vals.iter())
        .filter(|(rv, _)| matches_criteria(rv, &crit))
        .filter_map(|(_, sv)| to_float(sv).ok())
        .sum();
    Ok(as_integer_if_whole(total))
}

fn func_sumifs(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() % 2 == 0 { return Err("SUMIFS requires sum_range then pairs of (range,criteria)".into()); }
    let sum_vals = collect_values(&args[0], cells)?;
    let n = sum_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        let range_vals = collect_values(&args[i], cells)?;
        let crit = evaluate(&args[i + 1], cells)?;
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_criteria(rv, &crit) { mask[j] = false; }
        }
        i += 2;
    }
    let total: f64 = sum_vals.iter().enumerate()
        .filter(|(j, _)| mask[*j])
        .filter_map(|(_, v)| to_float(v).ok())
        .sum();
    Ok(as_integer_if_whole(total))
}

fn func_countifs(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() % 2 != 0 { return Err("COUNTIFS requires pairs of (range,criteria)".into()); }
    let first_vals = collect_values(&args[0], cells)?;
    let n = first_vals.len();
    let mut mask = vec![true; n];
    let crit0 = evaluate(&args[1], cells)?;
    for (j, rv) in first_vals.iter().enumerate() {
        if !matches_criteria(rv, &crit0) { mask[j] = false; }
    }
    let mut i = 2;
    while i + 1 < args.len() {
        let range_vals = collect_values(&args[i], cells)?;
        let crit = evaluate(&args[i + 1], cells)?;
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_criteria(rv, &crit) { mask[j] = false; }
        }
        i += 2;
    }
    Ok(Variant::Integer(mask.iter().filter(|&&b| b).count() as i64))
}

fn func_median(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("MEDIAN requires at least 1 argument".into()); }
    let mut nums: Vec<f64> = collect_all(args, cells)?.iter()
        .filter_map(as_f64)
        .collect();
    if nums.is_empty() { return Err("MEDIAN: no numeric values".into()); }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let mid = nums.len() / 2;
    let result = if nums.len() % 2 == 0 { (nums[mid - 1] + nums[mid]) / 2.0 } else { nums[mid] };
    Ok(as_integer_if_whole(result))
}

fn func_mode_mult(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("MODE.MULT requires at least 1 argument".into()); }
    let vals: Vec<i64> = collect_all(args, cells)?.into_iter()
        .filter_map(|v| match v { Variant::Integer(n) => Some(n), Variant::Float(f) if f.fract() == 0.0 => Some(f as i64), _ => None })
        .collect();
    if vals.is_empty() { return Err("MODE.MULT: no numeric values".into()); }
    let mut freq: std::collections::HashMap<i64, usize> = std::collections::HashMap::new();
    for &v in &vals { *freq.entry(v).or_insert(0) += 1; }
    let max_freq = *freq.values().max().unwrap();
    // Return the first value (in original order) that has max frequency
    let mode = vals.into_iter().find(|v| freq[v] == max_freq).unwrap();
    Ok(Variant::Integer(mode))
}

fn func_product(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("PRODUCT requires at least 1 argument".into()); }
    let product: f64 = collect_all(args, cells)?.iter()
        .filter_map(as_f64)
        .fold(1.0, |acc, x| acc * x);
    Ok(as_integer_if_whole(product))
}

fn func_rank(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("RANK requires 2 or 3 arguments".into()); }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let vals: Vec<f64> = collect_values(&args[1], cells)?.iter()
        .filter_map(as_f64)
        .collect();
    let asc = if args.len() == 3 { to_float(&evaluate(&args[2], cells)?)? != 0.0 } else { false };
    let rank = if asc {
        vals.iter().filter(|&&v| v < num).count() + 1
    } else {
        vals.iter().filter(|&&v| v > num).count() + 1
    };
    Ok(Variant::Integer(rank as i64))
}

// ── Conditional ───────────────────────────────────────────────────────────────

fn func_ifs(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() % 2 != 0 { return Err("IFS requires an even number of arguments".into()); }
    let mut i = 0;
    while i + 1 < args.len() {
        if is_truthy(&evaluate(&args[i], cells)?) {
            return evaluate(&args[i + 1], cells);
        }
        i += 2;
    }
    Err("IFS: no condition matched".into())
}

// ── Date helpers ──────────────────────────────────────────────────────────────

fn is_leap(y: i32) -> bool { (y % 4 == 0 && y % 100 != 0) || y % 400 == 0 }

fn days_in_month(y: i32, m: u32) -> u32 {
    match m {
        1|3|5|7|8|10|12 => 31,
        4|6|9|11        => 30,
        2               => if is_leap(y) { 29 } else { 28 },
        _               => 0,
    }
}

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

pub fn serial_to_ymd_pub(s: i64) -> (i32, u32, u32) { serial_to_ymd(s) }

fn serial_to_ymd(mut s: i64) -> (i32, u32, u32) {
    // Undo the Excel leap-year bug offset for dates after serial 60
    if s > 60 { s -= 1; }
    // s is now days since Jan 1 1900 (1-based)
    let mut y = 1900i32;
    loop {
        let days = if is_leap(y) { 366i64 } else { 365 };
        if s <= days { break; }
        s -= days;
        y += 1;
    }
    let mut m = 1u32;
    loop {
        let dim = days_in_month(y, m) as i64;
        if s <= dim { break; }
        s -= dim;
        m += 1;
    }
    (y, m, s as u32)
}

fn serial_to_display(s: i64) -> String {
    let (y, m, d) = serial_to_ymd(s);
    format!("{:04}-{:02}-{:02}", y, m, d)
}

fn today_serial() -> i64 {
    let unix_days = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs() / 86400;
    unix_days as i64 + 25569 // 25569 = Excel serial of 1970-01-01
}

/// Day of week from Excel serial: 0=Sun,1=Mon,...,6=Sat (like Excel WEEKDAY default)
fn serial_weekday(serial: i64) -> u32 {
    // Serial 1 (Jan 1 1900) was a Monday. (serial - 1 + 1) % 7 gives 0=Mon..6=Sun
    // Shift to 0=Sun: ((serial - 1 + 1) % 7 + 1) % 7
    ((serial % 7 + 6) % 7) as u32  // 0=Sun,1=Mon,...,5=Fri,6=Sat
}

fn func_date(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 3 { return Err("DATE requires 3 arguments".into()); }
    let y = to_float(&evaluate(&args[0], cells)?)? as i32;
    let m = to_float(&evaluate(&args[1], cells)?)? as u32;
    let d = to_float(&evaluate(&args[2], cells)?)? as u32;
    Ok(Variant::Date(date_to_serial(y, m, d)))
}

fn func_today(args: &[FormulaExpr], _cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if !args.is_empty() { return Err("TODAY takes no arguments".into()); }
    Ok(Variant::Date(today_serial()))
}

fn func_eomonth(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("EOMONTH requires 2 arguments".into()); }
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

fn func_networkdays(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("NETWORKDAYS requires 2 or 3 arguments".into()); }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let end   = to_float(&evaluate(&args[1], cells)?)? as i64;
    let holidays: std::collections::HashSet<i64> = if args.len() == 3 {
        collect_values(&args[2], cells)?.iter()
            .filter_map(|v| to_float(v).ok().map(|f| f as i64))
            .collect()
    } else {
        std::collections::HashSet::new()
    };
    let (lo, hi, sign) = if start <= end { (start, end, 1i64) } else { (end, start, -1) };
    let count: i64 = (lo..=hi)
        .filter(|&d| {
            let wd = serial_weekday(d);
            wd != 0 && wd != 6 && !holidays.contains(&d)  // 0=Sun, 6=Sat
        })
        .count() as i64;
    Ok(Variant::Integer(count * sign))
}

// ── ROW ───────────────────────────────────────────────────────────────────────

fn func_row(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    match args.first() {
        None => Ok(Variant::Integer(1)),
        Some(FormulaExpr::CellRef { row, .. }) => Ok(Variant::Integer(*row as i64)),
        Some(FormulaExpr::Range { r1, .. }) => Ok(Variant::Integer(*r1 as i64)),
        Some(other) => { evaluate(other, cells)?; Ok(Variant::Integer(1)) }
    }
}

// ── XLOOKUP ───────────────────────────────────────────────────────────────────

fn func_xlookup(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 6 { return Err("XLOOKUP requires 3 to 6 arguments".into()); }
    let key        = evaluate(&args[0], cells)?;
    let lookup     = collect_values(&args[1], cells)?;
    let return_arr = collect_values(&args[2], cells)?;
    let not_found  = if args.len() >= 4 { Some(evaluate(&args[3], cells)?) } else { None };
    let match_mode = if args.len() >= 5 { to_float(&evaluate(&args[4], cells)?)? as i32 } else { 0 };
    let search_mode= if args.len() >= 6 { to_float(&evaluate(&args[5], cells)?)? as i32 } else { 1 };

    let iter: Box<dyn Iterator<Item = usize>> = match search_mode {
        -1 => Box::new((0..lookup.len()).rev()),
        _  => Box::new(0..lookup.len()),
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
            for i in 0..lookup.len() {
                if let Ok(v) = to_float(&lookup[i]) {
                    if v <= key_f {
                        if best.map_or(true, |(_, bv)| v > bv) { best = Some((i, v)); }
                    }
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
            for i in 0..lookup.len() {
                if let Ok(v) = to_float(&lookup[i]) {
                    if v >= key_f {
                        if best.map_or(true, |(_, bv)| v < bv) { best = Some((i, v)); }
                    }
                }
            }
            if let Some((i, _)) = best {
                return Ok(return_arr.get(i).cloned().unwrap_or(Variant::Empty));
            }
        }
        m => return Err(format!("XLOOKUP: unsupported match_mode {}", m)),
    }

    Ok(not_found.unwrap_or(Variant::Error(ExcelError::NA)))
}

// ── SUBTOTAL ──────────────────────────────────────────────────────────────────

fn func_subtotal(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("SUBTOTAL requires at least 2 arguments".into()); }
    let fn_num = to_float(&evaluate(&args[0], cells)?)? as u32;
    let rest = &args[1..];
    // 101-111 = ignore hidden rows (same behavior here since no hidden rows)
    match fn_num % 100 {
        1 => func_average(rest, cells),
        2 => func_count(rest, cells),
        3 => func_counta(rest, cells),
        4 => func_max(rest, cells),
        5 => func_min(rest, cells),
        6 => func_product(rest, cells),
        9 => func_sum(rest, cells),
        n => Err(format!("SUBTOTAL: unsupported function_num {}", n)),
    }
}

// ── Additional Numerical ─────────────────────────────────────────────────────

fn func_averageif(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("AVERAGEIF requires 2 or 3 arguments".into()); }
    let range_vals = collect_values(&args[0], cells)?;
    let crit = evaluate(&args[1], cells)?;
    let avg_vals = if args.len() == 3 { collect_values(&args[2], cells)? } else { range_vals.clone() };
    let nums: Vec<f64> = range_vals.iter().zip(avg_vals.iter())
        .filter(|(rv, _)| matches_criteria(rv, &crit))
        .filter_map(|(_, av)| to_float(av).ok())
        .collect();
    if nums.is_empty() { return Err("AVERAGEIF: no matching values".into()); }
    Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
}

fn func_averageifs(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() % 2 == 0 { return Err("AVERAGEIFS requires avg_range then pairs".into()); }
    let avg_vals = collect_values(&args[0], cells)?;
    let n = avg_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        let range_vals = collect_values(&args[i], cells)?;
        let crit = evaluate(&args[i + 1], cells)?;
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_criteria(rv, &crit) { mask[j] = false; }
        }
        i += 2;
    }
    let nums: Vec<f64> = avg_vals.iter().enumerate()
        .filter(|(j, _)| mask[*j])
        .filter_map(|(_, v)| to_float(v).ok())
        .collect();
    if nums.is_empty() { return Err("AVERAGEIFS: no matching values".into()); }
    Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
}

fn func_int(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("INT requires 1 argument".into()); }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    Ok(as_integer_if_whole(n.floor()))
}

fn func_large(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("LARGE requires 2 arguments".into()); }
    let mut nums: Vec<f64> = collect_values(&args[0], cells)?.iter()
        .filter_map(as_f64)
        .collect();
    let k = to_float(&evaluate(&args[1], cells)?)? as usize;
    if k == 0 || k > nums.len() { return Err("LARGE: k out of range".into()); }
    nums.sort_by(|a, b| b.partial_cmp(a).unwrap());
    Ok(as_integer_if_whole(nums[k - 1]))
}

fn func_small(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("SMALL requires 2 arguments".into()); }
    let mut nums: Vec<f64> = collect_values(&args[0], cells)?.iter()
        .filter_map(as_f64)
        .collect();
    let k = to_float(&evaluate(&args[1], cells)?)? as usize;
    if k == 0 || k > nums.len() { return Err("SMALL: k out of range".into()); }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    Ok(as_integer_if_whole(nums[k - 1]))
}

fn func_maxifs(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() % 2 == 0 { return Err("MAXIFS requires max_range then pairs".into()); }
    let max_vals = collect_values(&args[0], cells)?;
    let n = max_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        let range_vals = collect_values(&args[i], cells)?;
        let crit = evaluate(&args[i + 1], cells)?;
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_criteria(rv, &crit) { mask[j] = false; }
        }
        i += 2;
    }
    let max = max_vals.iter().enumerate()
        .filter(|(j, _)| mask[*j])
        .filter_map(|(_, v)| as_f64(v))
        .reduce(f64::max);
    max.map(as_integer_if_whole).ok_or_else(|| "MAXIFS: no matching values".into())
}

fn func_minifs(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() % 2 == 0 { return Err("MINIFS requires min_range then pairs".into()); }
    let min_vals = collect_values(&args[0], cells)?;
    let n = min_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        let range_vals = collect_values(&args[i], cells)?;
        let crit = evaluate(&args[i + 1], cells)?;
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_criteria(rv, &crit) { mask[j] = false; }
        }
        i += 2;
    }
    let min = min_vals.iter().enumerate()
        .filter(|(j, _)| mask[*j])
        .filter_map(|(_, v)| as_f64(v))
        .reduce(f64::min);
    min.map(as_integer_if_whole).ok_or_else(|| "MINIFS: no matching values".into())
}

fn func_mod(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("MOD requires 2 arguments".into()); }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    let d = to_float(&evaluate(&args[1], cells)?)?;
    if d == 0.0 { return Err("MOD: division by zero".into()); }
    Ok(as_integer_if_whole(n - d * (n / d).floor()))
}

fn func_percentile(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("PERCENTILE requires 2 arguments".into()); }
    let mut nums: Vec<f64> = collect_values(&args[0], cells)?.iter()
        .filter_map(as_f64)
        .collect();
    let k = to_float(&evaluate(&args[1], cells)?)?;
    if !(0.0..=1.0).contains(&k) { return Err("PERCENTILE: k must be 0 to 1".into()); }
    if nums.is_empty() { return Err("PERCENTILE: no numeric values".into()); }
    nums.sort_by(|a, b| a.partial_cmp(b).unwrap());
    let pos = k * (nums.len() - 1) as f64;
    let lo = pos.floor() as usize;
    let hi = pos.ceil() as usize;
    let result = if lo == hi { nums[lo] } else { nums[lo] + (pos - lo as f64) * (nums[hi] - nums[lo]) };
    Ok(as_integer_if_whole(result))
}

fn func_percentrank(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("PERCENTRANK requires 2 or 3 arguments".into()); }
    let nums: Vec<f64> = collect_values(&args[0], cells)?.iter()
        .filter_map(as_f64)
        .collect();
    let x = to_float(&evaluate(&args[1], cells)?)?;
    let sig = if args.len() == 3 { to_float(&evaluate(&args[2], cells)?)? as usize } else { 3 };
    if nums.is_empty() { return Err("PERCENTRANK: no values".into()); }
    let below = nums.iter().filter(|&&v| v < x).count();
    let equal = nums.iter().filter(|&&v| v == x).count();
    if equal == 0 { return Err("PERCENTRANK: value not in array".into()); }
    let rank = below as f64 / (nums.len() - 1) as f64;
    let mult = 10_f64.powi(sig as i32);
    Ok(Variant::Float((rank * mult).floor() / mult))
}

fn pseudo_rand() -> f64 {
    let nanos = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .subsec_nanos() as u64;
    let x = nanos.wrapping_mul(6364136223846793005).wrapping_add(1442695040888963407);
    (x >> 11) as f64 / ((1u64 << 53) as f64)
}

fn func_rand(args: &[FormulaExpr], _cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if !args.is_empty() { return Err("RAND takes no arguments".into()); }
    Ok(Variant::Float(pseudo_rand()))
}

fn func_randbetween(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("RANDBETWEEN requires 2 arguments".into()); }
    let lo = to_float(&evaluate(&args[0], cells)?)? as i64;
    let hi = to_float(&evaluate(&args[1], cells)?)? as i64;
    if lo > hi { return Err("RANDBETWEEN: bottom > top".into()); }
    let range = (hi - lo + 1) as f64;
    Ok(Variant::Integer(lo + (pseudo_rand() * range) as i64))
}

fn func_sumproduct(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("SUMPRODUCT requires at least 1 argument".into()); }
    let arrays: Vec<Vec<f64>> = args.iter()
        .map(|a| collect_values(a, cells).map(|vals|
            vals.iter().map(|v| to_float(v).unwrap_or(0.0)).collect()
        ))
        .collect::<Result<_, _>>()?;
    let len = arrays[0].len();
    let sum: f64 = (0..len).map(|i| arrays.iter().map(|arr| arr.get(i).copied().unwrap_or(0.0)).product::<f64>()).sum();
    Ok(as_integer_if_whole(sum))
}

fn func_trunc(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 { return Err("TRUNC requires 1 or 2 arguments".into()); }
    let num    = to_float(&evaluate(&args[0], cells)?)?;
    let digits = if args.len() == 2 { to_float(&evaluate(&args[1], cells)?)? as i32 } else { 0 };
    let mult   = 10_f64.powi(digits);
    Ok(as_integer_if_whole(num.signum() * (num.abs() * mult).floor() / mult))
}

fn func_aggregate(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 { return Err("AGGREGATE requires at least 3 arguments".into()); }
    let fn_num  = to_float(&evaluate(&args[0], cells)?)? as u32;
    let options = to_float(&evaluate(&args[1], cells)?)? as u32;
    let rest    = &args[2..];
    // options & 6 != 0 means "ignore errors"
    let _ignore_errors = options & 6 != 0;
    let nums: Vec<f64> = collect_all(rest, cells)?.iter()
        .filter_map(|v| match v {
            Variant::Integer(n) => Some(*n as f64),
            Variant::Float(f)   => Some(*f),
            _ => None,
        })
        .collect();
    // For fn_num that need filtered nums, handle ignore_errors by already filtering non-numeric
    match fn_num % 100 {
        1  => { if nums.is_empty() { return Err("AGGREGATE: no values".into()); } Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64)) }
        2  => Ok(Variant::Integer(nums.len() as i64)),
        3  => {
            let count = collect_all(rest, cells)?.iter().filter(|v| !matches!(v, Variant::Empty)).count();
            Ok(Variant::Integer(count as i64))
        }
        4  => nums.iter().copied().reduce(f64::max).map(as_integer_if_whole).ok_or_else(|| "AGGREGATE: no values".into()),
        5  => nums.iter().copied().reduce(f64::min).map(as_integer_if_whole).ok_or_else(|| "AGGREGATE: no values".into()),
        6  => Ok(as_integer_if_whole(nums.iter().fold(1.0, |a, &x| a * x))),
        9  => Ok(as_integer_if_whole(nums.iter().sum::<f64>())),
        12 => { // MEDIAN
            let mut s = nums.clone(); s.sort_by(|a,b| a.partial_cmp(b).unwrap());
            if s.is_empty() { return Err("AGGREGATE: no values".into()); }
            let mid = s.len() / 2;
            let r = if s.len() % 2 == 0 { (s[mid-1]+s[mid])/2.0 } else { s[mid] };
            Ok(as_integer_if_whole(r))
        }
        14 => { // LARGE
            let mut s = nums.clone(); s.sort_by(|a,b| b.partial_cmp(a).unwrap());
            s.first().copied().map(as_integer_if_whole).ok_or_else(|| "AGGREGATE: no values".into())
        }
        15 => { // SMALL
            let mut s = nums.clone(); s.sort_by(|a,b| a.partial_cmp(b).unwrap());
            s.first().copied().map(as_integer_if_whole).ok_or_else(|| "AGGREGATE: no values".into())
        }
        16 => func_percentile(rest, cells),
        n  => Err(format!("AGGREGATE: unsupported function_num {}", n)),
    }
}

// ── String functions ──────────────────────────────────────────────────────────

fn func_upper(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("UPPER requires 1 argument".into()); }
    Ok(Variant::Str(to_str(&evaluate(&args[0], cells)?).to_uppercase()))
}

fn func_lower(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("LOWER requires 1 argument".into()); }
    Ok(Variant::Str(to_str(&evaluate(&args[0], cells)?).to_lowercase()))
}

fn func_proper(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("PROPER requires 1 argument".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    let mut cap = true;
    let result: String = s.chars().map(|c| {
        let out = if cap { c.to_uppercase().next().unwrap_or(c) } else { c.to_lowercase().next().unwrap_or(c) };
        cap = !c.is_alphanumeric();
        out
    }).collect();
    Ok(Variant::Str(result))
}

fn func_trim(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("TRIM requires 1 argument".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    let result = s.split_whitespace().collect::<Vec<_>>().join(" ");
    Ok(Variant::Str(result))
}

fn func_substitute(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 4 { return Err("SUBSTITUTE requires 3 or 4 arguments".into()); }
    let text = to_str(&evaluate(&args[0], cells)?);
    let old  = to_str(&evaluate(&args[1], cells)?);
    let new  = to_str(&evaluate(&args[2], cells)?);
    if old.is_empty() { return Ok(Variant::Str(text)); }
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

fn func_replace(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 4 { return Err("REPLACE requires 4 arguments".into()); }
    let text  = to_str(&evaluate(&args[0], cells)?);
    let start = (to_float(&evaluate(&args[1], cells)?)? as usize).saturating_sub(1);
    let len   = to_float(&evaluate(&args[2], cells)?)? as usize;
    let new   = to_str(&evaluate(&args[3], cells)?);
    let chars: Vec<char> = text.chars().collect();
    let end = (start + len).min(chars.len());
    let result: String = chars[..start].iter().chain(new.chars().collect::<Vec<_>>().iter()).chain(chars[end..].iter()).collect();
    Ok(Variant::Str(result))
}

fn func_find(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("FIND requires 2 or 3 arguments".into()); }
    let needle  = to_str(&evaluate(&args[0], cells)?);
    let haystack= to_str(&evaluate(&args[1], cells)?);
    let start   = if args.len() == 3 { (to_float(&evaluate(&args[2], cells)?)? as usize).saturating_sub(1) } else { 0 };
    let h_chars: Vec<char> = haystack.chars().collect();
    let n_chars: Vec<char> = needle.chars().collect();
    let pos = h_chars[start..].windows(n_chars.len()).position(|w| w == n_chars.as_slice());
    pos.map(|p| Variant::Integer((start + p + 1) as i64))
        .ok_or_else(|| "FIND: value not found".into())
}

fn func_search(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("SEARCH requires 2 or 3 arguments".into()); }
    let needle   = to_str(&evaluate(&args[0], cells)?).to_uppercase();
    let haystack = to_str(&evaluate(&args[1], cells)?);
    let start    = if args.len() == 3 { (to_float(&evaluate(&args[2], cells)?)? as usize).saturating_sub(1) } else { 0 };
    let h_chars: Vec<char> = haystack.chars().collect();
    let n_chars: Vec<char> = needle.chars().collect();
    // wildcard-aware case-insensitive search
    let h_upper: Vec<char> = h_chars.iter().map(|c| c.to_uppercase().next().unwrap_or(*c)).collect();
    for i in start..=h_upper.len() {
        if wildcard_match_prefix(&h_upper[i..], &n_chars) {
            return Ok(Variant::Integer((i + 1) as i64));
        }
    }
    Err("SEARCH: value not found".into())
}

fn func_exact(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("EXACT requires 2 arguments".into()); }
    let a = to_str(&evaluate(&args[0], cells)?);
    let b = to_str(&evaluate(&args[1], cells)?);
    Ok(Variant::Boolean(a == b))
}

fn func_textjoin(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 { return Err("TEXTJOIN requires at least 3 arguments".into()); }
    let delim       = to_str(&evaluate(&args[0], cells)?);
    let ignore_empty= is_truthy(&evaluate(&args[1], cells)?);
    let parts: Vec<String> = collect_all(&args[2..], cells)?.iter()
        .map(|v| to_str(v))
        .filter(|s| !ignore_empty || !s.is_empty())
        .collect();
    Ok(Variant::Str(parts.join(&delim)))
}

fn func_value(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("VALUE requires 1 argument".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    let cleaned = s.replace(',', "");
    cleaned.parse::<f64>()
        .map(as_integer_if_whole)
        .map_err(|_| format!("VALUE: cannot convert '{}' to number", s))
}

fn func_char(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("CHAR requires 1 argument".into()); }
    let n = to_float(&evaluate(&args[0], cells)?)? as u32;
    char::from_u32(n)
        .map(|c| Variant::Str(c.to_string()))
        .ok_or_else(|| format!("CHAR: invalid code {}", n))
}

fn func_code(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("CODE requires 1 argument".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    s.chars().next()
        .map(|c| Variant::Integer(c as i64))
        .ok_or_else(|| "CODE: empty string".into())
}

fn func_asc(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("ASC requires 1 argument".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    let result: String = s.chars().map(|c| {
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
                '\u{FF71}','\u{FF72}','\u{FF73}','\u{FF74}','\u{FF75}', // ア-オ
                '\u{FF76}','\u{FF77}','\u{FF78}','\u{FF79}','\u{FF7A}', // カ-コ
                '\u{FF7B}','\u{FF7C}','\u{FF7D}','\u{FF7E}','\u{FF7F}', // サ-ソ
                '\u{FF80}','\u{FF81}','\u{FF82}','\u{FF83}','\u{FF84}', // タ-ト
                '\u{FF85}','\u{FF86}','\u{FF87}','\u{FF88}','\u{FF89}', // ナ-ノ
                '\u{FF8A}','\u{FF8B}','\u{FF8C}','\u{FF8D}','\u{FF8E}', // ハ-ホ
                '\u{FF8F}','\u{FF90}','\u{FF91}','\u{FF92}','\u{FF93}', // マ-モ
                '\u{FF94}','\u{FF95}','\u{FF96}',                        // ヤユヨ
                '\u{FF97}','\u{FF98}','\u{FF99}','\u{FF9A}','\u{FF9B}', // ラ-ロ
                '\u{FF9C}','\u{FF9D}',                                   // ワン
            ];
            // Map full-width katakana code point to half-width index
            // Only map direct (non-voiced) correspondences; voiced/semi-voiced left as-is
            let idx_map: &[(u32, usize)] = &[
                (0x30A2,0),(0x30A4,1),(0x30A6,2),(0x30A8,3),(0x30AA,4),
                (0x30AB,5),(0x30AD,6),(0x30AF,7),(0x30B1,8),(0x30B3,9),
                (0x30B5,10),(0x30B7,11),(0x30B9,12),(0x30BB,13),(0x30BD,14),
                (0x30BF,15),(0x30C1,16),(0x30C4,17),(0x30C6,18),(0x30C8,19),
                (0x30CA,20),(0x30CB,21),(0x30CC,22),(0x30CD,23),(0x30CE,24),
                (0x30CF,25),(0x30D2,26),(0x30D5,27),(0x30D8,28),(0x30DB,29),
                (0x30DE,30),(0x30DF,31),(0x30E0,32),(0x30E1,33),(0x30E2,34),
                (0x30E4,35),(0x30E6,36),(0x30E8,37),
                (0x30E9,38),(0x30EA,39),(0x30EB,40),(0x30EC,41),(0x30ED,42),
                (0x30EF,43),(0x30F3,44),
            ];
            idx_map.iter().find(|&&(k,_)| k == cp).map(|&(_,i)| base[i]).unwrap_or(c)
        } else {
            c
        }
    }).collect();
    Ok(Variant::Str(result))
}

fn func_jis(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("JIS requires 1 argument".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    let result: String = s.chars().map(|c| {
        let cp = c as u32;
        // Half-width ASCII/punctuation U+0021-U+007E → full-width U+FF01-U+FF5E
        if (0x0021..=0x007E).contains(&cp) {
            char::from_u32(cp + 0xFEE0).unwrap_or(c)
        } else if cp == 0x0020 {
            '\u{3000}' // space → ideographic space
        } else {
            c
        }
    }).collect();
    Ok(Variant::Str(result))
}

// ── Date/Time functions ───────────────────────────────────────────────────────

fn func_year(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("YEAR requires 1 argument".into()); }
    let s = to_float(&evaluate(&args[0], cells)?)? as i64;
    Ok(Variant::Integer(serial_to_ymd(s).0 as i64))
}

fn func_month(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("MONTH requires 1 argument".into()); }
    let s = to_float(&evaluate(&args[0], cells)?)? as i64;
    Ok(Variant::Integer(serial_to_ymd(s).1 as i64))
}

fn func_day(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("DAY requires 1 argument".into()); }
    let s = to_float(&evaluate(&args[0], cells)?)? as i64;
    Ok(Variant::Integer(serial_to_ymd(s).2 as i64))
}

fn func_weekday(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() || args.len() > 2 { return Err("WEEKDAY requires 1 or 2 arguments".into()); }
    let serial = to_float(&evaluate(&args[0], cells)?)? as i64;
    let return_type = if args.len() == 2 { to_float(&evaluate(&args[1], cells)?)? as u32 } else { 1 };
    // serial_weekday: 0=Sun,1=Mon,...,5=Fri,6=Sat
    let wd = serial_weekday(serial);
    let result = match return_type {
        1 => wd + 1,           // Sun=1..Sat=7
        2 => (wd + 6) % 7 + 1, // Mon=1..Sun=7
        3 => (wd + 6) % 7,     // Mon=0..Sun=6
        _ => return Err(format!("WEEKDAY: unsupported return_type {}", return_type)),
    };
    Ok(Variant::Integer(result as i64))
}

fn func_days(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("DAYS requires 2 arguments".into()); }
    let end   = to_float(&evaluate(&args[0], cells)?)? as i64;
    let start = to_float(&evaluate(&args[1], cells)?)? as i64;
    Ok(Variant::Integer(end - start))
}

fn func_edate(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("EDATE requires 2 arguments".into()); }
    let start  = to_float(&evaluate(&args[0], cells)?)? as i64;
    let months = to_float(&evaluate(&args[1], cells)?)? as i32;
    let (mut y, mut m, d) = serial_to_ymd(start);
    let total = (m as i32 - 1) + months;
    y += total.div_euclid(12);
    m  = (total.rem_euclid(12) + 1) as u32;
    let d_clamped = d.min(days_in_month(y, m));
    Ok(Variant::Date(date_to_serial(y, m, d_clamped)))
}

fn func_datedif(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 3 { return Err("DATEDIF requires 3 arguments".into()); }
    let s1 = to_float(&evaluate(&args[0], cells)?)? as i64;
    let s2 = to_float(&evaluate(&args[1], cells)?)? as i64;
    let unit = to_str(&evaluate(&args[2], cells)?).to_uppercase();
    if s1 > s2 { return Err("DATEDIF: start_date > end_date".into()); }
    let (y1, m1, d1) = serial_to_ymd(s1);
    let (y2, m2, d2) = serial_to_ymd(s2);
    let result = match unit.as_str() {
        "Y"  => (y2 - y1 - if (m2, d2) < (m1, d1) { 1 } else { 0 }) as i64,
        "M"  => {
            let mut months = (y2 - y1) * 12 + (m2 as i32 - m1 as i32);
            if d2 < d1 { months -= 1; }
            months as i64
        }
        "D"  => s2 - s1,
        "MD" => {
            let d2i = d2 as i32; let d1i = d1 as i32;
            (if d2i >= d1i { d2i - d1i } else { days_in_month(y2, m2) as i32 - d1i + d2i }) as i64
        }
        "YM" => {
            let mut m = m2 as i32 - m1 as i32;
            if m < 0 { m += 12; }
            if d2 < d1 { m -= 1; if m < 0 { m += 12; } }
            m as i64
        }
        "YD" => {
            let base = date_to_serial(y2, m1, d1);
            if base <= s2 { s2 - base } else { s2 - date_to_serial(y2 - 1, m1, d1) }
        }
        u => return Err(format!("DATEDIF: unknown unit '{}'", u)),
    };
    Ok(Variant::Integer(result))
}

fn func_datevalue(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("DATEVALUE requires 1 argument".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    // Support YYYY/MM/DD and YYYY-MM-DD
    let parts: Vec<&str> = if s.contains('/') { s.splitn(3, '/').collect() } else { s.splitn(3, '-').collect() };
    if parts.len() == 3 {
        if let (Ok(y), Ok(m), Ok(d)) = (parts[0].trim().parse::<i32>(), parts[1].trim().parse::<u32>(), parts[2].trim().parse::<u32>()) {
            return Ok(Variant::Date(date_to_serial(y, m, d)));
        }
    }
    Err(format!("DATEVALUE: cannot parse '{}'", s))
}

fn func_now(args: &[FormulaExpr], _cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if !args.is_empty() { return Err("NOW takes no arguments".into()); }
    let secs = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs();
    let unix_days = secs / 86400;
    let frac = (secs % 86400) as f64 / 86400.0;
    Ok(Variant::Float(unix_days as f64 + 25569.0 + frac))
}

fn func_time_fn(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 3 { return Err("TIME requires 3 arguments".into()); }
    let h = to_float(&evaluate(&args[0], cells)?)?;
    let m = to_float(&evaluate(&args[1], cells)?)?;
    let s = to_float(&evaluate(&args[2], cells)?)?;
    Ok(Variant::Float((h * 3600.0 + m * 60.0 + s) / 86400.0))
}

fn func_timevalue(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("TIMEVALUE requires 1 argument".into()); }
    let s = to_str(&evaluate(&args[0], cells)?);
    let parts: Vec<&str> = s.splitn(3, ':').collect();
    if parts.len() >= 2 {
        if let (Ok(h), Ok(m)) = (parts[0].trim().parse::<f64>(), parts[1].trim().parse::<f64>()) {
            let sec = if parts.len() == 3 { parts[2].trim().parse::<f64>().unwrap_or(0.0) } else { 0.0 };
            return Ok(Variant::Float((h * 3600.0 + m * 60.0 + sec) / 86400.0));
        }
    }
    Err(format!("TIMEVALUE: cannot parse '{}'", s))
}

fn serial_frac(v: f64) -> f64 { v.fract().abs() }

fn func_hour(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("HOUR requires 1 argument".into()); }
    let v = to_float(&evaluate(&args[0], cells)?)?;
    Ok(Variant::Integer((serial_frac(v) * 24.0) as i64))
}

fn func_minute(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("MINUTE requires 1 argument".into()); }
    let v = to_float(&evaluate(&args[0], cells)?)?;
    Ok(Variant::Integer(((serial_frac(v) * 1440.0) % 60.0) as i64))
}

fn func_second(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("SECOND requires 1 argument".into()); }
    let v = to_float(&evaluate(&args[0], cells)?)?;
    Ok(Variant::Integer(((serial_frac(v) * 86400.0) % 60.0) as i64))
}

fn parse_weekend_mask(v: &Variant) -> Result<[bool; 7], String> {
    // Returns [Mon,Tue,Wed,Thu,Fri,Sat,Sun] true = is weekend
    match v {
        Variant::Str(s) if s.len() == 7 => {
            let mut mask = [false; 7];
            for (i, c) in s.chars().enumerate() { mask[i] = c == '1'; }
            Ok(mask)
        }
        _ => {
            let n = to_float(v)? as u32;
            Ok(match n {
                1  => [false,false,false,false,false,true,true],
                2  => [true,false,false,false,false,false,true],
                3  => [true,true,false,false,false,false,false],
                4  => [false,true,true,false,false,false,false],
                5  => [false,false,true,true,false,false,false],
                6  => [false,false,false,true,true,false,false],
                7  => [false,false,false,false,true,true,false],
                11 => [false,false,false,false,false,false,true],
                12 => [true,false,false,false,false,false,false],
                13 => [false,true,false,false,false,false,false],
                14 => [false,false,true,false,false,false,false],
                15 => [false,false,false,true,false,false,false],
                16 => [false,false,false,false,true,false,false],
                17 => [false,false,false,false,false,true,false],
                _  => return Err(format!("NETWORKDAYS.INTL: invalid weekend {}", n)),
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

fn func_networkdays_intl(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 4 { return Err("NETWORKDAYS.INTL requires 2 to 4 arguments".into()); }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let end   = to_float(&evaluate(&args[1], cells)?)? as i64;
    let mask = if args.len() >= 3 {
        parse_weekend_mask(&evaluate(&args[2], cells)?)?
    } else {
        [false,false,false,false,false,true,true] // default Sat+Sun
    };
    let holidays: std::collections::HashSet<i64> = if args.len() == 4 {
        collect_values(&args[3], cells)?.iter().filter_map(|v| to_float(v).ok().map(|f| f as i64)).collect()
    } else { std::collections::HashSet::new() };
    let (lo, hi, sign) = if start <= end { (start, end, 1i64) } else { (end, start, -1) };
    let count: i64 = (lo..=hi).filter(|&d| !is_weekend_intl(d, &mask) && !holidays.contains(&d)).count() as i64;
    Ok(Variant::Integer(count * sign))
}

fn func_workday_intl(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 4 { return Err("WORKDAY.INTL requires 2 to 4 arguments".into()); }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let days  = to_float(&evaluate(&args[1], cells)?)? as i64;
    let mask = if args.len() >= 3 {
        parse_weekend_mask(&evaluate(&args[2], cells)?)?
    } else {
        [false,false,false,false,false,true,true]
    };
    let holidays: std::collections::HashSet<i64> = if args.len() == 4 {
        collect_values(&args[3], cells)?.iter().filter_map(|v| to_float(v).ok().map(|f| f as i64)).collect()
    } else { std::collections::HashSet::new() };
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

fn func_switch(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 { return Err("SWITCH requires at least 3 arguments".into()); }
    let expr = evaluate(&args[0], cells)?;
    let mut i = 1;
    while i + 1 < args.len() {
        let v = evaluate(&args[i], cells)?;
        if variant_eq(&expr, &v) { return evaluate(&args[i + 1], cells); }
        i += 2;
    }
    // odd remaining arg = default
    if args.len() % 2 == 0 { evaluate(&args[args.len() - 1], cells) }
    else { Err("SWITCH: no match found".into()) }
}

fn func_xor(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("XOR requires at least 1 argument".into()); }
    let count = args.iter()
        .map(|a| evaluate(a, cells).map(|v| is_truthy(&v)))
        .collect::<Result<Vec<_>, _>>()?
        .into_iter()
        .filter(|&b| b)
        .count();
    Ok(Variant::Boolean(count % 2 == 1))
}

// ── Lookup ────────────────────────────────────────────────────────────────────

fn func_choose(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("CHOOSE requires at least 2 arguments".into()); }
    let idx = to_float(&evaluate(&args[0], cells)?)? as usize;
    if idx < 1 || idx >= args.len() { return Err("CHOOSE: index out of range".into()); }
    evaluate(&args[idx], cells)
}

fn func_column(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    match args.first() {
        None => Ok(Variant::Integer(1)),
        Some(FormulaExpr::CellRef { col, .. }) => Ok(Variant::Integer(*col as i64)),
        Some(FormulaExpr::Range { c1, .. }) => Ok(Variant::Integer(*c1 as i64)),
        Some(other) => { evaluate(other, cells)?; Ok(Variant::Integer(1)) }
    }
}

fn func_lookup(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("LOOKUP requires 2 or 3 arguments".into()); }
    let key    = evaluate(&args[0], cells)?;
    let lookup = collect_values(&args[1], cells)?;
    let result = if args.len() == 3 { collect_values(&args[2], cells)? } else { lookup.clone() };
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

fn func_xmatch(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 4 { return Err("XMATCH requires 2 to 4 arguments".into()); }
    let key        = evaluate(&args[0], cells)?;
    let lookup     = collect_values(&args[1], cells)?;
    let match_mode = if args.len() >= 3 { to_float(&evaluate(&args[2], cells)?)? as i32 } else { 0 };
    let search_mode= if args.len() >= 4 { to_float(&evaluate(&args[3], cells)?)? as i32 } else { 1 };

    let iter: Box<dyn Iterator<Item = usize>> = match search_mode {
        -1 => Box::new((0..lookup.len()).rev()),
        _  => Box::new(0..lookup.len()),
    };

    match match_mode {
        0 => {
            for i in iter {
                if variant_eq(&lookup[i], &key) { return Ok(Variant::Integer((i + 1) as i64)); }
            }
        }
        -1 => {
            let key_f = to_float(&key)?;
            let mut best: Option<(usize, f64)> = None;
            for i in 0..lookup.len() {
                if let Ok(v) = to_float(&lookup[i]) {
                    if v <= key_f && best.map_or(true, |(_,bv)| v > bv) { best = Some((i, v)); }
                }
            }
            if let Some((i,_)) = best { return Ok(Variant::Integer((i+1) as i64)); }
        }
        1 => {
            let key_f = to_float(&key)?;
            let mut best: Option<(usize, f64)> = None;
            for i in 0..lookup.len() {
                if let Ok(v) = to_float(&lookup[i]) {
                    if v >= key_f && best.map_or(true, |(_,bv)| v < bv) { best = Some((i, v)); }
                }
            }
            if let Some((i,_)) = best { return Ok(Variant::Integer((i+1) as i64)); }
        }
        m => return Err(format!("XMATCH: unsupported match_mode {}", m)),
    }
    Ok(Variant::Error(ExcelError::NA))
}

// ── Info functions ────────────────────────────────────────────────────────────

fn func_isblank(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("ISBLANK requires 1 argument".into()); }
    Ok(Variant::Boolean(matches!(evaluate(&args[0], cells)?, Variant::Empty)))
}

fn func_iserror(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("ISERROR requires 1 argument".into()); }
    Ok(Variant::Boolean(matches!(evaluate(&args[0], cells), Ok(Variant::Error(_)) | Err(_))))
}

fn func_isna(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("ISNA requires 1 argument".into()); }
    Ok(Variant::Boolean(matches!(evaluate(&args[0], cells), Ok(Variant::Error(ExcelError::NA)))))
}

fn func_isnumber(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("ISNUMBER requires 1 argument".into()); }
    Ok(Variant::Boolean(matches!(evaluate(&args[0], cells)?, Variant::Integer(_) | Variant::Float(_))))
}

fn func_istext(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("ISTEXT requires 1 argument".into()); }
    Ok(Variant::Boolean(matches!(evaluate(&args[0], cells)?, Variant::Str(_))))
}

fn func_islogical(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("ISLOGICAL requires 1 argument".into()); }
    Ok(Variant::Boolean(matches!(evaluate(&args[0], cells)?, Variant::Boolean(_))))
}

fn func_isnontext(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("ISNONTEXT requires 1 argument".into()); }
    Ok(Variant::Boolean(!matches!(evaluate(&args[0], cells).unwrap_or(Variant::Empty), Variant::Str(_))))
}

// ── Statistics ────────────────────────────────────────────────────────────────

fn collect_nums(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Vec<f64>, String> {
    Ok(collect_all(args, cells)?.into_iter()
        .filter_map(|v| match v { Variant::Integer(n) => Some(n as f64), Variant::Float(f) => Some(f), _ => None })
        .collect())
}

fn func_stdev_s(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let nums = collect_nums(args, cells)?;
    if nums.len() < 2 { return Err("STDEV requires at least 2 values".into()); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let var = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64;
    Ok(Variant::Float(var.sqrt()))
}

fn func_stdev_p(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let nums = collect_nums(args, cells)?;
    if nums.is_empty() { return Err("STDEVP requires at least 1 value".into()); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    let var = nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / nums.len() as f64;
    Ok(Variant::Float(var.sqrt()))
}

fn func_var_s(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let nums = collect_nums(args, cells)?;
    if nums.len() < 2 { return Err("VAR requires at least 2 values".into()); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    Ok(Variant::Float(nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / (nums.len() - 1) as f64))
}

fn func_var_p(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let nums = collect_nums(args, cells)?;
    if nums.is_empty() { return Err("VARP requires at least 1 value".into()); }
    let mean = nums.iter().sum::<f64>() / nums.len() as f64;
    Ok(Variant::Float(nums.iter().map(|x| (x - mean).powi(2)).sum::<f64>() / nums.len() as f64))
}

// ── Rounding ──────────────────────────────────────────────────────────────────

fn func_floor(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("FLOOR requires at least 1 argument".into()); }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let sig = if args.len() >= 2 { to_float(&evaluate(&args[1], cells)?)? } else { 1.0 };
    if sig == 0.0 { return Ok(Variant::Integer(0)); }
    Ok(as_integer_if_whole((num / sig).floor() * sig))
}

fn func_ceiling(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("CEILING requires at least 1 argument".into()); }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let sig = if args.len() >= 2 { to_float(&evaluate(&args[1], cells)?)? } else { 1.0 };
    if sig == 0.0 { return Ok(Variant::Integer(0)); }
    Ok(as_integer_if_whole((num / sig).ceil() * sig))
}

fn func_mround(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("MROUND requires 2 arguments".into()); }
    let num  = to_float(&evaluate(&args[0], cells)?)?;
    let mult = to_float(&evaluate(&args[1], cells)?)?;
    if mult == 0.0 { return Ok(Variant::Integer(0)); }
    if (num < 0.0) != (mult < 0.0) { return Ok(Variant::Error(ExcelError::Num)); }
    Ok(as_integer_if_whole((num / mult).round() * mult))
}

// ── Math ──────────────────────────────────────────────────────────────────────

fn func_abs(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("ABS requires 1 argument".into()); }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    Ok(as_integer_if_whole(n.abs()))
}

fn func_sqrt(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("SQRT requires 1 argument".into()); }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    if n < 0.0 { return Ok(Variant::Error(ExcelError::Num)); }
    Ok(as_integer_if_whole(n.sqrt()))
}

fn func_power(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("POWER requires 2 arguments".into()); }
    let base = to_float(&evaluate(&args[0], cells)?)?;
    let exp  = to_float(&evaluate(&args[1], cells)?)?;
    Ok(as_integer_if_whole(base.powf(exp)))
}

fn func_exp(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("EXP requires 1 argument".into()); }
    Ok(Variant::Float(to_float(&evaluate(&args[0], cells)?)?.exp()))
}

fn func_log(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("LOG requires at least 1 argument".into()); }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    if n <= 0.0 { return Ok(Variant::Error(ExcelError::Num)); }
    let base = if args.len() >= 2 { to_float(&evaluate(&args[1], cells)?)? } else { 10.0 };
    if base <= 0.0 || base == 1.0 { return Ok(Variant::Error(ExcelError::Num)); }
    Ok(Variant::Float(n.log(base)))
}

fn func_log10(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("LOG10 requires 1 argument".into()); }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    if n <= 0.0 { return Ok(Variant::Error(ExcelError::Num)); }
    Ok(Variant::Float(n.log10()))
}

fn func_ln(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 1 { return Err("LN requires 1 argument".into()); }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    if n <= 0.0 { return Ok(Variant::Error(ExcelError::Num)); }
    Ok(Variant::Float(n.ln()))
}

// ── Trigonometry ──────────────────────────────────────────────────────────────

fn func_pi(args: &[FormulaExpr], _cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if !args.is_empty() { return Err("PI takes no arguments".into()); }
    Ok(Variant::Float(std::f64::consts::PI))
}

fn func_trig1(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    f: fn(f64) -> f64,
) -> Result<Variant, String> {
    if args.len() != 1 { return Err("Trig function requires 1 argument".into()); }
    let n = to_float(&evaluate(&args[0], cells)?)?;
    Ok(as_integer_if_whole(f(n)))
}

fn func_atan2(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("ATAN2 requires 2 arguments".into()); }
    let x = to_float(&evaluate(&args[0], cells)?)?; // Excel: x-coordinate first
    let y = to_float(&evaluate(&args[1], cells)?)?; // then y-coordinate
    if x == 0.0 && y == 0.0 { return Ok(Variant::Error(ExcelError::DivZero)); }
    Ok(Variant::Float(y.atan2(x)))
}

// ── Info ──────────────────────────────────────────────────────────────────────

fn func_countblank(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let n = collect_all(args, cells)?.into_iter()
        .filter(|v| matches!(v, Variant::Empty) || matches!(v, Variant::Str(s) if s.is_empty()))
        .count();
    Ok(Variant::Integer(n as i64))
}

fn func_address(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("ADDRESS requires at least 2 arguments".into()); }
    let row = to_float(&evaluate(&args[0], cells)?)? as u32;
    let col = to_float(&evaluate(&args[1], cells)?)? as u32;
    let abs_num = if args.len() >= 3 { to_float(&evaluate(&args[2], cells)?)? as i32 } else { 1 };
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

fn func_let(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    // =LET(x, val1, y, val2, ..., result_expr)
    // Must have odd number of args: 2*n+1
    if args.len() < 3 || args.len() % 2 == 0 {
        return Err("LET requires an odd number of arguments: LET(name, val, ..., result)".into());
    }
    let mut frame: HashMap<String, Variant> = HashMap::new();
    let mut i = 0;
    while i < args.len() - 1 {
        let name = match &args[i] {
            FormulaExpr::FuncCall { name, args } if args.is_empty() => name.clone(),
            _ => return Err("LET: name arguments must be identifiers".into()),
        };
        let val = evaluate(&args[i + 1], cells)?; // evaluated with current scope
        frame.insert(name, val);
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
            if args.len() < 2 { return Err("LAMBDA requires at least 2 arguments".into()); }
            let params: Result<Vec<String>, String> = args[..args.len()-1].iter().map(|a| {
                match a {
                    FormulaExpr::FuncCall { name, args } if args.is_empty() => Ok(name.clone()),
                    _ => Err("LAMBDA: parameter names must be identifiers".into()),
                }
            }).collect();
            Ok((params?, &args[args.len() - 1]))
        }
        _ => Err("expected a LAMBDA expression".into()),
    }
}

/// LAMBDA(...) in expression context just returns a sentinel — use extract_lambda at call sites.
fn func_lambda(_args: &[FormulaExpr], _cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    Ok(Variant::Str("#LAMBDA".into())) // placeholder (not useful as a value by itself)
}

/// Call a LAMBDA with the given argument values.
fn call_lambda(lambda_expr: &FormulaExpr, arg_vals: Vec<Variant>, cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let (params, body) = extract_lambda(lambda_expr)?;
    if params.len() != arg_vals.len() {
        return Err(format!("LAMBDA: expected {} args, got {}", params.len(), arg_vals.len()));
    }
    let frame: HashMap<String, Variant> = params.into_iter().zip(arg_vals).collect();
    push_bindings(frame);
    let result = evaluate(body, cells);
    pop_bindings();
    result
}

// ── MAP ───────────────────────────────────────────────────────────────────────

fn func_map(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("MAP requires at least 2 arguments".into()); }
    let lambda_expr = &args[args.len() - 1];
    // Single-array case: MAP(array, LAMBDA(x, body))
    // Multi-array case: MAP(a1, a2, ..., LAMBDA(x,y,...,body))
    let arrays: Vec<Vec<Variant>> = (0..args.len()-1)
        .map(|i| collect_values(&args[i], cells))
        .collect::<Result<_, _>>()?;
    let len = arrays.first().map(|a| a.len()).unwrap_or(0);
    if arrays.iter().any(|a| a.len() != len) {
        return Err("MAP: all array arguments must have equal length".into());
    }
    let result: Result<Vec<Variant>, String> = (0..len).map(|i| {
        let vals: Vec<Variant> = arrays.iter().map(|a| a[i].clone()).collect();
        call_lambda(lambda_expr, vals, cells)
    }).collect();
    Ok(wrap_array(result?))
}

// ── REDUCE ────────────────────────────────────────────────────────────────────

fn func_reduce(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 { return Err("REDUCE requires 3 arguments: REDUCE(initial, array, LAMBDA(acc,x,body))".into()); }
    let mut acc = evaluate(&args[0], cells)?;
    let data    = collect_values(&args[1], cells)?;
    let lambda  = &args[2];
    for val in data {
        acc = call_lambda(lambda, vec![acc, val], cells)?;
    }
    Ok(acc)
}

// ── SCAN ──────────────────────────────────────────────────────────────────────

fn func_scan(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 { return Err("SCAN requires 3 arguments: SCAN(initial, array, LAMBDA(acc,x,body))".into()); }
    let mut acc = evaluate(&args[0], cells)?;
    let data    = collect_values(&args[1], cells)?;
    let lambda  = &args[2];
    let mut result = vec![];
    for val in data {
        acc = call_lambda(lambda, vec![acc.clone(), val], cells)?;
        result.push(acc.clone());
    }
    Ok(wrap_array(result))
}

// ── BYROW / BYCOL ─────────────────────────────────────────────────────────────

fn func_byrow(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("BYROW requires 2 arguments".into()); }
    let lambda = &args[1];
    match &args[0] {
        FormulaExpr::Range { c1, r1, c2, r2 } => {
            let result: Result<Vec<Variant>, String> = (*r1..=*r2).map(|row| {
                let row_vals: Vec<Variant> = (*c1..=*c2).map(|col| cell_val(cells, row, col)).collect();
                call_lambda(lambda, row_vals, cells)
            }).collect();
            Ok(wrap_array(result?))
        }
        _ => {
            let val = evaluate(&args[0], cells)?;
            call_lambda(lambda, vec![val], cells)
        }
    }
}

fn func_bycol(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("BYCOL requires 2 arguments".into()); }
    let lambda = &args[1];
    match &args[0] {
        FormulaExpr::Range { c1, r1, c2, r2 } => {
            let result: Result<Vec<Variant>, String> = (*c1..=*c2).map(|col| {
                let col_vals: Vec<Variant> = (*r1..=*r2).map(|row| cell_val(cells, row, col)).collect();
                call_lambda(lambda, col_vals, cells)
            }).collect();
            Ok(wrap_array(result?))
        }
        _ => {
            let val = evaluate(&args[0], cells)?;
            call_lambda(lambda, vec![val], cells)
        }
    }
}

// ── INDIRECT ──────────────────────────────────────────────────────────────────

fn func_indirect(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("INDIRECT requires 1 argument".into()); }
    let addr_str = match evaluate(&args[0], cells)? {
        Variant::Str(s) => s,
        other => return Err(format!("INDIRECT: expected string, got {}", other)),
    };
    // Resolve through vm's public parse_cell_addr / parse_range_addr
    let ((r1, c1), _) = crate::vm::parse_range_addr(addr_str.trim())
        .ok_or_else(|| format!("INDIRECT: invalid reference '{}'", addr_str))?;
    Ok(cells.get(&(r1, c1)).map(|c| c.value.clone()).unwrap_or(Variant::Empty))
}

// ── OFFSET ────────────────────────────────────────────────────────────────────

fn func_offset(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 { return Err("OFFSET requires at least 3 arguments".into()); }
    // First arg must be a cell/range reference expression — read coords without evaluating
    let (base_row, base_col): (u32, u32) = match args.first() {
        Some(FormulaExpr::CellRef { row, col }) => (*row, *col),
        Some(FormulaExpr::Range   { r1, c1, .. }) => (*r1, *c1),
        _ => return Err("OFFSET: first argument must be a cell reference".into()),
    };
    let row_off = to_float(&evaluate(&args[1], cells)?)? as i64;
    let col_off = to_float(&evaluate(&args[2], cells)?)? as i64;
    // height / width (args[3], args[4]): if > 1 would mean a range result; return top-left only
    let new_row_i = base_row as i64 + row_off;
    let new_col_i = base_col as i64 + col_off;
    if new_row_i < 1 || new_col_i < 1 {
        return Ok(Variant::Error(ExcelError::Ref));
    }
    let new_row = new_row_i as u32;
    let new_col = new_col_i as u32;
    Ok(cells.get(&(new_row, new_col)).map(|c| c.value.clone()).unwrap_or(Variant::Empty))
}

// ── Array / spill helpers ─────────────────────────────────────────────────────

/// Element-wise boolean comparison op (for array FILTER conditions).
fn compare_element(op: &BinOpKind, l: &Variant, r: &Variant) -> bool {
    match op {
        BinOpKind::Eq => variant_eq(l, r),
        BinOpKind::Ne => !variant_eq(l, r),
        BinOpKind::Lt => variant_cmp(l, r).map(|o| o == Ordering::Less).unwrap_or(false),
        BinOpKind::Le => variant_cmp(l, r).map(|o| o != Ordering::Greater).unwrap_or(false),
        BinOpKind::Gt => variant_cmp(l, r).map(|o| o == Ordering::Greater).unwrap_or(false),
        BinOpKind::Ge => variant_cmp(l, r).map(|o| o != Ordering::Less).unwrap_or(false),
        _ => is_truthy(l),
    }
}

/// Evaluate a formula expression to a Vec<bool>:
/// - Range → truthy check on each cell
/// - BinOp with a Range lhs → element-wise comparison
/// - Scalar → single-element vec
fn eval_as_bool_array(
    expr: &FormulaExpr,
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Vec<bool>, String> {
    match expr {
        FormulaExpr::Range { .. } => {
            Ok(collect_values(expr, cells)?.iter().map(|v| is_truthy(v)).collect())
        }
        FormulaExpr::BinOp { op, lhs, rhs } => {
            let lhs_vals = collect_values(lhs, cells)?;
            if lhs_vals.len() > 1 {
                let rhs_val = evaluate(rhs, cells)?;
                Ok(lhs_vals.iter().map(|l| compare_element(op, l, &rhs_val)).collect())
            } else {
                Ok(vec![is_truthy(&evaluate(expr, cells)?)])
            }
        }
        _ => Ok(vec![is_truthy(&evaluate(expr, cells)?)]),
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

fn func_filter(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("FILTER requires at least 2 arguments".into()); }
    let include = eval_as_bool_array(&args[1], cells)?;

    match &args[0] {
        FormulaExpr::Range { c1, r1, c2, r2 } => {
            let data_rows = (*r2 - *r1 + 1) as usize;
            if include.len() != data_rows {
                return Err(format!(
                    "FILTER: data has {} rows but include has {} elements",
                    data_rows, include.len()
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
                return if args.len() >= 3 { evaluate(&args[2], cells) }
                       else { Ok(Variant::Error(ExcelError::NA)) };
            }
            Ok(wrap_array(result))
        }
        _ => evaluate(&args[0], cells),
    }
}

// ── UNIQUE ────────────────────────────────────────────────────────────────────

fn func_unique(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("UNIQUE requires 1 argument".into()); }
    let vals = collect_values(&args[0], cells)?;
    let mut result: Vec<Variant> = vec![];
    for v in vals {
        if !result.contains(&v) { result.push(v); }
    }
    Ok(wrap_array(result))
}

// ── SORT ──────────────────────────────────────────────────────────────────────

fn func_sort(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("SORT requires 1 argument".into()); }
    let mut vals = collect_values(&args[0], cells)?;
    let order: i64 = if args.len() >= 3 { to_float(&evaluate(&args[2], cells)?)? as i64 } else { 1 };
    vals.sort_by(|a, b| {
        let af = to_float(a).unwrap_or(f64::INFINITY);
        let bf = to_float(b).unwrap_or(f64::INFINITY);
        let o = af.partial_cmp(&bf).unwrap_or(Ordering::Equal);
        if order < 0 { o.reverse() } else { o }
    });
    Ok(wrap_array(vals))
}

// ── SORTBY ────────────────────────────────────────────────────────────────────

fn func_sortby(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("SORTBY requires at least 2 arguments".into()); }
    let data    = collect_values(&args[0], cells)?;
    let by_vals = collect_values(&args[1], cells)?;
    let order: i64 = if args.len() >= 3 { to_float(&evaluate(&args[2], cells)?)? as i64 } else { 1 };
    if data.len() != by_vals.len() {
        return Err("SORTBY: data and sort-by arrays must have equal length".into());
    }
    let mut indexed: Vec<usize> = (0..data.len()).collect();
    indexed.sort_by(|&a, &b| {
        let af = to_float(&by_vals[a]).unwrap_or(f64::INFINITY);
        let bf = to_float(&by_vals[b]).unwrap_or(f64::INFINITY);
        let o = af.partial_cmp(&bf).unwrap_or(Ordering::Equal);
        if order < 0 { o.reverse() } else { o }
    });
    let result: Vec<Variant> = indexed.iter().map(|&i| data[i].clone()).collect();
    Ok(wrap_array(result))
}

// ── SEQUENCE ──────────────────────────────────────────────────────────────────

fn func_sequence(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("SEQUENCE requires at least 1 argument".into()); }
    let rows  = to_float(&evaluate(&args[0], cells)?)? as i64;
    let cols  = if args.len() >= 2 { to_float(&evaluate(&args[1], cells)?)? as i64 } else { 1 };
    let start = if args.len() >= 3 { to_float(&evaluate(&args[2], cells)?)? } else { 1.0 };
    let step  = if args.len() >= 4 { to_float(&evaluate(&args[3], cells)?)? } else { 1.0 };
    let count = (rows * cols).max(0) as usize;
    let result: Vec<Variant> = (0..count)
        .map(|i| as_integer_if_whole(start + i as f64 * step))
        .collect();
    Ok(wrap_array(result))
}

// ── TRANSPOSE ─────────────────────────────────────────────────────────────────

fn func_transpose(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("TRANSPOSE requires 1 argument".into()); }
    match &args[0] {
        FormulaExpr::Range { c1, r1, c2, r2 } => {
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
            if s == 0 { s = 1; }
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

fn func_randarray(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let rows        = if args.len() >= 1 { to_float(&evaluate(&args[0], cells)?)? as usize } else { 1 };
    let cols        = if args.len() >= 2 { to_float(&evaluate(&args[1], cells)?)? as usize } else { 1 };
    let min         = if args.len() >= 3 { to_float(&evaluate(&args[2], cells)?)? } else { 0.0 };
    let max         = if args.len() >= 4 { to_float(&evaluate(&args[3], cells)?)? } else { 1.0 };
    let whole       = if args.len() >= 5 { is_truthy(&evaluate(&args[4], cells)?) } else { false };
    if max < min { return Err("RANDARRAY: max must be >= min".into()); }
    let n = rows.max(1) * cols.max(1);
    let result: Vec<Variant> = (0..n).map(|_| {
        let v = min + next_rand_f64() * (max - min);
        if whole { as_integer_if_whole(v.floor()) } else { Variant::Float(v) }
    }).collect();
    Ok(wrap_array(result))
}

// ── WORKDAY ──────────────────────────────────────────────────────────────────

fn func_workday(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("WORKDAY requires 2 or 3 arguments".into()); }
    let start = to_float(&evaluate(&args[0], cells)?)? as i64;
    let days  = to_float(&evaluate(&args[1], cells)?)? as i64;
    let mask  = [false, false, false, false, false, true, true]; // Sat+Sun
    let holidays: std::collections::HashSet<i64> = if args.len() == 3 {
        collect_values(&args[2], cells)?.iter()
            .filter_map(|v| to_float(v).ok().map(|f| f as i64))
            .collect()
    } else { std::collections::HashSet::new() };
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

fn func_pmt(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 { return Err("PMT requires 3 to 5 arguments".into()); }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let nper = to_float(&evaluate(&args[1], cells)?)?;
    let pv   = to_float(&evaluate(&args[2], cells)?)?;
    let fv   = if args.len() >= 4 { to_float(&evaluate(&args[3], cells)?)? } else { 0.0 };
    let typ  = if args.len() >= 5 { to_float(&evaluate(&args[4], cells)?)? } else { 0.0 };
    if nper == 0.0 { return Err("PMT: nper cannot be 0".into()); }
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
    if nper == 0.0 { return f64::NAN; }
    if rate == 0.0 {
        -(pv + fv) / nper
    } else {
        let factor = (1.0 + rate).powf(nper);
        -(rate * (pv * factor + fv)) / ((factor - 1.0) * (1.0 + rate * typ))
    }
}

fn func_fv(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 { return Err("FV requires 3 to 5 arguments".into()); }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let nper = to_float(&evaluate(&args[1], cells)?)?;
    let pmt  = to_float(&evaluate(&args[2], cells)?)?;
    let pv   = if args.len() >= 4 { to_float(&evaluate(&args[3], cells)?)? } else { 0.0 };
    let typ  = if args.len() >= 5 { to_float(&evaluate(&args[4], cells)?)? } else { 0.0 };
    Ok(Variant::Float(annuity_fv(rate, nper, pmt, pv, typ)))
}

fn func_pv(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 { return Err("PV requires 3 to 5 arguments".into()); }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let nper = to_float(&evaluate(&args[1], cells)?)?;
    let pmt  = to_float(&evaluate(&args[2], cells)?)?;
    let fv   = if args.len() >= 4 { to_float(&evaluate(&args[3], cells)?)? } else { 0.0 };
    let typ  = if args.len() >= 5 { to_float(&evaluate(&args[4], cells)?)? } else { 0.0 };
    let result = if rate == 0.0 {
        -(fv + pmt * nper)
    } else {
        let factor = (1.0 + rate).powf(nper);
        -(fv / factor + pmt * (1.0 + rate * typ) * (1.0 - 1.0 / factor) / rate)
    };
    Ok(Variant::Float(result))
}

fn func_nper(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 5 { return Err("NPER requires 3 to 5 arguments".into()); }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let pmt  = to_float(&evaluate(&args[1], cells)?)?;
    let pv   = to_float(&evaluate(&args[2], cells)?)?;
    let fv   = if args.len() >= 4 { to_float(&evaluate(&args[3], cells)?)? } else { 0.0 };
    let typ  = if args.len() >= 5 { to_float(&evaluate(&args[4], cells)?)? } else { 0.0 };
    let result = if rate == 0.0 {
        if pmt == 0.0 { return Ok(Variant::Error(ExcelError::DivZero)); }
        -(pv + fv) / pmt
    } else {
        let z = pmt * (1.0 + rate * typ) / rate;
        let numer = z - fv;
        let denom = z + pv;
        if denom == 0.0 || numer / denom <= 0.0 { return Ok(Variant::Error(ExcelError::Num)); }
        (numer / denom).ln() / (1.0 + rate).ln()
    };
    Ok(Variant::Float(result))
}

fn func_rate(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 3 || args.len() > 6 { return Err("RATE requires 3 to 6 arguments".into()); }
    let nper  = to_float(&evaluate(&args[0], cells)?)?;
    let pmt   = to_float(&evaluate(&args[1], cells)?)?;
    let pv    = to_float(&evaluate(&args[2], cells)?)?;
    let fv    = if args.len() >= 4 { to_float(&evaluate(&args[3], cells)?)? } else { 0.0 };
    let typ   = if args.len() >= 5 { to_float(&evaluate(&args[4], cells)?)? } else { 0.0 };
    let guess = if args.len() >= 6 { to_float(&evaluate(&args[5], cells)?)? } else { 0.1 };
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
        if delta.abs() < 1e-10 { return Ok(Variant::Float(r)); }
    }
    Ok(Variant::Error(ExcelError::Num)) // did not converge
}

fn func_ipmt(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 6 { return Err("IPMT requires 4 to 6 arguments".into()); }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let per  = to_float(&evaluate(&args[1], cells)?)? as i64;
    let nper = to_float(&evaluate(&args[2], cells)?)?;
    let pv   = to_float(&evaluate(&args[3], cells)?)?;
    let fv   = if args.len() >= 5 { to_float(&evaluate(&args[4], cells)?)? } else { 0.0 };
    let typ  = if args.len() >= 6 { to_float(&evaluate(&args[5], cells)?)? } else { 0.0 };
    if per < 1 || per as f64 > nper { return Ok(Variant::Error(ExcelError::Num)); }
    let pmt_val = compute_pmt(rate, nper, pv, fv, typ);
    // Remaining balance after (per-1) periods = annuity_fv(rate, per-1, pmt, pv, typ)
    let balance = annuity_fv(rate, (per - 1) as f64, pmt_val, pv, typ);
    let mut ipmt = balance * rate;
    if typ == 1.0 {
        if per == 1 { return Ok(Variant::Float(0.0)); }
        ipmt /= 1.0 + rate;
    }
    Ok(Variant::Float(ipmt))
}

fn func_ppmt(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 4 || args.len() > 6 { return Err("PPMT requires 4 to 6 arguments".into()); }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let per  = to_float(&evaluate(&args[1], cells)?)? as i64;
    let nper = to_float(&evaluate(&args[2], cells)?)?;
    let pv   = to_float(&evaluate(&args[3], cells)?)?;
    let fv   = if args.len() >= 5 { to_float(&evaluate(&args[4], cells)?)? } else { 0.0 };
    let typ  = if args.len() >= 6 { to_float(&evaluate(&args[5], cells)?)? } else { 0.0 };
    if per < 1 || per as f64 > nper { return Ok(Variant::Error(ExcelError::Num)); }
    let pmt_val = compute_pmt(rate, nper, pv, fv, typ);
    let balance = annuity_fv(rate, (per - 1) as f64, pmt_val, pv, typ);
    let mut ipmt = balance * rate;
    if typ == 1.0 {
        if per == 1 { return Ok(Variant::Float(pmt_val)); }
        ipmt /= 1.0 + rate;
    }
    Ok(Variant::Float(pmt_val - ipmt))
}

fn func_npv(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("NPV requires at least 2 arguments".into()); }
    let rate = to_float(&evaluate(&args[0], cells)?)?;
    let values = collect_all(&args[1..], cells)?;
    let nums: Vec<f64> = values.iter()
        .filter_map(as_f64)
        .collect();
    let result = nums.iter().enumerate()
        .map(|(i, &v)| v / (1.0 + rate).powf((i + 1) as f64))
        .sum::<f64>();
    Ok(Variant::Float(result))
}

fn func_irr(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 1 || args.len() > 2 { return Err("IRR requires 1 to 2 arguments".into()); }
    let values = collect_values(&args[0], cells)?;
    let nums: Vec<f64> = values.iter()
        .filter_map(as_f64)
        .collect();
    if nums.is_empty() { return Ok(Variant::Error(ExcelError::Num)); }
    let guess = if args.len() >= 2 { to_float(&evaluate(&args[1], cells)?)? } else { 0.1 };
    // Newton-Raphson: find r where NPV = sum(v[i]/(1+r)^(i+1)) = 0
    let mut r = guess;
    for _ in 0..100 {
        // IRR: find r where sum(v[i]/(1+r)^i) = 0, i=0,1,...
        let f: f64  = nums.iter().enumerate().map(|(i, &v)| v / (1.0 + r).powf(i as f64)).sum();
        let df: f64 = nums.iter().enumerate().map(|(i, &v)| -(i as f64) * v / (1.0 + r).powf((i + 1) as f64)).sum();
        if df == 0.0 { break; }
        let delta = f / df;
        r -= delta;
        if delta.abs() < 1e-10 { return Ok(Variant::Float(r)); }
    }
    Ok(Variant::Error(ExcelError::Num))
}

fn func_mirr(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 3 { return Err("MIRR requires 3 arguments".into()); }
    let values = collect_values(&args[0], cells)?;
    let nums: Vec<f64> = values.iter()
        .filter_map(as_f64)
        .collect();
    let n = nums.len();
    if n < 2 { return Ok(Variant::Error(ExcelError::Num)); }
    let finance_rate  = to_float(&evaluate(&args[1], cells)?)?;
    let reinvest_rate = to_float(&evaluate(&args[2], cells)?)?;
    // PV of negative cash flows at finance_rate
    let pv_neg: f64 = nums.iter().enumerate()
        .map(|(i, &v)| if v < 0.0 { v / (1.0 + finance_rate).powf(i as f64) } else { 0.0 })
        .sum();
    // FV of positive cash flows at reinvest_rate
    let fv_pos: f64 = nums.iter().enumerate()
        .map(|(i, &v)| if v > 0.0 { v * (1.0 + reinvest_rate).powf((n - 1 - i) as f64) } else { 0.0 })
        .sum();
    if pv_neg == 0.0 || fv_pos == 0.0 { return Ok(Variant::Error(ExcelError::DivZero)); }
    let result = (fv_pos / (-pv_neg)).powf(1.0 / (n - 1) as f64) - 1.0;
    Ok(Variant::Float(result))
}

fn func_xnpv(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 3 { return Err("XNPV requires 3 arguments".into()); }
    let rate   = to_float(&evaluate(&args[0], cells)?)?;
    let values = collect_values(&args[1], cells)?;
    let dates  = collect_values(&args[2], cells)?;
    if values.len() != dates.len() || values.is_empty() { return Ok(Variant::Error(ExcelError::Value)); }
    let d0 = match as_f64(&dates[0]) { Some(d) => d, None => return Ok(Variant::Error(ExcelError::Value)) };
    let mut result = 0.0;
    for (v, d) in values.iter().zip(dates.iter()) {
        let val = match as_f64(v) { Some(x) => x, None => return Ok(Variant::Error(ExcelError::Value)) };
        let date = match as_f64(d) { Some(x) => x, None => return Ok(Variant::Error(ExcelError::Value)) };
        let exp = (date - d0) / 365.0;
        result += val / (1.0 + rate).powf(exp);
    }
    Ok(Variant::Float(result))
}

fn func_xirr(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("XIRR requires 2 to 3 arguments".into()); }
    let values = collect_values(&args[0], cells)?;
    let dates  = collect_values(&args[1], cells)?;
    if values.len() != dates.len() || values.is_empty() { return Ok(Variant::Error(ExcelError::Value)); }
    let guess = if args.len() >= 3 { to_float(&evaluate(&args[2], cells)?)? } else { 0.1 };
    let d0 = match as_f64(&dates[0]) { Some(d) => d, None => return Ok(Variant::Error(ExcelError::Value)) };
    let nums: Vec<f64> = values.iter().filter_map(as_f64).collect();
    let days: Vec<f64> = dates.iter().filter_map(as_f64).map(|d| (d - d0) / 365.0).collect();
    if nums.len() != values.len() || days.len() != dates.len() { return Ok(Variant::Error(ExcelError::Value)); }
    let mut r = guess;
    for _ in 0..100 {
        let f: f64  = nums.iter().zip(days.iter()).map(|(&v, &t)| v / (1.0 + r).powf(t)).sum();
        let df: f64 = nums.iter().zip(days.iter()).map(|(&v, &t)| -t * v / (1.0 + r).powf(t + 1.0)).sum();
        if df == 0.0 { break; }
        let delta = f / df;
        r -= delta;
        if delta.abs() < 1e-10 { return Ok(Variant::Float(r)); }
    }
    Ok(Variant::Error(ExcelError::Num))
}

// ── TEXTSPLIT ─────────────────────────────────────────────────────────────────

fn func_textsplit(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("TEXTSPLIT requires at least 2 arguments".into()); }
    let text  = to_str(&evaluate(&args[0], cells)?);
    let delim = to_str(&evaluate(&args[1], cells)?);
    if delim.is_empty() { return Err("TEXTSPLIT: delimiter cannot be empty".into()); }
    let ignore_empty    = args.len() >= 4 && is_truthy(&evaluate(&args[3], cells)?);
    let case_insensitive = args.len() >= 5 && is_truthy(&evaluate(&args[4], cells)?);

    // Use lowercase copies for searching while preserving original text for output.
    let (search_text, search_delim) = if case_insensitive {
        (text.to_lowercase(), delim.to_lowercase())
    } else {
        (text.clone(), delim.clone())
    };

    let mut result = vec![];
    let mut char_start = 0usize;
    while let Some(rel) = search_text[char_start..].find(&*search_delim) {
        let abs = char_start + rel;
        let piece = &text[char_start..abs];
        if !ignore_empty || !piece.is_empty() {
            result.push(Variant::Str(piece.to_string()));
        }
        char_start = abs + search_delim.len();
    }
    let last = &text[char_start..];
    if !ignore_empty || !last.is_empty() {
        result.push(Variant::Str(last.to_string()));
    }
    Ok(wrap_array(result))
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
    if args.len() < 2 { return Err(format!("{} requires at least 2 arguments", fname)); }
    let text  = to_str(&evaluate(&args[0], cells)?);
    let delim = to_str(&evaluate(&args[1], cells)?);
    if delim.is_empty() { return Err(format!("{}: delimiter cannot be empty", fname)); }
    let instance_num: i64 = if args.len() >= 3 { to_float(&evaluate(&args[2], cells)?)? as i64 } else { 1 };
    let case_insensitive  = args.len() >= 4 && is_truthy(&evaluate(&args[3], cells)?);

    let (search_text, search_delim) = if case_insensitive {
        (text.to_lowercase(), delim.to_lowercase())
    } else {
        (text.clone(), delim.clone())
    };

    let positions = find_all_occurrences(&search_text, &search_delim);

    // Resolve instance_num to an index
    let idx: Option<usize> = if instance_num > 0 {
        let n = (instance_num - 1) as usize;
        if n < positions.len() { Some(n) } else { None }
    } else if instance_num < 0 {
        let n = (-instance_num) as usize;
        if n <= positions.len() { Some(positions.len() - n) } else { None }
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
            } else {
                Ok(Variant::Str(text[byte_pos + delim.len()..].to_string()))
            }
        }
    }
}

fn func_textbefore(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    text_before_after(args, cells, true, "TEXTBEFORE")
}

fn func_textafter(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    text_before_after(args, cells, false, "TEXTAFTER")
}

// ── VALUETOTEXT ───────────────────────────────────────────────────────────────

fn func_valuetotext(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("VALUETOTEXT requires at least 1 argument".into()); }
    let val    = evaluate(&args[0], cells)?;
    let format = if args.len() >= 2 { to_float(&evaluate(&args[1], cells)?)? as i32 } else { 0 };
    let s = match &val {
        Variant::Str(s)     => if format == 1 { format!("\"{}\"", s.replace('"', "\"\"")) } else { s.clone() },
        Variant::Integer(n) => n.to_string(),
        Variant::Float(f)   => format!("{}", f),
        Variant::Boolean(b) => if *b { "TRUE".into() } else { "FALSE".into() },
        Variant::Empty      => String::new(),
        _                   => to_str(&val),
    };
    Ok(Variant::Str(s))
}

// ── TAKE / DROP ───────────────────────────────────────────────────────────────

fn func_take(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("TAKE requires at least 2 arguments".into()); }
    let vals  = flatten_array_vals(collect_values(&args[0], cells)?);
    let n     = to_float(&evaluate(&args[1], cells)?)? as i64;
    let n_abs = n.unsigned_abs() as usize;
    let result: Vec<Variant> = if n >= 0 {
        vals.into_iter().take(n_abs).collect()
    } else {
        let skip = vals.len().saturating_sub(n_abs);
        vals.into_iter().skip(skip).collect()
    };
    Ok(wrap_array(result))
}

fn func_drop(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("DROP requires at least 2 arguments".into()); }
    let vals  = flatten_array_vals(collect_values(&args[0], cells)?);
    let n     = to_float(&evaluate(&args[1], cells)?)? as i64;
    let n_abs = n.unsigned_abs() as usize;
    let result: Vec<Variant> = if n >= 0 {
        vals.into_iter().skip(n_abs).collect()
    } else {
        let keep = vals.len().saturating_sub(n_abs);
        vals.into_iter().take(keep).collect()
    };
    Ok(wrap_array(result))
}

// ── VSTACK / HSTACK ───────────────────────────────────────────────────────────

fn func_vstack(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("VSTACK requires at least 1 argument".into()); }
    let mut result = vec![];
    for arg in args {
        result.extend(flatten_array_vals(collect_values(arg, cells)?));
    }
    Ok(wrap_array(result))
}

fn func_hstack(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("HSTACK requires at least 1 argument".into()); }
    let mut result = vec![];
    for arg in args {
        result.extend(flatten_array_vals(collect_values(arg, cells)?));
    }
    Ok(wrap_array(result))
}

// ── CHOOSECOLS / CHOOSEROWS ───────────────────────────────────────────────────

fn choose_elements(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
    fname: &str,
) -> Result<Variant, String> {
    if args.len() < 2 { return Err(format!("{} requires at least 2 arguments", fname)); }
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    let len  = vals.len() as i64;
    let mut result = vec![];
    for i in 1..args.len() {
        let n = to_float(&evaluate(&args[i], cells)?)? as i64;
        let idx = if n > 0 { n - 1 } else { len + n };
        if idx < 0 || idx >= len {
            result.push(Variant::Error(ExcelError::Value));
        } else {
            result.push(vals[idx as usize].clone());
        }
    }
    Ok(wrap_array(result))
}

fn func_choosecols(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    choose_elements(args, cells, "CHOOSECOLS")
}

fn func_chooserows(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    choose_elements(args, cells, "CHOOSEROWS")
}

// ── COMBIN ───────────────────────────────────────────────────────────────────

fn func_combin(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("COMBIN requires 2 arguments".into()); }
    let n = to_float(&evaluate(&args[0], cells)?)? as i64;
    let k = to_float(&evaluate(&args[1], cells)?)? as i64;
    if n < 0 || k < 0 || k > n { return Ok(Variant::Error(ExcelError::Num)); }
    // Multiplicative formula: ∏ (n-i)/(i+1) for i in 0..k  — avoids factorial overflow
    let mut result = 1f64;
    for i in 0..k {
        result = result * (n - i) as f64 / (i + 1) as f64;
    }
    Ok(as_integer_if_whole(result.round()))
}

// ── DGET ─────────────────────────────────────────────────────────────────────

/// Resolve the `field` argument of a database function to an absolute column number.
/// `field` may be a 1-based column index (Integer/Float) or a header name (Str).
fn resolve_db_field(
    field_val: &Variant,
    cells: &HashMap<(u32, u32), CellContent>,
    c1: u32, c2: u32, header_row: u32,
) -> Result<u32, String> {
    match field_val {
        Variant::Integer(n) => {
            if *n < 1 { return Err("DGET: field index must be >= 1".into()); }
            Ok(c1 + (*n as u32) - 1)
        }
        Variant::Float(f) => {
            let n = *f as i64;
            if n < 1 { return Err("DGET: field index must be >= 1".into()); }
            Ok(c1 + (n as u32) - 1)
        }
        Variant::Str(s) => {
            (c1..=c2).find(|&c| {
                to_str(&cell_val(cells, header_row, c)).eq_ignore_ascii_case(s)
            }).ok_or_else(|| format!("DGET: field '{}' not found in database header", s))
        }
        _ => Err("DGET: field argument must be a column number or field name string".into()),
    }
}

/// Test whether a single database row satisfies the criteria range.
/// Criteria rows (excluding header) are OR-combined; columns within a row are AND-combined.
/// An empty criteria cell is treated as "match all" (wildcard).
fn db_row_matches_criteria(
    cells: &HashMap<(u32, u32), CellContent>,
    data_row: u32,
    db_c1: u32, db_c2: u32, db_header_row: u32,
    cr_c1: u32, cr_c2: u32, cr_r1: u32, cr_r2: u32,
) -> bool {
    // Each criteria row (cr_r1+1 .. cr_r2) is one OR-branch
    for cr_row in (cr_r1 + 1)..=cr_r2 {
        let mut row_match = true;
        for cr_col in cr_c1..=cr_c2 {
            let crit = cell_val(cells, cr_row, cr_col);
            if matches!(crit, Variant::Empty) { continue; } // blank criteria = match all
            // Find the database column this criteria column corresponds to
            let header_name = to_str(&cell_val(cells, cr_r1, cr_col));
            if header_name.is_empty() { continue; }
            // Locate the matching database column
            let db_col = match (db_c1..=db_c2).find(|&c| {
                to_str(&cell_val(cells, db_header_row, c)).eq_ignore_ascii_case(&header_name)
            }) {
                Some(c) => c,
                None => { row_match = false; break; }
            };
            let data_val = cell_val(cells, data_row, db_col);
            if !matches_criteria(&data_val, &crit) { row_match = false; break; }
        }
        if row_match { return true; }
    }
    false
}

fn func_dget(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 3 { return Err("DGET requires 3 arguments".into()); }

    let (db_c1, db_r1, db_c2, db_r2) = require_range(&args[0], "DGET")?;
    let field_val = evaluate(&args[1], cells)?;
    let field_col = resolve_db_field(&field_val, cells, db_c1, db_c2, db_r1)?;
    if field_col > db_c2 { return Ok(Variant::Error(ExcelError::Ref)); }

    let (cr_c1, cr_r1, cr_c2, cr_r2) = require_range(&args[2], "DGET")?;

    let mut matched: Vec<Variant> = vec![];
    for row in (db_r1 + 1)..=db_r2 {
        if db_row_matches_criteria(cells, row, db_c1, db_c2, db_r1, cr_c1, cr_c2, cr_r1, cr_r2) {
            matched.push(cell_val(cells, row, field_col));
        }
    }

    match matched.len() {
        0 => Ok(Variant::Error(ExcelError::Value)), // no match
        1 => Ok(matched.remove(0)),
        _ => Ok(Variant::Error(ExcelError::Num)),   // multiple matches
    }
}

// ── DSUM / DAVERAGE / DCOUNT / DCOUNTA / DMAX / DMIN ─────────────────────────

struct DbCtx { db_r1: u32, db_r2: u32, db_c1: u32, db_c2: u32, field_col: u32, cr_r1: u32, cr_r2: u32, cr_c1: u32, cr_c2: u32 }

fn db_resolve_args(args: &[FormulaExpr], cells: &HashMap<(u32,u32), CellContent>, fname: &str)
    -> Result<Option<DbCtx>, String>
{
    if args.len() != 3 { return Err(format!("{fname} requires 3 arguments")); }
    let (db_c1, db_r1, db_c2, db_r2) = require_range(&args[0], fname)?;
    let field_val = evaluate(&args[1], cells)?;
    let field_col = resolve_db_field(&field_val, cells, db_c1, db_c2, db_r1)?;
    if field_col > db_c2 { return Ok(None); }
    let (cr_c1, cr_r1, cr_c2, cr_r2) = require_range(&args[2], fname)?;
    Ok(Some(DbCtx { db_r1, db_r2, db_c1, db_c2, field_col, cr_r1, cr_r2, cr_c1, cr_c2 }))
}

fn db_matched_vals(ctx: &DbCtx, cells: &HashMap<(u32,u32), CellContent>) -> Vec<Variant> {
    (ctx.db_r1 + 1..=ctx.db_r2)
        .filter(|&row| db_row_matches_criteria(cells, row, ctx.db_c1, ctx.db_c2, ctx.db_r1, ctx.cr_c1, ctx.cr_c2, ctx.cr_r1, ctx.cr_r2))
        .map(|row| cell_val(cells, row, ctx.field_col))
        .collect()
}

fn func_dsum(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DSUM")? else { return Ok(Variant::Error(ExcelError::Ref)); };
    let sum: f64 = db_matched_vals(&ctx, cells).iter().filter_map(as_f64).sum();
    Ok(as_integer_if_whole(sum))
}

fn func_daverage(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DAVERAGE")? else { return Ok(Variant::Error(ExcelError::Ref)); };
    let nums: Vec<f64> = db_matched_vals(&ctx, cells).iter().filter_map(as_f64).collect();
    if nums.is_empty() { return Ok(Variant::Error(ExcelError::DivZero)); }
    Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
}

fn func_dcount(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DCOUNT")? else { return Ok(Variant::Error(ExcelError::Ref)); };
    let count = db_matched_vals(&ctx, cells).iter().filter(|v| as_f64(v).is_some()).count();
    Ok(Variant::Integer(count as i64))
}

fn func_dcounta(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DCOUNTA")? else { return Ok(Variant::Error(ExcelError::Ref)); };
    let count = db_matched_vals(&ctx, cells).iter().filter(|v| !matches!(v, Variant::Empty)).count();
    Ok(Variant::Integer(count as i64))
}

fn func_dmax(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DMAX")? else { return Ok(Variant::Error(ExcelError::Ref)); };
    let max = db_matched_vals(&ctx, cells).iter().filter_map(as_f64).reduce(f64::max);
    Ok(as_integer_if_whole(max.unwrap_or(0.0)))
}

fn func_dmin(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let Some(ctx) = db_resolve_args(args, cells, "DMIN")? else { return Ok(Variant::Error(ExcelError::Ref)); };
    let min = db_matched_vals(&ctx, cells).iter().filter_map(as_f64).reduce(f64::min);
    Ok(as_integer_if_whole(min.unwrap_or(0.0)))
}

// ── TOCOL / TOROW ─────────────────────────────────────────────────────────────

/// Flatten any `Variant::Array` items in the collected values so that functions
/// like `WRAPCOLS(SEQUENCE(6), 2)` work correctly even when the first argument
/// returns an Array variant rather than a cell range.
fn flatten_array_vals(vals: Vec<Variant>) -> Vec<Variant> {
    vals.into_iter().flat_map(|v| match v {
        Variant::Array(inner) => inner,
        other => vec![other],
    }).collect()
}

fn ignore_filter(vals: Vec<Variant>, ignore: u8) -> Vec<Variant> {
    vals.into_iter().filter(|v| {
        let skip_blank = ignore == 1 || ignore == 3;
        let skip_error = ignore == 2 || ignore == 3;
        !(skip_blank && matches!(v, Variant::Empty))
            && !(skip_error && matches!(v, Variant::Error(_)))
    }).collect()
}

fn func_tocol(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("TOCOL requires 1 argument".into()); }
    let ignore = if args.len() >= 2 { to_float(&evaluate(&args[1], cells)?)? as u8 } else { 0 };
    // args[2] = scan_by_column (bool); we treat range traversal as row-major regardless
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    Ok(wrap_array(ignore_filter(vals, ignore)))
}

fn func_torow(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.is_empty() { return Err("TOROW requires 1 argument".into()); }
    let ignore = if args.len() >= 2 { to_float(&evaluate(&args[1], cells)?)? as u8 } else { 0 };
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    Ok(wrap_array(ignore_filter(vals, ignore)))
}

// ── WRAPCOLS / WRAPROWS ───────────────────────────────────────────────────────

fn func_wrapcols(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("WRAPCOLS requires 2 arguments".into()); }
    let vals       = flatten_array_vals(collect_values(&args[0], cells)?);
    let wrap_count = to_float(&evaluate(&args[1], cells)?)? as usize;
    if wrap_count == 0 { return Err("WRAPCOLS: wrap_count must be > 0".into()); }
    let pad = if args.len() >= 3 { evaluate(&args[2], cells)? } else { Variant::Empty };
    let n_cols = vals.len().div_ceil(wrap_count);
    // Result is row-major of a (wrap_count × n_cols) 2-D table filled column-by-column:
    //   cell (row, col) = vals[col * wrap_count + row]
    let mut result = Vec::with_capacity(wrap_count * n_cols);
    for row in 0..wrap_count {
        for col in 0..n_cols {
            let idx = col * wrap_count + row;
            result.push(if idx < vals.len() { vals[idx].clone() } else { pad.clone() });
        }
    }
    Ok(wrap_array(result))
}

fn func_wraprows(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 { return Err("WRAPROWS requires 2 arguments".into()); }
    let vals       = flatten_array_vals(collect_values(&args[0], cells)?);
    let wrap_count = to_float(&evaluate(&args[1], cells)?)? as usize;
    if wrap_count == 0 { return Err("WRAPROWS: wrap_count must be > 0".into()); }
    let pad = if args.len() >= 3 { evaluate(&args[2], cells)? } else { Variant::Empty };
    let n_rows = vals.len().div_ceil(wrap_count);
    // Fill row-by-row; pad the last (partial) row if needed.
    let total = n_rows * wrap_count;
    let mut result = Vec::with_capacity(total);
    for i in 0..total {
        result.push(if i < vals.len() { vals[i].clone() } else { pad.clone() });
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
        pairs.iter().map(|(k, v)| (*k, CellContent { formula: None, value: v.clone() })).collect()
    }

    fn calc(formula: &str, cells: &HashMap<(u32, u32), CellContent>) -> Variant {
        evaluate(&fparse(formula).unwrap(), cells).unwrap()
    }

    #[test]
    fn test_arithmetic_ops() {
        let c = HashMap::new();
        assert_eq!(calc("=1+2",     &c), Variant::Integer(3));
        assert_eq!(calc("=10-3",    &c), Variant::Integer(7));
        assert_eq!(calc("=4*5",     &c), Variant::Integer(20));
        assert_eq!(calc("=10/4",    &c), Variant::Float(2.5));
        assert_eq!(calc("=(1+2)*3", &c), Variant::Integer(9));
    }

    #[test]
    fn test_cell_ref() {
        let c = cells_from(&[((1, 1), Variant::Integer(42))]);
        assert_eq!(calc("=A1", &c), Variant::Integer(42));
        assert_eq!(calc("=B1", &c), Variant::Empty);
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
    }

    #[test]
    fn test_average() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(20)),
            ((3, 1), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=AVERAGE(A1:A3)", &c), Variant::Float(20.0));
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
    }

    #[test]
    fn test_count_counta() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)),
            ((2, 1), Variant::Str("hi".into())),
            // (3,1) is absent → Empty by default
        ]);
        assert_eq!(calc("=COUNT(A1:A3)",  &c), Variant::Integer(1));
        assert_eq!(calc("=COUNTA(A1:A3)", &c), Variant::Integer(2));
    }

    #[test]
    fn test_if() {
        let c = cells_from(&[((1, 1), Variant::Integer(5))]);
        assert_eq!(calc("=IF(A1>3,\"yes\",\"no\")",  &c), Variant::Str("yes".into()));
        assert_eq!(calc("=IF(A1>10,\"yes\",\"no\")", &c), Variant::Str("no".into()));
    }

    #[test]
    fn test_and_or_not() {
        let c = HashMap::new();
        assert_eq!(calc("=AND(TRUE,TRUE)",  &c), Variant::Boolean(true));
        assert_eq!(calc("=AND(TRUE,FALSE)", &c), Variant::Boolean(false));
        assert_eq!(calc("=OR(FALSE,TRUE)",  &c), Variant::Boolean(true));
        assert_eq!(calc("=NOT(TRUE)",       &c), Variant::Boolean(false));
    }

    #[test]
    fn test_iferror() {
        let c = HashMap::new();
        assert_eq!(calc("=IFERROR(1/0,99)", &c), Variant::Integer(99));
        assert_eq!(calc("=IFERROR(10,99)",  &c), Variant::Integer(10));
    }

    #[test]
    fn test_string_functions() {
        let c = HashMap::new();
        assert_eq!(calc("=LEFT(\"hello\",3)",            &c), Variant::Str("hel".into()));
        assert_eq!(calc("=RIGHT(\"hello\",2)",           &c), Variant::Str("lo".into()));
        assert_eq!(calc("=MID(\"hello\",2,3)",           &c), Variant::Str("ell".into()));
        assert_eq!(calc("=LEN(\"hello\")",               &c), Variant::Integer(5));
        assert_eq!(calc("=CONCATENATE(\"A\",\"B\",\"C\")", &c), Variant::Str("ABC".into()));
    }

    #[test]
    fn test_concat_operator() {
        let c = HashMap::new();
        assert_eq!(calc("=\"Hello\"&\" \"&\"World\"", &c), Variant::Str("Hello World".into()));
    }

    #[test]
    fn test_text_format() {
        let c = HashMap::new();
        assert_eq!(calc("=TEXT(3.14159,\"0.00\")", &c), Variant::Str("3.14".into()));
        assert_eq!(calc("=TEXT(0.5,\"0%\")",       &c), Variant::Str("50%".into()));
    }

    #[test]
    fn test_vlookup_exact() {
        // (row, col): A1=(1,1), B1=(1,2), A2=(2,1), B2=(2,2), A3=(3,1), B3=(3,2)
        let c = cells_from(&[
            ((1, 1), Variant::Integer(1)), ((1, 2), Variant::Str("one".into())),
            ((2, 1), Variant::Integer(2)), ((2, 2), Variant::Str("two".into())),
            ((3, 1), Variant::Integer(3)), ((3, 2), Variant::Str("three".into())),
        ]);
        assert_eq!(calc("=VLOOKUP(2,A1:B3,2,FALSE)", &c), Variant::Str("two".into()));
    }

    #[test]
    fn test_index() {
        // (row, col): A1=(1,1)=10, B1=(1,2)=20, A2=(2,1)=30, B2=(2,2)=40
        let c = cells_from(&[
            ((1, 1), Variant::Integer(10)), ((1, 2), Variant::Integer(20)),
            ((2, 1), Variant::Integer(30)), ((2, 2), Variant::Integer(40)),
        ]);
        // INDEX(A1:B2, 2, 1) = row 2 col 1 of range = A2 = 30
        assert_eq!(calc("=INDEX(A1:B2,2,1)", &c), Variant::Integer(30));
    }

    #[test]
    fn test_match_exact() {
        let c = cells_from(&[
            ((1, 1), Variant::Integer(10)),
            ((2, 1), Variant::Integer(20)),
            ((3, 1), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=MATCH(20,A1:A3,0)", &c), Variant::Integer(2));
    }

    #[test]
    fn test_lenb() {
        let c = HashMap::new();
        assert_eq!(calc("=LENB(\"ABC\")",  &c), Variant::Integer(3));
        assert_eq!(calc("=LENB(\"日本語\")", &c), Variant::Integer(6));
        assert_eq!(calc("=LENB(\"A日\")",  &c), Variant::Integer(3));
        assert_eq!(calc("=LENB(\"\")",     &c), Variant::Integer(0));
    }

    #[test]
    fn test_leftb() {
        let c = HashMap::new();
        assert_eq!(calc("=LEFTB(\"ABC\",2)",   &c), Variant::Str("AB".into()));
        assert_eq!(calc("=LEFTB(\"日本語\",4)", &c), Variant::Str("日本".into()));
        assert_eq!(calc("=LEFTB(\"日本語\",3)", &c), Variant::Str("日".into()));
        assert_eq!(calc("=LEFTB(\"ABC\",0)",   &c), Variant::Str("".into()));
    }

    #[test]
    fn test_rightb() {
        let c = HashMap::new();
        assert_eq!(calc("=RIGHTB(\"ABC\",2)",   &c), Variant::Str("BC".into()));
        assert_eq!(calc("=RIGHTB(\"日本語\",4)", &c), Variant::Str("本語".into()));
        assert_eq!(calc("=RIGHTB(\"日本語\",3)", &c), Variant::Str("語".into()));
        assert_eq!(calc("=RIGHTB(\"ABC\",0)",   &c), Variant::Str("".into()));
    }

    #[test]
    fn test_midb() {
        let c = HashMap::new();
        assert_eq!(calc("=MIDB(\"日本語\",3,2)", &c), Variant::Str("本".into()));
        assert_eq!(calc("=MIDB(\"ABC\",2,2)",   &c), Variant::Str("BC".into()));
        assert_eq!(calc("=MIDB(\"日本語\",1,4)", &c), Variant::Str("日本".into()));
    }

    #[test]
    fn test_round() {
        let c = HashMap::new();
        assert_eq!(calc("=ROUND(2.5,0)",    &c), Variant::Integer(3));
        assert_eq!(calc("=ROUND(-2.5,0)",   &c), Variant::Integer(-3));
        assert_eq!(calc("=ROUND(2.15,1)",   &c), Variant::Float(2.2));
        assert_eq!(calc("=ROUND(1234,-2)",  &c), Variant::Integer(1200));
        assert_eq!(calc("=ROUND(3.0,0)",    &c), Variant::Integer(3));
    }

    #[test]
    fn test_roundup() {
        let c = HashMap::new();
        assert_eq!(calc("=ROUNDUP(2.1,0)",  &c), Variant::Integer(3));
        assert_eq!(calc("=ROUNDUP(-2.1,0)", &c), Variant::Integer(-3));
        assert_eq!(calc("=ROUNDUP(2.0,0)",  &c), Variant::Integer(2));
        assert_eq!(calc("=ROUNDUP(1.23,1)", &c), Variant::Float(1.3));
    }

    #[test]
    fn test_rounddown() {
        let c = HashMap::new();
        assert_eq!(calc("=ROUNDDOWN(2.9,0)",  &c), Variant::Integer(2));
        assert_eq!(calc("=ROUNDDOWN(-2.9,0)", &c), Variant::Integer(-2));
        assert_eq!(calc("=ROUNDDOWN(1.99,1)", &c), Variant::Float(1.9));
        assert_eq!(calc("=ROUNDDOWN(1234,-2)",&c), Variant::Integer(1200));
    }

    #[test]
    fn test_countif() {
        let c = cells_from(&[
            ((1,1), Variant::Integer(10)),
            ((2,1), Variant::Integer(20)),
            ((3,1), Variant::Integer(10)),
            ((4,1), Variant::Str("apple".into())),
        ]);
        assert_eq!(calc("=COUNTIF(A1:A3,10)",    &c), Variant::Integer(2));
        assert_eq!(calc("=COUNTIF(A1:A3,\">10\")", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNTIF(A1:A4,\"apple\")", &c), Variant::Integer(1));
    }

    #[test]
    fn test_sumif() {
        let c = cells_from(&[
            ((1,1), Variant::Str("a".into())), ((1,2), Variant::Integer(10)),
            ((2,1), Variant::Str("b".into())), ((2,2), Variant::Integer(20)),
            ((3,1), Variant::Str("a".into())), ((3,2), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=SUMIF(A1:A3,\"a\",B1:B3)", &c), Variant::Integer(40));
        assert_eq!(calc("=SUMIF(B1:B3,\">10\")",      &c), Variant::Integer(50));
    }

    #[test]
    fn test_sumifs_countifs() {
        let c = cells_from(&[
            ((1,1), Variant::Str("a".into())), ((1,2), Variant::Integer(10)),
            ((2,1), Variant::Str("b".into())), ((2,2), Variant::Integer(20)),
            ((3,1), Variant::Str("a".into())), ((3,2), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=SUMIFS(B1:B3,A1:A3,\"a\",B1:B3,\">10\")", &c), Variant::Integer(30));
        assert_eq!(calc("=COUNTIFS(A1:A3,\"a\",B1:B3,\">10\")",       &c), Variant::Integer(1));
    }

    #[test]
    fn test_median() {
        let c = HashMap::new();
        assert_eq!(calc("=MEDIAN(1,3,2)",   &c), Variant::Integer(2));
        assert_eq!(calc("=MEDIAN(1,2,3,4)", &c), Variant::Float(2.5));
    }

    #[test]
    fn test_mode_mult() {
        let c = HashMap::new();
        assert_eq!(calc("=MODE.MULT(1,2,2,3)", &c), Variant::Integer(2));
    }

    #[test]
    fn test_product() {
        let c = HashMap::new();
        assert_eq!(calc("=PRODUCT(2,3,4)", &c), Variant::Integer(24));
    }

    #[test]
    fn test_row() {
        let c = HashMap::new();
        assert_eq!(calc("=ROW(A5)",   &c), Variant::Integer(5));
        assert_eq!(calc("=ROW()",     &c), Variant::Integer(1));
        assert_eq!(calc("=ROW(B3:C7)",&c), Variant::Integer(3));
    }

    #[test]
    fn test_date_serial() {
        let c = HashMap::new();
        // Jan 1 1900 = 1
        assert_eq!(calc("=DATE(1900,1,1)", &c), Variant::Date(1));
        // Jan 1 2000 = 36526
        assert_eq!(calc("=DATE(2000,1,1)", &c), Variant::Date(36526));
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
            ((1,1), Variant::Integer(10)),
            ((2,1), Variant::Integer(30)),
            ((3,1), Variant::Integer(20)),
        ]);
        assert_eq!(calc("=RANK(20,A1:A3,0)", &c), Variant::Integer(2));
        assert_eq!(calc("=RANK(20,A1:A3,1)", &c), Variant::Integer(2));
        assert_eq!(calc("=RANK(30,A1:A3,0)", &c), Variant::Integer(1));
    }

    #[test]
    fn test_ifs() {
        let c = HashMap::new();
        assert_eq!(calc("=IFS(FALSE,\"a\",TRUE,\"b\")", &c), Variant::Str("b".into()));
        assert_eq!(calc("=IFS(TRUE,42,FALSE,99)",       &c), Variant::Integer(42));
    }

    #[test]
    fn test_xlookup() {
        let c = cells_from(&[
            ((1,1), Variant::Integer(1)), ((1,2), Variant::Str("one".into())),
            ((2,1), Variant::Integer(2)), ((2,2), Variant::Str("two".into())),
            ((3,1), Variant::Integer(3)), ((3,2), Variant::Str("three".into())),
        ]);
        assert_eq!(calc("=XLOOKUP(2,A1:A3,B1:B3)",         &c), Variant::Str("two".into()));
        assert_eq!(calc("=XLOOKUP(99,A1:A3,B1:B3,\"N/A\")", &c), Variant::Str("N/A".into()));
    }

    #[test]
    fn test_subtotal() {
        let c = cells_from(&[
            ((1,1), Variant::Integer(10)),
            ((2,1), Variant::Integer(20)),
            ((3,1), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=SUBTOTAL(9,A1:A3)",   &c), Variant::Integer(60));
        assert_eq!(calc("=SUBTOTAL(1,A1:A3)",   &c), Variant::Float(20.0));
        assert_eq!(calc("=SUBTOTAL(109,A1:A3)", &c), Variant::Integer(60));
    }

    #[test]
    fn test_concat() {
        let c = HashMap::new();
        assert_eq!(calc("=CONCAT(\"Hello\",\" \",\"World\")", &c), Variant::Str("Hello World".into()));
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
        assert_eq!(calc("=IFERROR(10,99)",  &c), Variant::Integer(10));
    }

    #[test]
    fn test_iserror_div_zero() {
        let c = HashMap::new();
        assert_eq!(calc("=ISERROR(1/0)",  &c), Variant::Boolean(true));
        assert_eq!(calc("=ISERROR(1)",    &c), Variant::Boolean(false));
    }

    #[test]
    fn test_isna_vlookup() {
        let c = cells_from(&[
            ((1,1), Variant::Integer(1)), ((1,2), Variant::Str("one".into())),
        ]);
        assert_eq!(calc("=ISNA(VLOOKUP(99,A1:B1,2,FALSE))", &c), Variant::Boolean(true));
        assert_eq!(calc("=ISNA(VLOOKUP(1,A1:B1,2,FALSE))",  &c), Variant::Boolean(false));
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
        assert_eq!(ExcelError::NA.as_str(),      "#N/A");
        assert_eq!(ExcelError::Value.as_str(),   "#VALUE!");
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
        assert_eq!(calc("=TEXT(DATE(2000,6,15),\"YYYY/MM/DD\")", &c), Variant::Str("2000/06/15".into()));
        assert_eq!(calc("=TEXT(DATE(2000,6,15),\"MM-DD-YYYY\")", &c), Variant::Str("06-15-2000".into()));
    }

    #[test]
    fn test_year_month_day_on_date_variant() {
        let c = HashMap::new();
        assert_eq!(calc("=YEAR(DATE(2000,6,15))",  &c), Variant::Integer(2000));
        assert_eq!(calc("=MONTH(DATE(2000,6,15))", &c), Variant::Integer(6));
        assert_eq!(calc("=DAY(DATE(2000,6,15))",   &c), Variant::Integer(15));
    }

    #[test]
    fn test_match_na_when_not_found() {
        let c = cells_from(&[
            ((1,1), Variant::Integer(1)),
            ((2,1), Variant::Integer(2)),
        ]);
        assert_eq!(calc("=MATCH(99,A1:A2,0)", &c), Variant::Error(ExcelError::NA));
    }

    #[test]
    fn test_averageif() {
        let c = cells_from(&[
            ((1,1), Variant::Str("a".into())), ((1,2), Variant::Integer(10)),
            ((2,1), Variant::Str("b".into())), ((2,2), Variant::Integer(20)),
            ((3,1), Variant::Str("a".into())), ((3,2), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=AVERAGEIF(A1:A3,\"a\",B1:B3)", &c), Variant::Float(20.0));
    }

    #[test]
    fn test_int_trunc_mod() {
        let c = HashMap::new();
        assert_eq!(calc("=INT(3.9)",    &c), Variant::Integer(3));
        assert_eq!(calc("=INT(-3.1)",   &c), Variant::Integer(-4));
        assert_eq!(calc("=TRUNC(3.9)",  &c), Variant::Integer(3));
        assert_eq!(calc("=TRUNC(-3.9)", &c), Variant::Integer(-3));
        assert_eq!(calc("=TRUNC(3.14159,2)", &c), Variant::Float(3.14));
        assert_eq!(calc("=MOD(10,3)",   &c), Variant::Integer(1));
        assert_eq!(calc("=MOD(-10,3)",  &c), Variant::Integer(2));
    }

    #[test]
    fn test_large_small() {
        let c2 = cells_from(&[
            ((1,1), Variant::Integer(5)),
            ((2,1), Variant::Integer(1)),
            ((3,1), Variant::Integer(3)),
        ]);
        assert_eq!(calc("=LARGE(A1:A3,1)", &c2), Variant::Integer(5));
        assert_eq!(calc("=LARGE(A1:A3,2)", &c2), Variant::Integer(3));
        assert_eq!(calc("=SMALL(A1:A3,1)", &c2), Variant::Integer(1));
        assert_eq!(calc("=SMALL(A1:A3,2)", &c2), Variant::Integer(3));
    }

    #[test]
    fn test_sumproduct() {
        let c = cells_from(&[
            ((1,1), Variant::Integer(2)), ((1,2), Variant::Integer(3)),
            ((2,1), Variant::Integer(4)), ((2,2), Variant::Integer(5)),
        ]);
        // SUMPRODUCT(A1:A2, B1:B2) = 2*3 + 4*5 = 26
        assert_eq!(calc("=SUMPRODUCT(A1:A2,B1:B2)", &c), Variant::Integer(26));
    }

    #[test]
    fn test_percentile() {
        let c2 = cells_from(&[
            ((1,1), Variant::Integer(1)),
            ((2,1), Variant::Integer(2)),
            ((3,1), Variant::Integer(3)),
            ((4,1), Variant::Integer(4)),
            ((5,1), Variant::Integer(5)),
        ]);
        assert_eq!(calc("=PERCENTILE(A1:A5,0.5)", &c2), Variant::Integer(3));
        let _ = c2;
    }

    #[test]
    fn test_maxifs_minifs() {
        let c = cells_from(&[
            ((1,1), Variant::Str("a".into())), ((1,2), Variant::Integer(10)),
            ((2,1), Variant::Str("b".into())), ((2,2), Variant::Integer(20)),
            ((3,1), Variant::Str("a".into())), ((3,2), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=MAXIFS(B1:B3,A1:A3,\"a\")", &c), Variant::Integer(30));
        assert_eq!(calc("=MINIFS(B1:B3,A1:A3,\"a\")", &c), Variant::Integer(10));
    }

    #[test]
    fn test_rand_randbetween() {
        let c = HashMap::new();
        if let Variant::Float(v) = calc("=RAND()", &c) {
            assert!((0.0..1.0).contains(&v));
        } else { panic!("RAND should return Float"); }
        if let Variant::Integer(v) = calc("=RANDBETWEEN(1,10)", &c) {
            assert!((1..=10).contains(&v));
        } else { panic!("RANDBETWEEN should return Integer"); }
    }

    #[test]
    fn test_string_upper_lower_proper_trim() {
        let c = HashMap::new();
        assert_eq!(calc("=UPPER(\"hello\")", &c), Variant::Str("HELLO".into()));
        assert_eq!(calc("=LOWER(\"HELLO\")", &c), Variant::Str("hello".into()));
        assert_eq!(calc("=PROPER(\"hello world\")", &c), Variant::Str("Hello World".into()));
        assert_eq!(calc("=TRIM(\"  hello   world  \")", &c), Variant::Str("hello world".into()));
    }

    #[test]
    fn test_substitute_replace() {
        let c = HashMap::new();
        assert_eq!(calc("=SUBSTITUTE(\"aabbaa\",\"a\",\"x\")",    &c), Variant::Str("xxbbxx".into()));
        assert_eq!(calc("=SUBSTITUTE(\"aabbaa\",\"a\",\"x\",2)",  &c), Variant::Str("axbbaa".into()));
        assert_eq!(calc("=REPLACE(\"Hello\",1,2,\"ZZ\")",          &c), Variant::Str("ZZllo".into()));
    }

    #[test]
    fn test_find_search() {
        let c = HashMap::new();
        assert_eq!(calc("=FIND(\"lo\",\"Hello\")",        &c), Variant::Integer(4));
        assert_eq!(calc("=SEARCH(\"LO\",\"Hello\")",      &c), Variant::Integer(4));
        assert_eq!(calc("=SEARCH(\"h*o\",\"Hello\")",     &c), Variant::Integer(1));
    }

    #[test]
    fn test_textjoin() {
        let c = HashMap::new();
        assert_eq!(calc("=TEXTJOIN(\"-\",TRUE,\"a\",\"\",\"b\")", &c), Variant::Str("a-b".into()));
        assert_eq!(calc("=TEXTJOIN(\"-\",FALSE,\"a\",\"\",\"b\")", &c), Variant::Str("a--b".into()));
    }

    #[test]
    fn test_char_code() {
        let c = HashMap::new();
        assert_eq!(calc("=CHAR(65)",      &c), Variant::Str("A".into()));
        assert_eq!(calc("=CODE(\"A\")",   &c), Variant::Integer(65));
        assert_eq!(calc("=UNICHAR(9786)", &c), Variant::Str("☺".into()));
        assert_eq!(calc("=UNICODE(\"☺\")", &c), Variant::Integer(9786));
    }

    #[test]
    fn test_exact() {
        let c = HashMap::new();
        assert_eq!(calc("=EXACT(\"Hello\",\"Hello\")", &c), Variant::Boolean(true));
        assert_eq!(calc("=EXACT(\"Hello\",\"hello\")", &c), Variant::Boolean(false));
    }

    #[test]
    fn test_value() {
        let c = HashMap::new();
        assert_eq!(calc("=VALUE(\"42\")",   &c), Variant::Integer(42));
        assert_eq!(calc("=VALUE(\"3.14\")", &c), Variant::Float(3.14));
    }

    #[test]
    fn test_asc_jis() {
        let c = HashMap::new();
        // Full-width A (U+FF21) → half-width A
        assert_eq!(calc("=ASC(\"Ａ\")", &c), Variant::Str("A".into()));
        // Half-width A → full-width A
        assert_eq!(calc("=JIS(\"A\")",  &c), Variant::Str("Ａ".into()));
    }

    #[test]
    fn test_year_month_day_weekday() {
        let c = HashMap::new();
        // DATE(2000,6,15) = June 15 2000
        assert_eq!(calc("=YEAR(DATE(2000,6,15))",    &c), Variant::Integer(2000));
        assert_eq!(calc("=MONTH(DATE(2000,6,15))",   &c), Variant::Integer(6));
        assert_eq!(calc("=DAY(DATE(2000,6,15))",     &c), Variant::Integer(15));
        // June 15 2000 was a Thursday; WEEKDAY type=2: Mon=1,...Thu=4
        assert_eq!(calc("=WEEKDAY(DATE(2000,6,15),2)", &c), Variant::Integer(4));
    }

    #[test]
    fn test_days_edate() {
        let c = HashMap::new();
        assert_eq!(calc("=DAYS(DATE(2000,2,1),DATE(2000,1,1))", &c), Variant::Integer(31));
        // EDATE(Jan 31 2000, 1) = Feb 29 2000 (leap year, clamped)
        assert_eq!(calc("=EDATE(DATE(2000,1,31),1)", &c), Variant::Date(date_to_serial(2000,2,29)));
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
        if let Variant::Float(v) = calc("=TIMEVALUE(\"12:00:00\")", &c) {
            assert!((v - 0.5).abs() < 1e-9);
        } else { panic!("TIMEVALUE should return Float"); }
    }

    #[test]
    fn test_time_hour_minute_second() {
        let c = HashMap::new();
        if let Variant::Float(v) = calc("=TIME(12,30,0)", &c) {
            assert!((v - 0.520833).abs() < 1e-5);
        }
        assert_eq!(calc("=HOUR(TIME(12,30,45))",   &c), Variant::Integer(12));
        assert_eq!(calc("=MINUTE(TIME(12,30,45))", &c), Variant::Integer(30));
        assert_eq!(calc("=SECOND(TIME(12,30,45))", &c), Variant::Integer(45));
    }

    #[test]
    fn test_networkdays_intl() {
        let c = HashMap::new();
        // Mon Jan 3 2000 to Fri Jan 7 2000, default weekend (Sat+Sun) = 5
        assert_eq!(calc("=NETWORKDAYS.INTL(36528,36532,1)", &c), Variant::Integer(5));
        // same range, weekend = Mon only (code 12): 4 days (Tue-Fri)
        assert_eq!(calc("=NETWORKDAYS.INTL(36528,36532,12)", &c), Variant::Integer(4));
    }

    #[test]
    fn test_workday_intl() {
        let c = HashMap::new();
        // 5 workdays after Mon Jan 3 2000 (serial 36528) = Fri Jan 7 2000 (36532)? no, it's Mon Jan 10 (36535)
        // Actually: Jan 3+1=Tue4, +2=Wed5, +3=Thu6, +4=Fri7, +5=Mon10 = 36535
        assert_eq!(calc("=WORKDAY.INTL(36528,5,1)", &c), Variant::Date(36535));
    }

    #[test]
    fn test_switch() {
        let c = HashMap::new();
        assert_eq!(calc("=SWITCH(2,1,\"one\",2,\"two\",3,\"three\")", &c), Variant::Str("two".into()));
        assert_eq!(calc("=SWITCH(99,1,\"one\",\"default\")",           &c), Variant::Str("default".into()));
    }

    #[test]
    fn test_xor() {
        let c = HashMap::new();
        assert_eq!(calc("=XOR(TRUE,FALSE)", &c), Variant::Boolean(true));
        assert_eq!(calc("=XOR(TRUE,TRUE)",  &c), Variant::Boolean(false));
        assert_eq!(calc("=XOR(TRUE,TRUE,TRUE)", &c), Variant::Boolean(true));
    }

    #[test]
    fn test_choose_column() {
        let c = HashMap::new();
        assert_eq!(calc("=CHOOSE(2,\"a\",\"b\",\"c\")", &c), Variant::Str("b".into()));
        assert_eq!(calc("=COLUMN(C1)", &c), Variant::Integer(3));
        assert_eq!(calc("=COLUMN()",   &c), Variant::Integer(1));
    }

    #[test]
    fn test_lookup_xmatch() {
        let c = cells_from(&[
            ((1,1), Variant::Integer(1)),
            ((2,1), Variant::Integer(2)),
            ((3,1), Variant::Integer(3)),
            ((1,2), Variant::Str("one".into())),
            ((2,2), Variant::Str("two".into())),
            ((3,2), Variant::Str("three".into())),
        ]);
        assert_eq!(calc("=LOOKUP(2,A1:A3,B1:B3)", &c), Variant::Str("two".into()));
        assert_eq!(calc("=XMATCH(2,A1:A3,0)",     &c), Variant::Integer(2));
        assert_eq!(calc("=XMATCH(2,A1:A3)",        &c), Variant::Integer(2));
    }

    #[test]
    fn test_is_functions() {
        let c = HashMap::new();
        assert_eq!(calc("=ISBLANK(\"\")",  &c), Variant::Boolean(false));
        assert_eq!(calc("=ISERROR(1/0)",   &c), Variant::Boolean(true));
        assert_eq!(calc("=ISERROR(1)",     &c), Variant::Boolean(false));
        assert_eq!(calc("=ISNUMBER(42)",   &c), Variant::Boolean(true));
        assert_eq!(calc("=ISNUMBER(\"x\")", &c), Variant::Boolean(false));
        assert_eq!(calc("=ISTEXT(\"a\")",  &c), Variant::Boolean(true));
        assert_eq!(calc("=ISLOGICAL(TRUE)",&c), Variant::Boolean(true));
        assert_eq!(calc("=ISNONTEXT(42)",  &c), Variant::Boolean(true));
        assert_eq!(calc("=ISNA(1)",        &c), Variant::Boolean(false));
    }

    #[test]
    fn test_aggregate() {
        let c = cells_from(&[
            ((1,1), Variant::Integer(10)),
            ((2,1), Variant::Integer(20)),
            ((3,1), Variant::Integer(30)),
        ]);
        assert_eq!(calc("=AGGREGATE(9,0,A1:A3)",  &c), Variant::Integer(60));
        assert_eq!(calc("=AGGREGATE(1,0,A1:A3)",  &c), Variant::Float(20.0));
        assert_eq!(calc("=AGGREGATE(4,0,A1:A3)",  &c), Variant::Integer(30));
        assert_eq!(calc("=AGGREGATE(12,0,A1:A3)", &c), Variant::Integer(20));
    }

    // ── Phase 10: numerical functions ─────────────────────────────────────────

    fn approx(v: Variant, expected: f64) {
        let f = match v { Variant::Float(f) => f, Variant::Integer(n) => n as f64, _ => panic!("not numeric") };
        assert!((f - expected).abs() < 1e-9, "expected {}, got {}", expected, f);
    }

    #[test]
    fn test_stdev_var() {
        let c = HashMap::new();
        // [2,4,4,4,5,5,7,9]: mean=5, sum_sq_dev=32
        // sample stdev = sqrt(32/7) ≈ 2.138
        approx(calc("=STDEV(2,4,4,4,5,5,7,9)", &c),   (32.0f64/7.0).sqrt());
        approx(calc("=STDEV.S(2,4,4,4,5,5,7,9)", &c), (32.0f64/7.0).sqrt());
        // population stdev = sqrt(32/8) = 2.0
        approx(calc("=STDEVP(2,4,4,4,5,5,7,9)", &c),  2.0);
        approx(calc("=STDEV.P(2,4,4,4,5,5,7,9)", &c), 2.0);
        // VAR: sample=32/7, population=32/8=4
        approx(calc("=VAR(2,4,4,4,5,5,7,9)", &c),  32.0/7.0);
        approx(calc("=VARP(2,4,4,4,5,5,7,9)", &c), 4.0);
    }

    #[test]
    fn test_floor_ceiling_mround() {
        let c = HashMap::new();
        assert_eq!(calc("=FLOOR(3.7,1)",    &c), Variant::Integer(3));
        approx(calc("=FLOOR(3.7,0.5)",      &c), 3.5);
        assert_eq!(calc("=CEILING(3.2,1)",  &c), Variant::Integer(4));
        approx(calc("=CEILING(3.2,0.5)",    &c), 3.5);
        assert_eq!(calc("=MROUND(10,3)",    &c), Variant::Integer(9));
        assert_eq!(calc("=MROUND(11,3)",    &c), Variant::Integer(12));
        approx(calc("=MROUND(1.4,0.5)",     &c), 1.5);
    }

    #[test]
    fn test_math_functions() {
        let c = HashMap::new();
        assert_eq!(calc("=ABS(-5)",     &c), Variant::Integer(5));
        approx(calc("=ABS(3.14)",       &c), 3.14);
        assert_eq!(calc("=SQRT(9)",     &c), Variant::Integer(3));
        approx(calc("=SQRT(2)",         &c), 2f64.sqrt());
        assert_eq!(calc("=POWER(2,10)", &c), Variant::Integer(1024));
        approx(calc("=EXP(0)",          &c), 1.0);
        approx(calc("=EXP(1)",          &c), std::f64::consts::E);
        approx(calc("=LOG(100)",        &c), 2.0);
        approx(calc("=LOG(8,2)",        &c), 3.0);
        approx(calc("=LOG10(1000)",     &c), 3.0);
        approx(calc("=LN(1)",           &c), 0.0);
    }

    #[test]
    fn test_trig_functions() {
        let c = HashMap::new();
        let pi = std::f64::consts::PI;
        approx(calc("=PI()",         &c), pi);
        approx(calc("=SIN(PI()/2)",  &c), 1.0);
        approx(calc("=COS(0)",       &c), 1.0);
        approx(calc("=TAN(0)",       &c), 0.0);
        approx(calc("=DEGREES(PI())", &c), 180.0);
        approx(calc("=RADIANS(180)", &c), pi);
        approx(calc("=ATAN2(1,1)",   &c), pi / 4.0);
        approx(calc("=ASIN(1)",      &c), pi / 2.0);
        approx(calc("=ACOS(1)",      &c), 0.0);
        approx(calc("=ATAN(1)",      &c), pi / 4.0);
    }

    #[test]
    fn test_countblank() {
        let c = cells_from(&[
            ((1,1), Variant::Integer(1)),
            ((2,1), Variant::Empty),
            ((3,1), Variant::Str("".into())),
        ]);
        assert_eq!(calc("=COUNTBLANK(A1:A3)", &c), Variant::Integer(2));
    }

    #[test]
    fn test_address() {
        let c = HashMap::new();
        assert_eq!(calc("=ADDRESS(1,1)",   &c), Variant::Str("$A$1".into()));
        assert_eq!(calc("=ADDRESS(2,3,4)", &c), Variant::Str("C2".into()));
        assert_eq!(calc("=ADDRESS(1,27)",  &c), Variant::Str("$AA$1".into()));
    }

    #[test]
    fn test_indirect() {
        let mut c = HashMap::new();
        c.insert((1, 1), CellContent { formula: None, value: Variant::Integer(42) });
        c.insert((3, 2), CellContent { formula: None, value: Variant::Str("hello".into()) });
        // INDIRECT("A1") → value at A1
        assert_eq!(calc("=INDIRECT(\"A1\")", &c), Variant::Integer(42));
        // INDIRECT("B3") → value at B3
        assert_eq!(calc("=INDIRECT(\"B3\")", &c), Variant::Str("hello".into()));
        // INDIRECT of empty cell → Empty
        assert_eq!(calc("=INDIRECT(\"C5\")", &c), Variant::Empty);
        // INDIRECT with range reference → top-left cell
        assert_eq!(calc("=INDIRECT(\"A1:B3\")", &c), Variant::Integer(42));
    }

    #[test]
    #[test]
    fn test_spill_sort() {
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(3) });
        c.insert((2,1), CellContent { formula: None, value: Variant::Integer(1) });
        c.insert((3,1), CellContent { formula: None, value: Variant::Integer(2) });
        assert_eq!(calc("=SORT(A1:A3)", &c),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(2), Variant::Integer(3)]));
        assert_eq!(calc("=SORT(A1:A3,1,-1)", &c),
            Variant::Array(vec![Variant::Integer(3), Variant::Integer(2), Variant::Integer(1)]));
    }

    #[test]
    fn test_spill_unique() {
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(2) });
        c.insert((2,1), CellContent { formula: None, value: Variant::Integer(1) });
        c.insert((3,1), CellContent { formula: None, value: Variant::Integer(2) });
        c.insert((4,1), CellContent { formula: None, value: Variant::Integer(3) });
        assert_eq!(calc("=UNIQUE(A1:A4)", &c),
            Variant::Array(vec![Variant::Integer(2), Variant::Integer(1), Variant::Integer(3)]));
    }

    #[test]
    fn test_spill_sequence() {
        let c = HashMap::new();
        assert_eq!(calc("=SEQUENCE(4)", &c),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(2),
                                Variant::Integer(3), Variant::Integer(4)]));
        assert_eq!(calc("=SEQUENCE(3,1,5,2)", &c),
            Variant::Array(vec![Variant::Integer(5), Variant::Integer(7), Variant::Integer(9)]));
    }

    #[test]
    fn test_spill_filter_range_condition() {
        // FILTER(A1:A4, B1:B4) — keep rows where B is truthy
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(10) });
        c.insert((2,1), CellContent { formula: None, value: Variant::Integer(20) });
        c.insert((3,1), CellContent { formula: None, value: Variant::Integer(30) });
        c.insert((4,1), CellContent { formula: None, value: Variant::Integer(40) });
        c.insert((1,2), CellContent { formula: None, value: Variant::Boolean(true) });
        c.insert((2,2), CellContent { formula: None, value: Variant::Boolean(false) });
        c.insert((3,2), CellContent { formula: None, value: Variant::Boolean(true) });
        c.insert((4,2), CellContent { formula: None, value: Variant::Boolean(false) });
        assert_eq!(calc("=FILTER(A1:A4, B1:B4)", &c),
            Variant::Array(vec![Variant::Integer(10), Variant::Integer(30)]));
    }

    #[test]
    fn test_spill_filter_inline_comparison() {
        // FILTER(A1:A4, A1:A4>15)
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(10) });
        c.insert((2,1), CellContent { formula: None, value: Variant::Integer(20) });
        c.insert((3,1), CellContent { formula: None, value: Variant::Integer(30) });
        c.insert((4,1), CellContent { formula: None, value: Variant::Integer(5) });
        assert_eq!(calc("=FILTER(A1:A4, A1:A4>15)", &c),
            Variant::Array(vec![Variant::Integer(20), Variant::Integer(30)]));
    }

    #[test]
    fn test_spill_transpose() {
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(1) });
        c.insert((1,2), CellContent { formula: None, value: Variant::Integer(2) });
        c.insert((1,3), CellContent { formula: None, value: Variant::Integer(3) });
        // TRANSPOSE(A1:C1) — 1 row × 3 cols → 3 rows × 1 col (flat = same values)
        assert_eq!(calc("=TRANSPOSE(A1:C1)", &c),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(2), Variant::Integer(3)]));
    }

    #[test]
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
    fn test_map_lambda() {
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(2) });
        c.insert((2,1), CellContent { formula: None, value: Variant::Integer(3) });
        c.insert((3,1), CellContent { formula: None, value: Variant::Integer(4) });
        // MAP(A1:A3, LAMBDA(x, x*2)) → [4, 6, 8]
        assert_eq!(calc("=MAP(A1:A3, LAMBDA(x, x*2))", &c),
            Variant::Array(vec![Variant::Integer(4), Variant::Integer(6), Variant::Integer(8)]));
    }

    #[test]
    fn test_reduce_lambda() {
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(1) });
        c.insert((2,1), CellContent { formula: None, value: Variant::Integer(2) });
        c.insert((3,1), CellContent { formula: None, value: Variant::Integer(3) });
        c.insert((4,1), CellContent { formula: None, value: Variant::Integer(4) });
        // REDUCE(0, A1:A4, LAMBDA(acc, x, acc+x)) → 10
        assert_eq!(calc("=REDUCE(0, A1:A4, LAMBDA(acc, x, acc+x))", &c), Variant::Integer(10));
        // REDUCE(1, A1:A4, LAMBDA(acc, x, acc*x)) → 24
        assert_eq!(calc("=REDUCE(1, A1:A4, LAMBDA(acc, x, acc*x))", &c), Variant::Integer(24));
    }

    #[test]
    fn test_scan_lambda() {
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(1) });
        c.insert((2,1), CellContent { formula: None, value: Variant::Integer(2) });
        c.insert((3,1), CellContent { formula: None, value: Variant::Integer(3) });
        // SCAN(0, A1:A3, LAMBDA(acc, x, acc+x)) → [1, 3, 6] (cumulative sum)
        assert_eq!(calc("=SCAN(0, A1:A3, LAMBDA(acc, x, acc+x))", &c),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(3), Variant::Integer(6)]));
    }

    #[test]
    fn test_let_with_map() {
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(10) });
        c.insert((2,1), CellContent { formula: None, value: Variant::Integer(20) });
        // LET(factor, 3, MAP(A1:A2, LAMBDA(x, x*factor))) → [30, 60]
        assert_eq!(calc("=LET(factor, 3, MAP(A1:A2, LAMBDA(x, x*factor)))", &c),
            Variant::Array(vec![Variant::Integer(30), Variant::Integer(60)]));
    }

    #[test]
    fn test_offset_negative_ref_error() {
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(10) });
        // OFFSET(A1, -1, 0) goes to row 0 → #REF!
        assert_eq!(calc("=OFFSET(A1,-1,0)", &c), Variant::Error(ExcelError::Ref));
        // OFFSET(A1, 0, -1) goes to col 0 → #REF!
        assert_eq!(calc("=OFFSET(A1,0,-1)", &c), Variant::Error(ExcelError::Ref));
    }

    #[test]
    fn test_vlookup_col_index_bounds() {
        let mut c = HashMap::new();
        c.insert((1,1), CellContent { formula: None, value: Variant::Integer(1) });
        c.insert((1,2), CellContent { formula: None, value: Variant::Integer(2) });
        c.insert((1,3), CellContent { formula: None, value: Variant::Integer(3) });
        // col_index=0 → #VALUE!
        assert_eq!(calc("=VLOOKUP(1,A1:C1,0,TRUE)", &c), Variant::Error(ExcelError::Value));
        // col_index > width (3) → #REF!
        assert_eq!(calc("=VLOOKUP(1,A1:C1,5,TRUE)", &c), Variant::Error(ExcelError::Ref));
        // col_index within range → normal
        assert_eq!(calc("=VLOOKUP(1,A1:C1,2,TRUE)", &c), Variant::Integer(2));
    }

    #[test]
    fn test_wildcard_match_many_stars() {
        let c = HashMap::new();
        // Many-star pattern must complete without stack overflow or timeout
        assert_eq!(calc("=COUNTIF(A1:A1,\"*****hello\")", &c), Variant::Integer(0));
        // Correct wildcard behavior
        let mut c2 = HashMap::new();
        c2.insert((1,1), CellContent { formula: None, value: Variant::Str("hello world".into()) });
        assert_eq!(calc("=COUNTIF(A1:A1,\"*world\")", &c2), Variant::Integer(1));
        assert_eq!(calc("=COUNTIF(A1:A1,\"*xyz*\")", &c2), Variant::Integer(0));
    }

    #[test]
    fn test_offset() {
        let mut c = HashMap::new();
        c.insert((1, 1), CellContent { formula: None, value: Variant::Integer(10) });
        c.insert((2, 1), CellContent { formula: None, value: Variant::Integer(20) });
        c.insert((1, 2), CellContent { formula: None, value: Variant::Integer(30) });
        c.insert((3, 3), CellContent { formula: None, value: Variant::Str("far".into()) });
        // OFFSET(A1, 1, 0) → A2 = 20
        assert_eq!(calc("=OFFSET(A1, 1, 0)", &c), Variant::Integer(20));
        // OFFSET(A1, 0, 1) → B1 = 30
        assert_eq!(calc("=OFFSET(A1, 0, 1)", &c), Variant::Integer(30));
        // OFFSET(A1, 2, 2) → C3 = "far"
        assert_eq!(calc("=OFFSET(A1, 2, 2)", &c), Variant::Str("far".into()));
        // OFFSET with negative offset: A2 → 1 row up = A1 = 10
        assert_eq!(calc("=OFFSET(A2, -1, 0)", &c), Variant::Integer(10));
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
        assert!((result + 188.71).abs() < 0.01, "PMT ≈ -188.71, got {}", result);
        // rate=0: PMT = -pv/nper
        assert_eq!(calc("=PMT(0,10,1000)", &c), Variant::Float(-100.0));
    }

    #[test]
    fn test_textsplit() {
        let c = HashMap::new();
        assert_eq!(calc("=TEXTSPLIT(\"a,b,c\",\",\")", &c),
            Variant::Array(vec![Variant::Str("a".into()), Variant::Str("b".into()), Variant::Str("c".into())]));
        // ignore_empty
        assert_eq!(calc("=TEXTSPLIT(\"a,,b\",\",\",\"\",TRUE)", &c),
            Variant::Array(vec![Variant::Str("a".into()), Variant::Str("b".into())]));
        // single result → scalar
        assert_eq!(calc("=TEXTSPLIT(\"hello\",\",\")", &c), Variant::Str("hello".into()));
    }

    #[test]
    fn test_textbefore_textafter() {
        let c = HashMap::new();
        assert_eq!(calc("=TEXTBEFORE(\"hello world\",\" \")", &c), Variant::Str("hello".into()));
        assert_eq!(calc("=TEXTAFTER(\"hello world\",\" \")", &c),  Variant::Str("world".into()));
        // instance_num = 2
        assert_eq!(calc("=TEXTBEFORE(\"a-b-c\",\"-\",2)", &c), Variant::Str("a-b".into()));
        assert_eq!(calc("=TEXTAFTER(\"a-b-c\",\"-\",2)", &c),  Variant::Str("c".into()));
        // negative instance_num (from end)
        assert_eq!(calc("=TEXTBEFORE(\"a-b-c\",\"-\",-1)", &c), Variant::Str("a-b".into()));
        // not found → #N/A
        assert_eq!(calc("=TEXTBEFORE(\"hello\",\",\")", &c), Variant::Error(ExcelError::NA));
    }

    #[test]
    fn test_valuetotext() {
        let c = HashMap::new();
        assert_eq!(calc("=VALUETOTEXT(42)",        &c), Variant::Str("42".into()));
        assert_eq!(calc("=VALUETOTEXT(3.14)",      &c), Variant::Str("3.14".into()));
        assert_eq!(calc("=VALUETOTEXT(TRUE)",       &c), Variant::Str("TRUE".into()));
        assert_eq!(calc("=VALUETOTEXT(\"hi\")",     &c), Variant::Str("hi".into()));
        assert_eq!(calc("=VALUETOTEXT(\"hi\",1)",   &c), Variant::Str("\"hi\"".into()));
    }

    #[test]
    fn test_take_drop() {
        let c = HashMap::new();
        // TAKE first 3 of [1,2,3,4,5]
        assert_eq!(calc("=TAKE(SEQUENCE(5),3)", &c),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(2), Variant::Integer(3)]));
        // TAKE last 2
        assert_eq!(calc("=TAKE(SEQUENCE(5),-2)", &c),
            Variant::Array(vec![Variant::Integer(4), Variant::Integer(5)]));
        // DROP first 2
        assert_eq!(calc("=DROP(SEQUENCE(5),2)", &c),
            Variant::Array(vec![Variant::Integer(3), Variant::Integer(4), Variant::Integer(5)]));
        // DROP last 3
        assert_eq!(calc("=DROP(SEQUENCE(5),-3)", &c),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(2)]));
        // TAKE more than available → all
        assert_eq!(calc("=TAKE(SEQUENCE(3),10)", &c),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(2), Variant::Integer(3)]));
    }

    #[test]
    fn test_vstack_hstack() {
        let c = HashMap::new();
        // VSTACK concatenates arrays
        assert_eq!(calc("=VSTACK(SEQUENCE(3),SEQUENCE(2))", &c),
            Variant::Array(vec![
                Variant::Integer(1), Variant::Integer(2), Variant::Integer(3),
                Variant::Integer(1), Variant::Integer(2),
            ]));
        // HSTACK same in 1D model
        assert_eq!(calc("=HSTACK(SEQUENCE(2),SEQUENCE(2))", &c),
            Variant::Array(vec![
                Variant::Integer(1), Variant::Integer(2),
                Variant::Integer(1), Variant::Integer(2),
            ]));
    }

    #[test]
    fn test_choosecols_chooserows() {
        let c = HashMap::new();
        // CHOOSECOLS(SEQUENCE(5), 1, 3, 5) → [1, 3, 5]
        assert_eq!(calc("=CHOOSECOLS(SEQUENCE(5),1,3,5)", &c),
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(3), Variant::Integer(5)]));
        // Negative index: -1 = last element
        assert_eq!(calc("=CHOOSECOLS(SEQUENCE(5),-1)", &c), Variant::Integer(5));
        // CHOOSEROWS (same logic)
        assert_eq!(calc("=CHOOSEROWS(SEQUENCE(4),2,4)", &c),
            Variant::Array(vec![Variant::Integer(2), Variant::Integer(4)]));
        // Out-of-bounds → #VALUE!
        assert_eq!(calc("=CHOOSECOLS(SEQUENCE(3),5)", &c), Variant::Error(ExcelError::Value));
    }

    #[test]
    fn test_combin() {
        let c = HashMap::new();
        assert_eq!(calc("=COMBIN(4,2)",  &c), Variant::Integer(6));
        assert_eq!(calc("=COMBIN(10,3)", &c), Variant::Integer(120));
        assert_eq!(calc("=COMBIN(0,0)",  &c), Variant::Integer(1));
        assert_eq!(calc("=COMBIN(5,0)",  &c), Variant::Integer(1));
        assert_eq!(calc("=COMBIN(5,5)",  &c), Variant::Integer(1));
        // Error cases
        assert_eq!(calc("=COMBIN(3,5)",  &c), Variant::Error(ExcelError::Num));
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
        c.insert((1,1), CellContent { formula: None, value: Variant::Str("Name".into()) });
        c.insert((1,2), CellContent { formula: None, value: Variant::Str("Score".into()) });
        // data rows
        c.insert((2,1), CellContent { formula: None, value: Variant::Str("Alice".into()) });
        c.insert((2,2), CellContent { formula: None, value: Variant::Integer(90) });
        c.insert((3,1), CellContent { formula: None, value: Variant::Str("Bob".into()) });
        c.insert((3,2), CellContent { formula: None, value: Variant::Integer(75) });
        c.insert((4,1), CellContent { formula: None, value: Variant::Str("Carol".into()) });
        c.insert((4,2), CellContent { formula: None, value: Variant::Integer(85) });
        // criteria: Name = Alice
        c.insert((1,4), CellContent { formula: None, value: Variant::Str("Name".into()) });
        c.insert((2,4), CellContent { formula: None, value: Variant::Str("Alice".into()) });

        // DGET by field name
        assert_eq!(calc("=DGET(A1:B4,\"Score\",D1:D2)", &c), Variant::Integer(90));
        // DGET by field index
        assert_eq!(calc("=DGET(A1:B4,2,D1:D2)", &c), Variant::Integer(90));

        // No match → #VALUE!
        c.insert((2,4), CellContent { formula: None, value: Variant::Str("Zara".into()) });
        assert_eq!(calc("=DGET(A1:B4,\"Score\",D1:D2)", &c), Variant::Error(ExcelError::Value));

        // Multiple matches → #NUM!
        c.insert((2,4), CellContent { formula: None, value: Variant::Str("Alice".into()) });
        c.insert((3,4), CellContent { formula: None, value: Variant::Str("Bob".into()) });
        // criteria D1:D3 with two rows: Alice OR Bob
        c.insert((1,4), CellContent { formula: None, value: Variant::Str("Name".into()) });
        assert_eq!(calc("=DGET(A1:B4,\"Score\",D1:D3)", &c), Variant::Error(ExcelError::Num));
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
        c.insert((1,1), CellContent { formula: None, value: Variant::Str("Name".into()) });
        c.insert((1,2), CellContent { formula: None, value: Variant::Str("Score".into()) });
        c.insert((2,1), CellContent { formula: None, value: Variant::Str("Alice".into()) });
        c.insert((2,2), CellContent { formula: None, value: Variant::Integer(90) });
        c.insert((3,1), CellContent { formula: None, value: Variant::Str("Bob".into()) });
        c.insert((3,2), CellContent { formula: None, value: Variant::Integer(70) });
        c.insert((4,1), CellContent { formula: None, value: Variant::Str("Alice".into()) });
        c.insert((4,2), CellContent { formula: None, value: Variant::Integer(80) });
        c.insert((5,1), CellContent { formula: None, value: Variant::Str("Carol".into()) });
        c.insert((5,2), CellContent { formula: None, value: Variant::Integer(60) });
        c.insert((1,4), CellContent { formula: None, value: Variant::Str("Name".into()) });
        c.insert((2,4), CellContent { formula: None, value: Variant::Str("Alice".into()) });

        // DSUM: 90 + 80 = 170
        assert_eq!(calc("=DSUM(A1:B5,\"Score\",D1:D2)", &c), Variant::Integer(170));
        assert_eq!(calc("=DSUM(A1:B5,2,D1:D2)", &c), Variant::Integer(170));
        // DAVERAGE: (90 + 80) / 2 = 85.0
        assert_eq!(calc("=DAVERAGE(A1:B5,\"Score\",D1:D2)", &c), Variant::Float(85.0));
        // DCOUNT: 2 numeric values
        assert_eq!(calc("=DCOUNT(A1:B5,\"Score\",D1:D2)", &c), Variant::Integer(2));
        // DCOUNTA: 2 non-empty values
        assert_eq!(calc("=DCOUNTA(A1:B5,\"Name\",D1:D2)", &c), Variant::Integer(2));
        // DMAX: 90
        assert_eq!(calc("=DMAX(A1:B5,\"Score\",D1:D2)", &c), Variant::Integer(90));
        // DMIN: 80
        assert_eq!(calc("=DMIN(A1:B5,\"Score\",D1:D2)", &c), Variant::Integer(80));

        // No matches → DSUM=0, DAVERAGE=#DIV/0!, DCOUNT=0, DMAX=0, DMIN=0
        c.insert((2,4), CellContent { formula: None, value: Variant::Str("Zara".into()) });
        assert_eq!(calc("=DSUM(A1:B5,\"Score\",D1:D2)", &c), Variant::Integer(0));
        assert_eq!(calc("=DAVERAGE(A1:B5,\"Score\",D1:D2)", &c), Variant::Error(ExcelError::DivZero));
        assert_eq!(calc("=DCOUNT(A1:B5,\"Score\",D1:D2)", &c), Variant::Integer(0));
        assert_eq!(calc("=DMAX(A1:B5,\"Score\",D1:D2)", &c), Variant::Integer(0));
        assert_eq!(calc("=DMIN(A1:B5,\"Score\",D1:D2)", &c), Variant::Integer(0));
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
            ((1,1), Variant::Float(50.0)),
            ((2,1), Variant::Float(60.0)),
            ((3,1), Variant::Float(70.0)),
        ]);
        match calc("=NPV(0.1,A1:A3)", &c_npv) {
            Variant::Float(f) => assert!((f - 147.63).abs() < 0.01),
            other => panic!("NPV unexpected: {:?}", other),
        }

        // IRR: cash flows [-100, 40, 60, 50] → find r where -100 + 40/(1+r) + 60/(1+r)^2 + 50/(1+r)^3 = 0
        // Verify: NPV at found rate ≈ 0
        let c_irr = cells_from(&[
            ((1,1), Variant::Float(-100.0)),
            ((2,1), Variant::Float(40.0)),
            ((3,1), Variant::Float(60.0)),
            ((4,1), Variant::Float(50.0)),
        ]);
        let irr_rate = match calc("=IRR(A1:A4)", &c_irr) {
            Variant::Float(f) => f,
            other => panic!("IRR unexpected: {:?}", other),
        };
        // Verify NPV ≈ 0 at that rate
        let npv_check = -100.0 + 40.0/(1.0+irr_rate) + 60.0/(1.0+irr_rate).powi(2) + 50.0/(1.0+irr_rate).powi(3);
        assert!(npv_check.abs() < 0.001, "IRR: NPV at found rate should be ~0, got {}", npv_check);

        // MIRR: [-120, 50, 60, 70] at finance_rate=10%, reinvest_rate=12%
        // pv_neg = -120; fv_pos = 50*1.12^2 + 60*1.12 + 70 = 199.92
        // MIRR = (199.92/120)^(1/3) - 1 ≈ 18.58%
        let c_mirr = cells_from(&[
            ((1,1), Variant::Float(-120.0)),
            ((2,1), Variant::Float(50.0)),
            ((3,1), Variant::Float(60.0)),
            ((4,1), Variant::Float(70.0)),
        ]);
        match calc("=MIRR(A1:A4,0.1,0.12)", &c_mirr) {
            Variant::Float(f) => assert!((f - 0.1858).abs() < 0.001),
            other => panic!("MIRR unexpected: {:?}", other),
        }

        // XNPV: values=[-100, 50, 60] at dates=[0, 180, 365] (serial days from base)
        // Using small day numbers to simplify: date0=1, date1=181, date2=366
        let c_xnpv = cells_from(&[
            ((1,1), Variant::Float(-100.0)), ((1,2), Variant::Integer(1)),
            ((2,1), Variant::Float(50.0)),   ((2,2), Variant::Integer(181)),
            ((3,1), Variant::Float(60.0)),   ((3,2), Variant::Integer(366)),
        ]);
        match calc("=XNPV(0.1,A1:A3,B1:B3)", &c_xnpv) {
            Variant::Float(f) => {
                let expected = -100.0
                    + 50.0 / 1.1f64.powf(180.0/365.0)
                    + 60.0 / 1.1f64.powf(365.0/365.0);
                assert!((f - expected).abs() < 0.01, "XNPV: got {}, expected {}", f, expected);
            }
            other => panic!("XNPV unexpected: {:?}", other),
        }

        // XIRR: same cash flows, find rate where XNPV=0
        match calc("=XIRR(A1:A3,B1:B3)", &c_xnpv) {
            Variant::Float(f) => assert!(f > 0.0 && f < 2.0, "XIRR out of range: {}", f),
            other => panic!("XIRR unexpected: {:?}", other),
        }
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
                } else { panic!("expected Float, got {:?}", v); }
            }
        } else { panic!("expected Array"); }

        // RANDARRAY(2, 3) → 6 elements
        let result2 = calc("=RANDARRAY(2, 3)", &c);
        if let Variant::Array(arr) = result2 { assert_eq!(arr.len(), 6); }
        else { panic!("expected Array"); }

        // RANDARRAY(5, 1, 1, 10, TRUE) → integers in [1, 10]
        let result3 = calc("=RANDARRAY(5, 1, 1, 10, TRUE)", &c);
        if let Variant::Array(arr) = result3 {
            assert_eq!(arr.len(), 5);
            for v in &arr {
                let n = match v {
                    Variant::Integer(i) => *i as f64,
                    Variant::Float(f)   => *f,
                    other => panic!("expected numeric, got {:?}", other),
                };
                assert!(n >= 1.0 && n <= 10.0, "out of [1,10]: {}", n);
            }
        } else { panic!("expected Array"); }
    }

    #[test]
    fn test_tocol_torow() {
        let c = HashMap::new();
        // TOCOL on a literal range of constants (flat sequence)
        assert_eq!(calc("=TOCOL(1)", &c), Variant::Integer(1));
        // with ignore=0 (none): empties pass through
        let mut c2 = HashMap::new();
        c2.insert((1,1), CellContent { formula: None, value: Variant::Integer(10) });
        c2.insert((2,1), CellContent { formula: None, value: Variant::Empty });
        c2.insert((3,1), CellContent { formula: None, value: Variant::Integer(30) });
        // ignore=1 → skip blanks
        assert_eq!(calc("=TOCOL(A1:A3, 1)", &c2),
            Variant::Array(vec![Variant::Integer(10), Variant::Integer(30)]));
        // TOROW same semantics
        assert_eq!(calc("=TOROW(A1:A3, 1)", &c2),
            Variant::Array(vec![Variant::Integer(10), Variant::Integer(30)]));
    }

    #[test]
    fn test_wraprows() {
        let c = HashMap::new();
        // SEQUENCE(1,6,1,1) → [1,2,3,4,5,6]
        // WRAPROWS([1,2,3,4,5,6], 3) → [[1,2,3],[4,5,6]] → flat: [1,2,3,4,5,6]
        assert_eq!(calc("=WRAPROWS(SEQUENCE(6),3)", &c),
            Variant::Array(vec![
                Variant::Integer(1), Variant::Integer(2), Variant::Integer(3),
                Variant::Integer(4), Variant::Integer(5), Variant::Integer(6),
            ]));
        // WRAPROWS with padding: [1,2,3,4,5] wrap=3 → [[1,2,3],[4,5,0]]
        let mut c2 = HashMap::new();
        for i in 1u32..=5 { c2.insert((i as u32,1), CellContent { formula: None, value: Variant::Integer(i as i64) }); }
        assert_eq!(calc("=WRAPROWS(A1:A5, 3, 0)", &c2),
            Variant::Array(vec![
                Variant::Integer(1), Variant::Integer(2), Variant::Integer(3),
                Variant::Integer(4), Variant::Integer(5), Variant::Integer(0),
            ]));
    }

    #[test]
    fn test_wrapcols() {
        let c = HashMap::new();
        // WRAPCOLS(SEQUENCE(6), 2):
        //   vals=[1,2,3,4,5,6], wrap_count=2, n_cols=3
        //   col0=[1,2], col1=[3,4], col2=[5,6]
        //   row-major 2D (2 rows × 3 cols): [1,3,5, 2,4,6]
        assert_eq!(calc("=WRAPCOLS(SEQUENCE(6),2)", &c),
            Variant::Array(vec![
                Variant::Integer(1), Variant::Integer(3), Variant::Integer(5),
                Variant::Integer(2), Variant::Integer(4), Variant::Integer(6),
            ]));
        // WRAPCOLS([1,2,3,4,5], 2, 0): n_cols=3, last col=[5,0]
        // row-major: [1,3,5, 2,4,0]
        let mut c2 = HashMap::new();
        for i in 1u32..=5 { c2.insert((i as u32,1), CellContent { formula: None, value: Variant::Integer(i as i64) }); }
        assert_eq!(calc("=WRAPCOLS(A1:A5, 2, 0)", &c2),
            Variant::Array(vec![
                Variant::Integer(1), Variant::Integer(3), Variant::Integer(5),
                Variant::Integer(2), Variant::Integer(4), Variant::Integer(0),
            ]));
    }
}
