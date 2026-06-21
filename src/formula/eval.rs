use std::cmp::Ordering;
use std::collections::HashMap;

use crate::vm::{CellContent, ExcelError, Variant};
use super::ast::{BinOpKind, FormulaExpr};

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
        FormulaExpr::FuncCall { name, args } => eval_func(name, args, cells),
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
        .filter_map(|v| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
        .collect();
    if nums.is_empty() { return Err("AVERAGE: no numeric values".into()); }
    Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
}

fn func_min(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let min = collect_all(args, cells)?.iter()
        .filter_map(|v| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
        .reduce(f64::min);
    min.map(as_integer_if_whole).ok_or_else(|| "MIN: no numeric values".into())
}

fn func_max(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    let max = collect_all(args, cells)?.iter()
        .filter_map(|v| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
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
    let (c1, r1, _c2, r2) = require_range(&args[1], "VLOOKUP")?;
    let col_n = to_float(&evaluate(&args[2], cells)?)? as u32;
    let exact = if args.len() == 4 { !is_truthy(&evaluate(&args[3], cells)?) } else { false };
    let ret_col = c1 + col_n - 1;

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
    let t: Vec<char> = text.chars().collect();
    let p: Vec<char> = pattern.chars().collect();
    fn wm(t: &[char], p: &[char]) -> bool {
        match (t, p) {
            (_, []) => t.is_empty(),
            (_, ['*', rest @ ..]) => wm(t, rest) || (!t.is_empty() && wm(&t[1..], p)),
            ([], _) => false,
            ([_, tr @ ..], ['?', pr @ ..]) => wm(tr, pr),
            ([tc, tr @ ..], [pc, pr @ ..]) if tc == pc => wm(tr, pr),
            _ => false,
        }
    }
    wm(&t, &p)
}

/// Like wildcard_match but pattern only needs to match a prefix of text (for SEARCH positioning).
fn wildcard_match_prefix(text: &[char], pattern: &[char]) -> bool {
    fn wm(t: &[char], p: &[char]) -> bool {
        match (t, p) {
            (_, []) => true,  // pattern consumed; remaining text is OK
            (_, ['*', rest @ ..]) => wm(t, rest) || (!t.is_empty() && wm(&t[1..], p)),
            ([], _) => false,
            ([_, tr @ ..], ['?', pr @ ..]) => wm(tr, pr),
            ([tc, tr @ ..], [pc, pr @ ..]) if tc == pc => wm(tr, pr),
            _ => false,
        }
    }
    wm(text, pattern)
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
        .filter_map(|v| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
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
        .filter_map(|v| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
        .fold(1.0, |acc, x| acc * x);
    Ok(as_integer_if_whole(product))
}

fn func_rank(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 { return Err("RANK requires 2 or 3 arguments".into()); }
    let num = to_float(&evaluate(&args[0], cells)?)?;
    let vals: Vec<f64> = collect_values(&args[1], cells)?.iter()
        .filter_map(|v| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
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
        .filter_map(|v| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
        .collect();
    let k = to_float(&evaluate(&args[1], cells)?)? as usize;
    if k == 0 || k > nums.len() { return Err("LARGE: k out of range".into()); }
    nums.sort_by(|a, b| b.partial_cmp(a).unwrap());
    Ok(as_integer_if_whole(nums[k - 1]))
}

fn func_small(args: &[FormulaExpr], cells: &HashMap<(u32, u32), CellContent>) -> Result<Variant, String> {
    if args.len() != 2 { return Err("SMALL requires 2 arguments".into()); }
    let mut nums: Vec<f64> = collect_values(&args[0], cells)?.iter()
        .filter_map(|v| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
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
        .filter_map(|(_, v)| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
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
        .filter_map(|(_, v)| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
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
        .filter_map(|v| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
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
        .filter_map(|v| match v { Variant::Integer(n) => Some(*n as f64), Variant::Float(f) => Some(*f), _ => None })
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
}
