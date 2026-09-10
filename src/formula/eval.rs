use std::cell::RefCell;
use std::cmp::Ordering;
use std::collections::HashMap;

use super::ast::{BinOpKind, FormulaExpr};
use crate::types::{days_in_month, is_leap, serial_to_ymd};
use crate::vm::{CellContent, ExcelError, Variant};

// ── LET/LAMBDA name-binding stack ────────────────────────────────────────────
// A stack of binding frames; each frame is pushed by LET or a lambda call.

thread_local! {
    static BINDINGS: RefCell<Vec<HashMap<String, Variant>>> = const { RefCell::new(vec![]) };
}

fn push_bindings(frame: HashMap<String, Variant>) {
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
            if let Some(v) = frame.get(name) {
                return Some(v.clone());
            }
        }
        None
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
        FormulaExpr::Number(_) | FormulaExpr::Str(_) | FormulaExpr::Bool(_) => false,
        FormulaExpr::CellRef { sheet, .. } => sheet.is_some(),
        FormulaExpr::Range { sheet, .. } => sheet.is_some(),
        FormulaExpr::BinOp { lhs, rhs, .. } => {
            references_another_sheet(lhs) || references_another_sheet(rhs)
        }
        FormulaExpr::UnaryMinus(inner) => references_another_sheet(inner),
        FormulaExpr::FuncCall { args, .. } => args.iter().any(references_another_sheet),
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
        FormulaExpr::CellRef { col, row, .. } => Ok(cells
            .get(&(*row, *col))
            .map(|c| c.value.clone())
            .unwrap_or(Variant::Empty)),
        FormulaExpr::Range { .. } => Err("Range cannot be used as a scalar value".into()),
        FormulaExpr::UnaryMinus(inner) => match evaluate(inner, cells)? {
            Variant::Integer(n) => Ok(Variant::Integer(-n)),
            Variant::Float(f) => Ok(Variant::Float(-f)),
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
    match op {
        BinOpKind::Add | BinOpKind::Sub | BinOpKind::Mul | BinOpKind::Div => {
            let lf = to_float(&l)?;
            let rf = to_float(&r)?;
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
        BinOpKind::Eq => Ok(Variant::Boolean(variant_eq(&l, &r))),
        BinOpKind::Ne => Ok(Variant::Boolean(!variant_eq(&l, &r))),
        BinOpKind::Lt => Ok(Variant::Boolean(variant_cmp(&l, &r)? == Ordering::Less)),
        BinOpKind::Le => Ok(Variant::Boolean(variant_cmp(&l, &r)? != Ordering::Greater)),
        BinOpKind::Gt => Ok(Variant::Boolean(variant_cmp(&l, &r)? == Ordering::Greater)),
        BinOpKind::Ge => Ok(Variant::Boolean(variant_cmp(&l, &r)? != Ordering::Less)),
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
            let (base_row, base_col) = match &args[0] {
                FormulaExpr::CellRef { row, col, .. } => (*row, *col),
                FormulaExpr::Range { r1, c1, .. } => (*r1, *c1),
                _ => return Ok(vec![evaluate(expr, cells)?]),
            };
            let row_offset = to_float(&evaluate(&args[1], cells)?)? as i64;
            let col_offset = to_float(&evaluate(&args[2], cells)?)? as i64;
            let height = args
                .get(3)
                .map(|arg| {
                    evaluate(arg, cells).and_then(|value| to_float(&value).map(|v| v as i64))
                })
                .transpose()?
                .unwrap_or(1);
            let width = args
                .get(4)
                .map(|arg| {
                    evaluate(arg, cells).and_then(|value| to_float(&value).map(|v| v as i64))
                })
                .transpose()?
                .unwrap_or(1);
            if height <= 0 || width <= 0 {
                return Err("OFFSET: height and width must be positive".to_string());
            }
            let row = (base_row as i64)
                .checked_add(row_offset)
                .ok_or_else(|| "OFFSET: row overflow".to_string())?;
            let col = (base_col as i64)
                .checked_add(col_offset)
                .ok_or_else(|| "OFFSET: column overflow".to_string())?;
            if row < 1 || col < 1 {
                return Err("OFFSET: reference is outside the worksheet".to_string());
            }
            let total = (height as u64)
                .checked_mul(width as u64)
                .ok_or_else(|| "OFFSET: range is too large".to_string())?;
            if total > 1_000_000 {
                return Err("OFFSET: range too large (maximum is 1,000,000 cells)".to_string());
            }
            let mut vals = Vec::with_capacity(total as usize);
            for row_offset in 0..height {
                for col_offset in 0..width {
                    let target_row = row
                        .checked_add(row_offset)
                        .ok_or_else(|| "OFFSET: row overflow".to_string())?;
                    let target_col = col
                        .checked_add(col_offset)
                        .ok_or_else(|| "OFFSET: column overflow".to_string())?;
                    if target_row > u32::MAX as i64 || target_col > u32::MAX as i64 {
                        return Err("OFFSET: reference is outside the worksheet".to_string());
                    }
                    vals.push(cell_val(cells, target_row as u32, target_col as u32));
                }
            }
            Ok(vals)
        }
        other => match evaluate(other, cells)? {
            Variant::Array(values) => Ok(values),
            Variant::VbaArray(array) => Ok(array.elements),
            value => Ok(vec![value]),
        },
    }
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
        "MIN" => func_min(args, cells),
        "MAX" => func_max(args, cells),
        "COUNT" => func_count(args, cells),
        "COUNTA" => func_counta(args, cells),
        "IF" => func_if(args, cells),
        "AND" => func_and(args, cells),
        "OR" => func_or(args, cells),
        "NOT" => func_not(args, cells),
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
        "MODE.SNGL" => func_mode_mult(args, cells, false),
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
        "LARGE" => func_large(args, cells),
        "MAXIFS" => func_maxifs(args, cells),
        "MINIFS" => func_minifs(args, cells),
        "MOD" => func_mod(args, cells),
        "PERCENTILE" | "PERCENTILE.INC" => func_percentile(args, cells),
        "PERCENTILE.EXC" => func_percentile_exc(args, cells),
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
        "TRUNC" => func_trunc(args, cells),
        // -- String --
        "ASC" => func_asc(args, cells),
        "CHAR" => func_char(args, cells),
        "CODE" => func_code(args, cells),
        "EXACT" => func_exact(args, cells),
        "FIND" => func_find(args, cells),
        "JIS" => func_jis(args, cells),
        "LOWER" => func_lower(args, cells),
        "PROPER" => func_proper(args, cells),
        "REPLACE" => func_replace(args, cells),
        "SEARCH" => func_search(args, cells),
        "SUBSTITUTE" => func_substitute(args, cells),
        "TEXTJOIN" => func_textjoin(args, cells),
        "TEXTSPLIT" => func_textsplit(args, cells),
        "TEXTBEFORE" => func_textbefore(args, cells),
        "TEXTAFTER" => func_textafter(args, cells),
        "VALUETOTEXT" => func_valuetotext(args, cells),
        "TRIM" => func_trim(args, cells),
        "CLEAN" => func_clean(args, cells),
        "T" => func_t(args, cells),
        "FIXED" => func_fixed(args, cells),
        "DOLLAR" => func_dollar(args, cells),
        "BAHTTEXT" => func_bahttext(args, cells),
        "BASE" => func_base(args, cells),
        "DECIMAL" => func_decimal(args, cells),
        "UNICHAR" => func_char(args, cells),
        "UNICODE" => func_code(args, cells),
        "UPPER" => func_upper(args, cells),
        "VALUE" => func_value(args, cells),
        "REPT" => func_rept(args, cells),
        "NUMBERVALUE" => func_numbervalue(args, cells),
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
        "VLOOKUP" => func_vlookup(args, cells),
        "HLOOKUP" => func_hlookup(args, cells),
        "INDEX" => func_index(args, cells),
        "MATCH" => func_match_fn(args, cells),
        // ── Statistics ───────────────────────────────────────────────────────
        "STDEV" | "STDEV.S" => func_stdev_s(args, cells),
        "STDEVP" | "STDEV.P" => func_stdev_p(args, cells),
        "VAR" | "VAR.S" => func_var_s(args, cells),
        "VARP" | "VAR.P" => func_var_p(args, cells),
        "CORREL" | "PEARSON" => func_correl(args, cells),
        "SLOPE" => func_slope(args, cells),
        "INTERCEPT" => func_intercept(args, cells),
        "RSQ" => func_rsq(args, cells),
        "FORECAST.LINEAR" | "FORECAST" => func_forecast_linear(args, cells),
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
        "NORM.DIST" | "NORMDIST" => func_norm_dist(args, cells),
        "NORM.INV" | "NORMINV" => func_norm_inv(args, cells),
        "NORM.S.DIST" | "NORMSDIST" => func_norm_s_dist(args, cells),
        "NORM.S.INV" | "NORMSINV" => func_norm_s_inv(args, cells),
        "BINOM.DIST" | "BINOMDIST" => func_binom_dist(args, cells),
        "BINOM.DIST.RANGE" => func_binom_dist_range(args, cells),
        "BINOM.INV" => func_binom_inv(args, cells),
        "NEGBINOM.DIST" | "NEGBINOMDIST" => func_negbinom_dist(args, cells),
        "HYPGEOM.DIST" | "HYPGEOMDIST" => func_hypgeom_dist(args, cells),
        "POISSON.DIST" | "POISSON" => func_poisson_dist(args, cells),
        "GAMMA" => func_gamma(args, cells),
        "GAMMALN" | "GAMMALN.PRECISE" => func_gammaln(args, cells),
        "GAMMA.DIST" => func_gamma_dist(args, cells),
        "GAMMA.INV" => func_gamma_inv(args, cells),
        "BETA.DIST" | "BETADIST" => func_beta_dist(args, cells),
        "BETA.INV" | "BETAINV" => func_beta_inv(args, cells),
        "CHISQ.DIST" | "CHIDIST" => func_chisq_dist(args, cells),
        "CHISQ.DIST.RT" => func_chisq_dist_rt(args, cells),
        "CHISQ.INV" => func_chisq_inv(args, cells),
        "CHISQ.INV.RT" => func_chisq_inv_rt(args, cells),
        "F.DIST" | "FDIST" => func_f_dist(args, cells),
        "F.DIST.RT" => func_f_dist_rt(args, cells),
        "F.DIST.2T" => func_f_dist_2t(args, cells),
        "F.INV" => func_f_inv(args, cells),
        "F.INV.RT" => func_f_inv_rt(args, cells),
        "WEIBULL.DIST" | "WEIBULL" => func_weibull_dist(args, cells),
        "EXPON.DIST" | "EXPONDIST" => func_expon_dist(args, cells),
        "LOGNORM.DIST" | "LOGNORMDIST" => func_lognorm_dist(args, cells),
        "LOGNORM.INV" | "LOGINV" => func_lognorm_inv(args, cells),
        "T.DIST" => func_t_dist(args, cells),
        "T.DIST.2T" => func_t_dist_2t(args, cells),
        "T.DIST.RT" => func_t_dist_rt(args, cells),
        "T.INV" => func_t_inv(args, cells),
        "T.INV.2T" => func_t_inv_2t(args, cells),
        // ── Rounding ─────────────────────────────────────────────────────────
        "FLOOR" | "FLOOR.MATH" => func_floor(args, cells),
        "CEILING" | "CEILING.MATH" => func_ceiling(args, cells),
        "FLOOR.PRECISE" | "ISO.FLOOR" => func_precise_round(args, cells, false),
        "CEILING.PRECISE" | "ISO.CEILING" => func_precise_round(args, cells, true),
        "EVEN" => func_even_odd(args, cells, true),
        "ODD" => func_even_odd(args, cells, false),
        "MROUND" => func_mround(args, cells),
        // ── Math ─────────────────────────────────────────────────────────────
        "ABS" => func_abs(args, cells),
        "SQRT" => func_sqrt(args, cells),
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
        "NPV" => func_npv(args, cells),
        "IRR" => func_irr(args, cells),
        "MIRR" => func_mirr(args, cells),
        "XNPV" => func_xnpv(args, cells),
        "XIRR" => func_xirr(args, cells),
        "SLN" => func_sln(args, cells),
        "SYD" => func_syd(args, cells),
        "DB" => func_db(args, cells),
        "DDB" => func_ddb(args, cells),
        "EFFECT" => func_effect(args, cells),
        "NOMINAL" => func_nominal(args, cells),
        "RRI" => func_rri(args, cells),
        "CUMIPMT" => func_cumipmt(args, cells),
        "CUMPRINC" => func_cumprinc(args, cells),
        "FVSCHEDULE" => func_fvschedule(args, cells),
        "DOLLARDE" => func_dollarde(args, cells),
        "DOLLARFR" => func_dollarfr(args, cells),
        "PDURATION" => func_pduration(args, cells),
        "PRICEDISC" => func_pricedisc(args, cells),
        "DISC" => func_disc(args, cells),
        "RECEIVED" => func_received(args, cells),
        "YIELDDISC" => func_yielddisc(args, cells),
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
        _ => Ok(Variant::Error(ExcelError::Name)),
    }
}

// ── Arithmetic ────────────────────────────────────────────────────────────────

/// Collect numeric values from a single Range argument without allocating `Vec<Variant>`.
/// Falls back to collect_all for non-Range args or multiple args.
macro_rules! range_nums_fast {
    ($args:expr, $cells:expr) => {{
        if $args.len() == 1 {
            if let FormulaExpr::Range { c1, r1, c2, r2, .. } = &$args[0] {
                let mut nums: Vec<f64> = vec![];
                for row in *r1..=*r2 {
                    for col in *c1..=*c2 {
                        if let Some(f) = cell_ref($cells, row, col).and_then(as_f64) {
                            nums.push(f);
                        }
                    }
                }
                nums
            } else {
                collect_all($args, $cells)?
                    .iter()
                    .filter_map(as_f64)
                    .collect::<Vec<_>>()
            }
        } else {
            collect_all($args, $cells)?
                .iter()
                .filter_map(as_f64)
                .collect::<Vec<_>>()
        }
    }};
}

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
                if let Some(f) = cell_ref(cells, row, col).and_then(as_f64) {
                    sum += f;
                }
            }
        }
        return Ok(as_integer_if_whole(sum));
    }
    let sum: f64 = collect_all(args, cells)?
        .iter()
        .filter_map(|v| {
            if matches!(v, Variant::Str(_)) {
                None
            } else {
                to_float(v).ok()
            }
        })
        .sum();
    Ok(as_integer_if_whole(sum))
}

fn func_average(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let nums = range_nums_fast!(args, cells);
    if nums.is_empty() {
        return Err("AVERAGE: no numeric values".into());
    }
    Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
}

fn func_min(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let min = range_nums_fast!(args, cells).into_iter().reduce(f64::min);
    min.map(as_integer_if_whole)
        .ok_or_else(|| "MIN: no numeric values".into())
}

fn func_max(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let max = range_nums_fast!(args, cells).into_iter().reduce(f64::max);
    max.map(as_integer_if_whole)
        .ok_or_else(|| "MAX: no numeric values".into())
}

fn func_count(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    Ok(Variant::Integer(
        collect_all(args, cells)?
            .iter()
            .filter(|v| matches!(v, Variant::Integer(_) | Variant::Float(_)))
            .count() as i64,
    ))
}

fn func_counta(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    Ok(Variant::Integer(
        collect_all(args, cells)?
            .iter()
            .filter(|v| !matches!(v, Variant::Empty))
            .count() as i64,
    ))
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
    if let Variant::Error(error) = value {
        return Ok(Variant::Error(error));
    }
    Ok(Variant::Boolean(!is_truthy(&value)))
}

fn func_iferror(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 2 {
        return Err("IFERROR requires 2 arguments".into());
    }
    match evaluate(&args[0], cells) {
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
        Ok(Variant::Error(ExcelError::NA)) => evaluate(&args[1], cells),
        Ok(value) => Ok(value),
        Err(error) => Err(error),
    }
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
    if args.len() < 2 || args.len() > 3 {
        return Err("INDEX requires 2 or 3 arguments".into());
    }
    let (c1, r1, c2, r2) = require_range(&args[0], "INDEX")?;
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
    let col_off = if args.len() == 3 {
        match integer_index(&args[2]) {
            Ok(value) => value,
            Err(_) => return Ok(Variant::Error(ExcelError::Value)),
        }
    } else {
        1i64
    };
    if row_off < 0 || col_off < 0 {
        return Ok(Variant::Error(ExcelError::Value));
    }
    let height = i64::from(r2 - r1 + 1);
    let width = i64::from(c2 - c1 + 1);
    if row_off > height || col_off > width {
        return Ok(Variant::Error(ExcelError::Ref));
    }
    if row_off == 0 || col_off == 0 {
        let rows = if row_off == 0 {
            r1..=r2
        } else {
            r1 + row_off as u32 - 1..=r1 + row_off as u32 - 1
        };
        let cols = if col_off == 0 {
            c1..=c2
        } else {
            c1 + col_off as u32 - 1..=c1 + col_off as u32 - 1
        };
        let values = rows
            .flat_map(|row| cols.clone().map(move |col| cell_val(cells, row, col)))
            .collect();
        return Ok(Variant::Array(values));
    }
    Ok(cell_val(
        cells,
        r1 + row_off as u32 - 1,
        c1 + col_off as u32 - 1,
    ))
}

fn func_match_fn(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("MATCH requires 2 or 3 arguments".into());
    }
    let key = evaluate(&args[0], cells)?;
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
    let vals = collect_values(&args[0], cells)?;
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
    let range_vals = collect_values(&args[0], cells)?;
    let pcrit = parse_criteria(&evaluate(&args[1], cells)?);
    let total: f64 = if args.len() == 3 {
        let sum_vals = collect_values(&args[2], cells)?;
        range_vals
            .iter()
            .zip(sum_vals.iter())
            .filter(|(rv, _)| matches_parsed(rv, &pcrit))
            .filter_map(|(_, sv)| to_float(sv).ok())
            .sum()
    } else {
        // 2-arg: criteria range = sum range — avoid cloning range_vals
        range_vals
            .iter()
            .filter(|rv| matches_parsed(rv, &pcrit))
            .filter_map(|v| to_float(v).ok())
            .sum()
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
    let sum_vals = collect_values(&args[0], cells)?;
    let n = sum_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        let range_vals = collect_values(&args[i], cells)?;
        let pcrit = parse_criteria(&evaluate(&args[i + 1], cells)?);
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_parsed(rv, &pcrit) {
                mask[j] = false;
            }
        }
        i += 2;
    }
    let total: f64 = sum_vals
        .iter()
        .enumerate()
        .filter(|(j, _)| mask[*j])
        .filter_map(|(_, v)| to_float(v).ok())
        .sum();
    Ok(as_integer_if_whole(total))
}

fn func_countifs(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || !args.len().is_multiple_of(2) {
        return Err("COUNTIFS requires pairs of (range,criteria)".into());
    }
    let first_vals = collect_values(&args[0], cells)?;
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
        let range_vals = collect_values(&args[i], cells)?;
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
    let mut nums: Vec<f64> = collect_all(args, cells)?
        .iter()
        .filter_map(as_f64)
        .collect();
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
    let product: f64 = collect_all(args, cells)?
        .iter()
        .filter_map(as_f64)
        .fold(1.0, |acc, x| acc * x);
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

fn func_date(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() != 3 {
        return Err("DATE requires 3 arguments".into());
    }
    let y = to_float(&evaluate(&args[0], cells)?)? as i32;
    let m = to_float(&evaluate(&args[1], cells)?)? as u32;
    let d = to_float(&evaluate(&args[2], cells)?)? as u32;
    Ok(Variant::Date(date_to_serial(y, m, d)))
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

fn func_subtotal(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 {
        return Err("SUBTOTAL requires at least 2 arguments".into());
    }
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

fn func_averageif(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 2 || args.len() > 3 {
        return Err("AVERAGEIF requires 2 or 3 arguments".into());
    }
    let range_vals = collect_values(&args[0], cells)?;
    let pcrit = parse_criteria(&evaluate(&args[1], cells)?);
    let avg_vals = if args.len() == 3 {
        collect_values(&args[2], cells)?
    } else {
        range_vals.clone()
    };
    let nums: Vec<f64> = range_vals
        .iter()
        .zip(avg_vals.iter())
        .filter(|(rv, _)| matches_parsed(rv, &pcrit))
        .filter_map(|(_, av)| to_float(av).ok())
        .collect();
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
    let avg_vals = collect_values(&args[0], cells)?;
    let n = avg_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        let range_vals = collect_values(&args[i], cells)?;
        let pcrit = parse_criteria(&evaluate(&args[i + 1], cells)?);
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_parsed(rv, &pcrit) {
                mask[j] = false;
            }
        }
        i += 2;
    }
    let nums: Vec<f64> = avg_vals
        .iter()
        .enumerate()
        .filter(|(j, _)| mask[*j])
        .filter_map(|(_, v)| to_float(v).ok())
        .collect();
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
    let max_vals = collect_values(&args[0], cells)?;
    let n = max_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        let range_vals = collect_values(&args[i], cells)?;
        let pcrit = parse_criteria(&evaluate(&args[i + 1], cells)?);
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_parsed(rv, &pcrit) {
                mask[j] = false;
            }
        }
        i += 2;
    }
    let max = max_vals
        .iter()
        .enumerate()
        .filter(|(j, _)| mask[*j])
        .filter_map(|(_, v)| as_f64(v))
        .reduce(f64::max);
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
    let min_vals = collect_values(&args[0], cells)?;
    let n = min_vals.len();
    let mut mask = vec![true; n];
    let mut i = 1;
    while i + 1 < args.len() {
        let range_vals = collect_values(&args[i], cells)?;
        let pcrit = parse_criteria(&evaluate(&args[i + 1], cells)?);
        for (j, rv) in range_vals.iter().enumerate() {
            if j < n && !matches_parsed(rv, &pcrit) {
                mask[j] = false;
            }
        }
        i += 2;
    }
    let min = min_vals
        .iter()
        .enumerate()
        .filter(|(j, _)| mask[*j])
        .filter_map(|(_, v)| as_f64(v))
        .reduce(f64::min);
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
    let options = to_float(&evaluate(&args[1], cells)?)? as u32;
    let rest = &args[2..];
    // options & 6 != 0 means "ignore errors"
    let _ignore_errors = options & 6 != 0;
    let nums: Vec<f64> = collect_all(rest, cells)?
        .iter()
        .filter_map(|v| match v {
            Variant::Integer(n) => Some(*n as f64),
            Variant::Float(f) => Some(*f),
            _ => None,
        })
        .collect();
    // For fn_num that need filtered nums, handle ignore_errors by already filtering non-numeric
    match fn_num % 100 {
        1 => {
            if nums.is_empty() {
                return Err("AGGREGATE: no values".into());
            }
            Ok(Variant::Float(nums.iter().sum::<f64>() / nums.len() as f64))
        }
        2 => Ok(Variant::Integer(nums.len() as i64)),
        3 => {
            let count = collect_all(rest, cells)?
                .iter()
                .filter(|v| !matches!(v, Variant::Empty))
                .count();
            Ok(Variant::Integer(count as i64))
        }
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
        9 => Ok(as_integer_if_whole(nums.iter().sum::<f64>())),
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
        14 => {
            // LARGE
            let mut s = nums.clone();
            s.sort_by(|a, b| b.partial_cmp(a).unwrap_or(Ordering::Equal));
            s.first()
                .copied()
                .map(as_integer_if_whole)
                .ok_or_else(|| "AGGREGATE: no values".into())
        }
        15 => {
            // SMALL
            let mut s = nums.clone();
            s.sort_by(|a, b| a.partial_cmp(b).unwrap_or(Ordering::Equal));
            s.first()
                .copied()
                .map(as_integer_if_whole)
                .ok_or_else(|| "AGGREGATE: no values".into())
        }
        16 => func_percentile(rest, cells),
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
        to_float(&evaluate(&args[1], cells)?)? as u32
    } else {
        1
    };
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
        to_float(&evaluate(&args[1], cells)?)? as u32
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
    let s = to_str(&evaluate(&args[0], cells)?);
    // Support YYYY/MM/DD and YYYY-MM-DD
    let parts: Vec<&str> = if s.contains('/') {
        s.splitn(3, '/').collect()
    } else {
        s.splitn(3, '-').collect()
    };
    if parts.len() == 3
        && let (Ok(y), Ok(m), Ok(d)) = (
            parts[0].trim().parse::<i32>(),
            parts[1].trim().parse::<u32>(),
            parts[2].trim().parse::<u32>(),
        )
    {
        return Ok(Variant::Date(date_to_serial(y, m, d)));
    }
    Err(format!("DATEVALUE: cannot parse '{}'", s))
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
    let s = to_str(&evaluate(&args[0], cells)?);
    let parts: Vec<&str> = s.splitn(3, ':').collect();
    if parts.len() >= 2
        && let (Ok(h), Ok(m)) = (
            parts[0].trim().parse::<f64>(),
            parts[1].trim().parse::<f64>(),
        )
    {
        let sec = if parts.len() == 3 {
            parts[2].trim().parse::<f64>().unwrap_or(0.0)
        } else {
            0.0
        };
        return Ok(Variant::Float((h * 3600.0 + m * 60.0 + sec) / 86400.0));
    }
    Err(format!("TIMEVALUE: cannot parse '{}'", s))
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
        Variant::Str(s) if s.len() == 7 => {
            let mut mask = [false; 7];
            for (i, c) in s.chars().enumerate() {
                mask[i] = c == '1';
            }
            Ok(mask)
        }
        _ => {
            let n = to_float(v)? as u32;
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
    let idx = to_float(&evaluate(&args[0], cells)?)? as usize;
    if idx < 1 || idx >= args.len() {
        return Err("CHOOSE: index out of range".into());
    }
    evaluate(&args[idx], cells)
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
    let lookup = collect_values(&args[1], cells)?;
    let result = if args.len() == 3 {
        collect_values(&args[2], cells)?
    } else {
        lookup.clone()
    };
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
    Ok(collect_all(args, cells)?
        .into_iter()
        .filter_map(|v| match v {
            Variant::Integer(n) => Some(n as f64),
            Variant::Float(f) => Some(f),
            _ => None,
        })
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
    let a: Vec<f64> = collect_values(&args[0], cells)?
        .into_iter()
        .filter_map(|v| as_f64(&v))
        .collect();
    let b: Vec<f64> = collect_values(&args[1], cells)?
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

struct RegressionInputs {
    known_y: Vec<f64>,
    known_x: Vec<f64>,
    new_x: Vec<f64>,
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
    if known_y.is_empty() || known_y.len() != known_x.len() {
        return Err(format!(
            "{name}: known arrays must have equal non-zero length"
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
    let constant = if args.len() == 4 {
        is_truthy(&evaluate(&args[3], cells)?)
    } else {
        true
    };
    Ok(RegressionInputs {
        known_y,
        known_x,
        new_x,
        constant,
    })
}

fn func_trend(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    let inputs = regression_inputs(args, cells, "TREND")?;
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
        return Err("LINEST: known arrays must have equal length".into());
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
    if known_x.len() != known_y.len() || known_x.iter().any(|value| !value.is_finite()) {
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
    let (params, body) = extract_lambda(lambda_expr)?;
    if params.len() != arg_vals.len() {
        return Err(format!(
            "LAMBDA: expected {} args, got {}",
            params.len(),
            arg_vals.len()
        ));
    }
    let frame: HashMap<String, Variant> = params.into_iter().zip(arg_vals).collect();
    push_bindings(frame);
    let result = evaluate(body, cells);
    pop_bindings();
    result
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
    if args.is_empty() {
        return Err("INDIRECT requires 1 argument".into());
    }
    let addr_str = match evaluate(&args[0], cells)? {
        Variant::Str(s) => s,
        other => return Err(format!("INDIRECT: expected string, got {}", other)),
    };
    // Resolve through elixcee-types's parse_cell_addr / parse_range_addr
    let ((r1, c1), _) = crate::types::parse_range_addr(addr_str.trim())
        .ok_or_else(|| format!("INDIRECT: invalid reference '{}'", addr_str))?;
    Ok(cells
        .get(&(r1, c1))
        .map(|c| c.value.clone())
        .unwrap_or(Variant::Empty))
}

// ── OFFSET ────────────────────────────────────────────────────────────────────

fn func_offset(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.len() < 3 {
        return Err("OFFSET requires at least 3 arguments".into());
    }
    // First arg must be a cell/range reference expression — read coords without evaluating
    let (base_row, base_col): (u32, u32) = match args.first() {
        Some(FormulaExpr::CellRef { row, col, .. }) => (*row, *col),
        Some(FormulaExpr::Range { r1, c1, .. }) => (*r1, *c1),
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
    Ok(cells
        .get(&(new_row, new_col))
        .map(|c| c.value.clone())
        .unwrap_or(Variant::Empty))
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
    let by_vals = flatten_array_vals(collect_values(&args[1], cells)?);
    let order = if args.len() >= 3 {
        match evaluate(&args[2], cells)? {
            Variant::Integer(value) if value == 1 || value == -1 => value,
            Variant::Float(value) if value == 1.0 || value == -1.0 => value as i64,
            _ => return Ok(Variant::Error(ExcelError::Value)),
        }
    } else {
        1
    };
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
    if args.len() < 2 {
        return Err("TEXTSPLIT requires at least 2 arguments".into());
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let delim = to_str(&evaluate(&args[1], cells)?);
    if delim.is_empty() {
        return Err("TEXTSPLIT: delimiter cannot be empty".into());
    }
    let ignore_empty = args.len() >= 4 && is_truthy(&evaluate(&args[3], cells)?);
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
    if args.len() < 2 {
        return Err(format!("{} requires at least 2 arguments", fname));
    }
    let text = to_str(&evaluate(&args[0], cells)?);
    let delim = to_str(&evaluate(&args[1], cells)?);
    if delim.is_empty() {
        return Err(format!("{}: delimiter cannot be empty", fname));
    }
    let instance_num: i64 = if args.len() >= 3 {
        to_float(&evaluate(&args[2], cells)?)? as i64
    } else {
        1
    };
    let case_insensitive = args.len() >= 4 && is_truthy(&evaluate(&args[3], cells)?);

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
    let criteria: Vec<Vec<(u32, ParsedCriteria)>> = (cr_r1 + 1..=cr_r2)
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
        .collect();

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
    if args.is_empty() {
        return Err("TOCOL requires 1 argument".into());
    }
    let ignore = if args.len() >= 2 {
        to_float(&evaluate(&args[1], cells)?)? as u8
    } else {
        0
    };
    // args[2] = scan_by_column (bool); we treat range traversal as row-major regardless
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    Ok(wrap_array(ignore_filter(vals, ignore)))
}

fn func_torow(
    args: &[FormulaExpr],
    cells: &HashMap<(u32, u32), CellContent>,
) -> Result<Variant, String> {
    if args.is_empty() {
        return Err("TOROW requires 1 argument".into());
    }
    let ignore = if args.len() >= 2 {
        to_float(&evaluate(&args[1], cells)?)? as u8
    } else {
        0
    };
    let vals = flatten_array_vals(collect_values(&args[0], cells)?);
    Ok(wrap_array(ignore_filter(vals, ignore)))
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
        assert_eq!(calc("=COUNT(A1:A3)", &c), Variant::Integer(1));
        assert_eq!(calc("=COUNTA(A1:A3)", &c), Variant::Integer(2));
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
        assert_eq!(calc("=AND(TRUE,TRUE)", &c), Variant::Boolean(true));
        assert_eq!(calc("=AND(TRUE,FALSE)", &c), Variant::Boolean(false));
        assert_eq!(calc("=OR(FALSE,TRUE)", &c), Variant::Boolean(true));
        assert_eq!(calc("=NOT(TRUE)", &c), Variant::Boolean(false));
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
            calc("=COUNTIFS(A1:A3,\"a\",B1:B3,\">10\")", &c),
            Variant::Integer(1)
        );
    }

    #[test]
    fn test_median() {
        let c = HashMap::new();
        assert_eq!(calc("=MEDIAN(1,3,2)", &c), Variant::Integer(2));
        assert_eq!(calc("=MEDIAN(1,2,3,4)", &c), Variant::Float(2.5));
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
        assert_eq!(calc("=ROWS(TRANSPOSE(B3:C7))", &c), Variant::Integer(2));
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
    fn test_asc_jis() {
        let c = HashMap::new();
        // Full-width A (U+FF21) → half-width A
        assert_eq!(calc("=ASC(\"Ａ\")", &c), Variant::Str("A".into()));
        // Half-width A → full-width A
        assert_eq!(calc("=JIS(\"A\")", &c), Variant::Str("Ａ".into()));
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
        if let Variant::Float(v) = calc("=TIMEVALUE(\"12:00:00\")", &c) {
            assert!((v - 0.5).abs() < 1e-9);
        } else {
            panic!("TIMEVALUE should return Float");
        }
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
        assert_eq!(
            calc("=SWITCH(2,1,\"one\",2,\"two\",3,\"three\")", &c),
            Variant::Str("two".into())
        );
        assert_eq!(
            calc("=SWITCH(99,1,\"one\",\"default\")", &c),
            Variant::Str("default".into())
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
        assert_eq!(calc("=XMATCH(2,A1:A3,0)", &c), Variant::Integer(2));
        assert_eq!(calc("=XMATCH(2,A1:A3)", &c), Variant::Integer(2));
        assert_eq!(calc("=XMATCH(\"t*\",B1:B3,2)", &c), Variant::Integer(2));
        assert_eq!(calc("=XMATCH(2,A1:A3,0,2)", &c), Variant::Integer(2));
        assert_eq!(calc("=XMATCH(4,A1:A3,-1,2)", &c), Variant::Integer(3));
        assert_eq!(
            calc("=XMATCH(2,A1:A3,0.5)", &c),
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
        // INDIRECT with range reference → top-left cell
        assert_eq!(calc("=INDIRECT(\"A1:B3\")", &c), Variant::Integer(42));
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
        assert_eq!(calc("=MODE.SNGL(1,2,2,3)", &c), Variant::Integer(2));
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
        match calc("=GAMMA(5)", &cn) {
            Variant::Float(f) => assert!((f - 24.0).abs() < 1e-6),
            other => panic!("GAMMA: {:?}", other),
        }
        match calc("=GAMMALN(5)", &cn) {
            Variant::Float(f) => assert!((f - 24.0_f64.ln()).abs() < 1e-9),
            other => panic!("GAMMALN: {:?}", other),
        }
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
