//! Workbook-wide formula evaluation for sheet-qualified references.
//!
//! The hot single-sheet path remains in [`super::eval::evaluate`].  This module
//! is deliberately a separate slow-path: it builds one dependency order only
//! when a workbook contains a qualified reference, then evaluates the existing
//! formula AST against a coordinate-remapped view of all sheets.  The remap
//! lets the mature scalar/range/function evaluator stay unchanged while making
//! the sheet dimension explicit and preventing accidental reads from the host
//! sheet.

use std::collections::{HashMap, HashSet, VecDeque};

use super::ast::{FormulaExpr, SheetQualifier};
use super::{eval, parse};
use crate::vm::CellContent;

type Position = (u32, u32);
type SheetCells = HashMap<Position, CellContent>;
type NodeKey = (String, u32, u32);

const SHEET_ROW_STRIDE: u32 = 2_000_000;

/// Recalculate every formula in a workbook containing at least one qualified
/// reference. Returns `Ok(true)` when this slow path handled the workbook and
/// `Ok(false)` when the caller should use its existing active-sheet plan.
pub(crate) fn recalculate(
    sheets: &mut HashMap<String, SheetCells>,
    named_ranges: &HashMap<String, String>,
    scoped_named_ranges: &HashMap<String, HashMap<String, String>>,
    structured_ranges: &HashMap<String, String>,
    dirty_cells: Option<&HashMap<String, HashSet<Position>>>,
    force_full: bool,
) -> Result<bool, String> {
    let mut parsed: HashMap<NodeKey, FormulaExpr> = HashMap::new();
    let mut needs_workbook_path = false;
    for (sheet, cells) in sheets.iter() {
        for (&(row, col), cell) in cells {
            let Some(source) = cell.formula.as_deref() else {
                continue;
            };
            let local_names = scoped_named_ranges.get(&sheet.to_ascii_lowercase());
            let normalized_source =
                normalize_structured_refs(source, structured_ranges, sheet, row);
            let Ok(raw_expr) = parse(&normalized_source) else {
                continue;
            };
            needs_workbook_path |= contains_qualified_ref(&raw_expr)
                || contains_named_range(&raw_expr, named_ranges, local_names);
            let Ok(expr) = expand_named_ranges(raw_expr, named_ranges, local_names) else {
                continue;
            };
            parsed.insert((sheet.to_ascii_lowercase(), row, col), expr);
        }
    }
    if !needs_workbook_path {
        return Ok(false);
    }
    if !force_full && dirty_cells.is_some_and(HashMap::is_empty) {
        return Ok(true);
    }

    let mut names = sheets.keys().cloned().collect::<Vec<_>>();
    names.sort_by_key(|name| name.to_ascii_lowercase());
    let mut offsets = HashMap::new();
    for (index, name) in names.iter().enumerate() {
        let offset = (index as u32)
            .checked_mul(SHEET_ROW_STRIDE)
            .ok_or_else(|| "too many worksheets for formula coordinate remapping".to_string())?;
        offsets.insert(name.to_ascii_lowercase(), offset);
    }

    let mut merged = HashMap::new();
    for (sheet, cells) in sheets.iter() {
        let offset = offsets[&sheet.to_ascii_lowercase()];
        for (&(row, col), cell) in cells {
            let mapped_row = row.checked_add(offset).ok_or_else(|| {
                "worksheet row overflows formula coordinate remapping".to_string()
            })?;
            merged.insert((mapped_row, col), cell.clone());
        }
    }

    let mut dependents: HashMap<NodeKey, Vec<NodeKey>> = HashMap::new();
    let mut range_dependents = Vec::new();
    let mut indegree: HashMap<NodeKey, usize> = parsed.keys().map(|key| (key.clone(), 0)).collect();
    let mut nodes_by_sheet: HashMap<String, Vec<NodeKey>> = HashMap::new();
    for key in parsed.keys() {
        nodes_by_sheet
            .entry(key.0.clone())
            .or_default()
            .push(key.clone());
    }
    for (key, expr) in &parsed {
        let mut refs = Vec::new();
        let mut ranges = Vec::new();
        collect_dependencies(expr, &key.0, &mut refs, &mut ranges);
        let mut unique = HashSet::new();
        for reference in refs {
            if unique.insert(reference.clone()) {
                if parsed.contains_key(&reference) {
                    *indegree.get_mut(key).expect("parsed node has indegree") += 1;
                }
                // Keep reverse edges for value cells too. They do not
                // participate in topological ordering, but a later write to
                // such a cell must still select the formulas that depend on
                // it during an incremental recalculation.
                dependents.entry(reference).or_default().push(key.clone());
            }
        }
        // Keep large ranges as intervals rather than expanding every covered
        // cell. Only formula nodes can affect ordering, so an interval lookup
        // over the sheet-local node index is sufficient and bounded by formula
        // count on that sheet.
        for (sheet, r1, c1, r2, c2) in ranges {
            range_dependents.push((sheet.clone(), r1, c1, r2, c2, key.clone()));
            for reference in nodes_by_sheet
                .get(&sheet)
                .into_iter()
                .flatten()
                .filter(|(_, row, col)| r1 <= *row && *row <= r2 && c1 <= *col && *col <= c2)
            {
                if unique.insert(reference.clone()) {
                    *indegree.get_mut(key).expect("parsed node has indegree") += 1;
                    dependents
                        .entry(reference.clone())
                        .or_default()
                        .push(key.clone());
                }
            }
        }
    }

    let mut queue = VecDeque::new();
    let mut ready = indegree
        .iter()
        .filter_map(|(key, &degree)| (degree == 0).then_some(key.clone()))
        .collect::<Vec<_>>();
    ready.sort();
    queue.extend(ready);
    let mut order = Vec::with_capacity(parsed.len());
    while let Some(key) = queue.pop_front() {
        order.push(key.clone());
        if let Some(children) = dependents.get(&key) {
            let mut children = children.clone();
            children.sort();
            for child in children {
                let degree = indegree.get_mut(&child).expect("dependent node exists");
                *degree -= 1;
                if *degree == 0 {
                    queue.push_back(child);
                }
            }
        }
    }
    // Preserve the existing best-effort cycle behavior: cyclic formulas are
    // evaluated after acyclic formulas using their previous cached values.
    let mut cyclic = parsed
        .keys()
        .filter(|key| !order.iter().any(|existing| existing == *key))
        .cloned()
        .collect::<Vec<_>>();
    cyclic.sort();
    for key in cyclic {
        order.push(key);
    }

    let selected = if force_full {
        parsed.keys().cloned().collect()
    } else {
        let mut selected = HashSet::new();
        let mut queue = VecDeque::new();
        if let Some(dirty_cells) = dirty_cells {
            for (sheet, cells) in dirty_cells {
                let sheet = sheet.to_ascii_lowercase();
                for &(row, col) in cells {
                    let input = (sheet.clone(), row, col);
                    if let Some(children) = dependents.get(&input) {
                        queue.extend(children.iter().cloned());
                    }
                    for (range_sheet, r1, c1, r2, c2, formula) in &range_dependents {
                        if range_sheet == &sheet
                            && *r1 <= row
                            && row <= *r2
                            && *c1 <= col
                            && col <= *c2
                        {
                            queue.push_back(formula.clone());
                        }
                    }
                }
            }
        }
        while let Some(key) = queue.pop_front() {
            if selected.insert(key.clone())
                && let Some(children) = dependents.get(&key)
            {
                queue.extend(children.iter().cloned());
            }
        }
        selected
    };
    for key in order {
        if !selected.contains(&key) {
            continue;
        }
        let expr = &parsed[&key];
        let mapped = remap_expr(expr, &key.0, &offsets)?;
        let value = eval::evaluate(&mapped, &merged)?;
        let offset = offsets[&key.0.to_ascii_lowercase()];
        let mapped_row = key
            .1
            .checked_add(offset)
            .ok_or_else(|| "worksheet row overflows formula coordinate remapping".to_string())?;
        merged.insert(
            (mapped_row, key.2),
            CellContent {
                formula: Some(String::new()),
                value: value.clone(),
            },
        );
        let actual_sheet = sheets
            .keys()
            .find(|name| name.eq_ignore_ascii_case(&key.0))
            .cloned();
        if let Some(cell) = actual_sheet
            .as_deref()
            .and_then(|sheet| sheets.get_mut(sheet))
            .and_then(|cells| cells.get_mut(&(key.1, key.2)))
        {
            cell.value = value;
        }
    }
    Ok(true)
}

/// Detect whether the formula dependency graph contains a cycle.
///
/// This is intentionally diagnostic-only: recalculation keeps its historical
/// best-effort behavior and evaluates cyclic nodes using their cached values.
pub(crate) fn has_formula_cycle(sheets: &HashMap<String, SheetCells>) -> bool {
    let mut formulas = HashMap::new();
    let mut nodes_by_sheet: HashMap<String, Vec<NodeKey>> = HashMap::new();
    for (sheet, cells) in sheets {
        let sheet_key = sheet.to_ascii_lowercase();
        for (&(row, col), cell) in cells {
            let Some(source) = cell.formula.as_deref() else {
                continue;
            };
            let Ok(expr) = parse(source) else {
                continue;
            };
            let key = (sheet_key.clone(), row, col);
            formulas.insert(key.clone(), expr);
            nodes_by_sheet
                .entry(sheet_key.clone())
                .or_default()
                .push(key);
        }
    }
    if formulas.is_empty() {
        return false;
    }

    let mut indegree: HashMap<NodeKey, usize> =
        formulas.keys().map(|key| (key.clone(), 0usize)).collect();
    let mut dependents: HashMap<NodeKey, Vec<NodeKey>> = HashMap::new();
    for (key, expr) in &formulas {
        let mut refs = Vec::new();
        let mut ranges = Vec::new();
        collect_dependencies(expr, &key.0, &mut refs, &mut ranges);
        let mut unique = HashSet::new();
        for reference in refs {
            if formulas.contains_key(&reference) && unique.insert(reference.clone()) {
                *indegree.get_mut(key).expect("formula node has indegree") += 1;
                dependents.entry(reference).or_default().push(key.clone());
            }
        }
        for (sheet, r1, c1, r2, c2) in ranges {
            for reference in nodes_by_sheet
                .get(&sheet)
                .into_iter()
                .flatten()
                .filter(|(_, row, col)| r1 <= *row && *row <= r2 && c1 <= *col && *col <= c2)
            {
                if unique.insert(reference.clone()) {
                    *indegree.get_mut(key).expect("formula node has indegree") += 1;
                    dependents
                        .entry(reference.clone())
                        .or_default()
                        .push(key.clone());
                }
            }
        }
    }

    let mut queue = VecDeque::new();
    let mut ready = indegree
        .iter()
        .filter_map(|(key, &degree)| (degree == 0).then_some(key.clone()))
        .collect::<Vec<_>>();
    ready.sort();
    queue.extend(ready);
    let mut visited = 0;
    while let Some(key) = queue.pop_front() {
        visited += 1;
        if let Some(children) = dependents.get(&key) {
            let mut children = children.clone();
            children.sort();
            for child in children {
                let degree = indegree.get_mut(&child).expect("dependent node exists");
                *degree -= 1;
                if *degree == 0 {
                    queue.push_back(child);
                }
            }
        }
    }
    visited != formulas.len()
}

fn contains_qualified_ref(expr: &FormulaExpr) -> bool {
    match expr {
        FormulaExpr::CellRef { sheet, .. } | FormulaExpr::Range { sheet, .. } => sheet.is_some(),
        FormulaExpr::BinOp { lhs, rhs, .. } => {
            contains_qualified_ref(lhs) || contains_qualified_ref(rhs)
        }
        FormulaExpr::UnaryMinus(inner) => contains_qualified_ref(inner),
        FormulaExpr::FuncCall { args, .. } => args.iter().any(contains_qualified_ref),
        FormulaExpr::Number(_) | FormulaExpr::Str(_) | FormulaExpr::Bool(_) => false,
    }
}

fn normalize_structured_refs(
    source: &str,
    structured_ranges: &HashMap<String, String>,
    host_sheet: &str,
    host_row: u32,
) -> String {
    let mut normalized = source.to_string();
    let mut patterns = structured_ranges.keys().collect::<Vec<_>>();
    patterns.sort_by_key(|pattern| std::cmp::Reverse(pattern.len()));
    for pattern in patterns {
        let replacement = &structured_ranges[pattern];
        let (scope, replacement) = replacement
            .strip_prefix("@sheet:")
            .and_then(|value| value.split_once('|'))
            .map_or((None, replacement.as_str()), |(scope, value)| {
                (Some(scope), value)
            });
        if let Some(scope) = scope {
            let Some((scope_sheet, row_band)) = scope.split_once(";rows:") else {
                continue;
            };
            let Some((first_row, last_row)) = row_band.split_once('-') else {
                continue;
            };
            let Ok(first_row) = first_row.parse::<u32>() else {
                continue;
            };
            let Ok(last_row) = last_row.parse::<u32>() else {
                continue;
            };
            if !scope_sheet.eq_ignore_ascii_case(host_sheet)
                || host_row < first_row
                || host_row > last_row
            {
                continue;
            }
        }
        let lower = normalized.to_ascii_lowercase();
        let pattern_lower = pattern.to_ascii_lowercase();
        let mut rebuilt = String::with_capacity(normalized.len());
        let mut cursor = 0;
        while let Some(relative) = lower[cursor..].find(&pattern_lower) {
            let start = cursor + relative;
            let end = start + pattern.len();
            let preceded_by_identifier = normalized[..start]
                .chars()
                .next_back()
                .is_some_and(|ch| ch.is_ascii_alphanumeric() || ch == '_');
            let followed_by_identifier = normalized[end..]
                .chars()
                .next()
                .is_some_and(|ch| ch.is_ascii_alphanumeric() || ch == '_');
            let inside_string = normalized[..start]
                .chars()
                .fold(false, |inside, ch| if ch == '"' { !inside } else { inside });
            rebuilt.push_str(&normalized[cursor..start]);
            if preceded_by_identifier || followed_by_identifier || inside_string {
                rebuilt.push_str(&normalized[start..end]);
            } else {
                rebuilt.push_str(&replacement.replace("{row}", &host_row.to_string()));
            }
            cursor = end;
        }
        if cursor != 0 {
            rebuilt.push_str(&normalized[cursor..]);
            normalized = rebuilt;
        }
    }
    normalized
}

fn contains_named_range(
    expr: &FormulaExpr,
    named_ranges: &HashMap<String, String>,
    scoped_named_ranges: Option<&HashMap<String, String>>,
) -> bool {
    match expr {
        FormulaExpr::FuncCall { name, args } => {
            (args.is_empty()
                && (scoped_named_ranges
                    .is_some_and(|local| local.contains_key(&name.to_ascii_lowercase()))
                    || named_ranges.contains_key(&name.to_ascii_lowercase())))
                || args
                    .iter()
                    .any(|arg| contains_named_range(arg, named_ranges, scoped_named_ranges))
        }
        FormulaExpr::BinOp { lhs, rhs, .. } => {
            contains_named_range(lhs, named_ranges, scoped_named_ranges)
                || contains_named_range(rhs, named_ranges, scoped_named_ranges)
        }
        FormulaExpr::UnaryMinus(inner) => {
            contains_named_range(inner, named_ranges, scoped_named_ranges)
        }
        FormulaExpr::Number(_)
        | FormulaExpr::Str(_)
        | FormulaExpr::Bool(_)
        | FormulaExpr::CellRef { .. }
        | FormulaExpr::Range { .. } => false,
    }
}

fn expand_named_ranges(
    expr: FormulaExpr,
    named_ranges: &HashMap<String, String>,
    scoped_named_ranges: Option<&HashMap<String, String>>,
) -> Result<FormulaExpr, String> {
    Ok(match expr {
        FormulaExpr::FuncCall { name, args } if args.is_empty() => {
            let address = scoped_named_ranges
                .and_then(|local| local.get(&name.to_ascii_lowercase()))
                .or_else(|| named_ranges.get(&name.to_ascii_lowercase()));
            let Some(address) = address else {
                return Ok(FormulaExpr::FuncCall { name, args });
            };
            if address.contains('!') || crate::vm::parse_range_addr(address).is_none() {
                parse(address).map_err(|_| format!("named range has an invalid address: {name}"))?
            } else {
                let ((r1, c1), (r2, c2)) = crate::vm::parse_range_addr(address)
                    .ok_or_else(|| format!("named range has an invalid address: {name}"))?;
                FormulaExpr::Range {
                    c1,
                    r1,
                    c2,
                    r2,
                    abs_c1: false,
                    abs_r1: false,
                    abs_c2: false,
                    abs_r2: false,
                    sheet: None,
                }
            }
        }
        FormulaExpr::BinOp { op, lhs, rhs } => FormulaExpr::BinOp {
            op,
            lhs: Box::new(expand_named_ranges(
                *lhs,
                named_ranges,
                scoped_named_ranges,
            )?),
            rhs: Box::new(expand_named_ranges(
                *rhs,
                named_ranges,
                scoped_named_ranges,
            )?),
        },
        FormulaExpr::UnaryMinus(inner) => FormulaExpr::UnaryMinus(Box::new(expand_named_ranges(
            *inner,
            named_ranges,
            scoped_named_ranges,
        )?)),
        FormulaExpr::FuncCall { name, args } => FormulaExpr::FuncCall {
            name,
            args: args
                .into_iter()
                .map(|arg| expand_named_ranges(arg, named_ranges, scoped_named_ranges))
                .collect::<Result<_, _>>()?,
        },
        other => other,
    })
}

fn target_sheet(host: &str, qualifier: Option<&SheetQualifier>) -> String {
    qualifier
        .map(|sheet| sheet.normalized_name.to_ascii_lowercase())
        .unwrap_or_else(|| host.to_ascii_lowercase())
}

fn collect_dependencies(
    expr: &FormulaExpr,
    host: &str,
    refs: &mut Vec<NodeKey>,
    ranges: &mut Vec<(String, u32, u32, u32, u32)>,
) {
    match expr {
        FormulaExpr::CellRef {
            col, row, sheet, ..
        } => {
            refs.push((target_sheet(host, sheet.as_ref()), *row, *col));
        }
        FormulaExpr::Range {
            c1,
            r1,
            c2,
            r2,
            sheet,
            ..
        } => {
            ranges.push((
                target_sheet(host, sheet.as_ref()),
                (*r1).min(*r2),
                (*c1).min(*c2),
                (*r1).max(*r2),
                (*c1).max(*c2),
            ));
        }
        FormulaExpr::BinOp { lhs, rhs, .. } => {
            collect_dependencies(lhs, host, refs, ranges);
            collect_dependencies(rhs, host, refs, ranges);
        }
        FormulaExpr::UnaryMinus(inner) => collect_dependencies(inner, host, refs, ranges),
        FormulaExpr::FuncCall { args, .. } => {
            for arg in args {
                collect_dependencies(arg, host, refs, ranges);
            }
        }
        FormulaExpr::Number(_) | FormulaExpr::Str(_) | FormulaExpr::Bool(_) => {}
    }
}

fn remap_expr(
    expr: &FormulaExpr,
    host: &str,
    offsets: &HashMap<String, u32>,
) -> Result<FormulaExpr, String> {
    let map_position = |sheet: Option<&SheetQualifier>, row: u32| -> Result<u32, String> {
        let name = target_sheet(host, sheet);
        let offset = offsets
            .get(&name)
            .ok_or_else(|| format!("unknown worksheet in formula reference: {name}"))?;
        row.checked_add(*offset)
            .ok_or_else(|| "worksheet row overflows formula coordinate remapping".to_string())
    };
    Ok(match expr {
        FormulaExpr::CellRef {
            col,
            row,
            abs_col,
            abs_row,
            sheet,
        } => FormulaExpr::CellRef {
            col: *col,
            row: map_position(sheet.as_ref(), *row)?,
            abs_col: *abs_col,
            abs_row: *abs_row,
            sheet: None,
        },
        FormulaExpr::Range {
            c1,
            r1,
            c2,
            r2,
            abs_c1,
            abs_r1,
            abs_c2,
            abs_r2,
            sheet,
        } => FormulaExpr::Range {
            c1: *c1,
            r1: map_position(sheet.as_ref(), *r1)?,
            c2: *c2,
            r2: map_position(sheet.as_ref(), *r2)?,
            abs_c1: *abs_c1,
            abs_r1: *abs_r1,
            abs_c2: *abs_c2,
            abs_r2: *abs_r2,
            sheet: None,
        },
        FormulaExpr::BinOp { op, lhs, rhs } => FormulaExpr::BinOp {
            op: op.clone(),
            lhs: Box::new(remap_expr(lhs, host, offsets)?),
            rhs: Box::new(remap_expr(rhs, host, offsets)?),
        },
        FormulaExpr::UnaryMinus(inner) => {
            FormulaExpr::UnaryMinus(Box::new(remap_expr(inner, host, offsets)?))
        }
        FormulaExpr::FuncCall { name, args } => FormulaExpr::FuncCall {
            name: name.clone(),
            args: args
                .iter()
                .map(|arg| remap_expr(arg, host, offsets))
                .collect::<Result<_, _>>()?,
        },
        FormulaExpr::Number(n) => FormulaExpr::Number(*n),
        FormulaExpr::Str(value) => FormulaExpr::Str(value.clone()),
        FormulaExpr::Bool(value) => FormulaExpr::Bool(*value),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::vm::Variant;

    fn cell(value: Variant, formula: Option<&str>) -> CellContent {
        CellContent {
            formula: formula.map(str::to_string),
            value,
        }
    }

    #[test]
    fn recalculates_qualified_cells_and_cross_sheet_formula_chains() {
        let mut sheets = HashMap::new();
        sheets.insert(
            "Sheet1".to_string(),
            HashMap::from([
                ((1, 1), cell(Variant::Integer(3), None)),
                ((1, 2), cell(Variant::Empty, Some("=Sheet2!A1+1"))),
            ]),
        );
        sheets.insert(
            "Sheet2".to_string(),
            HashMap::from([((1, 1), cell(Variant::Empty, Some("=Sheet1!A1*2")))]),
        );
        assert!(
            recalculate(
                &mut sheets,
                &HashMap::new(),
                &HashMap::new(),
                &HashMap::new(),
                None,
                true
            )
            .unwrap()
        );
        assert_eq!(sheets["Sheet2"][&(1, 1)].value, Variant::Integer(6));
        assert_eq!(sheets["Sheet1"][&(1, 2)].value, Variant::Integer(7));
    }

    #[test]
    fn unqualified_formulas_stay_on_their_host_sheet() {
        let mut sheets = HashMap::from([
            (
                "A".to_string(),
                HashMap::from([((1, 1), cell(Variant::Integer(2), None))]),
            ),
            (
                "B".to_string(),
                HashMap::from([((1, 1), cell(Variant::Integer(9), None))]),
            ),
        ]);
        assert!(
            !recalculate(
                &mut sheets,
                &HashMap::new(),
                &HashMap::new(),
                &HashMap::new(),
                None,
                true
            )
            .unwrap()
        );
    }

    #[test]
    fn expands_a_named_range_inside_a_cross_sheet_formula_workbook() {
        let mut sheets = HashMap::from([
            (
                "Sheet1".to_string(),
                HashMap::from([((1, 1), cell(Variant::Integer(2), None))]),
            ),
            (
                "Sheet2".to_string(),
                HashMap::from([
                    ((1, 1), cell(Variant::Integer(2), None)),
                    (
                        (2, 1),
                        cell(Variant::Empty, Some("=SUM(MyRange)+Sheet1!A1")),
                    ),
                ]),
            ),
        ]);
        let names = HashMap::from([(String::from("myrange"), String::from("A1:A1"))]);
        assert!(
            recalculate(
                &mut sheets,
                &names,
                &HashMap::new(),
                &HashMap::new(),
                None,
                true
            )
            .unwrap()
        );
        assert_eq!(sheets["Sheet2"][&(2, 1)].value, Variant::Integer(4));
    }

    #[test]
    fn sheet_local_named_range_shadows_global_name_on_its_host_sheet() {
        let mut sheets = HashMap::from([
            (
                "Sheet1".to_string(),
                HashMap::from([
                    ((1, 1), cell(Variant::Integer(2), None)),
                    ((2, 1), cell(Variant::Integer(9), None)),
                ]),
            ),
            (
                "Sheet2".to_string(),
                HashMap::from([
                    ((1, 1), cell(Variant::Empty, Some("=SUM(MyRange)"))),
                    ((2, 1), cell(Variant::Integer(9), None)),
                ]),
            ),
        ]);
        let global = HashMap::from([(String::from("myrange"), String::from("A1:A1"))]);
        let local = HashMap::from([(
            String::from("sheet2"),
            HashMap::from([(String::from("myrange"), String::from("A2:A2"))]),
        )]);
        assert!(recalculate(&mut sheets, &global, &local, &HashMap::new(), None, true).unwrap());
        assert_eq!(sheets["Sheet2"][&(1, 1)].value, Variant::Integer(9));
    }

    #[test]
    fn expands_a_qualified_named_range_without_rebinding_to_the_host_sheet() {
        let mut sheets = HashMap::from([
            (
                "Sheet1".to_string(),
                HashMap::from([((1, 1), cell(Variant::Integer(8), None))]),
            ),
            (
                "Sheet2".to_string(),
                HashMap::from([((1, 1), cell(Variant::Integer(3), None))]),
            ),
        ]);
        let names = HashMap::from([(String::from("source"), String::from("Sheet1!$A$1"))]);
        sheets
            .get_mut("Sheet2")
            .unwrap()
            .insert((2, 1), cell(Variant::Empty, Some("=Source+1")));
        assert!(
            recalculate(
                &mut sheets,
                &names,
                &HashMap::new(),
                &HashMap::new(),
                None,
                true
            )
            .unwrap()
        );
        assert_eq!(sheets["Sheet2"][&(2, 1)].value, Variant::Integer(9));
    }

    #[test]
    fn expands_a_dynamic_offset_named_range_for_aggregates() {
        let mut sheets = HashMap::from([
            (
                "Sheet1".to_string(),
                HashMap::from([
                    ((1, 1), cell(Variant::Integer(2), None)),
                    ((2, 1), cell(Variant::Integer(3), None)),
                    ((3, 1), cell(Variant::Integer(5), None)),
                ]),
            ),
            (
                "Sheet2".to_string(),
                HashMap::from([((1, 1), cell(Variant::Empty, Some("=SUM(DynamicRange)")))]),
            ),
        ]);
        let names = HashMap::from([(
            String::from("dynamicrange"),
            String::from("OFFSET(Sheet1!$A$1,0,0,3,1)"),
        )]);
        assert!(
            recalculate(
                &mut sheets,
                &names,
                &HashMap::new(),
                &HashMap::new(),
                None,
                true
            )
            .unwrap()
        );
        assert_eq!(sheets["Sheet2"][&(1, 1)].value, Variant::Integer(10));
    }

    #[test]
    fn formula_cycle_diagnostic_distinguishes_cyclic_and_acyclic_workbooks() {
        let mut cyclic = HashMap::from([("Sheet1".to_string(), HashMap::new())]);
        cyclic
            .get_mut("Sheet1")
            .unwrap()
            .insert((1, 1), cell(Variant::Empty, Some("=B1+1")));
        cyclic
            .get_mut("Sheet1")
            .unwrap()
            .insert((1, 2), cell(Variant::Empty, Some("=A1+1")));
        assert!(has_formula_cycle(&cyclic));

        let mut acyclic = cyclic.clone();
        acyclic
            .get_mut("Sheet1")
            .unwrap()
            .insert((1, 2), cell(Variant::Integer(1), None));
        assert!(!has_formula_cycle(&acyclic));
    }

    #[test]
    fn formula_cycle_diagnostic_crosses_qualified_sheet_references() {
        let mut sheets = HashMap::from([
            (
                "Input".to_string(),
                HashMap::from([((1, 1), cell(Variant::Empty, Some("=Calc!A1+1")))]),
            ),
            (
                "Calc".to_string(),
                HashMap::from([((1, 1), cell(Variant::Empty, Some("=Input!A1+1")))]),
            ),
        ]);
        assert!(has_formula_cycle(&sheets));

        sheets
            .get_mut("Calc")
            .unwrap()
            .insert((1, 1), cell(Variant::Integer(1), None));
        assert!(!has_formula_cycle(&sheets));
    }

    #[test]
    fn normalizes_a_simple_structured_column_reference() {
        let mut sheets = HashMap::from([
            (
                "Sheet1".to_string(),
                HashMap::from([
                    ((2, 2), cell(Variant::Integer(4), None)),
                    ((3, 2), cell(Variant::Integer(6), None)),
                ]),
            ),
            (
                "Sheet2".to_string(),
                HashMap::from([((1, 1), cell(Variant::Empty, Some("=SUM(Sales[Amount])")))]),
            ),
        ]);
        let structured = HashMap::from([(
            String::from("sales[amount]"),
            String::from("elixceetablesalesamount"),
        )]);
        let names = HashMap::from([(
            String::from("elixceetablesalesamount"),
            String::from("Sheet1!B2:B3"),
        )]);
        assert_eq!(
            normalize_structured_refs("=SUM(Sales[Amount])", &structured, "Sheet1", 1),
            "=SUM(elixceetablesalesamount)"
        );
        let normalized = parse("=SUM(ElixceeTableSalesAmount)").unwrap();
        assert!(contains_named_range(&normalized, &names, None));
        assert!(
            recalculate(
                &mut sheets,
                &names,
                &HashMap::new(),
                &structured,
                None,
                true
            )
            .unwrap()
        );
        assert_eq!(sheets["Sheet2"][&(1, 1)].value, Variant::Integer(10));
    }

    #[test]
    fn does_not_normalize_structured_text_inside_a_string_literal() {
        let structured = HashMap::from([(
            String::from("sales[amount]"),
            String::from("elixceetablesalesamount"),
        )]);

        assert_eq!(
            normalize_structured_refs("=\"Sales[Amount]\"", &structured, "Sheet1", 1),
            "=\"Sales[Amount]\""
        );
    }

    #[test]
    fn does_not_normalize_a_structured_fragment_inside_a_longer_identifier() {
        let structured = HashMap::from([(
            String::from("sales[amount]"),
            String::from("elixceetablesalesamount"),
        )]);

        assert_eq!(
            normalize_structured_refs("=MySales[Amount]", &structured, "Sheet1", 1),
            "=MySales[Amount]"
        );
    }

    #[test]
    fn normalizes_this_row_structured_reference_using_formula_host_row() {
        let mut sheets = HashMap::from([(
            "Sheet1".to_string(),
            HashMap::from([
                ((2, 2), cell(Variant::Integer(4), None)),
                ((3, 2), cell(Variant::Integer(6), None)),
                ((2, 3), cell(Variant::Empty, Some("=Sales[@Amount]*2"))),
            ]),
        )]);
        let structured = HashMap::from([(
            String::from("sales[@amount]"),
            String::from("@sheet:Sheet1;rows:2-3|Sheet1!B{row}"),
        )]);
        let names = HashMap::new();
        assert!(
            recalculate(
                &mut sheets,
                &names,
                &HashMap::new(),
                &structured,
                None,
                true
            )
            .unwrap()
        );
        assert_eq!(sheets["Sheet1"][&(2, 3)].value, Variant::Integer(8));
    }

    #[test]
    fn leaves_this_row_reference_unchanged_outside_the_table_data_rows() {
        let structured = HashMap::from([(
            String::from("sales[@amount]"),
            String::from("@sheet:Sheet1;rows:2-3|Sheet1!B{row}"),
        )]);
        assert_eq!(
            normalize_structured_refs("=Sales[@Amount]", &structured, "Sheet1", 1),
            "=Sales[@Amount]"
        );
    }

    #[test]
    fn normalizes_a_multi_column_structured_reference() {
        let mut sheets = HashMap::from([
            (
                "Sheet1".to_string(),
                HashMap::from([
                    ((2, 2), cell(Variant::Integer(4), None)),
                    ((2, 3), cell(Variant::Integer(6), None)),
                    ((3, 2), cell(Variant::Integer(1), None)),
                    ((3, 3), cell(Variant::Integer(2), None)),
                ]),
            ),
            (
                "Sheet2".to_string(),
                HashMap::from([(
                    (1, 1),
                    cell(Variant::Empty, Some("=SUM(Sales[[Amount]:[Tax]])")),
                )]),
            ),
        ]);
        let structured = HashMap::from([(
            String::from("sales[[amount]:[tax]]"),
            String::from("elixceetablesalesamounttaxrange"),
        )]);
        let names = HashMap::from([(
            String::from("elixceetablesalesamounttaxrange"),
            String::from("Sheet1!B2:C3"),
        )]);
        assert!(
            recalculate(
                &mut sheets,
                &names,
                &HashMap::new(),
                &structured,
                None,
                true
            )
            .unwrap()
        );
        assert_eq!(sheets["Sheet2"][&(1, 1)].value, Variant::Integer(13));
    }

    #[test]
    fn range_dependencies_order_formula_nodes_without_expanding_the_range() {
        let mut sheets = HashMap::from([
            (
                "Sheet1".to_string(),
                HashMap::from([
                    ((1, 1), cell(Variant::Integer(3), None)),
                    ((2, 1), cell(Variant::Empty, Some("=A1+1"))),
                ]),
            ),
            (
                "Sheet2".to_string(),
                HashMap::from([((1, 1), cell(Variant::Empty, Some("=SUM(Sheet1!A1:A2)")))]),
            ),
        ]);
        assert!(
            recalculate(
                &mut sheets,
                &HashMap::new(),
                &HashMap::new(),
                &HashMap::new(),
                None,
                true
            )
            .unwrap()
        );
        assert_eq!(sheets["Sheet1"][&(2, 1)].value, Variant::Integer(4));
        assert_eq!(sheets["Sheet2"][&(1, 1)].value, Variant::Integer(7));
    }
}
