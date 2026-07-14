//! Static analysis for the `elixcee check` subcommand: inspects a `.bas` file
//! without executing it. Deliberately narrow scope — see `tasks/todo.md`
//! Milestone B1/B1.1 for what's included here vs. still deferred and why.

use std::collections::HashSet;

use crate::diagnostics::{SourceLocation, json_string, locate};
use crate::parser::{self, CaseMatch, Expr, Program, SourceSpan, SpannedStmt, Stmt};
use crate::vm;

/// One static-analysis finding. `severity` "error" means the file can't (or
/// almost certainly won't) run correctly; "info" is a heads-up that doesn't
/// mean anything is broken (e.g. a macro that shows a MsgBox just isn't
/// fully headless).
pub struct Diagnostic {
    pub severity: &'static str,
    pub code: &'static str,
    pub kind: &'static str,
    pub message: String,
    pub location: Option<SourceLocation>,
}

/// Run every check this command currently supports against `source`.
/// `macro_name` is optional — pass `None` to check the file on its own
/// without asserting a particular entrypoint exists.
pub fn run_check(source: &str, file: &str, macro_name: Option<&str>) -> Vec<Diagnostic> {
    run_check_impl(source, file, macro_name, &HashSet::new())
}

/// Like `run_check`, but for one module within a multi-module project
/// (Milestone B2): `other_module_names` is every bare Sub/Function name
/// declared in *other* modules of the same project, so an unqualified call
/// to a name defined elsewhere in the project isn't misreported as
/// undefined (`is_resolvable` only ever sees this one module's own
/// `Program` otherwise). Does not itself check for cross-module name
/// collisions — see `parser::find_cross_module_sub_collisions`/
/// `find_cross_module_func_collisions`, surfaced separately by the caller.
pub fn run_check_in_project(
    source: &str,
    file: &str,
    macro_name: Option<&str>,
    other_module_names: &HashSet<String>,
) -> Vec<Diagnostic> {
    run_check_impl(source, file, macro_name, other_module_names)
}

fn run_check_impl(
    source: &str,
    file: &str,
    macro_name: Option<&str>,
    other_module_names: &HashSet<String>,
) -> Vec<Diagnostic> {
    let prog = match parser::parse_with_span(source) {
        Ok(prog) => prog,
        Err(e) => {
            let location = locate(source, file, e.span);
            return vec![Diagnostic {
                severity: "error",
                code: "E2001",
                kind: "parse_error",
                message: e.message,
                location: Some(location),
            }];
        }
    };

    let mut diags = Vec::new();

    if let Some(name) = macro_name {
        // Mirrors Vm::run_sub's exact lookup: the tokenizer lowercases every
        // identifier at parse time, so `SubDef.name` is always lowercase —
        // comparing against `name.to_lowercase()` reproduces run_sub's
        // case-insensitive match precisely (not just approximately).
        let found = prog.subs.iter().any(|s| s.name == name.to_lowercase());
        if !found {
            diags.push(Diagnostic {
                severity: "error",
                code: "E1002",
                kind: "undefined_sub_or_function",
                message: format!("Sub '{}' not found", name),
                location: None,
            });
        }
    }

    for (reason, span) in &prog.module_diagnostics {
        diags.push(Diagnostic {
            severity: "info",
            code: "I1002",
            kind: "unsupported_construct",
            message: reason.clone(),
            location: Some(locate(source, file, *span)),
        });
    }

    for sub in &prog.subs {
        let local_names = local_scope_names(&sub.name, &sub.params, &sub.body);
        walk_body(
            &sub.body,
            &prog,
            &local_names,
            other_module_names,
            source,
            file,
            &mut diags,
        );
    }
    for func in &prog.funcs {
        let local_names = local_scope_names(&func.name, &func.params, &func.body);
        walk_body(
            &func.body,
            &prog,
            &local_names,
            other_module_names,
            source,
            file,
            &mut diags,
        );
    }

    diags
}

/// Every name in scope for one Sub/Function: its own name (for recursion),
/// its parameters, and every variable/array/record name declared or
/// assigned anywhere in its body (VBA scoping is procedure-level, not
/// block-level, so a name introduced inside an `If`/`For` is visible for
/// the whole procedure — this collects across all nesting, not just the
/// top level).
fn local_scope_names(own_name: &str, params: &[String], body: &[SpannedStmt]) -> HashSet<String> {
    let mut names: HashSet<String> = params.iter().cloned().collect();
    names.insert(own_name.to_string());
    collect_declared_names(body, &mut names);
    names
}

/// `true` iff `name` resolves to an in-scope variable/array/record, a user
/// Sub, a user Function, or a built-in VBA/WorksheetFunction name — the
/// same places `Vm::run_sub`'s call resolution consults at runtime
/// (`src/vm/mod.rs`), checked here without executing anything. `name` is
/// always already lowercase by the time it reaches here (the tokenizer
/// lowercases every identifier), matching how `prog.subs`/`prog.funcs`/
/// declared names are keyed — no case conversion needed.
///
/// The variable check matters because `arr(i)` and `func(i)` are
/// syntactically identical in this AST (both `Expr::FuncCall` — there's no
/// separate "array index" expression variant), so an indexed read of any
/// local array/variable would otherwise be misreported as a call to an
/// undefined function.
///
/// `other_module_names` is every bare Sub/Function name declared in *other*
/// modules of the same project (empty for a single-file check) — without
/// it, a legitimate unqualified cross-module call would be misreported as
/// undefined, since this function otherwise only sees `prog`'s own module.
fn is_resolvable(
    name: &str,
    prog: &Program,
    local_names: &HashSet<String>,
    other_module_names: &HashSet<String>,
) -> bool {
    local_names.contains(name)
        || prog.subs.iter().any(|s| s.name == name)
        || prog.funcs.iter().any(|f| f.name == name)
        || other_module_names.contains(name)
        || vm::is_known_builtin_function(name)
}

/// Collect every name that a name-introducing statement declares or
/// assigns, recursing into nested bodies. Written as an exhaustive match
/// (no wildcard) — an under-collected name here is a false positive at the
/// call site, exactly the failure mode this feature exists to avoid, so a
/// new `Stmt` variant must be a deliberate decision, not a silent gap.
fn collect_declared_names(body: &[SpannedStmt], names: &mut HashSet<String>) {
    for s in body {
        match &s.stmt {
            Stmt::Assignment { var, .. } => {
                names.insert(var.clone());
            }
            Stmt::CellWrite { .. } => {}
            Stmt::SetCalcMode(_) => {}
            Stmt::SetAppProp { .. } => {}
            Stmt::RangeWrite { .. } => {}
            Stmt::RangeCopy { .. } => {}
            Stmt::RangeClear { .. } => {}
            Stmt::RangeOffsetWrite { .. } => {}
            Stmt::RangeDelete { .. } => {}
            Stmt::RangeInsert { .. } => {}
            Stmt::RangeSort { .. } => {}
            Stmt::RangeName { .. } => {}
            Stmt::SheetCellWrite { .. } => {}
            Stmt::WithSheet { body, .. } => collect_declared_names(body, names),
            Stmt::SheetsAdd => {}
            Stmt::SheetsDelete { .. } => {}
            Stmt::For { var, body, .. } => {
                names.insert(var.clone());
                collect_declared_names(body, names);
            }
            Stmt::ForEach { var, body, .. } => {
                names.insert(var.clone());
                collect_declared_names(body, names);
            }
            Stmt::If {
                then_body,
                else_body,
                ..
            } => {
                collect_declared_names(then_body, names);
                collect_declared_names(else_body, names);
            }
            Stmt::DoLoop { body, .. } => collect_declared_names(body, names),
            Stmt::SelectCase {
                cases, else_body, ..
            } => {
                for (_, case_body) in cases {
                    collect_declared_names(case_body, names);
                }
                collect_declared_names(else_body, names);
            }
            Stmt::ExitFor | Stmt::ExitDo | Stmt::ExitSub | Stmt::ExitFunction => {}
            Stmt::OnError { .. } => {}
            Stmt::OnErrorGoTo(_) => {}
            Stmt::Label(_) => {}
            Stmt::GoTo(_) => {}
            Stmt::Resume { .. } => {}
            Stmt::CallSub { .. } => {}
            Stmt::Dim => {}
            Stmt::DimArray { name, .. } => {
                names.insert(name.clone());
            }
            Stmt::ReDim { name, .. } => {
                names.insert(name.clone());
            }
            Stmt::ArrayWrite { name, .. } => {
                names.insert(name.clone());
            }
            Stmt::With { body } => collect_declared_names(body, names),
            Stmt::MsgBox { .. } => {}
            Stmt::RecordSet { var, .. } => {
                names.insert(var.clone());
            }
            Stmt::DimRecord { var, .. } => {
                names.insert(var.clone());
            }
            Stmt::DimArrayRecord { name, .. } => {
                names.insert(name.clone());
            }
            Stmt::RecordSetNested { var, .. } => {
                names.insert(var.clone());
            }
            Stmt::ArrayRecordSet { name, .. } => {
                names.insert(name.clone());
            }
            Stmt::WithRecord { var, body } => {
                names.insert(var.clone());
                collect_declared_names(body, names);
            }
            Stmt::Unsupported { .. } => {}
        }
    }
}

/// Recursively walk a statement list looking for interactive calls and
/// undefined Sub/Function calls. If you add a new `Stmt` variant with a
/// nested `Vec<SpannedStmt>` body, add a matching arm to the inner `match`
/// below too — its wildcard arm silently skips anything not listed there.
#[allow(clippy::too_many_arguments)]
fn walk_body(
    body: &[SpannedStmt],
    prog: &Program,
    local_names: &HashSet<String>,
    other_module_names: &HashSet<String>,
    source: &str,
    file: &str,
    diags: &mut Vec<Diagnostic>,
) {
    for s in body {
        if let Stmt::MsgBox { .. } = &s.stmt {
            diags.push(Diagnostic {
                severity: "info",
                code: "I1001",
                kind: "interactive_call",
                message: "MsgBox displays a dialog and blocks headless execution".to_string(),
                location: Some(locate(source, file, s.span)),
            });
        }

        if let Stmt::Unsupported { reason } = &s.stmt {
            diags.push(Diagnostic {
                severity: "info",
                code: "I1002",
                kind: "unsupported_construct",
                message: reason.clone(),
                location: Some(locate(source, file, s.span)),
            });
        }

        if let Stmt::CallSub { name, .. } = &s.stmt
            && !is_resolvable(name, prog, local_names, other_module_names)
        {
            diags.push(Diagnostic {
                severity: "error",
                code: "E1002",
                kind: "undefined_sub_or_function",
                message: format!("Sub/Function '{}' not found", name),
                location: Some(locate(source, file, s.span)),
            });
        }

        // Walk every expression reachable from this statement (assignment
        // values, cell indices, condition expressions, etc.) — not just
        // nested statement bodies — looking for undefined FuncCall targets
        // buried anywhere inside them.
        let mut exprs = Vec::new();
        collect_stmt_exprs(&s.stmt, &mut exprs);
        for e in exprs {
            walk_expr(
                e,
                prog,
                local_names,
                other_module_names,
                s.span,
                source,
                file,
                diags,
            );
        }

        match &s.stmt {
            Stmt::WithSheet { body, .. }
            | Stmt::For { body, .. }
            | Stmt::ForEach { body, .. }
            | Stmt::DoLoop { body, .. }
            | Stmt::With { body }
            | Stmt::WithRecord { body, .. } => walk_body(
                body,
                prog,
                local_names,
                other_module_names,
                source,
                file,
                diags,
            ),
            Stmt::If {
                then_body,
                else_body,
                ..
            } => {
                walk_body(
                    then_body,
                    prog,
                    local_names,
                    other_module_names,
                    source,
                    file,
                    diags,
                );
                walk_body(
                    else_body,
                    prog,
                    local_names,
                    other_module_names,
                    source,
                    file,
                    diags,
                );
            }
            Stmt::SelectCase {
                cases, else_body, ..
            } => {
                for (_, case_body) in cases {
                    walk_body(
                        case_body,
                        prog,
                        local_names,
                        other_module_names,
                        source,
                        file,
                        diags,
                    );
                }
                walk_body(
                    else_body,
                    prog,
                    local_names,
                    other_module_names,
                    source,
                    file,
                    diags,
                );
            }
            _ => {}
        }
    }
}

/// Push every `Expr` directly reachable from `stmt` (assignment values, cell
/// indices, condition expressions, `Select Case` match arms, etc.) into
/// `out` — nested statement bodies are walked separately by `walk_body`, not
/// here. Written as an exhaustive match (no wildcard arm) so adding a new
/// `Stmt` variant forces a decision about what expressions it carries,
/// instead of silently under-checking it.
fn collect_stmt_exprs<'a>(stmt: &'a Stmt, out: &mut Vec<&'a Expr>) {
    match stmt {
        Stmt::Assignment { value, .. } => out.push(value),
        Stmt::CellWrite { row, col, value } => {
            out.push(row);
            out.push(col);
            out.push(value);
        }
        Stmt::SetCalcMode(_) => {}
        Stmt::SetAppProp { value, .. } => out.push(value),
        Stmt::RangeWrite { value, .. } => out.push(value),
        Stmt::RangeCopy { .. } => {}
        Stmt::RangeClear { .. } => {}
        Stmt::RangeOffsetWrite {
            row_off,
            col_off,
            value,
            ..
        } => {
            out.push(row_off);
            out.push(col_off);
            out.push(value);
        }
        Stmt::RangeDelete { .. } => {}
        Stmt::RangeInsert { .. } => {}
        Stmt::RangeSort { .. } => {}
        Stmt::RangeName { .. } => {}
        Stmt::SheetCellWrite {
            sheet,
            row,
            col,
            value,
        } => {
            out.push(sheet);
            out.push(row);
            out.push(col);
            out.push(value);
        }
        Stmt::WithSheet { .. } => {}
        Stmt::SheetsAdd => {}
        Stmt::SheetsDelete { sheet } => out.push(sheet),
        Stmt::For { from, to, step, .. } => {
            out.push(from);
            out.push(to);
            if let Some(s) = step {
                out.push(s);
            }
        }
        Stmt::ForEach { .. } => {}
        Stmt::If { condition, .. } => out.push(condition),
        Stmt::DoLoop {
            pre_cond,
            post_cond,
            ..
        } => {
            if let Some((_, e)) = pre_cond {
                out.push(e);
            }
            if let Some((_, e)) = post_cond {
                out.push(e);
            }
        }
        Stmt::SelectCase { expr, cases, .. } => {
            out.push(expr);
            for (matches, _) in cases {
                for m in matches {
                    match m {
                        CaseMatch::Value(e) => out.push(e),
                        CaseMatch::Range(a, b) => {
                            out.push(a);
                            out.push(b);
                        }
                        CaseMatch::IsOp(_, e) => out.push(e),
                    }
                }
            }
        }
        Stmt::ExitFor | Stmt::ExitDo | Stmt::ExitSub | Stmt::ExitFunction => {}
        Stmt::OnError { .. } => {}
        Stmt::OnErrorGoTo(_) => {}
        Stmt::Label(_) => {}
        Stmt::GoTo(_) => {}
        Stmt::Resume { .. } => {}
        Stmt::CallSub { args, .. } => {
            for a in args {
                out.push(a);
            }
        }
        Stmt::Dim => {}
        Stmt::DimArray { sizes, .. } => {
            for s in sizes {
                out.push(s);
            }
        }
        Stmt::ReDim { sizes, .. } => {
            for s in sizes {
                out.push(s);
            }
        }
        Stmt::ArrayWrite { indices, value, .. } => {
            for i in indices {
                out.push(i);
            }
            out.push(value);
        }
        Stmt::With { .. } => {}
        Stmt::MsgBox { message } => out.push(message),
        Stmt::RecordSet { value, .. } => out.push(value),
        Stmt::DimRecord { .. } => {}
        Stmt::DimArrayRecord { sizes, .. } => {
            for s in sizes {
                out.push(s);
            }
        }
        Stmt::RecordSetNested { value, .. } => out.push(value),
        Stmt::ArrayRecordSet { indices, value, .. } => {
            for i in indices {
                out.push(i);
            }
            out.push(value);
        }
        Stmt::WithRecord { .. } => {}
        Stmt::Unsupported { .. } => {}
    }
}

/// Recursively walk an expression looking for calls to undefined
/// Sub/Function names, attributing any finding to `stmt_span` — expressions
/// don't carry their own span (Milestone A.5's statement-level granularity
/// decision: `location` points at the enclosing statement, not the exact
/// sub-expression, same as runtime error locations do).
#[allow(clippy::too_many_arguments)]
fn walk_expr(
    expr: &Expr,
    prog: &Program,
    local_names: &HashSet<String>,
    other_module_names: &HashSet<String>,
    stmt_span: SourceSpan,
    source: &str,
    file: &str,
    diags: &mut Vec<Diagnostic>,
) {
    match expr {
        Expr::FuncCall { name, args } => {
            if !is_resolvable(name, prog, local_names, other_module_names) {
                diags.push(Diagnostic {
                    severity: "error",
                    code: "E1002",
                    kind: "undefined_sub_or_function",
                    message: format!("Unknown VBA function: '{}'", name),
                    location: Some(locate(source, file, stmt_span)),
                });
            }
            for a in args {
                walk_expr(
                    a,
                    prog,
                    local_names,
                    other_module_names,
                    stmt_span,
                    source,
                    file,
                    diags,
                );
            }
        }
        Expr::BinOp { lhs, rhs, .. } => {
            walk_expr(
                lhs,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
            walk_expr(
                rhs,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
        }
        Expr::UnaryMinus(e) | Expr::UnaryNot(e) => walk_expr(
            e,
            prog,
            local_names,
            other_module_names,
            stmt_span,
            source,
            file,
            diags,
        ),
        Expr::CellRead { row, col } => {
            walk_expr(
                row,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
            walk_expr(
                col,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
        }
        Expr::RangeOffsetRead {
            row_off, col_off, ..
        } => {
            walk_expr(
                row_off,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
            walk_expr(
                col_off,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
        }
        Expr::CellsFind { what, .. } => walk_expr(
            what,
            prog,
            local_names,
            other_module_names,
            stmt_span,
            source,
            file,
            diags,
        ),
        Expr::SheetCellRead { sheet, row, col } => {
            walk_expr(
                sheet,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
            walk_expr(
                row,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
            walk_expr(
                col,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
        }
        Expr::CellsEndProp { row, col, .. } => {
            walk_expr(
                row,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
            walk_expr(
                col,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
        }
        Expr::ArrayRecordGet { indices, .. } => {
            for i in indices {
                walk_expr(
                    i,
                    prog,
                    local_names,
                    other_module_names,
                    stmt_span,
                    source,
                    file,
                    diags,
                );
            }
        }
        Expr::Integer(_)
        | Expr::Float(_)
        | Expr::Str(_)
        | Expr::Bool(_)
        | Expr::Var(_)
        | Expr::RangeRead { .. }
        | Expr::RowsCount
        | Expr::ColsCount
        | Expr::RecordGet { .. }
        | Expr::RecordGetNested { .. } => {}
    }
}

/// `true` iff no diagnostic has `severity == "error"`.
pub fn all_ok(diags: &[Diagnostic]) -> bool {
    !diags.iter().any(|d| d.severity == "error")
}

/// `{"schema_version":1,"ok":...,"diagnostics":[...]}` — the `check`
/// subcommand's own JSON shape (distinct from the run-mode success/error
/// shape in `src/diagnostics.rs`, since `check` reports a batch of findings
/// rather than one result).
pub fn diagnostics_to_json(diags: &[Diagnostic]) -> String {
    let items: Vec<String> = diags.iter().map(diagnostic_to_json).collect();
    format!(
        "{{\"schema_version\":1,\"ok\":{},\"diagnostics\":[{}]}}",
        all_ok(diags),
        items.join(","),
    )
}

fn diagnostic_to_json(d: &Diagnostic) -> String {
    let location_json = match &d.location {
        Some(loc) => format!(
            "{{\"file\":{},\"line\":{},\"column\":{}}}",
            json_string(&loc.file),
            loc.line,
            loc.column,
        ),
        None => "null".to_string(),
    };
    format!(
        "{{\"severity\":{},\"code\":{},\"kind\":{},\"message\":{},\"location\":{}}}",
        json_string(d.severity),
        json_string(d.code),
        json_string(d.kind),
        json_string(&d.message),
        location_json,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    fn codes(diags: &[Diagnostic]) -> Vec<&str> {
        diags.iter().map(|d| d.code).collect()
    }

    #[test]
    fn clean_program_has_no_diagnostics() {
        let diags = run_check(
            "Sub Main()\n    Cells(1, 1).Value = 1\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty());
        assert!(all_ok(&diags));
    }

    #[test]
    fn parse_error_short_circuits_everything_else() {
        let diags = run_check("Sub Main(\n    x = 1\n", "f.bas", Some("Main"));
        assert_eq!(codes(&diags), vec!["E2001"]);
        assert_eq!(diags[0].severity, "error");
        assert!(diags[0].location.is_some());
        assert!(!all_ok(&diags));
    }

    #[test]
    fn missing_entrypoint_is_reported() {
        let diags = run_check(
            "Sub Main()\n    x = 1\nEnd Sub\n",
            "f.bas",
            Some("DoesNotExist"),
        );
        assert_eq!(codes(&diags), vec!["E1002"]);
        assert_eq!(diags[0].kind, "undefined_sub_or_function");
        assert!(diags[0].location.is_none());
    }

    #[test]
    fn entrypoint_check_is_case_insensitive() {
        let diags = run_check("Sub Main()\n    x = 1\nEnd Sub\n", "f.bas", Some("MAIN"));
        assert!(diags.is_empty());
    }

    #[test]
    fn no_macro_name_skips_entrypoint_check() {
        let diags = run_check("Sub Main()\n    x = 1\nEnd Sub\n", "f.bas", None);
        assert!(diags.is_empty());
    }

    #[test]
    fn top_level_msgbox_is_detected() {
        let diags = run_check(
            "Sub Main()\n    MsgBox \"hi\"\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1001"]);
        assert_eq!(diags[0].severity, "info");
        assert_eq!(diags[0].location.as_ref().unwrap().line, 2);
        assert!(all_ok(&diags)); // info-only, still "ok"
    }

    #[test]
    fn msgbox_nested_inside_if_and_for_is_detected() {
        let diags = run_check(
            "Sub Main()\n\
             \x20   For i = 1 To 3\n\
             \x20       If i = 2 Then\n\
             \x20           MsgBox \"two\"\n\
             \x20       End If\n\
             \x20   Next i\n\
             End Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1001"]);
    }

    #[test]
    fn multiple_msgbox_calls_are_all_reported_in_order() {
        let diags = run_check(
            "Sub Main()\n    MsgBox \"one\"\n    MsgBox \"two\"\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1001", "I1001"]);
        assert_eq!(diags[0].location.as_ref().unwrap().line, 2);
        assert_eq!(diags[1].location.as_ref().unwrap().line, 3);
    }

    #[test]
    fn json_shape_round_trips_severity_and_ok() {
        let diags = run_check(
            "Sub Main()\n    MsgBox \"hi\"\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        let json = diagnostics_to_json(&diags);
        assert!(json.contains("\"ok\":true"));
        assert!(json.contains("\"severity\":\"info\""));
        assert!(json.contains("\"code\":\"I1001\""));
    }

    #[test]
    fn json_shape_reports_ok_false_when_an_error_is_present() {
        let diags = run_check("Sub Main()\n    x = 1\nEnd Sub\n", "f.bas", Some("Nope"));
        let json = diagnostics_to_json(&diags);
        assert!(json.contains("\"ok\":false"));
    }

    // ── undefined Sub/Function call detection (B1.1) ────────────────────────

    #[test]
    fn undefined_callsub_target_is_reported() {
        let diags = run_check(
            "Sub Main()\n    Call Bogus(1)\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1002"]);
        assert_eq!(diags[0].message, "Sub/Function 'bogus' not found");
        assert_eq!(diags[0].location.as_ref().unwrap().line, 2);
    }

    #[test]
    fn undefined_bare_call_target_is_reported() {
        // `Bogus()` without the `Call` keyword goes through a different
        // CallSub construction site (parse_ident_stmt) than `Call Bogus(1)`
        // (parse_call_stmt) — this parser doesn't support the paren-less
        // `Bogus 1` space-separated-args form (it parses as a no-op), so
        // the parenthesized bare form is what actually exercises that path.
        let diags = run_check("Sub Main()\n    Bogus()\nEnd Sub\n", "f.bas", Some("Main"));
        assert_eq!(codes(&diags), vec!["E1002"]);
    }

    #[test]
    fn undefined_funccall_target_at_top_level_is_reported() {
        let diags = run_check(
            "Sub Main()\n    x = Bogus(1)\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1002"]);
        assert_eq!(diags[0].message, "Unknown VBA function: 'bogus'");
    }

    #[test]
    fn undefined_call_nested_inside_an_expression_uses_the_statement_location() {
        let diags = run_check(
            "Sub Main()\n    x = 1 + Bogus(2)\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1002"]);
        // No expression-level spans (Milestone A.5 decision) — location is
        // the enclosing statement's line, same as runtime error locations.
        assert_eq!(diags[0].location.as_ref().unwrap().line, 2);
    }

    #[test]
    fn undefined_call_nested_inside_a_cells_index_is_reported() {
        let diags = run_check(
            "Sub Main()\n    Cells(Bogus(1), 2).Value = 1\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1002"]);
    }

    #[test]
    fn undefined_call_inside_select_case_condition_is_reported() {
        let diags = run_check(
            "Sub Main()\n\
             \x20   Select Case Bogus(1)\n\
             \x20       Case 1\n\
             \x20   End Select\n\
             End Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1002"]);
    }

    #[test]
    fn calling_a_real_user_sub_is_not_flagged() {
        let diags = run_check(
            "Sub Main()\n    Call Helper(1)\nEnd Sub\n\
             Sub Helper(x)\n    y = x\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty());
    }

    #[test]
    fn calling_a_real_user_function_is_not_flagged() {
        let diags = run_check(
            "Sub Main()\n    x = Helper(1)\nEnd Sub\n\
             Function Helper(n)\n    Helper = n\nEnd Function\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty());
    }

    #[test]
    fn builtin_vba_function_call_is_not_flagged() {
        let diags = run_check(
            "Sub Main()\n    x = Len(\"hi\")\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty());
    }

    #[test]
    fn worksheet_function_call_is_not_flagged() {
        let diags = run_check(
            "Sub Main()\n    x = WorksheetFunction.Sum(Range(\"A1:A2\"))\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty());
    }

    #[test]
    fn application_worksheet_function_call_is_not_flagged() {
        let diags = run_check(
            "Sub Main()\n    x = Application.WorksheetFunction.Sum(Range(\"A1:A2\"))\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty());
    }

    #[test]
    fn user_function_shadowing_a_builtin_name_is_not_flagged() {
        let diags = run_check(
            "Sub Main()\n    x = Len(1)\nEnd Sub\n\
             Function Len(n)\n    Len = n\nEnd Function\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty());
    }

    // ── array/variable reads must not be misread as undefined calls ────────
    // `arr(i)` and `func(i)` are syntactically identical in this AST (both
    // `Expr::FuncCall` — there's no separate array-index expression), so an
    // indexed read of a local variable is exactly the false-positive shape
    // this feature must never produce.

    #[test]
    fn indexing_a_split_result_is_not_flagged() {
        let diags = run_check(
            "Sub Main()\n    parts = Split(\"a,b,c\", \",\")\n    x = parts(0)\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty(), "{:?}", diags[0].message);
    }

    #[test]
    fn indexing_a_dim_array_is_not_flagged() {
        let diags = run_check(
            "Sub Main()\n    Dim arr(10)\n    arr(0) = 1\n    x = arr(0)\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty(), "{:?}", diags[0].message);
    }

    #[test]
    fn indexing_a_function_parameter_array_is_not_flagged() {
        let diags = run_check(
            "Sub Main()\n    x = Helper(1)\nEnd Sub\n\
             Function Helper(arr)\n    Helper = arr(0)\nEnd Function\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty(), "{:?}", diags[0].message);
    }

    #[test]
    fn a_genuinely_undefined_call_is_still_reported_alongside_a_real_array() {
        // Guards against the fix becoming so permissive it stops detecting
        // real typos once a program also happens to use arrays.
        let diags = run_check(
            "Sub Main()\n    Dim arr(10)\n    arr(0) = 1\n    x = Bogus(arr(0))\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1002"]);
        assert_eq!(diags[0].message, "Unknown VBA function: 'bogus'");
    }

    // ── unsupported-construct detection (I1002) ─────────────────────────────

    #[test]
    fn debug_print_is_an_unsupported_construct_diagnostic() {
        let diags = run_check(
            "Sub Main()\n    Debug.Print \"hi\"\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1002"]);
        assert_eq!(diags[0].severity, "info");
        assert_eq!(diags[0].kind, "unsupported_construct");
        assert!(
            diags[0].message.contains("Debug.Print"),
            "{:?}",
            diags[0].message
        );
        assert_eq!(diags[0].location.as_ref().unwrap().line, 2);
    }

    #[test]
    fn unrecognized_range_property_is_an_unsupported_construct_diagnostic() {
        let diags = run_check(
            "Sub Main()\n    Range(\"A1\").NumberFormat = \"0.00\"\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1002"]);
        assert!(
            diags[0].message.contains("numberformat"),
            "{:?}",
            diags[0].message
        );
    }

    #[test]
    fn unrecognized_sheets_method_is_an_unsupported_construct_diagnostic() {
        let diags = run_check(
            "Sub Main()\n    Sheets.Foo\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1002"]);
        assert!(
            diags[0].message.contains("Sheets.foo"),
            "{:?}",
            diags[0].message
        );
    }

    #[test]
    fn bare_statement_call_is_an_unsupported_construct_diagnostic() {
        let diags = run_check("Sub Main()\n    Foo\nEnd Sub\n", "f.bas", Some("Main"));
        assert_eq!(codes(&diags), vec!["I1002"]);
        assert!(diags[0].message.contains("'foo'"), "{:?}", diags[0].message);
    }

    #[test]
    fn unsupported_construct_nested_inside_if_and_for_is_detected() {
        let diags = run_check(
            "Sub Main()\n\
             \x20   For i = 1 To 3\n\
             \x20       If i = 2 Then\n\
             \x20           Debug.Print \"two\"\n\
             \x20       End If\n\
             \x20   Next i\n\
             End Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1002"]);
    }

    #[test]
    fn unsupported_construct_alone_is_still_ok() {
        // Info-only finding — the macro still runs to completion, this is
        // a heads-up, not a failure (mirrors the MsgBox info-only test).
        let diags = run_check(
            "Sub Main()\n    Debug.Print \"hi\"\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert!(all_ok(&diags));
    }

    #[test]
    fn unsupported_construct_coexists_with_a_real_error() {
        let diags = run_check(
            "Sub Main()\n    Debug.Print \"hi\"\n    x = Bogus(1)\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1002", "E1002"]);
        assert!(!all_ok(&diags));
    }

    #[test]
    fn module_level_const_with_modifier_is_an_unsupported_construct_diagnostic() {
        let diags = run_check(
            "Public Const MAX_RETRIES = 5\nSub Main()\n    a = 1\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1002"]);
        assert_eq!(diags[0].severity, "info");
        assert!(diags[0].message.contains("Const"), "{:?}", diags[0].message);
        assert_eq!(diags[0].location.as_ref().unwrap().line, 1);
        assert!(all_ok(&diags));
    }

    #[test]
    fn module_level_unrecognized_line_is_an_unsupported_construct_diagnostic() {
        let diags = run_check(
            "Declare Function Foo Lib \"x.dll\" ()\nSub Main()\n    a = 1\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1002"]);
        assert!(
            diags[0].message.contains("declare"),
            "{:?}",
            diags[0].message
        );
    }

    #[test]
    fn module_level_plain_declaration_is_not_flagged() {
        // Group A parity: no separate module scope exists at runtime, so a
        // plain `Public x`/`Dim x` with no value is a harmless no-op, same
        // as the already-excluded Sub-level case — not a gap worth a
        // diagnostic.
        let diags = run_check(
            "Public x As Long\nSub Main()\n    x = 1\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty());
    }

    #[test]
    fn unsupported_construct_nested_inside_with_record_is_detected() {
        let diags = run_check(
            "Sub Main()\n    With p\n        .Field\n    End With\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["I1002"]);
        assert!(
            diags[0].message.contains("p.field"),
            "{:?}",
            diags[0].message
        );
    }

    // ── run_check_in_project (Milestone B2) ─────────────────────────────────

    #[test]
    fn cross_module_call_is_not_flagged_when_other_module_names_are_given() {
        let mut others = HashSet::new();
        others.insert("helper".to_string());
        let diags = run_check_in_project(
            "Sub Main()\n    Call Helper()\nEnd Sub\n",
            "module2.bas",
            Some("Main"),
            &others,
        );
        assert!(
            diags.is_empty(),
            "expected no diagnostics, got {} entries",
            diags.len()
        );
    }

    #[test]
    fn genuinely_undefined_call_is_still_flagged_in_project_mode() {
        let others = HashSet::new();
        let diags = run_check_in_project(
            "Sub Main()\n    Call Bogus()\nEnd Sub\n",
            "module2.bas",
            Some("Main"),
            &others,
        );
        assert_eq!(codes(&diags), vec!["E1002"]);
    }

    #[test]
    fn run_check_without_project_context_still_flags_a_cross_module_name() {
        // Sanity check that run_check (the single-file wrapper) is
        // unaffected by the new project-aware path — a call this module
        // can't see is still (correctly, for a single-file check) flagged.
        let diags = run_check(
            "Sub Main()\n    Call Helper()\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1002"]);
    }
}
