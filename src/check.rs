//! Static analysis for the `elixcee check` subcommand: inspects a `.bas` file
//! without executing it. Deliberately narrow scope — see `tasks/todo.md`
//! Milestone B1/B1.1 for what's included here vs. still deferred and why.

use std::collections::HashSet;

use crate::diagnostics::{SourceLocation, json_string, locate};
use crate::parser::{
    self, CaseMatch, Expr, Program, SourceSpan, SpannedStmt, Stmt, WithMember, WithTarget,
};
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

    collect_extra_compile_diagnostics(&prog, source, file, &mut diags);

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
            Stmt::RangeObjectCopy { .. } => {}
            Stmt::Set { var, .. } => {
                names.insert(var.clone());
            }
            Stmt::RangePaste { .. } => {}
            Stmt::SheetRangePaste { .. } => {}
            Stmt::SheetProtection { .. } => {}
            Stmt::RangeClear { .. } => {}
            Stmt::RangeOffsetWrite { .. } => {}
            Stmt::RangeDelete { .. } => {}
            Stmt::RangeInsert { .. } => {}
            Stmt::RowColDelete { .. } => {}
            Stmt::RowColInsert { .. } => {}
            Stmt::RangeSort { .. } => {}
            Stmt::RangeAutoFilter { .. } => {}
            Stmt::RangeName { .. } => {}
            Stmt::SheetCellWrite { .. } => {}
            Stmt::SheetRangeWrite { .. } => {}
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
            Stmt::DimBare { var } => {
                names.insert(var.clone());
            }
            Stmt::DimArray { name, .. } => {
                names.insert(name.clone());
            }
            Stmt::ReDim { name, .. } => {
                names.insert(name.clone());
            }
            Stmt::ArrayWrite { name, .. } => {
                names.insert(name.clone());
            }
            Stmt::Erase { .. } => {}
            Stmt::With { target, body } => {
                // A bare-identifier With target names a variable (a UDT
                // record, or a Set-assigned object) — it must stay a
                // "declared name", exactly as the old `WithRecord` arm did,
                // or `check` starts reporting it as undefined.
                if let WithTarget::Var(var) = target {
                    names.insert(var.clone());
                }
                collect_declared_names(body, names);
            }
            Stmt::WithDot { .. } => {}
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
            Stmt::DimMulti(decls) => {
                for d in decls {
                    match d {
                        Stmt::DimRecord { var, .. } | Stmt::DimBare { var } => {
                            names.insert(var.clone());
                        }
                        Stmt::DimArray { name, .. } | Stmt::DimArrayRecord { name, .. } => {
                            names.insert(name.clone());
                        }
                        _ => {}
                    }
                }
            }
            Stmt::RecordSetNested { var, .. } => {
                names.insert(var.clone());
            }
            Stmt::ArrayRecordSet { name, .. } => {
                names.insert(name.clone());
            }
            Stmt::Unsupported { .. } => {}
            Stmt::ErrClear => {}
            Stmt::ErrRaise { .. } => {}
        }
    }
}

/// Every nested `Vec<SpannedStmt>` body directly inside `stmt`, in
/// declaration order — the single source of truth for "which `Stmt`
/// variants carry a nested body", shared by every walker in this module
/// that must not silently under-recurse: `collect_labels`,
/// `check_body_for_compile_errors`, and `collect_extra_compile_diagnostics_body`.
/// (`collect_declared_names`/`walk_body`/`walk_expr` predate this helper and
/// keep their own separate exhaustive matches — not worth the churn of
/// switching already-stable, already-tested code to it.) Exhaustive over
/// every `Stmt` variant (no wildcard): `compile_check_errors` treats a
/// `GoTo`/`On Error GoTo` target `collect_labels` fails to find as a
/// *pre-flight, uncatchable* error that blocks the whole run, so under-
/// recursing here wouldn't just misreport a diagnostic, it would stop a
/// legitimate program from running at all. A new `Stmt` variant with a
/// nested body must be a deliberate addition here, not a silent gap.
fn nested_bodies(stmt: &Stmt) -> Vec<&[SpannedStmt]> {
    match stmt {
        Stmt::WithSheet { body, .. }
        | Stmt::For { body, .. }
        | Stmt::ForEach { body, .. }
        | Stmt::DoLoop { body, .. }
        | Stmt::With { body, .. } => vec![body],
        Stmt::If {
            then_body,
            else_body,
            ..
        } => vec![then_body, else_body],
        Stmt::SelectCase {
            cases, else_body, ..
        } => {
            let mut bodies: Vec<&[SpannedStmt]> = cases.iter().map(|(_, b)| b.as_slice()).collect();
            bodies.push(else_body);
            bodies
        }
        Stmt::Assignment { .. }
        | Stmt::CellWrite { .. }
        | Stmt::SetCalcMode(_)
        | Stmt::SetAppProp { .. }
        | Stmt::RangeWrite { .. }
        | Stmt::RangeCopy { .. }
        | Stmt::RangeObjectCopy { .. }
        | Stmt::Set { .. }
        | Stmt::RangePaste { .. }
        | Stmt::SheetRangePaste { .. }
        | Stmt::SheetProtection { .. }
        | Stmt::RangeClear { .. }
        | Stmt::RangeOffsetWrite { .. }
        | Stmt::RangeDelete { .. }
        | Stmt::RangeInsert { .. }
        | Stmt::RowColDelete { .. }
        | Stmt::RowColInsert { .. }
        | Stmt::RangeSort { .. }
        | Stmt::RangeAutoFilter { .. }
        | Stmt::RangeName { .. }
        | Stmt::SheetCellWrite { .. }
        | Stmt::SheetRangeWrite { .. }
        | Stmt::SheetsAdd
        | Stmt::SheetsDelete { .. }
        | Stmt::ExitFor
        | Stmt::ExitDo
        | Stmt::ExitSub
        | Stmt::ExitFunction
        | Stmt::OnError { .. }
        | Stmt::OnErrorGoTo(_)
        | Stmt::GoTo(_)
        | Stmt::Label(_)
        | Stmt::Resume { .. }
        | Stmt::CallSub { .. }
        | Stmt::Dim
        | Stmt::DimBare { .. }
        | Stmt::DimArray { .. }
        | Stmt::ReDim { .. }
        | Stmt::ArrayWrite { .. }
        | Stmt::Erase { .. }
        | Stmt::WithDot { .. }
        | Stmt::MsgBox { .. }
        | Stmt::RecordSet { .. }
        | Stmt::DimRecord { .. }
        | Stmt::DimArrayRecord { .. }
        | Stmt::DimMulti(_)
        | Stmt::RecordSetNested { .. }
        | Stmt::ArrayRecordSet { .. }
        | Stmt::Unsupported { .. }
        | Stmt::ErrClear
        | Stmt::ErrRaise { .. } => vec![],
    }
}

/// Every label declared anywhere in a Sub/Function body (VBA `GoTo`/`On
/// Error GoTo` scope is the whole procedure, not just the current block —
/// same reasoning as `local_scope_names` collecting across all nesting, not
/// just the top level).
fn collect_labels(body: &[SpannedStmt], labels: &mut HashSet<String>) {
    for s in body {
        if let Stmt::Label(name) = &s.stmt {
            labels.insert(name.clone());
        }
        for nested in nested_bodies(&s.stmt) {
            collect_labels(nested, labels);
        }
    }
}

/// Every `Expr::FuncCall` reachable from `expr`, including `expr` itself —
/// used by `compile_check_errors` to find calls buried inside a larger
/// expression (`Foo(Bar(1))`), not just ones that are a whole statement's
/// value on their own. Not exhaustive over every `Expr` variant the way
/// `collect_declared_names`/`collect_labels` are: missing a rare nested
/// container here only means `compile_check_errors` fails to catch that one
/// case early — it still surfaces as a normal runtime error later (the
/// existing, pre-this-phase behavior), not a false positive, so the lower
/// rigor bar is safe.
fn collect_func_calls<'a>(expr: &'a Expr, out: &mut Vec<&'a Expr>) {
    if let Expr::FuncCall { args, .. } = expr {
        out.push(expr);
        for a in args {
            collect_func_calls(a, out);
        }
        return;
    }
    match expr {
        Expr::BinOp { lhs, rhs, .. } => {
            collect_func_calls(lhs, out);
            collect_func_calls(rhs, out);
        }
        Expr::UnaryMinus(e) | Expr::UnaryNot(e) => collect_func_calls(e, out),
        Expr::CellRead { row, col } => {
            collect_func_calls(row, out);
            collect_func_calls(col, out);
        }
        Expr::RangeOffsetRead {
            row_off, col_off, ..
        } => {
            collect_func_calls(row_off, out);
            collect_func_calls(col_off, out);
        }
        Expr::CellsFind { what, .. } => collect_func_calls(what, out),
        Expr::SheetCellRead { sheet, row, col } => {
            collect_func_calls(sheet, out);
            collect_func_calls(row, out);
            collect_func_calls(col, out);
        }
        Expr::SheetRangeRead { sheet, .. } => collect_func_calls(sheet, out),
        Expr::WorkbookQualifiedSheet { workbook, sheet } => {
            collect_func_calls(workbook, out);
            collect_func_calls(sheet, out);
        }
        Expr::CellsEndProp { row, col, .. } => {
            collect_func_calls(row, out);
            collect_func_calls(col, out);
        }
        Expr::ArrayRecordGet { indices, .. } => {
            for i in indices {
                collect_func_calls(i, out);
            }
        }
        _ => {}
    }
}

/// `Some(params.len())` iff `name` resolves to a Sub or Function declared in
/// `prog` itself (not shadowed by a local variable/array/param — mirroring
/// `is_resolvable`'s own local-first precedence). `None` for a local name, a
/// cross-module call, or a builtin — `compile_check_errors`'s argument-count
/// check only fires when the callee's own declared arity is actually known,
/// which this project only tracks for a Sub/Function defined in the same
/// module being checked.
fn resolved_user_proc_arity(
    name: &str,
    prog: &Program,
    local_names: &HashSet<String>,
) -> Option<usize> {
    if local_names.contains(name) {
        return None;
    }
    if let Some(s) = prog.subs.iter().find(|s| s.name == name) {
        return Some(s.params.len());
    }
    if let Some(f) = prog.funcs.iter().find(|f| f.name == name) {
        return Some(f.params.len());
    }
    None
}

/// The subset of this module's own findings that are genuine VBA
/// *compile*-time errors — real VBA refuses to run a macro at all when one
/// of these is present anywhere in the project, and (being compile, not
/// runtime, errors) none of them can be trapped by `On Error`.
/// `Vm::run_sub`/`run_sub_multi` call this once per invocation, over the
/// whole program, before executing any statement — which is what makes
/// "uncatchable by `On Error`" true for free: the check runs before any
/// `On Error` statement has had a chance to take effect.
///
/// Checks exactly three things, chosen because each is syntactically
/// decidable without the type inference a real VBA compiler needs:
/// - an undefined Sub/Function call (reuses `is_resolvable`, the same logic
///   `run_check`'s own E1002 diagnostic already uses);
/// - an argument-count mismatch against a same-module callee's own declared
///   arity (`resolved_user_proc_arity` — cross-module calls aren't checked,
///   since this function only ever sees one module's own `Program`);
/// - a `GoTo`/`On Error GoTo` target that isn't a label anywhere in the
///   same procedure.
///
/// Returns the first violation found (message, span), walking `prog.subs`
/// then `prog.funcs` in declaration order, each depth-first through nested
/// blocks — deterministic, but the order carries no severity meaning.
///
/// Deliberately omits the fourth item from this phase's own spec, "invalid
/// assignment target": `name(args) = value` parses unconditionally as
/// `Stmt::ArrayWrite` (see `parse_ident_stmt`) regardless of whether `name`
/// is a real declared array or (invalidly) a Function name — telling those
/// apart isn't syntactically decidable without the type-inference this
/// project stays out of by design.
pub fn compile_check_errors(
    prog: &Program,
    other_module_names: &HashSet<String>,
) -> Option<(String, SourceSpan)> {
    for sub in &prog.subs {
        let local_names = local_scope_names(&sub.name, &sub.params, &sub.body);
        let mut labels = HashSet::new();
        collect_labels(&sub.body, &mut labels);
        if let Some(v) = check_body_for_compile_errors(
            &sub.body,
            prog,
            &local_names,
            other_module_names,
            &labels,
        ) {
            return Some(v);
        }
    }
    for func in &prog.funcs {
        let local_names = local_scope_names(&func.name, &func.params, &func.body);
        let mut labels = HashSet::new();
        collect_labels(&func.body, &mut labels);
        if let Some(v) = check_body_for_compile_errors(
            &func.body,
            prog,
            &local_names,
            other_module_names,
            &labels,
        ) {
            return Some(v);
        }
    }
    None
}

fn check_body_for_compile_errors(
    body: &[SpannedStmt],
    prog: &Program,
    local_names: &HashSet<String>,
    other_module_names: &HashSet<String>,
    labels: &HashSet<String>,
) -> Option<(String, SourceSpan)> {
    for s in body {
        match &s.stmt {
            Stmt::GoTo(label) if !labels.contains(label) => {
                return Some((format!("GoTo: label '{}' not found", label), s.span));
            }
            Stmt::OnErrorGoTo(label) if !labels.contains(label) => {
                return Some((
                    format!("On Error GoTo: label '{}' not found", label),
                    s.span,
                ));
            }
            Stmt::CallSub { name, args } => {
                if !is_resolvable(name, prog, local_names, other_module_names) {
                    return Some((format!("Sub/Function '{}' not found", name), s.span));
                }
                if let Some(arity) = resolved_user_proc_arity(name, prog, local_names)
                    && args.len() != arity
                {
                    return Some((
                        format!(
                            "'{}' expects {} argument(s), got {}",
                            name,
                            arity,
                            args.len()
                        ),
                        s.span,
                    ));
                }
            }
            _ => {}
        }

        let mut exprs = Vec::new();
        collect_stmt_exprs(&s.stmt, &mut exprs);
        for e in exprs {
            let mut calls = Vec::new();
            collect_func_calls(e, &mut calls);
            for call in calls {
                let Expr::FuncCall { name, args } = call else {
                    continue;
                };
                if !is_resolvable(name, prog, local_names, other_module_names) {
                    // `Expr::FuncCall` in expression position (unlike
                    // `Stmt::CallSub`) reaches the real builtin dispatch at
                    // runtime, so its actual failure text depends on which
                    // dispatch arm the name matches — e.g. `wsf_textjoin`
                    // fails inside `eval_wsf` with "WorksheetFunction.
                    // textjoin is not implemented", not the generic
                    // "Unknown VBA function" text a name matching no arm at
                    // all gets. `vm::builtin_call_error` asks the VM
                    // itself, so this always matches word-for-word.
                    let msg = vm::builtin_call_error(name)
                        .unwrap_or_else(|| format!("Unknown VBA function: '{}'", name));
                    return Some((msg, s.span));
                }
                if let Some(arity) = resolved_user_proc_arity(name, prog, local_names)
                    && args.len() != arity
                {
                    return Some((
                        format!(
                            "'{}' expects {} argument(s), got {}",
                            name,
                            arity,
                            args.len()
                        ),
                        s.span,
                    ));
                }
            }
        }

        for nested in nested_bodies(&s.stmt) {
            if let Some(v) =
                check_body_for_compile_errors(nested, prog, local_names, other_module_names, labels)
            {
                return Some(v);
            }
        }
    }
    None
}

/// `run_check`'s own way of surfacing the argument-count and GoTo-label
/// checks `compile_check_errors` also runs (see that function's doc for the
/// exact rules) — without this, `elixcee check` could report a program
/// clean that `Vm::run_sub`'s pre-flight pass would then refuse to run a
/// single statement of, since neither check was previously known to this
/// command at all. A parallel walk rather than a shared one with
/// `check_body_for_compile_errors`: that function stops at the first
/// violation (right for a pre-flight gate); this one collects every
/// violation as a located `Diagnostic` (right for a report meant to be read
/// in full). Deliberately does NOT re-check undefined-name resolution —
/// `walk_body`/`walk_expr` already report that as E1002, and running it
/// again here would double-report the same finding under two codes; that's
/// also why, unlike `compile_check_errors`, this never needs
/// `other_module_names` — the only thing that parameter is for.
fn collect_extra_compile_diagnostics(
    prog: &Program,
    source: &str,
    file: &str,
    diags: &mut Vec<Diagnostic>,
) {
    for sub in &prog.subs {
        let local_names = local_scope_names(&sub.name, &sub.params, &sub.body);
        let mut labels = HashSet::new();
        collect_labels(&sub.body, &mut labels);
        collect_extra_compile_diagnostics_body(
            &sub.body,
            prog,
            &local_names,
            &labels,
            source,
            file,
            diags,
        );
    }
    for func in &prog.funcs {
        let local_names = local_scope_names(&func.name, &func.params, &func.body);
        let mut labels = HashSet::new();
        collect_labels(&func.body, &mut labels);
        collect_extra_compile_diagnostics_body(
            &func.body,
            prog,
            &local_names,
            &labels,
            source,
            file,
            diags,
        );
    }
}

fn collect_extra_compile_diagnostics_body(
    body: &[SpannedStmt],
    prog: &Program,
    local_names: &HashSet<String>,
    labels: &HashSet<String>,
    source: &str,
    file: &str,
    diags: &mut Vec<Diagnostic>,
) {
    for s in body {
        match &s.stmt {
            Stmt::GoTo(label) if !labels.contains(label) => {
                diags.push(Diagnostic {
                    severity: "error",
                    code: "E1009",
                    kind: "undefined_label",
                    message: format!("GoTo: label '{}' not found", label),
                    location: Some(locate(source, file, s.span)),
                });
            }
            Stmt::OnErrorGoTo(label) if !labels.contains(label) => {
                diags.push(Diagnostic {
                    severity: "error",
                    code: "E1009",
                    kind: "undefined_label",
                    message: format!("On Error GoTo: label '{}' not found", label),
                    location: Some(locate(source, file, s.span)),
                });
            }
            Stmt::CallSub { name, args } => {
                if let Some(arity) = resolved_user_proc_arity(name, prog, local_names)
                    && args.len() != arity
                {
                    diags.push(Diagnostic {
                        severity: "error",
                        code: "E1008",
                        kind: "argument_count_mismatch",
                        message: format!(
                            "'{}' expects {} argument(s), got {}",
                            name,
                            arity,
                            args.len()
                        ),
                        location: Some(locate(source, file, s.span)),
                    });
                }
            }
            _ => {}
        }

        let mut exprs = Vec::new();
        collect_stmt_exprs(&s.stmt, &mut exprs);
        for e in exprs {
            let mut calls = Vec::new();
            collect_func_calls(e, &mut calls);
            for call in calls {
                let Expr::FuncCall { name, args } = call else {
                    continue;
                };
                if let Some(arity) = resolved_user_proc_arity(name, prog, local_names)
                    && args.len() != arity
                {
                    diags.push(Diagnostic {
                        severity: "error",
                        code: "E1008",
                        kind: "argument_count_mismatch",
                        message: format!(
                            "'{}' expects {} argument(s), got {}",
                            name,
                            arity,
                            args.len()
                        ),
                        location: Some(locate(source, file, s.span)),
                    });
                }
            }
        }

        for nested in nested_bodies(&s.stmt) {
            collect_extra_compile_diagnostics_body(
                nested,
                prog,
                local_names,
                labels,
                source,
                file,
                diags,
            );
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
            | Stmt::With { body, .. } => walk_body(
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
        Stmt::RangeObjectCopy { .. } => {}
        // `Set`'s RHS is an `ObjectExpr`, a separate AST from `Expr` — its
        // one nested `Expr` (an `Areas(n)` index) isn't walked for
        // undefined-function detection; out of scope for this pass.
        Stmt::Set { .. } => {}
        Stmt::RangePaste { transpose, .. } => {
            if let Some(e) = transpose {
                out.push(e);
            }
        }
        Stmt::SheetRangePaste { sheet, .. } => out.push(sheet),
        Stmt::SheetProtection { sheet, ui_only, .. } => {
            out.push(sheet);
            if let Some(e) = ui_only {
                out.push(e);
            }
        }
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
        Stmt::RowColDelete { index, .. } => out.push(index),
        Stmt::RowColInsert { index, .. } => out.push(index),
        Stmt::RangeSort { .. } => {}
        Stmt::RangeAutoFilter {
            field, criteria1, ..
        } => {
            if let Some(e) = field {
                out.push(e);
            }
            if let Some(e) = criteria1 {
                out.push(e);
            }
        }
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
        Stmt::SheetRangeWrite { sheet, value, .. } => {
            out.push(sheet);
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
        Stmt::DimBare { .. } => {}
        Stmt::DimArray { sizes, .. } => {
            for d in sizes {
                out.push(&d.upper);
                if let Some(lo) = &d.lower {
                    out.push(lo);
                }
            }
        }
        Stmt::ReDim { sizes, .. } => {
            for d in sizes {
                out.push(&d.upper);
                if let Some(lo) = &d.lower {
                    out.push(lo);
                }
            }
        }
        Stmt::Erase { .. } => {}
        Stmt::ArrayWrite { indices, value, .. } => {
            for i in indices {
                out.push(i);
            }
            out.push(value);
        }
        Stmt::With { target, .. } => {
            // A `With Cells(r, c)` target's index expressions are real
            // expressions and must be walked like any other.
            if let WithTarget::Cells(row, col) = target {
                out.push(row);
                out.push(col);
            }
        }
        Stmt::WithDot { member, value } => {
            if let WithMember::Cells { row, col, .. } = member {
                out.push(row);
                out.push(col);
            }
            out.push(value);
        }
        Stmt::MsgBox { message } => out.push(message),
        Stmt::RecordSet { value, .. } => out.push(value),
        Stmt::DimRecord { .. } => {}
        Stmt::DimArrayRecord { sizes, .. } => {
            for s in sizes {
                out.push(s);
            }
        }
        Stmt::DimMulti(decls) => {
            for d in decls {
                collect_stmt_exprs(d, out);
            }
        }
        Stmt::RecordSetNested { value, .. } => out.push(value),
        Stmt::ArrayRecordSet { indices, value, .. } => {
            for i in indices {
                out.push(i);
            }
            out.push(value);
        }
        Stmt::Unsupported { .. } => {}
        Stmt::ErrClear => {}
        Stmt::ErrRaise {
            number,
            source,
            description,
            help_file,
            help_context,
        } => {
            out.push(number);
            if let Some(e) = source {
                out.push(e);
            }
            if let Some(e) = description {
                out.push(e);
            }
            if let Some(e) = help_file {
                out.push(e);
            }
            if let Some(e) = help_context {
                out.push(e);
            }
        }
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
        Expr::SheetRangeRead { sheet, .. } => {
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
        }
        Expr::WorkbookQualifiedSheet { workbook, sheet } => {
            walk_expr(
                workbook,
                prog,
                local_names,
                other_module_names,
                stmt_span,
                source,
                file,
                diags,
            );
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
        | Expr::ActiveSheetRef
        | Expr::ObjectVarSheet(_)
        | Expr::RecordGet { .. }
        | Expr::RecordGetNested { .. }
        // `<var> Is Nothing` holds only a variable name, and a bare
        // `.member` read holds only field names — no sub-expression and no
        // callable in either, so nothing for the undefined-Sub/Function walk.
        | Expr::IsNothing(_)
        | Expr::WithDot(_)
        | Expr::ErrNumber
        | Expr::ErrDescription
        | Expr::ErrSource
        | Expr::ErrHelpFile
        | Expr::ErrHelpContext => {}
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
        // The message names the member, not `p.field`: the With target is no
        // longer substituted into the body at parse time (it's resolved at
        // runtime), so the parser genuinely doesn't know the variable's name
        // here. The diagnostic still points at the same statement.
        assert!(
            diags[0].message.contains(".field"),
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

    // ── compile_check_errors: Vm::run_sub's pre-flight compile-error pass ──

    fn parse_ok(src: &str) -> Program {
        parser::parse_with_span(src).unwrap_or_else(|e| panic!("should parse: {}", e.message))
    }

    fn compile_errors(src: &str) -> Option<(String, SourceSpan)> {
        compile_check_errors(&parse_ok(src), &HashSet::new())
    }

    #[test]
    fn clean_program_has_no_compile_errors() {
        assert!(compile_errors("Sub Main()\n    x = 1\nEnd Sub\n").is_none());
    }

    #[test]
    fn undefined_sub_call_is_a_compile_error() {
        let (msg, _) = compile_errors("Sub Main()\n    Call Helper()\nEnd Sub\n").unwrap();
        assert_eq!(msg, "Sub/Function 'helper' not found");
    }

    #[test]
    fn undefined_function_used_in_an_expression_is_a_compile_error() {
        let (msg, _) = compile_errors("Sub Main()\n    x = Helper(1)\nEnd Sub\n").unwrap();
        assert_eq!(msg, "Unknown VBA function: 'helper'");
    }

    #[test]
    fn undefined_call_nested_inside_another_call_is_still_caught() {
        // Foo is defined but Bar isn't — the outer call resolves, the
        // buried inner one doesn't; collect_func_calls must recurse into
        // Foo's own argument list to find it.
        let (msg, _) = compile_errors(
            "Function Foo(n)\n    Foo = n\nEnd Function\nSub Main()\n    x = Foo(Bar(1))\nEnd Sub\n",
        )
        .unwrap();
        assert_eq!(msg, "Unknown VBA function: 'bar'");
    }

    #[test]
    fn defined_sub_call_with_the_right_arg_count_is_not_a_compile_error() {
        assert!(compile_errors(
            "Sub Helper(a, b)\n    x = a + b\nEnd Sub\nSub Main()\n    Call Helper(1, 2)\nEnd Sub\n"
        )
        .is_none());
    }

    #[test]
    fn too_few_arguments_to_a_same_module_sub_is_a_compile_error() {
        let (msg, _) = compile_errors(
            "Sub Helper(a, b)\n    x = a + b\nEnd Sub\nSub Main()\n    Call Helper(1)\nEnd Sub\n",
        )
        .unwrap();
        assert_eq!(msg, "'helper' expects 2 argument(s), got 1");
    }

    #[test]
    fn too_many_arguments_to_a_same_module_function_is_a_compile_error() {
        let (msg, _) = compile_errors(
            "Function Helper(a)\n    Helper = a\nEnd Function\nSub Main()\n    x = Helper(1, 2)\nEnd Sub\n",
        )
        .unwrap();
        assert_eq!(msg, "'helper' expects 1 argument(s), got 2");
    }

    #[test]
    fn array_index_read_is_not_mistaken_for_an_argument_count_mismatch() {
        // `arr(1, 2)` and a 2-arg function call are syntactically identical
        // (`Expr::FuncCall`) — `arr` being a locally-Dim'd array must take
        // precedence, exactly as `is_resolvable` already establishes for
        // the undefined-name check.
        assert!(
            compile_errors(
                "Sub Main()\n    Dim arr(3, 3)\n    arr(1, 1) = 5\n    x = arr(1, 1)\nEnd Sub\n"
            )
            .is_none()
        );
    }

    #[test]
    fn goto_to_an_undefined_label_is_a_compile_error() {
        let (msg, _) = compile_errors("Sub Main()\n    GoTo Nowhere\nEnd Sub\n").unwrap();
        assert_eq!(msg, "GoTo: label 'nowhere' not found");
    }

    #[test]
    fn goto_to_a_label_declared_later_in_the_same_body_is_not_a_compile_error() {
        assert!(
            compile_errors("Sub Main()\n    GoTo Skip\n    x = 1\nSkip:\n    y = 2\nEnd Sub\n")
                .is_none()
        );
    }

    #[test]
    fn goto_to_a_label_nested_inside_an_if_block_is_not_a_compile_error() {
        // Real VBA GoTo scope is the whole procedure, not the current
        // block — a label inside a sibling `If` branch is a valid target.
        assert!(
            compile_errors(concat!(
                "Sub Main()\n",
                "    If True Then\n",
                "        GoTo Inner\n",
                "    End If\n",
                "    If False Then\n",
                "Inner:\n",
                "        y = 2\n",
                "    End If\n",
                "End Sub\n",
            ))
            .is_none()
        );
    }

    #[test]
    fn on_error_goto_an_undefined_label_is_a_compile_error() {
        let (msg, _) = compile_errors("Sub Main()\n    On Error GoTo Nowhere\nEnd Sub\n").unwrap();
        assert_eq!(msg, "On Error GoTo: label 'nowhere' not found");
    }

    #[test]
    fn on_error_goto_a_real_label_is_not_a_compile_error() {
        assert!(compile_errors(
            "Sub Main()\n    On Error GoTo Handler\n    x = 1\n    Exit Sub\nHandler:\n    y = 2\nEnd Sub\n"
        )
        .is_none());
    }

    #[test]
    fn a_deliberately_unimplemented_worksheet_function_is_not_flagged() {
        // wsf_-prefixed names always reach the real dispatch table at
        // runtime (`eval_wsf`'s own catch-all), never the generic "Unknown
        // VBA function" fallback — `is_known_builtin_function` (and so
        // `is_resolvable`) already treats every not-actually-implemented
        // wsf_ name as unresolvable, and this compile-check inherits that
        // via `is_resolvable` too, so a not-yet-implemented
        // WorksheetFunction call is reported (not silently missed) — this
        // test exists to pin the exact wording, which must come from
        // `vm::builtin_call_error` (the real dispatch), not be invented.
        let (msg, _) = compile_errors(
            "Sub Main()\n    x = WorksheetFunction.TextJoin(\",\", True, \"a\", \"b\")\nEnd Sub\n",
        )
        .unwrap();
        assert_eq!(msg, "WorksheetFunction.textjoin is not implemented");
    }

    #[test]
    fn a_call_to_a_name_in_another_module_is_not_flagged_when_registered() {
        let mut others = HashSet::new();
        others.insert("helper".to_string());
        let prog = parse_ok("Sub Main()\n    Call Helper()\nEnd Sub\n");
        assert!(compile_check_errors(&prog, &others).is_none());
    }

    #[test]
    fn a_cross_module_call_is_not_arg_count_checked_since_its_arity_is_unknown_here() {
        // `resolved_user_proc_arity` only ever looks at `prog`'s own
        // subs/funcs — a name registered only via `other_module_names`
        // (this function's only visibility into other modules) has no
        // known arity here, so no arg-count diagnostic can fire for it,
        // regardless of how many arguments the call site actually passes.
        let mut others = HashSet::new();
        others.insert("helper".to_string());
        let prog = parse_ok("Sub Main()\n    Call Helper(1, 2, 3)\nEnd Sub\n");
        assert!(compile_check_errors(&prog, &others).is_none());
    }

    #[test]
    fn first_violation_wins_when_a_program_has_several() {
        // Sub declaration order (Main, then Second) — Main's own undefined
        // call is found before Second's is ever reached.
        let prog = parse_ok(concat!(
            "Sub Main()\n",
            "    Call FirstUndefined()\n",
            "End Sub\n",
            "Sub Second()\n",
            "    Call SecondUndefined()\n",
            "End Sub\n",
        ));
        let (msg, _) = compile_check_errors(&prog, &HashSet::new()).unwrap();
        assert_eq!(msg, "Sub/Function 'firstundefined' not found");
    }

    // ── run_check itself must agree with compile_check_errors — otherwise
    // `elixcee check` can report a program clean that `Vm::run_sub`'s
    // pre-flight pass then refuses to run a single statement of. ──────────

    #[test]
    fn run_check_reports_an_argument_count_mismatch_as_e1008() {
        let diags = run_check(
            "Sub Helper(a, b)\n    x = a + b\nEnd Sub\nSub Main()\n    Call Helper(1)\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1008"]);
        assert_eq!(diags[0].kind, "argument_count_mismatch");
        assert_eq!(diags[0].message, "'helper' expects 2 argument(s), got 1");
        assert!(diags[0].location.is_some());
    }

    #[test]
    fn run_check_reports_an_undefined_goto_label_as_e1009() {
        let diags = run_check(
            "Sub Main()\n    GoTo Nowhere\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1009"]);
        assert_eq!(diags[0].kind, "undefined_label");
        assert_eq!(diags[0].message, "GoTo: label 'nowhere' not found");
    }

    #[test]
    fn run_check_reports_an_undefined_on_error_goto_label_as_e1009() {
        let diags = run_check(
            "Sub Main()\n    On Error GoTo Nowhere\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1009"]);
        assert_eq!(diags[0].message, "On Error GoTo: label 'nowhere' not found");
    }

    #[test]
    fn run_check_does_not_double_report_an_undefined_call_as_both_e1002_and_e1008() {
        let diags = run_check(
            "Sub Main()\n    Call DoesNotExist(1, 2, 3)\nEnd Sub\n",
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1002"]);
    }

    #[test]
    fn run_check_finds_every_argument_count_mismatch_not_just_the_first() {
        let diags = run_check(
            concat!(
                "Sub Helper(a, b)\n",
                "    x = a + b\n",
                "End Sub\n",
                "Sub Main()\n",
                "    Call Helper(1)\n",
                "    Call Helper(1, 2, 3)\n",
                "End Sub\n",
            ),
            "f.bas",
            Some("Main"),
        );
        assert_eq!(codes(&diags), vec!["E1008", "E1008"]);
    }

    #[test]
    fn run_check_does_not_flag_a_correct_program_with_the_new_checks() {
        let diags = run_check(
            concat!(
                "Sub Helper(a, b)\n",
                "    x = a + b\n",
                "End Sub\n",
                "Sub Main()\n",
                "    On Error GoTo Handler\n",
                "    Call Helper(1, 2)\n",
                "    GoTo Skip\n",
                "Handler:\n",
                "    y = 1\n",
                "Skip:\n",
                "    z = 2\n",
                "End Sub\n",
            ),
            "f.bas",
            Some("Main"),
        );
        assert!(diags.is_empty());
    }

    #[test]
    fn run_check_and_compile_check_errors_agree_on_the_exact_same_program() {
        // The regression this whole test group exists to prevent: `elixcee
        // check` reporting "ok" for a program `Vm::run_sub`'s pre-flight
        // check would then refuse to run.
        let src =
            "Sub Helper(a, b)\n    x = a + b\nEnd Sub\nSub Main()\n    Call Helper(1)\nEnd Sub\n";
        let diags = run_check(src, "f.bas", Some("Main"));
        let (compile_msg, _) = compile_errors(src).unwrap();
        assert_eq!(diags.len(), 1);
        assert_eq!(diags[0].message, compile_msg);
    }
}
