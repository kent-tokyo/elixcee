//! Milestone B6a: the `diagnose` subcommand — classifies *why* Excel would
//! reject an operation (missing worksheet/workbook, array out of bounds)
//! with concrete evidence, instead of only reporting a bare runtime-error
//! string. Runs a macro exactly once, with `Vm::strict_resolution` turned
//! on (see that field's doc comment in `src/vm/mod.rs`) so a reference to a
//! nonexistent sheet/workbook is a classified failure rather than being
//! papered over by elixcee's usual auto-vivify/silent-`Empty` convenience —
//! and with `On Error` not honored, so the *first* such failure always
//! propagates instead of being swallowed by the macro's own error handling.
//!
//! Own JSON contract, not `crate::diagnostics::ElixceeError`'s flat
//! `{code, kind, message, location}` shape — same reasoning as
//! `test-workbook` getting its own contract in Milestone B5a: a
//! `root_causes` array with per-kind evidence fields doesn't fit a flat
//! object, and every other subcommand's contract stays untouched.
//!
//! Explicit non-goals (see `docs/agent-contract.md` for the full list):
//! Copy/Paste shape validation, Clipboard/merged-cell/multi-area modeling,
//! a real `Collection` object, real multi-workbook execution, and
//! `Dim arr(1 To N)` non-zero-lower-bound tracking are all out of scope for
//! this milestone. `root_causes` carries at most one entry today (the first
//! failure) but is an array, not a bare object, because a later milestone's
//! ranked-candidate model reuses this exact shape.

use crate::diagnostics::{SourceLocation, json_string};
use crate::parser::{Program, ast::SourceSpan};
use crate::vm::{ResolutionEvidence, ResolutionFailureKind, Vm};

/// The outcome of running one macro under strict-resolution diagnosis.
#[derive(Debug)]
pub struct Diagnosis {
    pub ok: bool,
    /// The raw runtime-error string, if the run failed (whether or not it
    /// was a classified resolution failure — an unrelated runtime error
    /// still gets reported, just without a `root_causes` entry).
    pub message: Option<String>,
    /// Where in the source the failure happened, if it happened after at
    /// least one statement started executing. Resolving this into a
    /// file/line/column `SourceLocation` needs the source text, which this
    /// module doesn't have — the caller (mirroring run-mode's own
    /// convention in `main.rs`) does that conversion via
    /// `diagnostics::locate`.
    pub span: Option<SourceSpan>,
    /// The span of the `.Copy` statement that populated the clipboard, when
    /// `root_cause` is a `PASTE_SHAPE_MISMATCH` (Milestone B6b) — `span`
    /// above already points at the failing *Paste* statement, so this lets
    /// a diagnosis report both locations. `None` for every other kind.
    pub copy_span: Option<SourceSpan>,
    pub root_cause: Option<RootCause>,
    pub messages: Vec<String>,
}

#[derive(Debug)]
pub struct RootCause {
    pub code: &'static str,
    pub certainty: &'static str,
    pub kind: ResolutionFailureKind,
}

impl RootCause {
    fn from_kind(kind: ResolutionFailureKind) -> Self {
        let code = match &kind {
            ResolutionFailureKind::WorksheetNotFound(_) => "WORKSHEET_NOT_FOUND",
            ResolutionFailureKind::WorkbookNotFound(_) => "WORKBOOK_NOT_FOUND",
            ResolutionFailureKind::ArrayIndexOutOfBounds { .. } => "ARRAY_INDEX_OUT_OF_BOUNDS",
            ResolutionFailureKind::PasteShapeMismatch { .. } => "PASTE_SHAPE_MISMATCH",
            ResolutionFailureKind::PasteWithoutCopy { .. } => "PASTE_WITHOUT_COPY",
            ResolutionFailureKind::SheetProtected { .. } => "SHEET_PROTECTED",
        };
        RootCause {
            code,
            certainty: "definite",
            kind,
        }
    }

    /// Plain-English fix candidates. Kept modest and mechanically derived
    /// from the evidence — not a source-level rewrite suggestion (that
    /// depth of "here's the exact line to add" reasoning is out of scope
    /// for this milestone; see the module doc's non-goals list).
    fn suggestions(&self) -> Vec<String> {
        match &self.kind {
            ResolutionFailureKind::WorksheetNotFound(e)
            | ResolutionFailureKind::WorkbookNotFound(e) => match &e.suggested {
                Some(s) => vec![format!("did you mean '{}'?", s)],
                None if !e.available.is_empty() => {
                    vec![format!(
                        "check the available names: {}",
                        e.available.join(", ")
                    )]
                }
                None => vec![],
            },
            ResolutionFailureKind::ArrayIndexOutOfBounds {
                name,
                index,
                lower,
                upper,
            } => {
                vec![format!(
                    "check that '{}' is large enough for index {} (valid range is {} To {})",
                    name, index, lower, upper
                )]
            }
            ResolutionFailureKind::PasteShapeMismatch {
                source_rows,
                source_cols,
                dest_row1,
                dest_col1,
                transpose,
                ..
            } => {
                let (rows, cols) = if *transpose {
                    (*source_cols, *source_rows)
                } else {
                    (*source_rows, *source_cols)
                };
                let anchor = format!("{}{}", col_to_letters(*dest_col1), dest_row1);
                let bottom_right = format!(
                    "{}{}",
                    col_to_letters(dest_col1 + cols - 1),
                    dest_row1 + rows - 1
                );
                vec![
                    format!("resize the destination to {}:{}", anchor, bottom_right),
                    format!("or specify only the top-left cell {}", anchor),
                ]
            }
            ResolutionFailureKind::PasteWithoutCopy { .. } => vec![
                "add a Range(...).Copy before this Paste, or check whether \
                 Application.CutCopyMode was cleared first"
                    .to_string(),
            ],
            ResolutionFailureKind::SheetProtected { sheet } => vec![format!(
                "unprotect the sheet first: Worksheets(\"{}\").Unprotect",
                sheet
            )],
        }
    }
}

/// Renders a 1-based column number as its Excel letter form (`1` -> `"A"`).
/// A small private copy of the same tiny helper already independently
/// duplicated per-module in `snapshot.rs`/`main.rs`/`testworkbook.rs`, not a
/// shared `utils` module — matches existing project convention.
fn col_to_letters(mut col: u32) -> String {
    let mut bytes = Vec::new();
    while col > 0 {
        col -= 1;
        bytes.push(b'A' + (col % 26) as u8);
        col /= 26;
    }
    bytes.reverse();
    String::from_utf8(bytes).unwrap()
}

/// Runs `entrypoint` once against the workbook at `workbook_path`, in
/// strict-resolution mode. `Err` is only for setup failures before the
/// macro could start (the workbook file couldn't be read, or has no
/// sheets) — mirrors `Vm::load_workbook_file`'s own error shape, which the
/// caller already knows how to classify (`E3001`/`E3002`, same as run-mode
/// and `test-workbook`).
pub fn run_diagnosis(
    programs: &[(String, Program)],
    workbook_path: &str,
    entrypoint: &str,
) -> Result<Diagnosis, String> {
    let mut vm = Vm::new();
    vm.strict_resolution = true;
    vm.load_workbook_file(workbook_path)?;

    let run_result = if programs.len() == 1 {
        vm.run_sub(&programs[0].1, entrypoint)
    } else {
        vm.run_sub_multi(programs, entrypoint)
    };

    match run_result {
        Ok(()) => Ok(Diagnosis {
            ok: true,
            message: None,
            span: None,
            copy_span: None,
            root_cause: None,
            messages: vm.take_messages(),
        }),
        Err(message) => {
            let root_cause = vm.take_resolution_failure().map(RootCause::from_kind);
            let span = vm.current_span();
            let copy_span = root_cause.as_ref().and_then(|rc| match &rc.kind {
                ResolutionFailureKind::PasteShapeMismatch { copy_span, .. } => *copy_span,
                _ => None,
            });
            Ok(Diagnosis {
                ok: false,
                message: Some(message),
                span,
                copy_span,
                root_cause,
                messages: vm.take_messages(),
            })
        }
    }
}

fn location_json(location: Option<&SourceLocation>) -> String {
    match location {
        Some(loc) => format!(
            "{{\"file\":{},\"line\":{},\"column\":{}}}",
            json_string(&loc.file),
            loc.line,
            loc.column,
        ),
        None => "null".to_string(),
    }
}

fn evidence_json(e: &ResolutionEvidence) -> String {
    let available = format!(
        "[{}]",
        e.available
            .iter()
            .map(|s| json_string(s))
            .collect::<Vec<_>>()
            .join(",")
    );
    let suggested = match &e.suggested {
        Some(s) => json_string(s),
        None => "null".to_string(),
    };
    format!(
        "\"expression\":{},\"requested\":{},\"available\":{},\"suggested\":{}",
        json_string(&e.expression),
        json_string(&e.requested),
        available,
        suggested,
    )
}

fn root_cause_json(rc: &RootCause, copy_location: Option<&SourceLocation>) -> String {
    let suggestions = format!(
        "[{}]",
        rc.suggestions()
            .iter()
            .map(|s| json_string(s))
            .collect::<Vec<_>>()
            .join(",")
    );
    let fields = match &rc.kind {
        ResolutionFailureKind::WorksheetNotFound(e)
        | ResolutionFailureKind::WorkbookNotFound(e) => evidence_json(e),
        ResolutionFailureKind::ArrayIndexOutOfBounds {
            name,
            index,
            lower,
            upper,
        } => format!(
            "\"name\":{},\"index\":{},\"lower\":{},\"upper\":{}",
            json_string(name),
            index,
            lower,
            upper,
        ),
        ResolutionFailureKind::PasteShapeMismatch {
            source_addr,
            source_rows,
            source_cols,
            dest_addr,
            dest_rows,
            dest_cols,
            transpose,
            ..
        } => format!(
            "\"source_addr\":{},\"source_rows\":{},\"source_cols\":{},\
             \"dest_addr\":{},\"dest_rows\":{},\"dest_cols\":{},\"transpose\":{},\
             \"copy_location\":{}",
            json_string(source_addr),
            source_rows,
            source_cols,
            json_string(dest_addr),
            dest_rows,
            dest_cols,
            transpose,
            location_json(copy_location),
        ),
        ResolutionFailureKind::PasteWithoutCopy { dest_addr } => {
            format!("\"dest_addr\":{}", json_string(dest_addr))
        }
        ResolutionFailureKind::SheetProtected { sheet } => {
            format!("\"sheet\":{}", json_string(sheet))
        }
    };
    format!(
        "{{\"code\":\"{}\",\"certainty\":\"{}\",{},\"suggestions\":{}}}",
        rc.code, rc.certainty, fields, suggestions,
    )
}

/// `{"schema_version":1,"ok":true,"messages":[...]}` on success, or
/// `{"schema_version":1,"ok":false,"message":...,"location":...,"root_causes":[...],"messages":[...]}`
/// on failure. `location`/`copy_location` are resolved by the caller (see
/// `Diagnosis::span`/`copy_span`'s doc comments) — `None` when the caller
/// couldn't or didn't resolve one. `copy_location` only ever appears nested
/// inside a `PASTE_SHAPE_MISMATCH` root cause; it's accepted here regardless
/// of `diag.root_cause`'s kind so callers don't need to inspect it first.
pub fn to_json(
    diag: &Diagnosis,
    location: Option<&SourceLocation>,
    copy_location: Option<&SourceLocation>,
) -> String {
    let messages_json = format!(
        "[{}]",
        diag.messages
            .iter()
            .map(|m| json_string(m))
            .collect::<Vec<_>>()
            .join(",")
    );
    if diag.ok {
        return format!(
            "{{\"schema_version\":1,\"ok\":true,\"messages\":{}}}",
            messages_json
        );
    }
    let root_causes = match &diag.root_cause {
        Some(rc) => format!("[{}]", root_cause_json(rc, copy_location)),
        None => "[]".to_string(),
    };
    format!(
        "{{\"schema_version\":1,\"ok\":false,\"message\":{},\"location\":{},\"root_causes\":{},\"messages\":{}}}",
        json_string(diag.message.as_deref().unwrap_or("")),
        location_json(location),
        root_causes,
        messages_json,
    )
}

/// Plain-text summary for non-`--json` invocations — mirrors the level of
/// detail `test-workbook`'s `to_plain_text` gives, not a full replica of
/// the JSON shape.
pub fn to_plain_text(
    diag: &Diagnosis,
    location: Option<&SourceLocation>,
    copy_location: Option<&SourceLocation>,
) -> String {
    if diag.ok {
        return "OK: no resolution failure detected".to_string();
    }
    let mut out = format!(
        "FAILED: {}",
        diag.message.as_deref().unwrap_or("(unknown error)")
    );
    if let Some(loc) = location {
        out.push_str(&format!("\n  at {}:{}:{}", loc.file, loc.line, loc.column));
    }
    if let Some(rc) = &diag.root_cause {
        out.push_str(&format!("\n\n{}: ", rc.code));
        match &rc.kind {
            ResolutionFailureKind::WorksheetNotFound(e)
            | ResolutionFailureKind::WorkbookNotFound(e) => {
                out.push_str(&format!(
                    "{}\n  requested: {}\n  available: {}",
                    e.expression,
                    e.requested,
                    e.available.join(", "),
                ));
                if let Some(s) = &e.suggested {
                    out.push_str(&format!("\n  did you mean: {}", s));
                }
            }
            ResolutionFailureKind::ArrayIndexOutOfBounds {
                name,
                index,
                lower,
                upper,
            } => {
                out.push_str(&format!(
                    "array '{}' index {} out of bounds (valid range: {} To {})",
                    name, index, lower, upper
                ));
            }
            ResolutionFailureKind::PasteShapeMismatch {
                source_addr,
                source_rows,
                source_cols,
                dest_addr,
                dest_rows,
                dest_cols,
                transpose,
                ..
            } => {
                out.push_str(&format!(
                    "copy source {} ({}x{}) does not match paste destination {} ({}x{}){}",
                    source_addr,
                    source_rows,
                    source_cols,
                    dest_addr,
                    dest_rows,
                    dest_cols,
                    if *transpose { ", Transpose:=True" } else { "" },
                ));
                if let Some(loc) = copy_location {
                    out.push_str(&format!(
                        "\n  copied at {}:{}:{}",
                        loc.file, loc.line, loc.column
                    ));
                }
            }
            ResolutionFailureKind::PasteWithoutCopy { dest_addr } => {
                out.push_str(&format!(
                    "Paste to {} attempted with an empty clipboard",
                    dest_addr
                ));
            }
            ResolutionFailureKind::SheetProtected { sheet } => {
                out.push_str(&format!("sheet '{}' is protected", sheet));
            }
        }
        for s in rc.suggestions() {
            out.push_str(&format!("\n  suggestion: {}", s));
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser;

    fn programs_from(src: &str) -> Vec<(String, Program)> {
        vec![("main".to_string(), parser::parse(src).unwrap())]
    }

    fn build_workbook(path: &str) {
        let vm = Vm::new();
        crate::save_workbook(&vm, path).unwrap();
    }

    #[test]
    fn passing_macro_reports_ok_true() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_ok.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from("Sub Main()\n    x = 1\nEnd Sub\n");
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(diag.ok);
        assert!(diag.root_cause.is_none());
        assert!(to_json(&diag, None, None).contains("\"ok\":true"));
    }

    #[test]
    fn missing_worksheet_reports_a_worksheet_not_found_root_cause_with_a_suggestion() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_missing_sheet.xlsx");
        let mut source_vm = Vm::new();
        source_vm.ensure_sheet("入力");
        source_vm.ensure_sheet("売上2026");
        crate::save_workbook(&source_vm, out_path.to_str().unwrap()).unwrap();

        let programs = programs_from(
            "Sub Main()\n    Worksheets(\"売上2025\").Range(\"A1\").Value = 1\nEnd Sub\n",
        );
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        let rc = diag
            .root_cause
            .as_ref()
            .expect("should classify a root cause");
        assert_eq!(rc.code, "WORKSHEET_NOT_FOUND");
        match &rc.kind {
            ResolutionFailureKind::WorksheetNotFound(e) => {
                assert_eq!(e.requested, "売上2025");
                assert_eq!(e.suggested.as_deref(), Some("売上2026"));
            }
            other => panic!("expected WorksheetNotFound, got {:?}", other),
        }
        let json = to_json(&diag, None, None);
        assert!(json.contains("WORKSHEET_NOT_FOUND"));
        assert!(json.contains("売上2026"));
    }

    #[test]
    fn a_non_resolution_failure_reports_the_raw_message_with_no_root_cause() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_undefined_var.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from("Sub Main()\n    Cells(1, 1).Value = x\nEnd Sub\n");
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        assert!(diag.root_cause.is_none());
        assert!(
            diag.message
                .as_deref()
                .unwrap_or("")
                .contains("Undefined variable")
        );
        assert!(to_json(&diag, None, None).contains("\"root_causes\":[]"));
    }

    #[test]
    fn missing_workbook_reports_a_workbook_not_found_root_cause() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_missing_workbook.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from(
            "Sub Main()\n    Workbooks(\"data.xlsx\").Worksheets(1).Cells(1,1).Value = 1\nEnd Sub\n",
        );
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        let rc = diag
            .root_cause
            .as_ref()
            .expect("should classify a root cause");
        assert_eq!(rc.code, "WORKBOOK_NOT_FOUND");
    }

    #[test]
    fn array_out_of_bounds_reports_evidence_with_true_zero_based_bounds() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_array_oob.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from("Sub Main()\n    Dim values(5)\n    values(6) = 1\nEnd Sub\n");
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        let rc = diag
            .root_cause
            .as_ref()
            .expect("should classify a root cause");
        assert_eq!(rc.code, "ARRAY_INDEX_OUT_OF_BOUNDS");
        match &rc.kind {
            ResolutionFailureKind::ArrayIndexOutOfBounds {
                name,
                index,
                lower,
                upper,
            } => {
                assert_eq!(name, "values");
                assert_eq!(*index, 6);
                assert_eq!(*lower, 0);
                assert_eq!(*upper, 5);
            }
            other => panic!("expected ArrayIndexOutOfBounds, got {:?}", other),
        }
    }

    #[test]
    fn on_error_resume_next_does_not_swallow_the_failure() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_on_error.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from(
            "Sub Main()\n    On Error Resume Next\n    Worksheets(\"NoSuchSheet\").Range(\"A1\").Value = 1\nEnd Sub\n",
        );
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(
            !diag.ok,
            "diagnose must not let On Error Resume Next hide the failure"
        );
        assert_eq!(
            diag.root_cause.as_ref().map(|rc| rc.code),
            Some("WORKSHEET_NOT_FOUND")
        );
    }

    #[test]
    fn missing_workbook_file_is_a_setup_error_not_a_panic() {
        let programs = programs_from("Sub Main()\nEnd Sub\n");
        let err = run_diagnosis(&programs, "/nonexistent/path.xlsx", "Main").unwrap_err();
        assert!(err.starts_with("cannot read"), "{:?}", err);
    }

    #[test]
    fn paste_shape_mismatch_reports_both_locations_and_a_resize_suggestion() {
        // The user's own literal example: A1:C10 (10x3) copied, E1:F10
        // (10x2) pasted into — column counts differ.
        let out_path = std::env::temp_dir().join("elixcee_diagnose_paste_shape_mismatch.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from(
            "Sub Main()\n    Range(\"A1:C10\").Copy\n    Range(\"E1:F10\").PasteSpecial\nEnd Sub\n",
        );
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        let rc = diag
            .root_cause
            .as_ref()
            .expect("should classify a root cause");
        assert_eq!(rc.code, "PASTE_SHAPE_MISMATCH");
        assert!(diag.span.is_some(), "must locate the failing Paste");
        assert!(diag.copy_span.is_some(), "must also locate the Copy");
        assert_eq!(
            rc.suggestions(),
            vec![
                "resize the destination to E1:G10".to_string(),
                "or specify only the top-left cell E1".to_string(),
            ]
        );
        let json = to_json(&diag, None, None);
        assert!(json.contains("PASTE_SHAPE_MISMATCH"));
        assert!(json.contains("\"source_addr\":\"A1:C10\""));
        assert!(json.contains("\"dest_addr\":\"E1:F10\""));
        assert!(json.contains("resize the destination to E1:G10"));
    }

    #[test]
    fn paste_without_copy_reports_a_root_cause_with_a_fix_suggestion() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_paste_without_copy.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from("Sub Main()\n    Range(\"A1\").PasteSpecial\nEnd Sub\n");
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        let rc = diag
            .root_cause
            .as_ref()
            .expect("should classify a root cause");
        assert_eq!(rc.code, "PASTE_WITHOUT_COPY");
        assert!(diag.copy_span.is_none());
        let json = to_json(&diag, None, None);
        assert!(json.contains("PASTE_WITHOUT_COPY"));
        assert!(json.contains("\"dest_addr\":\"A1\""));
    }

    #[test]
    fn sheet_protected_reports_a_root_cause_with_an_unprotect_suggestion() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_sheet_protected.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from(
            "Sub Main()\n    Worksheets(\"Sheet1\").Protect\n    Cells(1,1).Value = 1\nEnd Sub\n",
        );
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        let rc = diag
            .root_cause
            .as_ref()
            .expect("should classify a root cause");
        assert_eq!(rc.code, "SHEET_PROTECTED");
        let json = to_json(&diag, None, None);
        assert!(json.contains("SHEET_PROTECTED"));
        assert!(json.contains("unprotect the sheet first"));
    }
}
