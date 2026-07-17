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
//! hidden/filtered rows, Excel Tables, a real `Collection` object, real
//! multi-workbook execution, and `Dim arr(1 To N)` non-zero-lower-bound
//! tracking are all out of scope. (Copy/Paste shape/clipboard validation
//! shipped in Milestone B6b, merged-cell-aware Paste conflicts in Milestone
//! B6c2, and a multi-area (`Areas`) Range foundation — diagnose-only, never
//! completes a multi-area paste — in Milestone B7a; all three superseding
//! this comment's original B6a-era non-goals list.) `root_causes` carries
//! at most one entry today (the first failure) but is an array, not a bare
//! object, because a later milestone's ranked-candidate model reuses this
//! exact shape.

use crate::diagnostics::{SourceLocation, json_string};
use crate::parser::{Program, ast::SourceSpan};
use crate::vm::{
    HiddenCellsObservation, Interval, Rect, ResolutionEvidence, ResolutionFailureKind, Vm,
};

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
    /// `root_cause` is any Paste-related kind — `PASTE_SHAPE_MISMATCH`
    /// (Milestone B6b) or `PASTE_INTO_NON_ANCHOR_MERGED_CELL`/
    /// `PASTE_PARTIAL_MERGED_RANGE`/`PASTE_MERGE_LAYOUT_MISMATCH`
    /// (Milestone B6c2) — `span` above already points at the failing
    /// *Paste* statement, so this lets a diagnosis report both locations.
    /// `None` for every other kind.
    pub copy_span: Option<SourceSpan>,
    pub root_cause: Option<RootCause>,
    /// The `RANGE_CONTAINS_HIDDEN_CELLS` observation (Milestone B7b), if
    /// the last `.Copy`'d range overlapped hidden rows/columns — populated
    /// regardless of `ok`, since this isn't a failure at all (see
    /// `Vm::hidden_cells_observation`'s doc comment for exactly when it's
    /// `None`). Rendered as a separate `observations` JSON field, not
    /// folded into `root_causes` (which means "why it failed").
    pub hidden_cells: Option<HiddenCellsObservation>,
    pub messages: Vec<String>,
}

#[derive(Debug)]
pub struct RootCause {
    pub code: &'static str,
    pub certainty: &'static str,
    pub kind: ResolutionFailureKind,
}

impl RootCause {
    /// `pub(crate)` (not fully private) so `diagnose-workbook` (Milestone
    /// B6d) can classify a `ResolutionFailureKind` captured from a
    /// generated test-workbook case and read `.code`/`.kind` directly for
    /// plain-text rendering — the JSON-rendering internals
    /// (`.suggestions()`, `root_cause_json`) stay private; `root_causes_json`
    /// below is the one JSON-shaped entry point.
    pub(crate) fn from_kind(kind: ResolutionFailureKind) -> Self {
        let code = match &kind {
            ResolutionFailureKind::WorksheetNotFound(_) => "WORKSHEET_NOT_FOUND",
            ResolutionFailureKind::WorkbookNotFound(_) => "WORKBOOK_NOT_FOUND",
            ResolutionFailureKind::ArrayIndexOutOfBounds { .. } => "ARRAY_INDEX_OUT_OF_BOUNDS",
            ResolutionFailureKind::PasteShapeMismatch { .. } => "PASTE_SHAPE_MISMATCH",
            ResolutionFailureKind::PasteWithoutCopy { .. } => "PASTE_WITHOUT_COPY",
            ResolutionFailureKind::SheetProtected { .. } => "SHEET_PROTECTED",
            ResolutionFailureKind::PasteIntoNonAnchorMergedCell { .. } => {
                "PASTE_INTO_NON_ANCHOR_MERGED_CELL"
            }
            ResolutionFailureKind::PastePartialMergedRange { .. } => "PASTE_PARTIAL_MERGED_RANGE",
            ResolutionFailureKind::PasteMergeLayoutMismatch { .. } => "PASTE_MERGE_LAYOUT_MISMATCH",
            ResolutionFailureKind::MultiAreaToSingleAreaPaste { .. } => {
                "MULTI_AREA_TO_SINGLE_AREA_PASTE"
            }
            ResolutionFailureKind::MultiAreaCountMismatch { .. } => "MULTI_AREA_COUNT_MISMATCH",
            ResolutionFailureKind::MultiAreaShapeMismatch { .. } => "MULTI_AREA_SHAPE_MISMATCH",
            ResolutionFailureKind::MultiAreaPasteUnsupported { .. } => {
                "MULTI_AREA_PASTE_UNSUPPORTED"
            }
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
            ResolutionFailureKind::PasteIntoNonAnchorMergedCell { merged_range, .. } => {
                let addr = rect_addr(*merged_range);
                let anchor = rect_addr((merged_range.0, merged_range.0));
                vec![
                    format!("unmerge {} before pasting", addr),
                    format!("or paste into its top-left cell {} instead", anchor),
                ]
            }
            ResolutionFailureKind::PastePartialMergedRange { conflicts, .. } => {
                let addrs = conflicts.iter().map(|r| rect_addr(*r)).collect::<Vec<_>>().join(", ");
                vec![format!(
                    "unmerge {} before pasting, or resize the destination to fully contain it",
                    addrs
                )]
            }
            ResolutionFailureKind::PasteMergeLayoutMismatch { conflicts, .. } => {
                let addrs = conflicts.iter().map(|r| rect_addr(*r)).collect::<Vec<_>>().join(", ");
                vec![
                    format!("unmerge {} before pasting", addrs),
                    "or make the source and destination merge layouts identical".to_string(),
                ]
            }
            ResolutionFailureKind::MultiAreaToSingleAreaPaste { .. } => vec![
                "paste each source area separately".to_string(),
                "copy a contiguous rectangular range".to_string(),
                "use destination areas with matching count and shapes".to_string(),
            ],
            ResolutionFailureKind::MultiAreaCountMismatch { .. } => vec![
                "match the destination area count to the source area count".to_string(),
                "paste each source area separately".to_string(),
            ],
            ResolutionFailureKind::MultiAreaShapeMismatch { area_index, .. } => vec![
                format!(
                    "resize destination area {} to match the source area's shape",
                    area_index
                ),
                "paste each source area separately".to_string(),
            ],
            ResolutionFailureKind::MultiAreaPasteUnsupported { .. } => {
                vec!["paste each source area separately".to_string()]
            }
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

/// A 1-based inclusive `((row1,col1),(row2,col2))` rect (Milestone B6c2) —
/// a private per-module alias, not a shared type, matching this codebase's
/// existing per-module `col_to_letters` duplication convention rather than
/// a cross-module `utils` dependency.
type MergeRect = ((u32, u32), (u32, u32));

/// Renders a rect as an address string, e.g. `((1,5),(1,7))` -> `"E1:G1"` —
/// a single-cell rect renders without the `:` (`"E1"`, not `"E1:E1"`).
fn rect_addr(r: MergeRect) -> String {
    let ((r1, c1), (r2, c2)) = r;
    if (r1, c1) == (r2, c2) {
        format!("{}{}", col_to_letters(c1), r1)
    } else {
        format!("{}{}:{}{}", col_to_letters(c1), r1, col_to_letters(c2), r2)
    }
}

/// Formats a `vm::Rect` (Milestone B7a) via the same `rect_addr` used for
/// `MergeRect` — the two types share the same 1-based inclusive bounds,
/// just a named struct vs. a bare tuple.
fn area_addr(r: Rect) -> String {
    rect_addr(((r.start_row, r.start_col), (r.end_row, r.end_col)))
}

/// One multi-area evidence object (Milestone B7a): `{"address":...,
/// "rows":N,"columns":N}` — the completion-condition JSON's own per-area
/// shape, deliberately not `evidence_json`'s flat-field convention (see the
/// B7a plan's "Advisor-reviewed scope decisions" #2).
fn area_json(r: Rect) -> String {
    format!(
        "{{\"address\":{},\"rows\":{},\"columns\":{}}}",
        json_string(&area_addr(r)),
        r.rows(),
        r.cols(),
    )
}

/// A JSON array of `area_json` objects, same `format!("[{}]", ...)` idiom
/// as `conflicts_json`.
fn areas_json(areas: &[Rect]) -> String {
    format!(
        "[{}]",
        areas.iter().map(|r| area_json(*r)).collect::<Vec<_>>().join(",")
    )
}

/// Renders a hidden-row interval as `"11:14"` (Milestone B7b) — always
/// `"{start}:{end}"`, even for a single row, matching the completion
/// condition's own `"B:B"`-style column rendering rather than
/// `rect_addr`'s single-cell-omits-the-colon convention.
fn row_interval_addr(iv: &Interval) -> String {
    format!("{}:{}", iv.start, iv.end)
}

/// Renders a hidden-column interval as `"B:B"` (Milestone B7b).
fn col_interval_addr(iv: &Interval) -> String {
    format!("{}:{}", col_to_letters(iv.start), col_to_letters(iv.end))
}

fn row_intervals_json(intervals: &[Interval]) -> String {
    format!(
        "[{}]",
        intervals.iter().map(|iv| json_string(&row_interval_addr(iv))).collect::<Vec<_>>().join(",")
    )
}

fn col_intervals_json(intervals: &[Interval]) -> String {
    format!(
        "[{}]",
        intervals.iter().map(|iv| json_string(&col_interval_addr(iv))).collect::<Vec<_>>().join(",")
    )
}

/// Renders the `"observations":[...]` fragment (Milestone B7b) — `"[]"`
/// for `None`, a one-item array for `Some`. Callers (`to_json`,
/// `diagnose-workbook`) omit the whole `observations` field rather than
/// emit `"observations":[]"` when this is `None` — see `Diagnosis::
/// hidden_cells`'s doc comment for why (this is a non-failure observation,
/// not a `root_causes`-shaped "why it failed" entry, and an
/// always-present empty array would break every existing `--json`
/// fixture in `tests/blackbox.rs`).
pub(crate) fn observations_json(obs: Option<&HiddenCellsObservation>) -> String {
    match obs {
        Some(o) => format!(
            "[{{\"code\":\"RANGE_CONTAINS_HIDDEN_CELLS\",\"certainty\":\"observed\",\
             \"range\":{{\"sheet\":{},\"address\":{},\"rows\":{},\"columns\":{}}},\
             \"visibility\":{{\"hidden_rows\":{},\"hidden_columns\":{},\
             \"total_cells\":{},\"visible_cells\":{}}},\"message\":{}}}]",
            json_string(&o.sheet),
            json_string(&o.address),
            o.rows,
            o.columns,
            row_intervals_json(&o.hidden_rows),
            col_intervals_json(&o.hidden_columns),
            o.total_cells,
            o.visible_cells,
            json_string(
                "The range contains hidden rows or columns. Excel operations using \
                 visible cells only may produce a multi-area range."
            ),
        ),
        None => "[]".to_string(),
    }
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
            hidden_cells: vm.hidden_cells_observation(),
            messages: vm.take_messages(),
        }),
        Err(message) => {
            let root_cause = vm.take_resolution_failure().map(RootCause::from_kind);
            let span = vm.current_span();
            let copy_span = root_cause.as_ref().and_then(|rc| match &rc.kind {
                ResolutionFailureKind::PasteShapeMismatch { copy_span, .. }
                | ResolutionFailureKind::PasteIntoNonAnchorMergedCell { copy_span, .. }
                | ResolutionFailureKind::PastePartialMergedRange { copy_span, .. }
                | ResolutionFailureKind::PasteMergeLayoutMismatch { copy_span, .. } => *copy_span,
                _ => None,
            });
            let hidden_cells = vm.hidden_cells_observation();
            Ok(Diagnosis {
                ok: false,
                message: Some(message),
                span,
                copy_span,
                root_cause,
                hidden_cells,
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

/// Renders a list of merge-conflict rects as a JSON array of address
/// strings (Milestone B6c2), same `format!("[{}]", ...join(","))` idiom as
/// `evidence_json`'s `available` field.
fn conflicts_json(conflicts: &[MergeRect]) -> String {
    format!(
        "[{}]",
        conflicts
            .iter()
            .map(|r| json_string(&rect_addr(*r)))
            .collect::<Vec<_>>()
            .join(",")
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
        ResolutionFailureKind::PasteIntoNonAnchorMergedCell {
            dest_addr,
            dest_sheet,
            merged_range,
            ..
        } => format!(
            "\"dest_addr\":{},\"dest_sheet\":{},\"merged_range\":{},\"copy_location\":{}",
            json_string(dest_addr),
            json_string(dest_sheet),
            json_string(&rect_addr(*merged_range)),
            location_json(copy_location),
        ),
        ResolutionFailureKind::PastePartialMergedRange {
            dest_addr,
            dest_sheet,
            conflicts,
            ..
        } => format!(
            "\"dest_addr\":{},\"dest_sheet\":{},\"conflicts\":{},\"copy_location\":{}",
            json_string(dest_addr),
            json_string(dest_sheet),
            conflicts_json(conflicts),
            location_json(copy_location),
        ),
        ResolutionFailureKind::PasteMergeLayoutMismatch {
            source_addr,
            source_sheet,
            dest_addr,
            dest_sheet,
            conflicts,
            ..
        } => format!(
            "\"source_addr\":{},\"source_sheet\":{},\"dest_addr\":{},\"dest_sheet\":{},\
             \"conflicts\":{},\"copy_location\":{}",
            json_string(source_addr),
            json_string(source_sheet),
            json_string(dest_addr),
            json_string(dest_sheet),
            conflicts_json(conflicts),
            location_json(copy_location),
        ),
        ResolutionFailureKind::MultiAreaToSingleAreaPaste { source_areas, destination_areas }
        | ResolutionFailureKind::MultiAreaCountMismatch { source_areas, destination_areas }
        | ResolutionFailureKind::MultiAreaPasteUnsupported { source_areas, destination_areas } => {
            format!(
                "\"source_areas\":{},\"destination_areas\":{}",
                areas_json(source_areas),
                areas_json(destination_areas),
            )
        }
        ResolutionFailureKind::MultiAreaShapeMismatch {
            area_index,
            source_area,
            destination_area,
        } => format!(
            "\"area_index\":{},\"source_area\":{},\"destination_area\":{}",
            area_index,
            area_json(*source_area),
            area_json(*destination_area),
        ),
    };
    format!(
        "{{\"code\":\"{}\",\"certainty\":\"{}\",{},\"suggestions\":{}}}",
        rc.code, rc.certainty, fields, suggestions,
    )
}

/// Renders the same `"root_causes":[...]` fragment `to_json` produces —
/// `"[]"` for `None`, a one-item array for `Some` — reusing the exact same
/// field spellings (Milestone B6d) so `diagnose-workbook` never reports the
/// same `code` with different evidence field names than plain `diagnose`
/// does. No `copy_location` is available here (there's no per-case source
/// text/location resolution in the generated-case search), so paste-related
/// kinds render `"copy_location":null`, same as `to_json` does whenever the
/// caller doesn't resolve one.
pub(crate) fn root_causes_json(kind: Option<&ResolutionFailureKind>) -> String {
    match kind {
        Some(k) => format!("[{}]", root_cause_json(&RootCause::from_kind(k.clone()), None)),
        None => "[]".to_string(),
    }
}

/// `{"schema_version":1,"ok":true,"messages":[...]}` on success, or
/// `{"schema_version":1,"ok":false,"message":...,"location":...,"root_causes":[...],"messages":[...]}`
/// on failure. `location`/`copy_location` are resolved by the caller (see
/// `Diagnosis::span`/`copy_span`'s doc comments) — `None` when the caller
/// couldn't or didn't resolve one. `copy_location` only ever appears nested
/// inside a Paste-related root cause (see `Diagnosis::copy_span`'s doc
/// comment for the full list); it's accepted here regardless of
/// `diag.root_cause`'s kind so callers don't need to inspect it first.
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
    // Milestone B7b: a sibling field, present only when `Some` — never
    // `"observations":[]"`, since that would break every existing
    // `--json` fixture in `tests/blackbox.rs` that predates this field.
    let observations_field = match &diag.hidden_cells {
        Some(obs) => format!(",\"observations\":{}", observations_json(Some(obs))),
        None => String::new(),
    };
    if diag.ok {
        return format!(
            "{{\"schema_version\":1,\"ok\":true,\"messages\":{}{}}}",
            messages_json, observations_field
        );
    }
    let root_causes = match &diag.root_cause {
        Some(rc) => format!("[{}]", root_cause_json(rc, copy_location)),
        None => "[]".to_string(),
    };
    format!(
        "{{\"schema_version\":1,\"ok\":false,\"message\":{},\"location\":{},\"root_causes\":{},\"messages\":{}{}}}",
        json_string(diag.message.as_deref().unwrap_or("")),
        location_json(location),
        root_causes,
        messages_json,
        observations_field,
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
        let mut out = "OK: no resolution failure detected".to_string();
        if let Some(obs) = &diag.hidden_cells {
            out.push_str(&hidden_cells_plain_text_note(obs));
        }
        return out;
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
            ResolutionFailureKind::PasteIntoNonAnchorMergedCell {
                dest_addr,
                dest_sheet,
                merged_range,
                ..
            } => {
                out.push_str(&format!(
                    "paste destination {} (sheet '{}') is inside merged range {} but isn't its top-left cell",
                    dest_addr, dest_sheet, rect_addr(*merged_range),
                ));
                if let Some(loc) = copy_location {
                    out.push_str(&format!("\n  copied at {}:{}:{}", loc.file, loc.line, loc.column));
                }
            }
            ResolutionFailureKind::PastePartialMergedRange {
                dest_addr,
                dest_sheet,
                conflicts,
                ..
            } => {
                let addrs: Vec<String> = conflicts.iter().map(|r| rect_addr(*r)).collect();
                out.push_str(&format!(
                    "paste destination {} (sheet '{}') partially overlaps merged range(s): {}",
                    dest_addr, dest_sheet, addrs.join(", "),
                ));
                if let Some(loc) = copy_location {
                    out.push_str(&format!("\n  copied at {}:{}:{}", loc.file, loc.line, loc.column));
                }
            }
            ResolutionFailureKind::PasteMergeLayoutMismatch {
                source_addr,
                source_sheet,
                dest_addr,
                dest_sheet,
                conflicts,
                ..
            } => {
                let addrs: Vec<String> = conflicts.iter().map(|r| rect_addr(*r)).collect();
                out.push_str(&format!(
                    "copy source {} (sheet '{}') and paste destination {} (sheet '{}') have different merged-cell layouts: {}",
                    source_addr, source_sheet, dest_addr, dest_sheet, addrs.join(", "),
                ));
                if let Some(loc) = copy_location {
                    out.push_str(&format!("\n  copied at {}:{}:{}", loc.file, loc.line, loc.column));
                }
            }
            ResolutionFailureKind::MultiAreaToSingleAreaPaste { source_areas, destination_areas }
            | ResolutionFailureKind::MultiAreaCountMismatch { source_areas, destination_areas }
            | ResolutionFailureKind::MultiAreaPasteUnsupported { source_areas, destination_areas } => {
                let src: Vec<String> = source_areas.iter().map(|r| area_addr(*r)).collect();
                let dst: Vec<String> = destination_areas.iter().map(|r| area_addr(*r)).collect();
                out.push_str(&format!(
                    "copy source has {} area(s) ({}), paste destination has {} area(s) ({})",
                    src.len(),
                    src.join(", "),
                    dst.len(),
                    dst.join(", "),
                ));
            }
            ResolutionFailureKind::MultiAreaShapeMismatch {
                area_index,
                source_area,
                destination_area,
            } => {
                out.push_str(&format!(
                    "area {} shape mismatch: source {} ({}x{}) vs destination {} ({}x{})",
                    area_index,
                    area_addr(*source_area),
                    source_area.rows(),
                    source_area.cols(),
                    area_addr(*destination_area),
                    destination_area.rows(),
                    destination_area.cols(),
                ));
            }
        }
        for s in rc.suggestions() {
            out.push_str(&format!("\n  suggestion: {}", s));
        }
    }
    if let Some(obs) = &diag.hidden_cells {
        out.push_str(&hidden_cells_plain_text_note(obs));
    }
    out
}

/// Plain-text rendering of the `RANGE_CONTAINS_HIDDEN_CELLS` observation
/// (Milestone B7b), shared by both `to_plain_text`'s success and failure
/// branches.
fn hidden_cells_plain_text_note(obs: &HiddenCellsObservation) -> String {
    format!(
        "\n\nRANGE_CONTAINS_HIDDEN_CELLS: {} ({}) has {} of {} cells visible \
         (hidden rows: {}; hidden columns: {})",
        obs.address,
        obs.sheet,
        obs.visible_cells,
        obs.total_cells,
        if obs.hidden_rows.is_empty() {
            "none".to_string()
        } else {
            obs.hidden_rows.iter().map(row_interval_addr).collect::<Vec<_>>().join(", ")
        },
        if obs.hidden_columns.is_empty() {
            "none".to_string()
        } else {
            obs.hidden_columns.iter().map(col_interval_addr).collect::<Vec<_>>().join(", ")
        },
    )
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

    // ── Milestone B7a: multi-area Range foundation ───────────────────────────

    #[test]
    fn multi_area_to_single_area_paste_reports_the_completion_condition_json() {
        // The B7a plan's own completion-condition scenario.
        let out_path = std::env::temp_dir().join("elixcee_diagnose_multi_area_to_single.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from(
            "Sub Main()\n    Range(\"A1:A10,C1:C10\").Copy\n    Range(\"E1:F10\").PasteSpecial\nEnd Sub\n",
        );
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        let rc = diag.root_cause.as_ref().expect("should classify a root cause");
        assert_eq!(rc.code, "MULTI_AREA_TO_SINGLE_AREA_PASTE");
        assert_eq!(
            rc.suggestions(),
            vec![
                "paste each source area separately".to_string(),
                "copy a contiguous rectangular range".to_string(),
                "use destination areas with matching count and shapes".to_string(),
            ]
        );
        let json = to_json(&diag, None, None);
        assert!(json.contains("\"code\":\"MULTI_AREA_TO_SINGLE_AREA_PASTE\""));
        assert!(json.contains("\"certainty\":\"definite\""));
        assert!(json.contains(
            "\"source_areas\":[{\"address\":\"A1:A10\",\"rows\":10,\"columns\":1},\
             {\"address\":\"C1:C10\",\"rows\":10,\"columns\":1}]"
        ));
        assert!(json.contains(
            "\"destination_areas\":[{\"address\":\"E1:F10\",\"rows\":10,\"columns\":2}]"
        ));
        assert!(json.contains("paste each source area separately"));
    }

    #[test]
    fn multi_area_count_mismatch_reports_a_root_cause() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_multi_area_count_mismatch.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from(
            "Sub Main()\n    Range(\"A1:A10,C1:C10,E1:E10\").Copy\n    \
             Range(\"G1:G10,I1:I10\").PasteSpecial\nEnd Sub\n",
        );
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        let rc = diag.root_cause.as_ref().expect("should classify a root cause");
        assert_eq!(rc.code, "MULTI_AREA_COUNT_MISMATCH");
        let json = to_json(&diag, None, None);
        assert!(json.contains("\"source_areas\":[{\"address\":\"A1:A10\""));
        assert!(json.contains("\"destination_areas\":[{\"address\":\"G1:G10\""));
        assert!(json.contains("match the destination area count to the source area count"));
    }

    #[test]
    fn multi_area_shape_mismatch_reports_the_first_mismatching_area() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_multi_area_shape_mismatch.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from(
            "Sub Main()\n    Range(\"A1:A10,C1:C10\").Copy\n    \
             Range(\"G1:G10,I1:J10\").PasteSpecial\nEnd Sub\n",
        );
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        let rc = diag.root_cause.as_ref().expect("should classify a root cause");
        assert_eq!(rc.code, "MULTI_AREA_SHAPE_MISMATCH");
        let json = to_json(&diag, None, None);
        assert!(json.contains("\"area_index\":2"));
        assert!(json.contains("\"source_area\":{\"address\":\"C1:C10\",\"rows\":10,\"columns\":1}"));
        assert!(json.contains(
            "\"destination_area\":{\"address\":\"I1:J10\",\"rows\":10,\"columns\":2}"
        ));
        assert!(json.contains("resize destination area 2 to match the source area's shape"));
    }

    #[test]
    fn multi_area_paste_unsupported_reports_a_root_cause_for_a_single_to_multi_paste() {
        let out_path = std::env::temp_dir().join("elixcee_diagnose_multi_area_unsupported.xlsx");
        build_workbook(out_path.to_str().unwrap());
        let programs = programs_from(
            "Sub Main()\n    Range(\"A1:B10\").Copy\n    Range(\"E1:E10,G1:G10\").PasteSpecial\nEnd Sub\n",
        );
        let diag = run_diagnosis(&programs, out_path.to_str().unwrap(), "Main").unwrap();
        assert!(!diag.ok);
        let rc = diag.root_cause.as_ref().expect("should classify a root cause");
        assert_eq!(rc.code, "MULTI_AREA_PASTE_UNSUPPORTED");
        let json = to_json(&diag, None, None);
        assert!(json.contains("\"source_areas\":[{\"address\":\"A1:B10\",\"rows\":10,\"columns\":2}]"));
        assert!(json.contains("\"destination_areas\":[{\"address\":\"E1:E10\""));
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

    // ── Milestone B6c2: merged-cell-aware Paste diagnostics ─────────────────
    //
    // Unlike every kind above, these can't be reproduced through
    // `run_diagnosis`'s full file-loading pipeline: there's no VBA
    // construct to create a merge, and `save_workbook` doesn't write
    // `<mergeCells>` (out of scope — reading merges was this milestone's
    // job, not writing them). So these test `RootCause`/`root_cause_json`
    // directly against a hand-built `ResolutionFailureKind`, the same
    // pattern the VM-level tests use to bypass the lack of a `.Merge`
    // VBA statement.

    #[test]
    fn paste_into_non_anchor_merged_cell_reports_the_correct_root_cause() {
        let rc = RootCause::from_kind(ResolutionFailureKind::PasteIntoNonAnchorMergedCell {
            dest_addr: "C1".to_string(),
            dest_sheet: "sheet1".to_string(),
            merged_range: ((1, 2), (1, 4)),
            copy_span: None,
        });
        assert_eq!(rc.code, "PASTE_INTO_NON_ANCHOR_MERGED_CELL");
        assert_eq!(
            rc.suggestions(),
            vec![
                "unmerge B1:D1 before pasting".to_string(),
                "or paste into its top-left cell B1 instead".to_string(),
            ]
        );
        let json = root_cause_json(&rc, None);
        assert!(json.contains("PASTE_INTO_NON_ANCHOR_MERGED_CELL"));
        assert!(json.contains("\"dest_addr\":\"C1\""));
        assert!(json.contains("\"merged_range\":\"B1:D1\""));
    }

    #[test]
    fn paste_partial_merged_range_reports_the_correct_root_cause() {
        let rc = RootCause::from_kind(ResolutionFailureKind::PastePartialMergedRange {
            dest_addr: "B1:C1".to_string(),
            dest_sheet: "sheet1".to_string(),
            conflicts: vec![((1, 2), (1, 4))],
            copy_span: None,
        });
        assert_eq!(rc.code, "PASTE_PARTIAL_MERGED_RANGE");
        assert_eq!(
            rc.suggestions(),
            vec![
                "unmerge B1:D1 before pasting, or resize the destination to fully contain it"
                    .to_string()
            ]
        );
        let json = root_cause_json(&rc, None);
        assert!(json.contains("PASTE_PARTIAL_MERGED_RANGE"));
        assert!(json.contains("\"conflicts\":[\"B1:D1\"]"));
    }

    #[test]
    fn paste_merge_layout_mismatch_reports_the_correct_root_cause_directly() {
        let rc = RootCause::from_kind(ResolutionFailureKind::PasteMergeLayoutMismatch {
            source_addr: "A1:C10".to_string(),
            source_sheet: "sheet1".to_string(),
            dest_addr: "E1:G10".to_string(),
            dest_sheet: "sheet1".to_string(),
            conflicts: vec![((1, 5), (1, 7))],
            copy_span: None,
        });
        assert_eq!(rc.code, "PASTE_MERGE_LAYOUT_MISMATCH");
        assert_eq!(
            rc.suggestions(),
            vec![
                "unmerge E1:G1 before pasting".to_string(),
                "or make the source and destination merge layouts identical".to_string(),
            ]
        );
        let json = root_cause_json(&rc, None);
        assert!(json.contains("PASTE_MERGE_LAYOUT_MISMATCH"));
        assert!(json.contains("\"source_addr\":\"A1:C10\""));
        assert!(json.contains("\"dest_addr\":\"E1:G10\""));
        assert!(json.contains("\"conflicts\":[\"E1:G1\"]"));
    }

    // ── Milestone B6d: root_causes_json (the diagnose-workbook entry point) ─

    #[test]
    fn root_causes_json_renders_an_empty_array_for_none() {
        assert_eq!(root_causes_json(None), "[]");
    }

    #[test]
    fn root_causes_json_matches_diagnosis_to_json_exactly_for_the_same_kind() {
        let kind = ResolutionFailureKind::SheetProtected {
            sheet: "sheet1".to_string(),
        };
        let diag = Diagnosis {
            ok: false,
            message: Some("Cannot write: sheet is protected".to_string()),
            span: None,
            copy_span: None,
            root_cause: Some(RootCause::from_kind(kind.clone())),
            hidden_cells: None,
            messages: vec![],
        };
        let full_json = to_json(&diag, None, None);
        let root_causes_fragment = format!(
            "\"root_causes\":{}",
            root_causes_json(Some(&kind))
        );
        assert!(
            full_json.contains(&root_causes_fragment),
            "expected {} to contain {}",
            full_json,
            root_causes_fragment
        );
    }

    // ── Milestone B7b: hidden row/column evidence ────────────────────────────

    fn sample_observation() -> HiddenCellsObservation {
        // The B7b plan's own worked example — note its `visible_cells`
        // (172, hand-verified: (100-14)*(3-1)) intentionally differs from
        // the request's own inconsistent `162`; see the plan's decision 5.
        HiddenCellsObservation {
            sheet: "sheet1".to_string(),
            address: "A1:C100".to_string(),
            rows: 100,
            columns: 3,
            hidden_rows: vec![
                Interval { start: 11, end: 14 },
                Interval { start: 30, end: 39 },
            ],
            hidden_columns: vec![Interval { start: 2, end: 2 }],
            total_cells: 300,
            visible_cells: 172,
        }
    }

    #[test]
    fn observations_json_renders_an_empty_array_for_none() {
        assert_eq!(observations_json(None), "[]");
    }

    #[test]
    fn observations_json_renders_the_completion_condition_shape() {
        let json = observations_json(Some(&sample_observation()));
        assert!(json.contains("\"code\":\"RANGE_CONTAINS_HIDDEN_CELLS\""));
        assert!(json.contains("\"certainty\":\"observed\""));
        assert!(json.contains(
            "\"range\":{\"sheet\":\"sheet1\",\"address\":\"A1:C100\",\"rows\":100,\"columns\":3}"
        ));
        assert!(json.contains("\"hidden_rows\":[\"11:14\",\"30:39\"]"));
        assert!(json.contains("\"hidden_columns\":[\"B:B\"]"));
        assert!(json.contains("\"total_cells\":300"));
        assert!(json.contains("\"visible_cells\":172"));
    }

    fn diagnosis_with_hidden_cells(ok: bool, hidden_cells: Option<HiddenCellsObservation>) -> Diagnosis {
        Diagnosis {
            ok,
            message: if ok { None } else { Some("Sheet 'x' not found".to_string()) },
            span: None,
            copy_span: None,
            root_cause: if ok {
                None
            } else {
                Some(RootCause::from_kind(ResolutionFailureKind::WorksheetNotFound(
                    ResolutionEvidence {
                        expression: "Worksheets(\"x\")".to_string(),
                        requested: "x".to_string(),
                        available: vec![],
                        suggested: None,
                    },
                )))
            },
            hidden_cells,
            messages: vec![],
        }
    }

    #[test]
    fn to_json_omits_the_observations_field_when_there_is_nothing_to_observe() {
        assert!(!to_json(&diagnosis_with_hidden_cells(true, None), None, None)
            .contains("observations"));
        assert!(!to_json(&diagnosis_with_hidden_cells(false, None), None, None)
            .contains("observations"));
    }

    #[test]
    fn to_json_includes_observations_on_a_successful_run() {
        let json = to_json(
            &diagnosis_with_hidden_cells(true, Some(sample_observation())),
            None,
            None,
        );
        assert!(json.contains("\"ok\":true"));
        assert!(json.contains("\"observations\":[{\"code\":\"RANGE_CONTAINS_HIDDEN_CELLS\""));
    }

    #[test]
    fn to_json_includes_observations_alongside_root_causes_on_a_failing_run() {
        let json = to_json(
            &diagnosis_with_hidden_cells(false, Some(sample_observation())),
            None,
            None,
        );
        assert!(json.contains("\"root_causes\":[{\"code\":\"WORKSHEET_NOT_FOUND\""));
        assert!(json.contains("\"observations\":[{\"code\":\"RANGE_CONTAINS_HIDDEN_CELLS\""));
    }

    #[test]
    fn to_plain_text_appends_the_hidden_cells_note_on_success() {
        let text = to_plain_text(
            &diagnosis_with_hidden_cells(true, Some(sample_observation())),
            None,
            None,
        );
        assert!(text.starts_with("OK: no resolution failure detected"));
        assert!(text.contains("RANGE_CONTAINS_HIDDEN_CELLS"));
        assert!(text.contains("172"));
    }

    #[test]
    fn to_plain_text_appends_the_hidden_cells_note_on_failure() {
        let text = to_plain_text(
            &diagnosis_with_hidden_cells(false, Some(sample_observation())),
            None,
            None,
        );
        assert!(text.contains("WORKSHEET_NOT_FOUND"));
        assert!(text.contains("RANGE_CONTAINS_HIDDEN_CELLS"));
    }
}
