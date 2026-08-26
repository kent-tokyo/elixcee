use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};

use crate::check;
use crate::formula;
use crate::parser::ast::{
    ArrayDim, Axis, CalcModeValue, CaseMatch, Expr, FuncDef, ObjectExpr, Program, SourceSpan,
    SpannedStmt, Stmt, SubDef, VbaBinOp, WithMember, WithTarget, XlDir, XlEndProp,
};
use crate::parser::{self, EntrypointResolution};
use crate::reader::{self, SheetCell, WorkbookSheet};

/// `ExcelError`/`Variant`/`CellContent`/`serial_to_display` and the range
/// address helpers below are physically defined in `elixcee-types` (Phase
/// 2A) — re-exported here so every existing `vm::X` / `crate::vm::X`
/// reference across the codebase keeps resolving unchanged.
pub use crate::types::{
    ArrayBound, CellContent, ExcelError, MAX_ARRAY_ELEMENTS, Variant, VbaArray, parse_cell_addr,
    parse_range_addr, serial_to_display,
};

/// A procedure's own `On Error` state — real VBA scopes this per Sub/
/// Function, not globally, which is exactly what `Vm::call_stack` models by
/// giving each `CallFrame` its own `ErrorMode`. Consumed (reset to
/// `Disabled`) the moment a `GoTo` handler actually fires, matching real
/// VBA: a second failure while already inside the handler (without the
/// handler itself running a fresh `On Error`) propagates to the caller
/// instead of re-entering the same handler.
#[derive(Debug, Clone, PartialEq)]
enum ErrorMode {
    Disabled,
    ResumeNext,
    GoTo(String),
}

/// One entry in `Vm::call_stack` — the currently-executing procedure's name
/// (for future diagnostics; not yet surfaced anywhere, so genuinely unread —
/// confirmed by removing the `allow` and rebuilding, which reintroduces the
/// dead-code warning) and its own `ErrorMode`.
#[derive(Debug, Clone)]
struct CallFrame {
    #[allow(dead_code)]
    procedure_name: String,
    error_mode: ErrorMode,
}

/// The full set of values `Err.Raise` supplies — real VBA's own five
/// `Err` properties, minus `Number`/`Description`'s own doc (see
/// `Vm::err_number`) since those two aren't new here.
#[derive(Debug, Clone)]
struct RaisedError {
    number: i64,
    description: String,
    source: String,
    help_file: String,
    help_context: i64,
}

/// Evidence for a resolution failure (Milestone B6a's `diagnose`
/// subcommand) — the requested key, what was actually available, and (for
/// name lookups) the closest match by edit distance, if any.
#[derive(Debug, Clone, PartialEq)]
pub struct ResolutionEvidence {
    pub expression: String,
    pub requested: String,
    pub available: Vec<String>,
    pub suggested: Option<String>,
}

/// Why a VBA "Subscript out of range" (Error 9)-shaped operation failed,
/// classified with evidence instead of only a formatted message string.
/// Set on `Vm::last_resolution_failure` immediately before the matching
/// `Err(String)` is returned — a side channel, same pattern as
/// `current_span`/`take_messages()` — so `diagnose` (or any caller) can
/// read structured detail after `run_sub`/`run_sub_multi` fails, while
/// every other caller that only wants the plain string is unaffected.
#[derive(Debug, Clone, PartialEq)]
pub enum ResolutionFailureKind {
    WorksheetNotFound(ResolutionEvidence),
    WorkbookNotFound(ResolutionEvidence),
    ArrayIndexOutOfBounds {
        name: String,
        index: i64,
        lower: i64,
        upper: i64,
    },
    /// A `.Paste`/`.PasteSpecial` destination's shape doesn't match what
    /// was copied (Milestone B6b) — `dest_row1`/`dest_col1` are the
    /// destination's 1-based anchor cell, used to render a "resize to..."
    /// suggestion. `copy_span` is the *Copy* statement's span (the Paste
    /// statement's own span is already `Vm::current_span()` by the time
    /// this fires), so a diagnosis can point at both statements.
    PasteShapeMismatch {
        source_addr: String,
        source_rows: u32,
        source_cols: u32,
        dest_addr: String,
        dest_rows: u32,
        dest_cols: u32,
        dest_row1: u32,
        dest_col1: u32,
        transpose: bool,
        copy_span: Option<SourceSpan>,
    },
    /// A `.Paste`/`.PasteSpecial` was attempted with nothing copied — either
    /// no prior `.Copy` ran, or `Application.CutCopyMode` was cleared since
    /// (Milestone B6b).
    PasteWithoutCopy {
        dest_addr: String,
    },
    /// A cell-mutating statement targeted a sheet that's been `.Protect`ed
    /// (Milestone B6c) — real Excel blocks any write/clear/insert/sort/
    /// paste/delete on a protected sheet, unconditionally.
    SheetProtected {
        sheet: String,
    },
    /// A `.Paste`/`.PasteSpecial` anchor cell is a covered (non-top-left)
    /// cell of an existing merged range on the destination sheet
    /// (Milestone B6c2) — real Excel refuses to paste directly into one.
    /// `merged_range` is the raw rect (not pre-formatted), matching
    /// `PasteShapeMismatch`'s `dest_row1`/`dest_col1` convention of leaving
    /// address formatting to `diagnose.rs`'s own `col_to_letters`.
    PasteIntoNonAnchorMergedCell {
        dest_addr: String,
        dest_sheet: String,
        merged_range: MergeRect,
        copy_span: Option<SourceSpan>,
    },
    /// A `.Paste`/`.PasteSpecial` destination range only partially
    /// overlaps one or more merged ranges on the destination sheet
    /// (Milestone B6c2) — pasting would split an existing merge.
    PastePartialMergedRange {
        dest_addr: String,
        dest_sheet: String,
        conflicts: Vec<MergeRect>,
        copy_span: Option<SourceSpan>,
    },
    /// The copied range and the paste destination have matching row/column
    /// counts (Milestone B6b's shape check passed) but differ in which
    /// relative cells are merged (Milestone B6c2) — e.g. the destination
    /// has a merged row where the source has none.
    PasteMergeLayoutMismatch {
        source_addr: String,
        source_sheet: String,
        dest_addr: String,
        dest_sheet: String,
        conflicts: Vec<MergeRect>,
        copy_span: Option<SourceSpan>,
    },
    /// A multi-area source (`Range("A1:A3,C1:C3")`, `Areas.Count > 1`) was
    /// pasted into a single-area destination (Milestone B7a) — real Excel
    /// can't determine a unique area-to-area correspondence. This is the
    /// foundation milestone's completion-condition scenario. Raw `Rect`s,
    /// not pre-formatted addresses — same precedent as
    /// `PasteIntoNonAnchorMergedCell`'s `MergeRect`; `diagnose.rs` formats
    /// them (via its existing private `rect_addr`) at JSON-serialization
    /// time.
    MultiAreaToSingleAreaPaste {
        source_areas: Vec<Rect>,
        /// Always exactly 1 element — kept as a `Vec` for field-shape
        /// symmetry with the other 3 multi-area kinds below.
        destination_areas: Vec<Rect>,
    },
    /// Both source and destination are multi-area, but their `Areas.Count`
    /// differ (Milestone B7a).
    MultiAreaCountMismatch {
        source_areas: Vec<Rect>,
        destination_areas: Vec<Rect>,
    },
    /// Both source and destination are multi-area with matching
    /// `Areas.Count`, but at least one area pair (by position) differs in
    /// rows/columns (Milestone B7a). Reports the first mismatching pair;
    /// `area_index` is 1-based, matching VBA's own `Areas(1)`-style
    /// indexing.
    MultiAreaShapeMismatch {
        area_index: usize,
        source_area: Rect,
        destination_area: Rect,
    },
    /// A multi-area paste shape that isn't diagnosed as structurally
    /// *wrong* but also still isn't executed: a single-area source into a
    /// multi-area destination, or (as of Milestone B7c item 5) the
    /// opposite, a multi-area source into a single-area destination. Real
    /// Excel would complete either of these; elixcee only executes the one
    /// shape both sides are multi-area with matching `Areas.Count` and
    /// matching per-area shapes (see `do_paste`'s B7c comment) — this
    /// variant reports the remaining limitation plainly rather than
    /// silently doing nothing or misreporting a mismatch that isn't there.
    MultiAreaPasteUnsupported {
        source_areas: Vec<Rect>,
        destination_areas: Vec<Rect>,
    },
}

/// The VM's clipboard state, populated by `.Copy` and consumed by
/// `.Paste`/`.PasteSpecial` (Milestone B6b). Values are snapshotted at copy
/// time (`cells`), not re-read from the source range at paste time — this
/// matches real Excel's copy-then-mutate-then-paste semantics now that Copy
/// and Paste can be separate statements.
#[derive(Debug, Clone)]
struct ClipboardState {
    source_addr: String,
    /// The sheet `.Copy` ran against (Milestone B6c2) — `.Copy` always
    /// targets `self.active_sheet`, but nothing captured that name before
    /// this was needed to look up the source sheet's merged ranges.
    src_sheet: String,
    rows: u32,
    cols: u32,
    cells: Vec<Vec<Variant>>, // [row][col], 0-based offsets from the source's top-left
    span: SourceSpan,
    /// Full area geometry of what was copied (Milestone B7a) — `rows`/
    /// `cols` above are only the *first* area's dimensions, kept so the
    /// existing `areas.len() == 1` fast path (every pre-B7a macro) is
    /// byte-identical to before. `cells` is only ever populated when
    /// `areas.len() == 1`: a multi-area paste is diagnose-only in v1 (see
    /// `do_paste`) and never reads per-area cell values, so none are
    /// snapshotted for `areas.len() > 1`.
    areas: Vec<Rect>,
    /// Every area's own cell snapshot, `[area_index][row][col]`, parallel
    /// to `areas` (Milestone B7c item 5) — unlike `cells` (first-area-only,
    /// and only when `areas.len() == 1`, an existing-test invariant left
    /// untouched), this is always fully populated, feeding the one
    /// multi-area Paste shape that's now actually executed (matching
    /// `Areas.Count`s, matching per-area shapes) rather than only
    /// diagnosed — see `do_paste`.
    area_cells: Vec<Vec<Vec<Variant>>>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CalculationMode {
    Automatic,
    Manual,
}

/// Signals emitted by Exit For / Exit Do / Exit Sub / Exit Function.
#[derive(Debug, Clone, PartialEq)]
pub enum ExitKind {
    For,
    Do,
    Sub,
    Function,
}

/// A 1-based inclusive `((row1,col1),(row2,col2))` rect (Milestone B6c2) —
/// a private per-module alias, not a shared type, matching this codebase's
/// existing per-module `col_to_letters` duplication convention rather than
/// a cross-module `utils` dependency.
type MergeRect = ((u32, u32), (u32, u32));

/// A 1-based inclusive rectangular area (Milestone B7a) — the multi-area
/// foundation's basic building block. Same bounds convention as
/// `MergeRect`/`SheetRange`, but a named struct (not a bare tuple) since
/// `RangeRef` needs a `Vec` of these.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Rect {
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

impl Rect {
    pub fn rows(&self) -> u32 {
        self.end_row - self.start_row + 1
    }

    pub fn cols(&self) -> u32 {
        self.end_col - self.start_col + 1
    }
}

/// A possibly-disjoint Excel range on one sheet — one or more rectangular
/// areas (Milestone B7a's foundation for `Areas.Count`-shaped ranges like
/// `"A1:A3,C1:C3"`, and eventually a filtered `SpecialCells
/// (xlCellTypeVisible)` result in B7c). Every existing single-rect call
/// site (`SheetRange`, `parse_range_addr`, every VBA statement except
/// Copy/Paste) is unchanged and keeps using the bare-tuple representation
/// — only Copy/Paste resolve through `Rect`/`RangeRef`, and only their
/// `areas.len() == 1` fast path is exercised by pre-B7a macros.
#[derive(Debug, Clone, PartialEq)]
pub struct RangeRef {
    pub sheet: String,
    pub areas: Vec<Rect>,
}

impl RangeRef {
    pub fn single(sheet: String, rect: Rect) -> Self {
        RangeRef {
            sheet,
            areas: vec![rect],
        }
    }

    pub fn is_single_area(&self) -> bool {
        self.areas.len() == 1
    }

    pub fn single_rect(&self) -> Option<&Rect> {
        if self.is_single_area() {
            self.areas.first()
        } else {
            None
        }
    }

    pub fn cell_count(&self) -> usize {
        self.areas
            .iter()
            .map(|r| r.rows() as usize * r.cols() as usize)
            .sum()
    }
}

/// A 1-based inclusive `[start, end]` row or column interval (Milestone
/// B7b). Deliberately one type for both rows and columns (structurally
/// identical), not two.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Interval {
    pub start: u32,
    pub end: u32,
}

impl Interval {
    /// Clips this interval to `[lo, hi]`, returning `None` if they don't
    /// overlap at all (Milestone B7b).
    fn clip(&self, lo: u32, hi: u32) -> Option<Interval> {
        let start = self.start.max(lo);
        let end = self.end.min(hi);
        if start <= end {
            Some(Interval { start, end })
        } else {
            None
        }
    }
}

/// Splits `[lo, hi]` into its maximal visible-only sub-intervals, given a
/// set of hidden intervals (not necessarily sorted, clipped, or
/// non-overlapping) on the same axis (Milestone B7c's `SpecialCells
/// (xlCellTypeVisible)`). Pure interval math shared by both the row and
/// column axis — see `Vm::visible_areas`.
fn visible_runs(lo: u32, hi: u32, hidden: &[Interval]) -> Vec<Interval> {
    let mut clipped: Vec<Interval> = hidden.iter().filter_map(|iv| iv.clip(lo, hi)).collect();
    clipped.sort_by_key(|iv| iv.start);
    let mut runs = Vec::new();
    let mut cursor = lo;
    for h in &clipped {
        if h.start > cursor {
            runs.push(Interval {
                start: cursor,
                end: h.start - 1,
            });
        }
        cursor = cursor.max(h.end.saturating_add(1));
        if cursor > hi {
            break;
        }
    }
    if cursor <= hi {
        runs.push(Interval {
            start: cursor,
            end: hi,
        });
    }
    runs
}

/// Which rows/columns are hidden on one sheet (Milestone B7b), read from
/// XLSX's `<row hidden="1">`/`<col min=".." max=".." hidden="1">` (ODS is
/// deferred — see `docs/agent-contract.md`). Threaded into
/// `Vm.sheet_visibility` the same way `merged_ranges` already is.
#[derive(Debug, Clone, Default)]
pub struct SheetVisibility {
    pub hidden_rows: Vec<Interval>,
    pub hidden_columns: Vec<Interval>,
}

/// A loaded sheet's origin facts from its source file (0.10.0-A) — threaded from
/// `reader::WorkbookSheet`'s own `sheet_id`/`workbook_rel_id`/`source_part_name` fields
/// (see their doc comments) the same way `merged_ranges`/`SheetVisibility` already are,
/// via `Vm.worksheet_origins`. Exists so `save_xlsx_impl` (`src/lib.rs`) can preserve a
/// sheet's original `sheetId` on save instead of unconditionally renumbering it from
/// current position — see `docs/xlsx-worksheet-preservation-0.10.0-design.md` §2/§6 for
/// why that matters (it's the same `sheetId` `snapshot.rs`'s `stable_id` already treats
/// as the one cross-save-stable identifier a real `.xlsx` can offer).
///
/// Deliberately has no separate rename-stable "VM internal identity" field:
/// `Vm::rename_sheet` re-keys this map (and every other per-sheet map --
/// `merged_ranges`, `sheet_visibility`, `cell_style_indices`, `cell_number_formats`,
/// `protected_sheets`) atomically in one call, including updating this struct's own
/// `original_display_name` to the new name, so the sheet's lowercased name stays a
/// valid stable key across a rename -- no separate identity needed because the re-key
/// is atomic, not because rename doesn't happen.
#[derive(Debug, Clone, Default)]
pub struct WorksheetOrigin {
    pub original_sheet_id: Option<String>,
    pub original_workbook_rel_id: Option<String>,
    pub original_part_name: Option<String>,
    /// The sheet's name exactly as written in the source `<sheet name="...">`
    /// (not lowercased) -- every other per-sheet `Vm` map, including this
    /// struct's own home (`Vm.worksheet_origins`), is keyed by the
    /// lowercased name, so without this field a save had no way to recover
    /// the original casing: `Sheet1` round-tripped as `sheet1`, a visible
    /// tab-label change on every save with zero macro involvement (same
    /// class of bug as `Vm::sheet_order`, found via the same fixture).
    /// `Ensure_sheet` (backing both `Sheets.Add` and Python's `set_sheet()`) also
    /// populates this for a sheet with no loaded-file origin, from the caller's
    /// as-written name -- GitHub #2 was exactly this field being `None` for such a
    /// sheet and the writer falling back to the lowercased key instead. `None` here
    /// in practice only for `Vm::new()`'s own provisional default sheet, which
    /// `populate_from_sheets` always replaces before any save can observe it.
    pub original_display_name: Option<String>,
}

/// Evidence for the `RANGE_CONTAINS_HIDDEN_CELLS` observation (Milestone
/// B7b) — computed on demand by `Vm::hidden_cells_observation`, not stored
/// as a side channel. `hidden_rows`/`hidden_columns` are already clipped to
/// this range (not the sheet's full hidden-row/column list).
#[derive(Debug, Clone, PartialEq)]
pub struct HiddenCellsObservation {
    pub sheet: String,
    pub address: String,
    pub rows: u32,
    pub columns: u32,
    pub hidden_rows: Vec<Interval>,
    pub hidden_columns: Vec<Interval>,
    pub total_cells: u64,
    pub visible_cells: u64,
}

/// An object reference held by a `Set`-assigned variable (Milestone B7c;
/// `Worksheet`/`Workbook` added Phase 2C items 7/8). Kept as an enum (not a
/// bare `RangeRef`) so `Set ws = ActiveSheet`/`Set wb = ThisWorkbook` can
/// live in the same namespace as `Set rng = Range(...)`. Deliberately *not*
/// a variant on the shared `Variant` type from `elixcee-types` — see
/// `Vm::object_variables`'s doc for why.
#[derive(Debug, Clone, PartialEq)]
pub enum ObjectRef {
    Range(RangeRef),
    /// `Set ws = ActiveSheet` — the lowercase sheet key `ws` now refers to
    /// (a snapshot of whichever sheet was active at `Set`-time, same as
    /// real VBA fixing a Worksheet reference's identity at assignment, not
    /// at each later access — same convention `ObjectRef::Range`'s doc on
    /// `RangeLit` already established for `Set`).
    Worksheet(String),
    /// `Set wb = ThisWorkbook` (or `= ActiveWorkbook`). No payload: elixcee
    /// only ever has one workbook loaded (see `Expr::WorkbookQualifiedSheet`
    /// 's doc), so there's nothing to distinguish — this variant exists
    /// only so the *variable* is a real Workbook-typed object reference
    /// (`wb.Worksheets(...)`/`.Sheets(...)` resolve through it) instead of
    /// the pre-Phase-2C silent no-op.
    Workbook,
    /// The null object reference: an object variable that exists but holds
    /// no live object. Two ways in — `Dim r As Range` (declared, never
    /// `Set`) and an explicit `Set r = Nothing` — and real VBA can't tell
    /// them apart either (`r Is Nothing` is `True` for both). Distinct from
    /// *absence* from `object_variables`, which means "not an object
    /// variable at all" (an ordinary scalar/UDT name, or an undeclared one)
    /// and must keep its pre-existing non-object behavior: making a missing
    /// key mean Nothing would turn every `p.field = 1` on an
    /// undeclared-but-legal UDT into a runtime error.
    Nothing,
}

/// Real VBA's error 91 text, raised for any member access through an object
/// variable that holds no live reference. One constant rather than the
/// literal repeated per call site, so the wording can't drift between the
/// read path, the write path and the `.Copy`/sheet-qualifier paths.
pub const OBJECT_NOT_SET: &str = "Object variable or With block variable not set";

/// Default `Err.Description` text for a well-known VBA error number — what
/// real VBA fills in automatically when `Err.Raise <number>` is called
/// without an explicit description. Only covers the numbers
/// `classify_vba_error_number` below confidently matches; anything else
/// gets 1004's own generic text, real VBA's own catch-all.
fn default_description_for_vba_error_number(number: i64) -> &'static str {
    match number {
        5 => "Invalid procedure call or argument",
        6 => "Overflow",
        9 => "Subscript out of range",
        11 => "Division by zero",
        13 => "Type mismatch",
        91 => OBJECT_NOT_SET,
        94 => "Invalid use of Null",
        _ => "Application-defined or object-defined error",
    }
}

/// Maps one of elixcee's own internal runtime-error message strings to real
/// VBA's `Err.Number`/`Err.Description`. The numbers below are confirmed
/// matches against Microsoft's own long-stable, publicly documented VBA
/// runtime error constants (unchanged since VB6 — a fact independent of
/// this project's own lack of a live Excel/VBA oracle, see `ROADMAP.md`'s
/// "Known gaps" #1). Every other elixcee-internal condition — undefined
/// variable, sheet/sub/workbook not found, and so on — has no single
/// confidently-correct real VBA number (several, like calling an undefined
/// Sub/Function, would actually be a *compile*-time failure in real VBA,
/// never reaching `On Error` at runtime at all — a known, disclosed
/// divergence, not fixed here) and defaults to 1004, real VBA's own
/// generic "Application-defined or object-defined error" — the number it
/// commonly raises for Excel-object-related failures itself.
fn classify_vba_error_number(msg: &str) -> (i64, String) {
    let mapped = if msg == "Division by zero" {
        Some(11)
    } else if msg == "Subscript out of range" {
        Some(9)
    } else if msg == "Type mismatch" {
        Some(13)
    } else if msg == "Invalid procedure call or argument" {
        Some(5)
    } else if msg == "Invalid use of Null" {
        Some(94)
    } else if msg == OBJECT_NOT_SET {
        Some(91)
    } else if msg == "Integer division overflow" {
        // elixcee's own wording (i64-based overflow, not real VBA's native
        // 32-bit Long overflow) — the *number* still matches VBA's own
        // Overflow (6), the description text is kept as elixcee's own
        // rather than substituted, since it wasn't independently confirmed.
        Some(6)
    } else {
        None
    };
    match mapped {
        Some(number) => (number, msg.to_string()),
        None => (1004, msg.to_string()),
    }
}

/// One entry on the VM's runtime `With` stack: the *already-evaluated*
/// target of an active `With` block. Pushed on entry, popped on exit
/// (including on an early `Exit Sub`/`Exit For` or a runtime error), so
/// nesting works and an outer target is restored exactly.
///
/// This is what replaced the parser's old textual With rewrite. The target
/// expression is evaluated once, here, when the block is entered — not
/// re-evaluated per `.member` access, and not substituted into the body's
/// statements at parse time.
#[derive(Debug, Clone, PartialEq)]
enum WithValue {
    /// `With Range("A1")`, `With Cells(r, c)`, or a `Set`-assigned Range
    /// object variable.
    Range(RangeRef),
    /// `With Worksheets("X")`, or a `Set`-assigned Worksheet object
    /// variable. Holds the lowercase sheet key.
    Sheet(String),
    /// `With p` where `p` is a UDT record — `.field` resolves against
    /// `Vm::variables`, not `object_variables`.
    Record(String),
    /// A target elixcee doesn't model (`With Application`, an unset object
    /// variable, an unrecognized header). The body still runs; every bare
    /// `.member` inside it is a no-op, matching the pre-existing behavior
    /// for an unrecognized `With` header.
    Unmodeled,
}

/// Narrows an `ObjectRef` to its `Range` payload, or a descriptive error —
/// `Union`/`.Areas(n)`/`.SpecialCells(...)` only make sense on a Range
/// object; a `Worksheet`/`Workbook` reference reaching one of them (e.g.
/// `Union(ws, Range("A1"))` where `ws` came from `Set ws = ActiveSheet`) is
/// a real VBA type error, not a case any of the three should silently
/// mishandle.
fn expect_range_ref(obj: ObjectRef, context: &str) -> Result<RangeRef, String> {
    match obj {
        ObjectRef::Range(r) => Ok(r),
        ObjectRef::Worksheet(_) => Err(format!(
            "{}: expected a Range object, got a Worksheet reference",
            context
        )),
        ObjectRef::Workbook => Err(format!(
            "{}: expected a Range object, got a Workbook reference",
            context
        )),
        ObjectRef::Nothing => Err(OBJECT_NOT_SET.to_string()),
    }
}

pub struct Vm {
    /// Per-sheet cell storage. Key is sheet name (lowercase for lookup).
    sheets: HashMap<String, HashMap<(u32, u32), CellContent>>,
    /// Lowercased sheet names in the order they entered this `Vm` — source
    /// file order for a loaded workbook (`populate_from_sheets`), append
    /// order for anything created afterward (`ensure_sheet`, the single
    /// choke point behind every sheet-introducing call site: `Sheets.Add`,
    /// `Sheets("New").Cells(...) = ...` auto-vivification, `With
    /// Sheets("New")`). Kept in sync with `sheets`' key set by
    /// `ensure_sheet` (push on first insert) and `Stmt::SheetsDelete`
    /// (remove on delete); `populate_from_sheets` clears both together.
    /// `rename_sheet` swaps a slot's value in place (position preserved);
    /// `move_sheet` removes and re-inserts a slot at a new position (the
    /// only primitive that actually reorders an existing sheet).
    ///
    /// This is deliberately a *second* source of sheet order, kept apart
    /// from `sheet_names()` (which still sorts alphabetically) — real order
    /// exists so `save_xlsx_impl` can write sheets back in their original
    /// tab order (0.10.0 "Lossless Worksheet Preservation": found via a
    /// synthetic "Zebra"/"Alpha" fixture that round-tripped as "Alpha"/
    /// "Zebra"). `sheet_names()`'s alphabetical order is a separate,
    /// already-documented VBA-runtime fidelity gap (`Sheets(i)`/
    /// `Worksheets(i)` numeric indexing, `docs/agent-contract.md`) —
    /// changing that is a distinct decision, not made here.
    pub(crate) sheet_order: Vec<String>,
    /// Currently active sheet name (lowercase).
    pub active_sheet: String,
    pub variables: HashMap<String, Variant>,
    pub calc_mode: CalculationMode,
    pub error_on_msgbox: bool,
    pub print_msgbox: bool,
    /// Every MsgBox message shown during the current `run_sub` call, in
    /// order — populated regardless of `print_msgbox`, so callers (e.g. the
    /// `--json` CLI path) can surface them without relying on stdout
    /// printing. Cleared at the start of each `run_sub`; use
    /// `take_messages()` to read (and drain) it. Private so external callers
    /// can't mutate it directly.
    msgbox_log: Vec<String>,
    /// Span of the statement currently executing (set on every `exec_stmt`
    /// call, at every nesting level) — so a caller can locate where a
    /// runtime error happened via `current_span()` after `run_sub` fails.
    /// `None` until the first statement actually starts executing (e.g. a
    /// "Sub not found" failure happens before this is ever set).
    current_span: Option<SourceSpan>,
    pub exit_flag: Option<ExitKind>,
    /// Pending unconditional jump target (`GoTo <label>`).
    pending_goto: Option<String>,
    /// One frame per currently-executing Sub/Function call, innermost last —
    /// pushed/popped around every `call_sub_def`/`call_func_def` invocation
    /// (including the entrypoint's own call, so this is never empty while a
    /// statement is executing). Each frame's own `error_mode` is what makes
    /// `On Error GoTo`/`Resume Next` a per-procedure scope instead of a
    /// single VM-wide flag a callee could see and mistakenly resolve its
    /// label against — see `exec_body`'s doc comment for the bug this
    /// replaced.
    call_stack: Vec<CallFrame>,
    user_funcs: HashMap<String, FuncDef>,
    user_subs: HashMap<String, SubDef>,
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
    /// Set to true by `move_sheet` and `rename_sheet`; once true, `save_xlsx_impl`
    /// drops any `<definedNames>` passthrough even if no sheet was deleted. Two
    /// distinct reasons feed this one flag: a `<definedName localSheetId="N">` is
    /// positional, so `move_sheet` reordering `sheet_order` can silently invalidate
    /// it; separately, a `<definedName>`'s TEXT can reference a sheet by name (e.g.
    /// `Sheet1!$F$5`), so `rename_sheet` can leave that text dangling even though
    /// nothing about *position* changed. Neither case is worth trying to fix
    /// surgically (rewriting `localSheetId`s, or parsing and rewriting a sheet name
    /// out of arbitrary defined-name formula text) -- dropped wholesale instead,
    /// matching the exact same choice already made for a deleted sheet. Never reset
    /// -- once a session has reordered or renamed a sheet, that workbook's original
    /// `<definedNames>` can no longer be trusted for its remaining lifetime.
    pub(crate) defined_names_may_be_stale: bool,
    /// Wall-clock deadline for loop execution (Milestone B5a's `test-workbook`
    /// timeout guard). `None` (the default) means no limit — every existing
    /// caller (run-mode, `check`, `snapshot`, Python bindings) is unaffected.
    pub deadline: Option<std::time::Instant>,
    /// Counts outer-loop iterations across `For`/`ForEach`/`DoLoop` so the
    /// deadline is only actually checked (a real `Instant::now()` call)
    /// every 256th iteration, not every one.
    loop_iters: u64,
    /// Milestone B6a's `diagnose` opt-in mode. `false` (the default) is
    /// today's existing behavior for every caller (`run`, `check`,
    /// `snapshot`, `test-workbook`, Python bindings): a missing
    /// `Sheets("X")`/`Worksheets("X")` name auto-creates the sheet on write
    /// and silently reads as `Empty`. `true` (set only by `diagnose`) is
    /// the more Excel-faithful behavior a diagnostic tool needs: a missing
    /// name is a resolution failure (see `ResolutionFailureKind`), and `On
    /// Error Resume Next`/`GoTo` no longer swallow/redirect the first error
    /// — it propagates so `diagnose` can report it.
    pub strict_resolution: bool,
    /// Set immediately before returning an `Err` for a resolution failure
    /// (missing worksheet/workbook, array out of bounds) — a side channel
    /// read by `diagnose` after `run_sub`/`run_sub_multi` fails, same
    /// pattern as `current_span`. Cleared at the start of each `run_sub`.
    last_resolution_failure: Option<ResolutionFailureKind>,
    /// The file name (not full path) of the workbook loaded via
    /// `load_workbook_file`, if any — elixcee only ever has one workbook
    /// loaded at a time, so this is only enough to detect a `Workbooks("x")`
    /// reference that doesn't match it (Milestone B6a), not to model real
    /// multi-workbook switching.
    loaded_workbook_name: Option<String>,
    /// Full path of the file loaded via `load_workbook_file` (or, for the
    /// Python bindings, the module-level `load_workbook()` free function), if
    /// any. Used only by `save_xlsx_impl` (`src/lib.rs`) to re-open the
    /// original ZIP for unknown-part passthrough at save time — internal
    /// plumbing between `vm` and `lib.rs`, not a public API.
    pub(crate) loaded_workbook_path: Option<String>,
    /// The clipboard populated by `.Copy` and consumed by
    /// `.Paste`/`.PasteSpecial` (Milestone B6b). `None` initially, and
    /// whenever `Application.CutCopyMode` is set to `False`.
    clipboard: Option<ClipboardState>,
    /// Lowercase sheet keys currently `.Protect`ed (Milestone B6c) — same
    /// key space as `sheets`/`active_sheet`/`ensure_sheet`. Empty by
    /// default; blocks any cell-mutating statement on that sheet.
    protected_sheets: HashSet<String>,
    /// Merged cell ranges per sheet (Milestone B6c2), keyed the same way as
    /// `protected_sheets` — lowercase sheet name → its merged ranges as
    /// `((row1,col1),(row2,col2))`, 1-based inclusive. Populated by
    /// `populate_from_sheets` from the reader's `WorkbookSheet::
    /// merged_ranges`; empty for any sheet built purely in-VBA (`Sheets.Add`
    /// has no merge concept). `pub(crate)`: read directly by
    /// `save_xlsx_impl` (`src/lib.rs`, safe round-trip milestone) to
    /// re-emit `<mergeCells>` on save.
    pub(crate) merged_ranges: HashMap<String, Vec<MergeRect>>,
    /// Hidden row/column metadata per sheet (Milestone B7b), keyed the same
    /// way as `merged_ranges`/`protected_sheets`. Populated by
    /// `populate_from_sheets` from the reader's `WorkbookSheet::
    /// hidden_rows`/`hidden_columns` (XLSX only — ODS is deferred); empty
    /// for any sheet built purely in-VBA. `pub(crate)`: read directly by
    /// `save_xlsx_impl` (`src/lib.rs`, safe round-trip milestone) to
    /// re-emit `<row hidden="1">`/`<col hidden="1">` on save.
    pub(crate) sheet_visibility: HashMap<String, SheetVisibility>,
    /// Per-sheet, per-cell raw `s="N"` style index (Milestone: safe
    /// round-trip), keyed the same way as `merged_ranges`. Populated by
    /// `populate_from_sheets` from the reader's `WorkbookSheet::
    /// raw_style_indices`; used only by `save_xlsx_impl` (`src/lib.rs`) to
    /// re-emit each surviving cell's original style index unchanged --
    /// always correct for this VM to do, since no VBA statement here ever
    /// mutates a cell's style (`Range.Interior.Color =`/`.NumberFormat =`
    /// are explicit no-ops, see the tests of those same names below). Empty
    /// for any sheet built purely in-VBA.
    pub(crate) cell_style_indices: HashMap<String, HashMap<(u32, u32), u32>>,
    /// Per-sheet, per-cell resolved number-format code string (GitHub #4), keyed the
    /// same way as `merged_ranges`/`cell_style_indices`. Populated by
    /// `populate_from_sheets` from the reader's `WorkbookSheet::cell_number_formats`;
    /// read by `get_cell_number_format` (below) so a Python caller can tell a date-
    /// formatted cell's serial number apart from a plain number without a second,
    /// format-aware Excel library. Empty for any sheet built purely in-VBA or loaded
    /// from `.ods`.
    pub(crate) cell_number_formats: HashMap<String, HashMap<(u32, u32), String>>,
    /// Per-sheet origin facts (0.10.0-A), keyed the same way as `merged_ranges`.
    /// Populated unconditionally by `populate_from_sheets` for every sheet that came from
    /// a real `WorkbookSheet` (unlike `merged_ranges`/`sheet_visibility`/
    /// `cell_style_indices`, which only get an entry when there's non-empty data to
    /// store) — an all-`None` `WorksheetOrigin` is itself meaningful here (e.g. an `.ods`
    /// source has no `sheetId` concept at all, distinct from "never loaded"). A sheet
    /// with no entry at all was created purely in-VBA and has no source-file identity to
    /// preserve. `pub(crate)`: read directly by `save_xlsx_impl` (`src/lib.rs`) to
    /// preserve a sheet's original `sheetId` on save instead of renumbering it — see
    /// `WorksheetOrigin`'s own doc comment above.
    pub(crate) worksheet_origins: HashMap<String, WorksheetOrigin>,
    /// `Set`-assigned object variables (Milestone B7c) — lowercase name →
    /// `ObjectRef`, a namespace deliberately separate from `Vm::variables`
    /// (`Variant`s), matching VBA's own distinction between plain `=` and
    /// `Set`. Because the *cells themselves* live in `Vm::sheets` (keyed by
    /// coordinates, not by variable), storing the same `RangeRef` — sheet +
    /// area coordinates, nothing else — in two variables already gives
    /// real `Set` reference semantics for free: a write through one
    /// variable is a write to the shared cell store, immediately visible
    /// through the other. No `Rc<RefCell<_>>` indirection needed.
    object_variables: HashMap<String, ObjectRef>,
    /// Runtime `With` stack — the already-evaluated target of each active
    /// `With` block, innermost last. Pushed on block entry, popped on exit
    /// (including on `Exit Sub`/`Exit For` and on a runtime error, so it
    /// can't leak into whatever runs next), which is what makes nesting
    /// restore the outer target exactly. Every bare `.member` statement or
    /// expression resolves against `last()`, wherever in the AST it sits.
    with_stack: Vec<WithValue>,
    /// Real VBA `Err.Number` — the runtime error number of the most recent
    /// failure caught by `On Error Resume Next`/`On Error GoTo <label>`, or
    /// set directly by `Err.Raise`. 0 means no error since the start of
    /// this `run_sub` call or the last `Err.Clear`. See
    /// `classify_vba_error_number` for which numbers are a confirmed match
    /// against Microsoft's own long-stable, publicly documented VBA
    /// runtime error constants versus a disclosed default (1004) for
    /// conditions elixcee itself raises that don't map to one of those.
    err_number: i64,
    /// Real VBA `Err.Description`, paired with `err_number`.
    err_description: String,
    /// Real VBA `Err.Source` — "" unless the most recent error was raised
    /// via `Err.Raise` with an explicit Source argument. This project
    /// doesn't model a VBA project/class name, so an internally-raised
    /// runtime error (division by zero, subscript out of range, ...) never
    /// sets this to anything but "".
    err_source: String,
    /// Real VBA `Err.HelpFile` — "" unless set by `Err.Raise`.
    err_help_file: String,
    /// Real VBA `Err.HelpContext` — 0 unless set by `Err.Raise`.
    err_help_context: i64,
    /// Set by `Err.Raise` immediately before it returns its `Err(String)`,
    /// so the next `On Error Resume Next`/`On Error GoTo` catch site uses
    /// the number/description/source/help-file/help-context the user
    /// actually raised instead of running that message text back through
    /// `classify_vba_error_number` (which would misclassify it as a
    /// generic 1004, since a raised description is arbitrary user text,
    /// not one of elixcee's own known message strings). Taken (cleared) by
    /// the first catch site that consumes it.
    pending_raised_error: Option<RaisedError>,
    /// The current module's `Option Base` value (real VBA: 0 or 1),
    /// default 0. Set from `Program::option_base` at the start of
    /// `run_sub`/`run_sub_multi` — see `eval_array_bounds`. `run_sub_multi`
    /// takes the first module that sets one rather than modeling true
    /// per-module `Option Base` scoping (this codebase's execution model
    /// is already a single flat `Vm` across all loaded modules).
    option_base: i64,
}

impl Vm {
    pub fn new() -> Self {
        let mut sheets = HashMap::new();
        sheets.insert("sheet1".into(), HashMap::new());
        Vm {
            sheets,
            sheet_order: vec!["sheet1".into()],
            active_sheet: "sheet1".into(),
            variables: HashMap::new(),
            calc_mode: CalculationMode::Automatic,
            error_on_msgbox: false,
            print_msgbox: false,
            msgbox_log: Vec::new(),
            current_span: None,
            exit_flag: None,
            pending_goto: None,
            call_stack: Vec::new(),
            user_funcs: HashMap::new(),
            user_subs: HashMap::new(),
            named_ranges: HashMap::new(),
            type_defs: HashMap::new(),
            col_rows: HashMap::new(),
            row_cols: HashMap::new(),
            cell_index_dirty: true,
            defined_names_may_be_stale: false,
            deadline: None,
            loop_iters: 0,
            strict_resolution: false,
            last_resolution_failure: None,
            loaded_workbook_name: None,
            loaded_workbook_path: None,
            clipboard: None,
            protected_sheets: HashSet::new(),
            merged_ranges: HashMap::new(),
            sheet_visibility: HashMap::new(),
            cell_style_indices: HashMap::new(),
            cell_number_formats: HashMap::new(),
            worksheet_origins: HashMap::new(),
            object_variables: HashMap::new(),
            with_stack: Vec::new(),
            err_number: 0,
            err_description: String::new(),
            err_source: String::new(),
            err_help_file: String::new(),
            err_help_context: 0,
            pending_raised_error: None,
            option_base: 0,
        }
    }

    /// Records a caught runtime error into every `Err` property — called at
    /// every point an `On Error Resume Next`/`On Error GoTo <label>`
    /// actually catches an error (see `exec_stmt`/`exec_body`). Uses
    /// `pending_raised_error` if `Err.Raise` set it (the values the user
    /// actually specified); otherwise classifies the plain error message
    /// via `classify_vba_error_number` and leaves Source/HelpFile/
    /// HelpContext at their zero values — this project doesn't model a VBA
    /// project/class name to default Source to for an internally-raised
    /// error.
    fn record_error(&mut self, msg: &str) {
        match self.pending_raised_error.take() {
            Some(r) => {
                self.err_number = r.number;
                self.err_description = r.description;
                self.err_source = r.source;
                self.err_help_file = r.help_file;
                self.err_help_context = r.help_context;
            }
            None => {
                let (number, description) = classify_vba_error_number(msg);
                self.err_number = number;
                self.err_description = description;
                self.err_source.clear();
                self.err_help_file.clear();
                self.err_help_context = 0;
            }
        }
    }

    /// Drains the resolution-failure evidence set by the most recent failed
    /// `run_sub`/`run_sub_multi` call, if the failure was a classified
    /// resolution failure (missing worksheet/workbook, array out of
    /// bounds) rather than some other runtime error. `None` either if the
    /// run succeeded or if it failed for an unrelated reason.
    pub fn take_resolution_failure(&mut self) -> Option<ResolutionFailureKind> {
        self.last_resolution_failure.take()
    }

    /// Checked once per outer-loop iteration by `For`/`ForEach`/`DoLoop` —
    /// not a per-statement check, so it doesn't touch the interpreter's hot
    /// path outside loop constructs. Only actually calls `Instant::now()`
    /// every 256th iteration (cheap counter increment otherwise), so a
    /// single slow iteration can overshoot the deadline by at most ~256
    /// iterations' worth of time, not indefinitely.
    fn check_deadline(&mut self) -> Result<(), String> {
        self.loop_iters = self.loop_iters.wrapping_add(1);
        if self.loop_iters.is_multiple_of(256)
            && let Some(deadline) = self.deadline
            && std::time::Instant::now() >= deadline
        {
            return Err("TIMEOUT: loop execution exceeded the configured deadline".to_string());
        }
        Ok(())
    }

    /// Resolve a range address, expanding named ranges if needed.
    /// Accepts both "A1:B3" syntax and registered range names (case-insensitive).
    fn resolve_range_addr(&self, addr: &str) -> Option<((u32, u32), (u32, u32))> {
        if let Some(r) = parse_range_addr(addr) {
            return Some(r);
        }
        self.named_ranges
            .get(&addr.to_lowercase())
            .and_then(|real| parse_range_addr(real))
    }

    /// Multi-area sibling of `resolve_range_addr` (Milestone B7a) — used
    /// only by Copy-source and Paste-destination resolution, not any of the
    /// other ~11 `parse_range_addr`/`resolve_range_addr` call sites (sort,
    /// sheet-range write, formula evaluation), which have no need for
    /// comma-separated addresses.
    fn resolve_multi_area_addr(&self, addr: &str) -> Option<Vec<Rect>> {
        if let Some(r) = parse_multi_area_addr(addr) {
            return Some(r);
        }
        self.named_ranges
            .get(&addr.to_lowercase())
            .and_then(|real| parse_multi_area_addr(real))
    }

    pub fn cells(&self) -> &HashMap<(u32, u32), CellContent> {
        self.sheets
            .get(&self.active_sheet)
            .expect("active sheet must exist")
    }

    pub fn cells_mut(&mut self) -> &mut HashMap<(u32, u32), CellContent> {
        self.cell_index_dirty = true;
        self.sheets
            .get_mut(&self.active_sheet)
            .expect("active sheet must exist")
    }

    /// `insert_rows`'s sheet-parameterized sibling (backs Python's
    /// `insert_rows(..., sheet=None)`). `insert_rows` below just forwards here with
    /// `key = active_sheet` -- VBA's `RowColInsert`/`RangeInsert` call sites don't
    /// change at all. Does NOT shift `merged_ranges`/`sheet_visibility`/
    /// `cell_style_indices`/`cell_number_formats`/formula text -- a pre-existing
    /// VBA-engine limitation, now Python-reachable; see ROADMAP.md's known gaps.
    pub fn insert_rows_on_sheet(&mut self, key: &str, first: u32, count: u32) {
        let to_move: Vec<((u32, u32), CellContent)> = self
            .get_sheet_cells(key)
            .into_iter()
            .flatten()
            .filter(|((r, _), _)| *r >= first)
            .map(|((r, c), v)| ((*r, *c), v.clone()))
            .collect();
        let Some(cells) = self.sheet_cells_mut(key) else {
            return;
        };
        for ((r, c), _) in &to_move {
            cells.remove(&(*r, *c));
        }
        for ((r, c), v) in to_move {
            cells.insert((r + count, c), v);
        }
    }

    /// Inserts `count` blank rows at 1-based `first` on the active sheet, shifting
    /// `first` and everything below it down by `count`. Shared by
    /// `Stmt::RangeInsert`'s `Axis::Row` and `Stmt::RowColInsert` (`Rows(n).Insert`).
    fn insert_rows(&mut self, first: u32, count: u32) {
        let key = self.active_sheet.clone();
        self.insert_rows_on_sheet(&key, first, count);
    }

    /// `delete_rows`'s sheet-parameterized sibling. Single-pass `retain` that drops
    /// everything at `row >= first` (both the deleted band `[first, last]` *and* the
    /// cells about to be reinserted shifted) before reinserting `to_move` at its new
    /// position -- NOT `row < first || row > last`, which would keep the `row > last`
    /// cells at their stale original position while ALSO inserting a second copy at
    /// the shifted position, duplicating data. See
    /// `delete_rows_on_sheet_removes_the_stale_entry_at_the_pre_shift_row` for the
    /// regression test this exists to keep passing.
    pub fn delete_rows_on_sheet(&mut self, key: &str, first: u32, count: u32) {
        let last = first + count - 1;
        let to_move: Vec<((u32, u32), CellContent)> = self
            .get_sheet_cells(key)
            .into_iter()
            .flatten()
            .filter(|((r, _), _)| *r > last)
            .map(|((r, c), v)| ((*r, *c), v.clone()))
            .collect();
        let Some(cells) = self.sheet_cells_mut(key) else {
            return;
        };
        cells.retain(|&(row, _), _| row < first);
        for ((r, c), v) in to_move {
            cells.insert((r - count, c), v);
        }
    }

    /// Deletes `count` rows starting at 1-based `first` (inclusive) on the active
    /// sheet, shifting every row below the deleted range up by `count`. Shared by
    /// `Stmt::RangeDelete`'s `Axis::Row` (`Range(addr).Delete`/`EntireRow.Delete`) and
    /// `Stmt::RowColDelete` (`Rows(n).Delete`) so the two syntactic forms can't drift
    /// apart. Mirrored by `delete_cols` below -- keep the two in sync if either changes.
    fn delete_rows(&mut self, first: u32, count: u32) {
        let key = self.active_sheet.clone();
        self.delete_rows_on_sheet(&key, first, count);
    }

    /// `insert_cols_on_sheet`'s row-axis sibling is `insert_rows_on_sheet` above --
    /// this is `delete_cols`'s sheet-parameterized version, the column-axis mirror of
    /// `delete_rows_on_sheet` (same single-pass `retain(col < first)` correctness
    /// reasoning, on the column instead of the row).
    pub fn delete_cols_on_sheet(&mut self, key: &str, first: u32, count: u32) {
        let last = first + count - 1;
        let to_move: Vec<((u32, u32), CellContent)> = self
            .get_sheet_cells(key)
            .into_iter()
            .flatten()
            .filter(|((_, c), _)| *c > last)
            .map(|((r, c), v)| ((*r, *c), v.clone()))
            .collect();
        let Some(cells) = self.sheet_cells_mut(key) else {
            return;
        };
        cells.retain(|&(_, col), _| col < first);
        for ((r, c), v) in to_move {
            cells.insert((r, c - count), v);
        }
    }

    /// `delete_rows`'s column-axis mirror -- deletes `count` columns starting at 1-based
    /// `first` on the active sheet, shifting every column to its right left by `count`.
    fn delete_cols(&mut self, first: u32, count: u32) {
        let key = self.active_sheet.clone();
        self.delete_cols_on_sheet(&key, first, count);
    }

    /// `insert_rows_on_sheet`'s column-axis mirror.
    pub fn insert_cols_on_sheet(&mut self, key: &str, first: u32, count: u32) {
        let to_move: Vec<((u32, u32), CellContent)> = self
            .get_sheet_cells(key)
            .into_iter()
            .flatten()
            .filter(|((_, c), _)| *c >= first)
            .map(|((r, c), v)| ((*r, *c), v.clone()))
            .collect();
        let Some(cells) = self.sheet_cells_mut(key) else {
            return;
        };
        for ((r, c), _) in &to_move {
            cells.remove(&(*r, *c));
        }
        for ((r, c), v) in to_move {
            cells.insert((r, c + count), v);
        }
    }

    /// `insert_rows`'s column-axis mirror -- inserts `count` blank columns at 1-based
    /// `first` on the active sheet, shifting `first` and everything to its right right
    /// by `count`.
    fn insert_cols(&mut self, first: u32, count: u32) {
        let key = self.active_sheet.clone();
        self.insert_cols_on_sheet(&key, first, count);
    }

    fn rebuild_cell_index(&mut self) {
        let pairs: Vec<(u32, u32)> = self
            .cells()
            .iter()
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
        self.ensure_sheet_at(name, None);
    }

    /// `ensure_sheet`, with an optional insertion position (GitHub #3) -- `None` keeps
    /// `ensure_sheet`'s existing append-at-the-end behavior; `Some(i)` inserts the new
    /// sheet at 0-based position `i` in `sheet_order` (clamped to the current length, so
    /// an out-of-range index appends rather than panicking). Ignored if `name` already
    /// exists -- this only controls where a *newly created* sheet lands, not reordering
    /// an existing one; use `move_sheet` to reposition a sheet that already exists.
    pub fn ensure_sheet_at(&mut self, name: &str, index: Option<usize>) {
        let key = name.to_lowercase();
        if !self.sheets.contains_key(&key) {
            match index {
                Some(i) => self
                    .sheet_order
                    .insert(i.min(self.sheet_order.len()), key.clone()),
                None => self.sheet_order.push(key.clone()),
            }
            // GitHub #2: a sheet created here (via `Sheets.Add`, or Python's
            // `set_sheet()`) has no `WorksheetOrigin` from a loaded file, so
            // `save_xlsx_impl`'s display-name fallback used the lowercased `key`
            // itself -- silently lowercasing any ASCII name on save (non-ASCII names
            // were never affected: `to_lowercase()` is a no-op on e.g. Japanese,
            // which is exactly why this went unnoticed until a plain ASCII name was
            // tried). Recording the caller's real casing here is what that fallback
            // needs to find instead.
            self.worksheet_origins.insert(
                key.clone(),
                WorksheetOrigin {
                    original_display_name: Some(name.to_string()),
                    ..Default::default()
                },
            );
        }
        self.sheets.entry(key).or_default();
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

    /// Resolves `sheet` (`None` = active sheet) to its internal lowercase key,
    /// for the Python-binding bulk range/row API (`get_range`/`set_range`/
    /// `append_row`/`iter_rows`/`max_row`/`max_column`/`calculate_dimension`).
    /// Deliberately does not reuse `require_sheet_exists` (used by
    /// `delete_sheet`/VBA sheet resolution): that one is `&mut self` and
    /// records `last_resolution_failure` for the `diagnose` subcommand's
    /// VBA-side diagnostics -- a side channel Python callers never read.
    /// `get_sheet_cells` alone isn't enough either: it returns `None` silently
    /// on an unknown name, and every caller here needs an explicit error.
    pub fn resolve_sheet_key(&self, sheet: Option<&str>) -> Result<String, String> {
        match sheet {
            None => Ok(self.active_sheet.clone()),
            Some(name) => {
                let key = name.to_lowercase();
                if self.sheets.contains_key(&key) {
                    Ok(key)
                } else {
                    Err(format!("Sheet '{name}' not found"))
                }
            }
        }
    }

    /// The 1-based inclusive bounding box of every non-`Empty` cell in `key`
    /// (an already-resolved, lowercased sheet key) -- `None` if the sheet has
    /// no non-empty cells at all. Follows `cells()`/`get_sheet()`'s
    /// Empty-exclusion convention (the one actually surfaced to Python via
    /// `get_range`/`iter_rows`/`max_row`/`max_column`/`calculate_dimension`),
    /// not `cells_df`'s divergent one (`src/lib.rs`'s `cells_df` includes
    /// `Variant::Empty` map entries in its own max) -- a pre-existing,
    /// disclosed inconsistency this doesn't reconcile; see
    /// docs/openpyxl-gap-audit.md.
    ///
    /// NOTE: this excludes `Variant::Empty` only, not `Variant::Null` -- a
    /// `Null`-valued cell (VBA `Null`, distinct from an uninitialized
    /// `Empty`) counts as "non-empty" here and inflates the bounding box,
    /// even though it also crosses into Python as `None` (see
    /// `variant_to_py`). This exactly matches `cells()`/`get_sheet()`'s
    /// existing behavior; it is not a new inconsistency introduced here, but
    /// it is a real, disclosed surprise -- see the gap-audit doc.
    pub fn sheet_used_range(&self, key: &str) -> Option<((u32, u32), (u32, u32))> {
        let cells = self.get_sheet_cells(key)?;
        let mut bounds: Option<((u32, u32), (u32, u32))> = None;
        for (&(r, c), content) in cells {
            if matches!(content.value, Variant::Empty) {
                continue;
            }
            bounds = Some(match bounds {
                None => ((r, c), (r, c)),
                Some(((r1, c1), (r2, c2))) => ((r1.min(r), c1.min(c)), (r2.max(r), c2.max(c))),
            });
        }
        bounds
    }

    /// The 1-based row `append_row` should write to: one past the sheet's
    /// current max used row, or row 1 if the sheet is empty/all-empty. Uses
    /// `sheet_used_range`'s real max, not a populated-row count -- correct on
    /// a sparse sheet (data only at row 50 appends at row 51).
    pub fn next_append_row(&self, key: &str) -> u32 {
        self.sheet_used_range(key)
            .map_or(1, |(_, (max_r, _))| max_r + 1)
    }

    /// Reads a rectangular region (1-based inclusive `r1..=r2`, `c1..=c2`) of
    /// `key` as a row-major grid, `Variant::Empty` for any cell with no
    /// `CellContent` entry. No validation of `key`'s existence (callers
    /// resolve via `resolve_sheet_key` first) or of the rect's shape (callers
    /// guarantee `r1<=r2`, `c1<=c2`, both >=1) -- a pure mechanical read,
    /// matching the established inline-nested-loop style used throughout this
    /// file and `src/formula/eval.rs` rather than a new range-iteration
    /// abstraction.
    pub fn read_rect(&self, key: &str, r1: u32, c1: u32, r2: u32, c2: u32) -> Vec<Vec<Variant>> {
        let empty = HashMap::new();
        let cells = self.get_sheet_cells(key).unwrap_or(&empty);
        (r1..=r2)
            .map(|r| {
                (c1..=c2)
                    .map(|c| {
                        cells
                            .get(&(r, c))
                            .map(|ct| ct.value.clone())
                            .unwrap_or(Variant::Empty)
                    })
                    .collect()
            })
            .collect()
    }

    /// Writes `values` (already validated by the caller: rectangular, exact
    /// target shape, every element already a `Variant`) at `top_left` in
    /// `key`. No shape check here -- the PyO3 glue (`set_range`/`append_row`
    /// in `src/lib.rs`) validates the *entire* input against the target shape
    /// and converts every value before calling this, so by the time this runs
    /// there is nothing left that can fail partway through.
    ///
    /// Deliberately does NOT call `check_sheet_not_protected` and does NOT
    /// consult `merged_ranges` -- matches `PyVm::set_cell`'s existing,
    /// equally unchecked behavior. Sheet protection is a VBA-statement
    /// concept today (14 call sites, all inside this file's statement
    /// handlers; `protected_sheets` isn't even reachable from `src/lib.rs`).
    /// Merge-conflict checking (`check_merge_conflicts`) exists only on the
    /// VBA Copy/Paste path, matching real Excel's own distinction between
    /// Paste's stricter semantics and plain `.Value=` assignment, which
    /// really does store a value on a non-anchor merged cell (just never
    /// displays it). Adding either restriction here would be a new
    /// Python-only behavior with no VBA-side precedent -- see
    /// docs/openpyxl-gap-audit.md.
    ///
    /// Never touches `self.active_sheet`.
    pub fn write_rect(&mut self, key: &str, top_left: (u32, u32), values: &[Vec<Variant>]) {
        let (r1, c1) = top_left;
        let Some(cells) = self.sheet_cells_mut(key) else {
            return;
        };
        for (i, row) in values.iter().enumerate() {
            for (j, v) in row.iter().enumerate() {
                cells.insert(
                    (r1 + i as u32, c1 + j as u32),
                    CellContent {
                        formula: None,
                        value: v.clone(),
                    },
                );
            }
        }
    }

    /// Core of the Python `iter_rows` API: `max_row`/`max_col` of `None` mean
    /// "default to the sheet's used range." If the sheet has NO non-empty
    /// cells at all and the caller didn't pin `max_row` down explicitly,
    /// there is no used range to iterate -- returns zero rows, not one row of
    /// `Empty`s. Only `max_row`'s explicitness matters for this
    /// short-circuit, not `max_col`'s. An explicit `max_row` is always
    /// honored even on an empty sheet (an explicit ask for N rows of Emptys
    /// is not the ambiguous case this guards against).
    pub fn iter_rows_values(
        &self,
        key: &str,
        min_row: u32,
        max_row: Option<u32>,
        min_col: u32,
        max_col: Option<u32>,
    ) -> Vec<Vec<Variant>> {
        let bounds = self.sheet_used_range(key);
        let resolved_max_row = match max_row {
            Some(r) => r,
            None => match bounds {
                Some((_, (r2, _))) => r2,
                None => return Vec::new(),
            },
        };
        let resolved_max_col = max_col.unwrap_or_else(|| bounds.map_or(min_col, |(_, (_, c2))| c2));
        if resolved_max_row < min_row || resolved_max_col < min_col {
            return Vec::new();
        }
        self.read_rect(key, min_row, min_col, resolved_max_row, resolved_max_col)
    }

    /// Core of the Python `iter_cols` API: the column-major transpose of
    /// `iter_rows_values`. Same short-circuit shape, but keyed off
    /// `max_col`'s explicitness instead of `max_row`'s -- an empty sheet with
    /// no explicit `max_col` returns zero columns, while an explicit
    /// `max_col` always forces N columns even on an empty sheet.
    pub fn iter_cols_values(
        &self,
        key: &str,
        min_row: u32,
        max_row: Option<u32>,
        min_col: u32,
        max_col: Option<u32>,
    ) -> Vec<Vec<Variant>> {
        let bounds = self.sheet_used_range(key);
        let resolved_max_col = match max_col {
            Some(c) => c,
            None => match bounds {
                Some((_, (_, c2))) => c2,
                None => return Vec::new(),
            },
        };
        let resolved_max_row = max_row.unwrap_or_else(|| bounds.map_or(min_row, |(_, (r2, _))| r2));
        if resolved_max_row < min_row || resolved_max_col < min_col {
            return Vec::new();
        }
        let grid = self.read_rect(key, min_row, min_col, resolved_max_row, resolved_max_col);
        let num_cols = (resolved_max_col - min_col + 1) as usize;
        (0..num_cols)
            .map(|ci| grid.iter().map(|row| row[ci].clone()).collect())
            .collect()
    }

    /// Sorts a rectangular range on `key`, in place, by a single 1-based absolute
    /// key column. Extracted from `Stmt::RangeSort`'s formerly-inline body so both
    /// the VBA statement and PyVm's `sort_range` share one implementation.
    /// `header` excludes the range's first row from the sort (data starts at
    /// `r1+1`) without moving it. `key_col` outside `c1..=c2` silently clamps via
    /// `saturating_sub` -- a pre-existing behavior from the original inline code,
    /// preserved as-is for the VBA path; PyVm's own `sort_range` validates
    /// `key_col` explicitly instead of inheriting this silent clamp.
    ///
    /// Deliberately does NOT check sheet protection -- matches R1's
    /// `write_rect`/`set_range` precedent (a bulk cell-value write bypasses
    /// protection in the Python API even though VBA's own equivalent path does
    /// check it -- confirmed at `write_range_ref_value`). The VBA
    /// `Stmt::RangeSort` arm keeps its own protection check before calling
    /// this, so VBA's existing tested behavior is unaffected.
    #[allow(clippy::too_many_arguments)]
    pub fn sort_range_on_sheet(
        &mut self,
        key: &str,
        r1: u32,
        c1: u32,
        r2: u32,
        c2: u32,
        key_col: u32,
        descending: bool,
        header: bool,
    ) {
        let data_r1 = if header { r1 + 1 } else { r1 };
        let key_off = key_col.saturating_sub(c1) as usize;
        let mut rows = self.read_rect(key, data_r1, c1, r2, c2);
        rows.sort_by(|a, b| {
            let va = a.get(key_off).unwrap_or(&Variant::Empty);
            let vb = b.get(key_off).unwrap_or(&Variant::Empty);
            let ord = cmp_variants(va, vb);
            if descending { ord.reverse() } else { ord }
        });
        self.write_rect(key, (data_r1, c1), &rows);
    }

    /// Creates a merge over `(r1,c1)..(r2,c2)` on `key`. Rejects a single-cell
    /// "merge" (nothing would actually be merged) and rejects any merge that
    /// would overlap an existing one on the same sheet -- reusing
    /// `rects_overlap` (Milestone B6c2's Copy/Paste conflict-detection
    /// primitive, already sheet-agnostic and side-channel-free) rather than
    /// `check_merge_conflicts` (Copy/Paste-specific: `&mut self`, writes
    /// `last_resolution_failure`, a diagnostic side channel Python callers
    /// don't read). Two overlapping `<mergeCell>` elements is genuinely
    /// invalid OOXML, not just a fidelity gap, so this is a hard error, not a
    /// disclosed limitation.
    ///
    /// Does NOT touch cell values -- whatever is in the covered cells (if
    /// anything) stays exactly as it was. This VM's merge geometry and cell
    /// values are already orthogonal by design (`write_rect`/`set_range`
    /// explicitly allow writing into a non-anchor merged cell without error;
    /// this is the same precedent applied in the other direction).
    pub fn merge_cells(
        &mut self,
        key: &str,
        r1: u32,
        c1: u32,
        r2: u32,
        c2: u32,
    ) -> Result<(), String> {
        if r1 == r2 && c1 == c2 {
            return Err("a merge must span at least 2 cells".to_string());
        }
        let new_rect = ((r1, c1), (r2, c2));
        // Check BEFORE touching the map -- `.entry().or_default()` would
        // insert an empty Vec for `key` even on a rejected merge, a state
        // mutation on a failure path this project's validate-before-
        // committing convention rules out.
        if let Some(existing) = self.merged_ranges.get(key)
            && let Some(&conflict) = existing.iter().find(|&&m| rects_overlap(m, new_rect))
        {
            return Err(format!(
                "merge {} would overlap an existing merge {}",
                crate::merge_rect_to_a1(&new_rect),
                crate::merge_rect_to_a1(&conflict)
            ));
        }
        self.merged_ranges
            .entry(key.to_string())
            .or_default()
            .push(new_rect);
        Ok(())
    }

    /// Removes a merge on `key` whose rect exactly matches `(r1,c1)..(r2,c2)`
    /// -- an inexact/partial match is rejected rather than silently
    /// no-opping, matching this project's "must not silently no-op on
    /// failure" house rule (`rename_sheet`/`move_sheet`/`delete_sheet` all
    /// reject an unknown target the same way).
    pub fn unmerge_cells(
        &mut self,
        key: &str,
        r1: u32,
        c1: u32,
        r2: u32,
        c2: u32,
    ) -> Result<(), String> {
        let target = ((r1, c1), (r2, c2));
        let before = self.merged_ranges.get(key).map_or(0, |v| v.len());
        if let Some(merges) = self.merged_ranges.get_mut(key) {
            merges.retain(|&m| m != target);
        }
        let after = self.merged_ranges.get(key).map_or(0, |v| v.len());
        if after == before {
            return Err(format!(
                "no merge found at {} on sheet '{}'",
                crate::merge_rect_to_a1(&target),
                key
            ));
        }
        Ok(())
    }

    /// `true` iff `requested` identifies the one workbook `load_workbook_file`
    /// loaded (by name, case-insensitively, or by the numeric index `1` —
    /// elixcee never has more than one workbook open, so any other index is
    /// always a mismatch). No workbook loaded yet is never a match.
    fn workbook_matches(&self, requested: &Variant) -> bool {
        match requested {
            Variant::Integer(1) => self.loaded_workbook_name.is_some(),
            Variant::Integer(_) => false,
            other => {
                let name = vba_to_str(other).to_lowercase();
                self.loaded_workbook_name
                    .as_deref()
                    .is_some_and(|n| n.to_lowercase() == name)
            }
        }
    }

    // ── Runtime With stack (see `WithValue`) ────────────────────────────────

    /// Evaluates a `With` block's target once, on block entry.
    fn eval_with_target(&mut self, target: &WithTarget) -> Result<WithValue, String> {
        Ok(match target {
            WithTarget::Object(obj) => match self.eval_object_expr(obj)? {
                ObjectRef::Range(r) => WithValue::Range(r),
                ObjectRef::Worksheet(key) => WithValue::Sheet(key),
                ObjectRef::Workbook | ObjectRef::Nothing => WithValue::Unmodeled,
            },
            WithTarget::Cells(row, col) => {
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                WithValue::Range(RangeRef::single(
                    self.active_sheet.clone(),
                    Rect {
                        start_row: r,
                        start_col: c,
                        end_row: r,
                        end_col: c,
                    },
                ))
            }
            // A bare identifier: the parser can't know its type, so decide
            // here. An object variable wins over a same-named record, the
            // same precedence `Stmt::RecordSet`/`Expr::RecordGet` already
            // use. A name that is neither is left Unmodeled (a no-op body)
            // rather than erroring — the pre-existing behavior for a
            // `With <unknown>` target.
            WithTarget::Var(name) => match self.object_variables.get(name) {
                Some(ObjectRef::Range(r)) => WithValue::Range(r.clone()),
                Some(ObjectRef::Worksheet(key)) => WithValue::Sheet(key.clone()),
                Some(ObjectRef::Nothing) => return Err(OBJECT_NOT_SET.to_string()),
                Some(ObjectRef::Workbook) => WithValue::Unmodeled,
                None => WithValue::Record(name.clone()),
            },
            WithTarget::Unmodeled => WithValue::Unmodeled,
        })
    }

    /// Runs a `With` body with `value` pushed as the innermost target.
    ///
    /// The pop happens on *every* exit path — normal completion, an early
    /// `Exit Sub`/`Exit For`, and a runtime error (which is why the error is
    /// held and returned after popping rather than propagated with `?`).
    /// Leaking an entry here would silently mis-resolve whatever `.member`
    /// ran next, which is exactly the kind of bug a stack invites.
    fn run_with_body(&mut self, value: WithValue, body: &[SpannedStmt]) -> Result<(), String> {
        self.with_stack.push(value);
        let mut result = Ok(());
        for s in body {
            result = self.exec_stmt(s);
            if result.is_err() || self.exit_flag.is_some() {
                break;
            }
        }
        self.with_stack.pop();
        result
    }

    /// The innermost active `With` target, or the error real VBA's own
    /// error-91 text describes ("… or With block variable not set") when a
    /// bare `.member` appears with no enclosing `With` at all. Before the
    /// runtime stack existed this was a *parse* error; it stays an error,
    /// just a runtime one, since a bare `.member` can now appear anywhere a
    /// statement or expression can.
    fn current_with(&self) -> Result<WithValue, String> {
        self.with_stack
            .last()
            .cloned()
            .ok_or_else(|| OBJECT_NOT_SET.to_string())
    }

    /// Resolves a `.Cells(r, c)` / `.Range("addr")` qualifier inside a With
    /// body to the `(sheet_key, row, col)` it addresses.
    ///
    /// For a Worksheet target these are that sheet's own cells — the fix
    /// for `With ws` + `.Cells(i, 1)`, which used to silently write to the
    /// *active* sheet instead. For a Range target they keep their
    /// pre-existing meaning (an independent reference on the active sheet,
    /// pinned by `with_range_nested_range_reference_still_works`); real
    /// VBA's `Range.Range`/`Range.Cells` are relative to the base range's
    /// top-left, which elixcee does not model — see this project's
    /// disclosure list.
    fn resolve_with_qualified_cell(
        &mut self,
        member: &WithMember,
    ) -> Result<Option<(String, u32, u32)>, String> {
        let sheet = match self.current_with()? {
            WithValue::Sheet(key) => key,
            // Every non-Worksheet target keeps the pre-existing behavior: a
            // `.Cells(...)`/`.Range(...)` qualifier inside a With body was
            // always an independent, absolute reference on the active sheet
            // (that's what the parse-time rewrite emitted for it, whatever
            // the target was). Notably that includes `With Sheet1`, where
            // `Sheet1` is a worksheet *code name* elixcee doesn't model —
            // the active sheet is the closest available reading.
            WithValue::Range(_) | WithValue::Record(_) | WithValue::Unmodeled => {
                self.active_sheet.clone()
            }
        };
        Ok(match member {
            WithMember::Cells { row, col, .. } => {
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                Some((sheet, r, c))
            }
            WithMember::Range { addr, .. } => {
                let ((r, c), _) = parse_range_addr(addr)
                    .ok_or_else(|| format!("Invalid range address '{}'", addr))?;
                Some((sheet, r, c))
            }
            WithMember::Fields(_) => None,
        })
    }

    /// `.member = value` inside a With body.
    fn write_with_member(&mut self, member: &WithMember, v: Variant) -> Result<(), String> {
        if !matches!(member, WithMember::Fields(_)) {
            let is_formula = matches!(member,
                WithMember::Cells { fields, .. } | WithMember::Range { fields, .. }
                    if fields.last().map(String::as_str) == Some("formula"));
            if let Some((sheet, r, c)) = self.resolve_with_qualified_cell(member)? {
                let target = RangeRef::single(
                    sheet,
                    Rect {
                        start_row: r,
                        start_col: c,
                        end_row: r,
                        end_col: c,
                    },
                );
                self.write_range_ref_value(&target, is_formula, &v)?;
            }
            return Ok(());
        }
        let WithMember::Fields(fields) = member else {
            unreachable!("guarded above")
        };
        match self.current_with()? {
            WithValue::Range(r) => {
                let f = fields.first().map(String::as_str).unwrap_or("");
                if f == "value" || f == "formula" {
                    self.write_range_ref_value(&r, f == "formula", &v)?;
                }
                // Any other property write on a Range is a harmless no-op,
                // same leniency `Stmt::RecordSet`'s object-variable path has.
            }
            // A worksheet property write (`.Name = "x"`, …) isn't modeled.
            WithValue::Sheet(_) | WithValue::Unmodeled => {}
            WithValue::Record(var) => {
                // `.a = 1` / `.a.b = 1` on a UDT target — the same
                // `nested_set` path `Stmt::RecordSetNested` uses, which also
                // covers the single-field case.
                let target = self
                    .variables
                    .entry(var)
                    .or_insert_with(|| Variant::Record(HashMap::new()));
                nested_set(target, fields, v);
            }
        }
        Ok(())
    }

    /// Rejects a member access through an object variable that holds no live
    /// reference (`Dim r As Range` with no `Set`, or an explicit `Set r =
    /// Nothing`) with real VBA's own error-91 wording.
    ///
    /// Deliberately checks only for a registered `ObjectRef::Nothing`, never
    /// for *absence* from `object_variables`: a name that isn't an object
    /// variable at all is an ordinary scalar/UDT name, and `p.field = 1` on
    /// one of those must keep its pre-existing behavior (auto-create the
    /// record), not start erroring. That is the whole difference between
    /// "declared object variable, unset" and "not an object variable".
    fn require_live_object(&self, var: &str) -> Result<(), String> {
        if matches!(self.object_variables.get(var), Some(ObjectRef::Nothing)) {
            return Err(OBJECT_NOT_SET.to_string());
        }
        Ok(())
    }

    /// Resolves a sheet-identifying `Expr` — a string name, a 1-based
    /// numeric index, or a `Workbooks(...).Worksheets(...)` qualifier — to
    /// `(key, display)`: the lowercase key used to index `self.sheets`, and
    /// the human-readable form to show in evidence/error messages (the
    /// as-written name, or the numeric index as a string). Both the
    /// numeric-index and `Workbooks(...)` forms are new in Milestone B6a, so
    /// unlike plain-name lookups there is no pre-B6a lenient behavior to
    /// preserve for them: a workbook mismatch or an out-of-range index is
    /// always a hard error (evidence recorded via `last_resolution_failure`),
    /// in every mode, not just `strict_resolution`.
    ///
    /// The returned key is **not** guaranteed to exist in `self.sheets` for
    /// a plain-name lookup — each of the four sheet-access call sites
    /// (`SheetCellRead`/`SheetRangeRead`/`SheetCellWrite`/`SheetRangeWrite`)
    /// checks that separately via `check_strict_sheet_exists`, since each
    /// has its own pre-B6a fallback (auto-vivify on write, silent `Empty`
    /// on read) that only applies when `strict_resolution` is off.
    fn resolve_sheet_expr(&mut self, sheet_expr: &Expr) -> Result<(String, String), String> {
        // `ActiveSheet` (Milestone B7c item 6) resolves directly to
        // `self.active_sheet` — it's already exactly the key/display pair
        // every other branch below is working to produce, and (unlike a
        // name/index lookup) it can't fail to resolve.
        if let Expr::ActiveSheetRef = sheet_expr {
            let key = self.active_sheet.clone();
            return Ok((key.clone(), key));
        }
        // `<var>.Range(...)`/`.Cells(...)` where `var` was `Set`-assigned a
        // Worksheet reference (Phase 2C item 7, e.g. `Set ws = ActiveSheet`)
        // — resolved against `object_variables` here, at runtime, since the
        // parser can't tell `ws` apart from an ordinary variable at parse
        // time (see `Expr::ObjectVarSheet`'s doc).
        if let Expr::ObjectVarSheet(name) = sheet_expr {
            return match self.object_variables.get(name).cloned() {
                Some(ObjectRef::Worksheet(key)) => Ok((key.clone(), key)),
                Some(ObjectRef::Workbook) => Err(format!(
                    "'{}' is a Workbook object — use '{}.Worksheets(name)', not '.Range(...)'/'.Cells(...)' directly",
                    name, name
                )),
                Some(ObjectRef::Range(_)) => {
                    Err(format!("'{}' is a Range object, not a Worksheet", name))
                }
                Some(ObjectRef::Nothing) => Err(OBJECT_NOT_SET.to_string()),
                None => Err(format!("'{}' is Nothing — Set was never called", name)),
            };
        }
        let plain = match sheet_expr {
            Expr::WorkbookQualifiedSheet { workbook, sheet } => {
                let wb_val = self.eval_expr(workbook)?;
                if !self.workbook_matches(&wb_val) {
                    let requested = vba_to_str(&wb_val);
                    let available = match &self.loaded_workbook_name {
                        Some(n) => vec![n.clone()],
                        None => vec![],
                    };
                    let evidence = ResolutionEvidence {
                        expression: format!("Workbooks({})", requested),
                        requested: requested.clone(),
                        suggested: closest_match(&requested, &available),
                        available,
                    };
                    self.last_resolution_failure =
                        Some(ResolutionFailureKind::WorkbookNotFound(evidence));
                    return Err(format!("Workbook '{}' not found", requested));
                }
                sheet.as_ref()
            }
            other => other,
        };

        let val = self.eval_expr(plain)?;
        match val {
            Variant::Integer(n) => {
                let names = self.sheet_names();
                let idx = n - 1;
                if idx >= 0 && (idx as usize) < names.len() {
                    let key = names[idx as usize].clone();
                    Ok((key, n.to_string()))
                } else {
                    let evidence = ResolutionEvidence {
                        expression: format!("Worksheets({})", n),
                        requested: n.to_string(),
                        available: names,
                        suggested: None,
                    };
                    self.last_resolution_failure =
                        Some(ResolutionFailureKind::WorksheetNotFound(evidence));
                    Err(format!("Sheet index {} not found", n))
                }
            }
            other => {
                let display = vba_to_str(&other);
                let key = display.to_lowercase();
                Ok((key, display))
            }
        }
    }

    /// Records `ArrayIndexOutOfBounds` evidence and returns the same
    /// message string every array-access site has always returned — a pure
    /// addition (the error was already unconditionally hard, in every
    /// mode, before Milestone B6a), so existing callers/tests see byte-
    /// identical output. `vba_idx`/`lower` are the VBA-facing values (the
    /// index actually written in source, and the array's real lower bound
    /// — 0 for the UDT-array call sites, which don't track an explicit
    /// bound), so `lower`/`upper` here are elixcee's true bounds, not a
    /// fabricated always-0-based guess.
    fn array_oob_error(&mut self, name: &str, vba_idx: i64, lower: i64, len: usize) -> String {
        self.last_resolution_failure = Some(ResolutionFailureKind::ArrayIndexOutOfBounds {
            name: name.to_string(),
            index: vba_idx,
            lower,
            upper: lower + len as i64 - 1,
        });
        // Real VBA's runtime error 9 message, verbatim — no array
        // name/index/length embellishment, matching this codebase's own
        // established convention for other runtime errors (e.g. "Division
        // by zero" carries no extra detail either). The rich per-failure
        // detail this used to put inline is not lost: `diagnose`/
        // `diagnose-workbook` already read it from the structured
        // `last_resolution_failure` side channel set just above, not by
        // parsing this string — see docs/agent-contract.md's own note that
        // `message` is free text, not a stable/matchable field.
        "Subscript out of range".to_string()
    }

    /// Same idea as `array_oob_error`, for a real (possibly multi-dim)
    /// `VbaArray` subscript failure: records evidence for the first
    /// dimension whose index is actually out of range, and returns
    /// "Subscript out of range" either way. A wrong subscript *count*
    /// (`arr(1)` on a 2-D array) has no single out-of-range dimension to
    /// report — `ArrayIndexOutOfBounds` evidence has no shape for that
    /// failure mode — so evidence is explicitly cleared for that case
    /// rather than left at whatever an earlier, unrelated operation set;
    /// the returned message is still correct, just without extra evidence.
    fn vba_array_oob_error_for(
        &mut self,
        name: &str,
        indices: &[i64],
        bounds: &[ArrayBound],
    ) -> String {
        if indices.len() == bounds.len() {
            for (&sub, bound) in indices.iter().zip(bounds.iter()) {
                if sub < bound.lower || sub > bound.upper {
                    return self.array_oob_error(name, sub, bound.lower, bound.len() as usize);
                }
            }
        }
        self.last_resolution_failure = None;
        "Subscript out of range".to_string()
    }

    /// Evaluates every dimension of a `Dim`/`ReDim` array declarator into
    /// real per-dimension bounds. Each dimension's explicit `lo To hi`
    /// supplies its own lower bound; a bare upper-bound expression defaults
    /// to `Option Base`'s value, independently per dimension (real VBA:
    /// `Option Base 1` followed by `Dim arr(3, 2)` means both dimensions
    /// start at 1, not just the first).
    fn eval_array_bounds(&mut self, dims: &[ArrayDim]) -> Result<Vec<ArrayBound>, String> {
        dims.iter()
            .map(|dim| {
                let lower = match &dim.lower {
                    Some(e) => to_f64(&self.eval_expr(e)?)? as i64,
                    None => self.option_base,
                };
                let upper = to_f64(&self.eval_expr(&dim.upper)?)? as i64;
                Ok(ArrayBound { lower, upper })
            })
            .collect()
    }

    /// Evaluates `arr(i, j, ...)`-style subscript `Expr`s into the `i64`
    /// indices `VbaArray::get`/`set`/`linear_index` take.
    fn eval_array_indices(&mut self, indices: &[Expr]) -> Result<Vec<i64>, String> {
        indices
            .iter()
            .map(|e| Ok(to_f64(&self.eval_expr(e)?)? as i64))
            .collect()
    }

    /// Evaluates `LBound`/`UBound`'s optional second (dimension) argument —
    /// 1 if omitted, matching real VBA's own default. Negative results clamp
    /// to 0 rather than erroring here — `VbaArray::lbound`/`ubound` already
    /// reject `dimension == 0` with the same "Subscript out of range" a
    /// too-large dimension gets, so there's no second error shape to keep
    /// consistent with.
    fn array_func_dimension(&mut self, expr: Option<&Expr>) -> Result<usize, String> {
        let n = match expr {
            Some(e) => to_f64(&self.eval_expr(e)?)? as i64,
            None => 1,
        };
        Ok(n.max(0) as usize)
    }

    /// If `strict_resolution` is on and `key` doesn't name an existing
    /// sheet, records `WorksheetNotFound` evidence (with a "did you mean"
    /// suggestion, if any) and returns the matching error. Callers only
    /// invoke this when they're about to do something that pre-B6a leniency
    /// (auto-vivify on write / silent `Empty` on read) would otherwise paper
    /// over — see `resolve_sheet_expr`'s doc comment.
    fn check_strict_sheet_exists(&mut self, requested: &str, key: &str) -> Result<(), String> {
        if self.strict_resolution && !self.sheets.contains_key(key) {
            let available = self.sheet_names();
            let evidence = ResolutionEvidence {
                expression: format!("Worksheets(\"{}\")", requested),
                requested: requested.to_string(),
                suggested: closest_match(requested, &available),
                available,
            };
            self.last_resolution_failure = Some(ResolutionFailureKind::WorksheetNotFound(evidence));
            return Err(format!("Sheet '{}' not found", requested));
        }
        Ok(())
    }

    /// Unconditional sheet-must-exist check (every mode, not gated behind
    /// `strict_resolution`) — used by `.Protect`/`.Unprotect` (Milestone
    /// B6c), which is a brand-new construct with no pre-existing lenient
    /// behavior to preserve, same reasoning as `WorkbookQualifiedSheet`'s
    /// mismatch check in `resolve_sheet_expr`.
    fn require_sheet_exists(&mut self, requested: &str, key: &str) -> Result<(), String> {
        if !self.sheets.contains_key(key) {
            let available = self.sheet_names();
            let evidence = ResolutionEvidence {
                expression: format!("Worksheets(\"{}\")", requested),
                requested: requested.to_string(),
                suggested: closest_match(requested, &available),
                available,
            };
            self.last_resolution_failure = Some(ResolutionFailureKind::WorksheetNotFound(evidence));
            return Err(format!("Sheet '{}' not found", requested));
        }
        Ok(())
    }

    /// If `key` names a `.Protect`ed sheet, records `SheetProtected`
    /// evidence and returns the matching error — unconditional in every
    /// mode (Milestone B6c), since real Excel blocks any cell-content
    /// mutation on a protected sheet regardless of error-handling state,
    /// and nothing pre-existing relied on writes to a "protected" sheet
    /// succeeding (the concept didn't exist before this milestone).
    fn check_sheet_not_protected(&mut self, key: &str, display: &str) -> Result<(), String> {
        if self.protected_sheets.contains(key) {
            self.last_resolution_failure = Some(ResolutionFailureKind::SheetProtected {
                sheet: display.to_string(),
            });
            return Err(format!("Cannot edit: sheet '{}' is protected", display));
        }
        Ok(())
    }

    /// The actual removal, shared by `Stmt::SheetsDelete` (`sheet` already resolved to
    /// `key`/`display` by `resolve_sheet_expr`) and `delete_sheet` below (GitHub #3) --
    /// one removal path so 0.10.0-D4's save-time reachability pruning, which keys off
    /// `self.sheets`/`self.sheet_order`, behaves identically regardless of which
    /// caller removed the sheet. Silently no-ops on `key == self.active_sheet` -- a
    /// pre-existing, deliberate limitation of the VBA path (real Excel allows deleting
    /// the active sheet and re-activates another one; this VM doesn't attempt that),
    /// preserved as-is rather than fixed as a side effect of this refactor.
    fn remove_sheet(&mut self, key: &str, display: &str) -> Result<(), String> {
        self.check_sheet_not_protected(key, display)?;
        if key != self.active_sheet {
            self.sheets.remove(key);
            self.sheet_order.retain(|n| n != key);
        }
        Ok(())
    }

    /// Deletes the sheet named `name` -- the direct counterpart to `ensure_sheet`'s
    /// create-on-demand, for a caller (Python's `delete_sheet()`) that isn't going
    /// through a VBA `Sheets(name).Delete` statement. Unlike the VBA path (whose sheet
    /// name comes from `resolve_sheet_expr`, which never validates existence itself --
    /// an existing, separate gap not touched here), this direct entry point DOES
    /// validate `name` exists first: a caller building a structural "delete this sheet"
    /// action wants a clear error on a typo, not `Sheets("Typo").Delete`'s pre-existing
    /// silent no-op. The removal itself is `remove_sheet`, shared with the VBA path.
    pub fn delete_sheet(&mut self, name: &str) -> Result<(), String> {
        let key = name.to_lowercase();
        self.require_sheet_exists(name, &key)?;
        self.remove_sheet(&key, name)
    }

    /// Renames a sheet, atomically re-keying all eight lowercase-keyed per-sheet `Vm`
    /// maps that a rename can touch (`sheets`, `sheet_order`, `active_sheet`,
    /// `merged_ranges`, `sheet_visibility`, `cell_style_indices`,
    /// `cell_number_formats`, `worksheet_origins`). Each gets one explicit
    /// remove+insert line rather than a generic "walk every map" helper: the maps
    /// have different value types, a truly generic helper needs a macro or
    /// trait-object indirection to cross that, and with exactly one call site a
    /// helper would save nothing. `remove_sheet`'s own, separate, pre-existing leak
    /// of 6 of these maps on delete is NOT touched here (see ROADMAP.md's known
    /// gaps -- left unfixed by deliberate choice, not missed).
    ///
    /// `protected_sheets` (a `HashSet`) is NOT re-keyed here -- it doesn't need to
    /// be. Renaming a protected sheet is rejected outright below, so `old_key` is
    /// guaranteed absent from `protected_sheets` by the time any re-keying runs; a
    /// re-key step for it would be unreachable dead code, not a defensive one.
    ///
    /// Renaming the ACTIVE sheet is a normal, supported case (updates `active_sheet`
    /// itself) -- NOT rejected, and NOT the silent no-op `remove_sheet` uses for
    /// deleting the active sheet: a real rename request must actually rename.
    /// Skipping `active_sheet` here would leave `cells()`/`cells_mut()`'s
    /// `.expect("active sheet must exist")` pointing at a key no longer present in
    /// `sheets`, panicking on the very next cell access.
    ///
    /// Does not touch `cell_index_dirty`/`col_rows`/`row_cols`: that lazy index is
    /// keyed off `sheets[active_sheet]`'s CONTENTS, which are unchanged by a rename
    /// (same cell map, just re-keyed) -- and `active_sheet` itself is re-pointed in
    /// the same call, so the index still resolves to the same data.
    ///
    /// Renaming to the same name (including a pure-case change, e.g. "Sheet1" ->
    /// "SHEET1") is allowed and updates `original_display_name`'s casing -- this is
    /// NOT a collision with itself; the collision check only fires when `new_key`
    /// names a DIFFERENT existing sheet.
    ///
    /// Known, deliberate non-goals (see ROADMAP.md known gaps): does not validate
    /// Excel's 31-char length limit, illegal characters (`: \ / ? * [ ]`), or
    /// reserved names -- matches `set_sheet`'s pre-existing total lack of name
    /// validation. Does not rewrite any formula or `<definedName>` text that refers
    /// to this sheet by its OLD name -- only the `<sheet name="...">` tab label
    /// changes.
    pub fn rename_sheet(&mut self, old_name: &str, new_name: &str) -> Result<(), String> {
        let old_key = old_name.to_lowercase();
        if !self.sheets.contains_key(&old_key) {
            return Err(format!("Sheet '{}' not found", old_name));
        }
        if new_name.trim().is_empty() {
            return Err("Sheet name must not be empty".to_string());
        }
        if self.protected_sheets.contains(&old_key) {
            return Err(format!("Cannot rename: sheet '{}' is protected", old_name));
        }
        let new_key = new_name.to_lowercase();
        if new_key != old_key && self.sheets.contains_key(&new_key) {
            return Err(format!("Sheet '{}' already exists", new_name));
        }

        // 1. `sheets` -- the cell map itself.
        if let Some(cells) = self.sheets.remove(&old_key) {
            self.sheets.insert(new_key.clone(), cells);
        }
        // 2. `sheet_order` -- IN-PLACE value swap, not remove+push, so tab position
        //    is preserved.
        if let Some(slot) = self.sheet_order.iter_mut().find(|k| **k == old_key) {
            *slot = new_key.clone();
        }
        // 3. `active_sheet`.
        if self.active_sheet == old_key {
            self.active_sheet = new_key.clone();
        }
        // 4-7: merged_ranges / sheet_visibility / cell_style_indices / cell_number_formats.
        if let Some(v) = self.merged_ranges.remove(&old_key) {
            self.merged_ranges.insert(new_key.clone(), v);
        }
        if let Some(v) = self.sheet_visibility.remove(&old_key) {
            self.sheet_visibility.insert(new_key.clone(), v);
        }
        if let Some(v) = self.cell_style_indices.remove(&old_key) {
            self.cell_style_indices.insert(new_key.clone(), v);
        }
        if let Some(v) = self.cell_number_formats.remove(&old_key) {
            self.cell_number_formats.insert(new_key.clone(), v);
        }
        // 8. `worksheet_origins` -- re-key AND update `original_display_name` to the
        //    NEW name; `save_xlsx_impl` reads this field (not the lowercased key) to
        //    write `<sheet name="...">` on save.
        let mut origin = self.worksheet_origins.remove(&old_key).unwrap_or_default();
        origin.original_display_name = Some(new_name.to_string());
        self.worksheet_origins.insert(new_key, origin);
        // `protected_sheets` needs no re-key: the protection check above already
        // rejected this call if `old_key` were a member, so it's guaranteed absent
        // here -- a re-key step would be unreachable dead code, not a defensive one.
        //
        // A <definedName>'s TEXT can reference a sheet by name (e.g. "Sheet1!$F$5"),
        // and this rename doesn't rewrite that text -- so any surviving
        // <definedNames> passthrough could now dangle. Set unconditionally, even for
        // a same-name/case-only rename, matching this flag's existing coarse-grained
        // "drop rather than try to prove it's still safe" precedent.
        self.defined_names_may_be_stale = true;
        Ok(())
    }

    /// Repositions an EXISTING sheet in `sheet_order` -- the missing complement to
    /// `ensure_sheet_at` (which only positions a NEWLY created sheet; no primitive
    /// reorders an existing one). Touches ONLY `sheet_order`: no other per-sheet map
    /// is keyed by position, so unlike `rename_sheet` this needs no re-keying
    /// anywhere else.
    ///
    /// `new_index` is an ABSOLUTE 0-based target position (matching `set_sheet`'s
    /// existing `index` convention), not a relative offset like openpyxl's
    /// `Worksheet.move_sheet(offset)`. Clamped to `0..=sheet_order.len()`
    /// post-removal, matching `ensure_sheet_at`'s own clamp -- an out-of-range index
    /// moves to the nearest end rather than erroring.
    ///
    /// Deliberately does NOT check `protected_sheets`: real Excel's per-sheet
    /// "Protect Sheet" does not gate tab reordering (that's the entirely separate,
    /// unimplemented "Protect Workbook" structure flag) -- no precedent in this
    /// codebase to check anything here.
    ///
    /// Sets `defined_names_may_be_stale`, which gates `<definedNames>` passthrough
    /// on save (see `save_xlsx_impl`) -- a `<definedName localSheetId="N">` is
    /// positional, and reordering can silently invalidate it otherwise.
    pub fn move_sheet(&mut self, name: &str, new_index: usize) -> Result<(), String> {
        let key = name.to_lowercase();
        if !self.sheets.contains_key(&key) {
            return Err(format!("Sheet '{}' not found", name));
        }
        self.sheet_order.retain(|k| k != &key);
        let idx = new_index.min(self.sheet_order.len());
        self.sheet_order.insert(idx, key);
        self.defined_names_may_be_stale = true;
        Ok(())
    }

    /// Evaluates an `ObjectExpr` to the `ObjectRef` it names (Milestone
    /// B7c) — the object-typed sibling of `eval_expr`. `Range("...")`
    /// resolves against `self.active_sheet` *now* (fixed into the returned
    /// `RangeRef` from this point on — real VBA captures a Range object's
    /// parent worksheet at creation, not at each later `.Value` access).
    fn eval_object_expr(&mut self, expr: &ObjectExpr) -> Result<ObjectRef, String> {
        match expr {
            ObjectExpr::RangeLit(addr) => {
                let areas = self
                    .resolve_multi_area_addr(addr)
                    .ok_or_else(|| format!("Range: invalid address '{}'", addr))?;
                Ok(ObjectRef::Range(RangeRef {
                    sheet: self.active_sheet.clone(),
                    areas,
                }))
            }
            ObjectExpr::Var(name) => self.object_variables.get(name).cloned().ok_or_else(|| {
                format!(
                    "Object variable '{}' is not set (Set was never called, or it holds Nothing)",
                    name
                )
            }),
            ObjectExpr::Union(parts) => {
                let mut areas: Vec<Rect> = Vec::new();
                let mut sheet: Option<String> = None;
                for p in parts {
                    let r = expect_range_ref(self.eval_object_expr(p)?, "Union")?;
                    match &sheet {
                        None => sheet = Some(r.sheet.clone()),
                        Some(s) if *s != r.sheet => {
                            return Err("Union: all ranges must be on the same sheet".to_string());
                        }
                        Some(_) => {}
                    }
                    areas.extend(r.areas);
                }
                Ok(ObjectRef::Range(RangeRef {
                    sheet: sheet.unwrap_or_else(|| self.active_sheet.clone()),
                    areas,
                }))
            }
            ObjectExpr::Area(base, index) => {
                let r = expect_range_ref(self.eval_object_expr(base)?, "Areas")?;
                let i = to_f64(&self.eval_expr(index)?)? as i64;
                if i < 1 || i as usize > r.areas.len() {
                    return Err(format!(
                        "Areas: index {} out of range (range has {} area{})",
                        i,
                        r.areas.len(),
                        if r.areas.len() == 1 { "" } else { "s" }
                    ));
                }
                Ok(ObjectRef::Range(RangeRef::single(
                    r.sheet.clone(),
                    r.areas[(i - 1) as usize],
                )))
            }
            ObjectExpr::SpecialCellsVisible(base) => {
                let r = expect_range_ref(self.eval_object_expr(base)?, "SpecialCells")?;
                let areas = self.visible_areas(&r.sheet, &r.areas);
                if areas.is_empty() {
                    return Err(
                        "SpecialCells: no visible cells were found (Error 1004)".to_string()
                    );
                }
                Ok(ObjectRef::Range(RangeRef {
                    sheet: r.sheet,
                    areas,
                }))
            }
        }
    }

    /// `SpecialCells(xlCellTypeVisible)`'s geometry (Milestone B7c item 4):
    /// every `area` split along both axes by `sheet`'s hidden row/column
    /// metadata (Milestone B7b's `sheet_visibility`, consumed here rather
    /// than re-read). Since only *whole* rows/columns can be hidden (never
    /// partial cells), each area's visible region decomposes exactly into
    /// the Cartesian product of its maximal visible row-bands and
    /// column-bands — every resulting `Rect` is genuinely all-visible, and
    /// their union is exactly the visible-cell set. This matches real
    /// Excel's own `Areas` grouping when only one axis has hidden spans;
    /// when both axes do, real Excel's own area-merging heuristic can
    /// differ in *how many* Areas it reports for the same visible cells —
    /// unmodeled, since it doesn't change which cells are visible.
    fn visible_areas(&self, sheet: &str, areas: &[Rect]) -> Vec<Rect> {
        let empty = SheetVisibility::default();
        let vis = self.sheet_visibility.get(sheet).unwrap_or(&empty);
        let mut out = Vec::new();
        for r in areas {
            let row_runs = visible_runs(r.start_row, r.end_row, &vis.hidden_rows);
            let col_runs = visible_runs(r.start_col, r.end_col, &vis.hidden_columns);
            for rr in &row_runs {
                for cc in &col_runs {
                    out.push(Rect {
                        start_row: rr.start,
                        start_col: cc.start,
                        end_row: rr.end,
                        end_col: cc.end,
                    });
                }
            }
        }
        out
    }

    /// Reads `.Value` through an object variable's `RangeRef` (Milestone
    /// B7c) — mirrors `Expr::RangeRead`'s single-cell/array split, but
    /// resolves against the range's own captured `sheet` (fixed at `Set`
    /// time — see `eval_object_expr`) rather than `self.active_sheet`,
    /// which may have changed since. A multi-area range has no single
    /// `.Value` in real VBA either (`Areas(n)` first).
    fn read_range_ref_value(&self, r: &RangeRef) -> Result<Variant, String> {
        let area = *r.single_rect().ok_or_else(|| {
            "Range.Value: multi-area range has no single .Value — read through .Areas(n) instead".to_string()
        })?;
        let sheet_cells = self.sheets.get(&r.sheet);
        let get = |row: u32, col: u32| {
            sheet_cells
                .and_then(|s| s.get(&(row, col)))
                .map(|c| c.value.clone())
                .unwrap_or(Variant::Empty)
        };
        if area.start_row == area.end_row && area.start_col == area.end_col {
            Ok(get(area.start_row, area.start_col))
        } else {
            let arr = (area.start_row..=area.end_row)
                .flat_map(|row| (area.start_col..=area.end_col).map(move |col| (row, col)))
                .map(|(row, col)| get(row, col))
                .collect();
            Ok(Variant::Array(arr))
        }
    }

    /// Writes `.Value`/`.Formula` through an object variable's `RangeRef`
    /// (Milestone B7c) — the write-side twin of `read_range_ref_value`,
    /// mirroring `Stmt::RangeWrite`'s fill-the-whole-rect behavior but
    /// against the range's own captured sheet.
    fn write_range_ref_value(
        &mut self,
        r: &RangeRef,
        is_formula: bool,
        v: &Variant,
    ) -> Result<(), String> {
        let area = *r.single_rect().ok_or_else(|| {
            "Range.Value: multi-area range has no single .Value — write through .Areas(n) instead".to_string()
        })?;
        self.check_sheet_not_protected(&r.sheet, &r.sheet)?;
        if is_formula {
            let s = vba_to_str(v);
            let prev = self.active_sheet.clone();
            self.active_sheet = r.sheet.clone();
            self.cell_index_dirty = true;
            let result = (|| -> Result<(), String> {
                for row in area.start_row..=area.end_row {
                    for col in area.start_col..=area.end_col {
                        self.set_cell_formula(row, col, &s)?;
                    }
                }
                Ok(())
            })();
            self.active_sheet = prev;
            self.cell_index_dirty = true;
            result
        } else {
            if let Some(cells) = self.sheets.get_mut(&r.sheet) {
                for row in area.start_row..=area.end_row {
                    for col in area.start_col..=area.end_col {
                        cells.insert(
                            (row, col),
                            CellContent {
                                formula: None,
                                value: v.clone(),
                            },
                        );
                    }
                }
            }
            self.cell_index_dirty = true;
            Ok(())
        }
    }

    /// Populates the clipboard from `areas` on `sheet` (Milestone B7c
    /// refactor of what was `Stmt::RangeCopy`'s inline body) — shared by
    /// the literal-address `Stmt::RangeCopy` (`sheet` is always
    /// `self.active_sheet`) and the new object-variable
    /// `Stmt::RangeObjectCopy` (`sheet` is the variable's own captured
    /// sheet, which may differ from the currently active one — see
    /// `read_range_ref_value`'s doc for why that's correct VBA behavior).
    /// `cells`'s "first area only, and only when `areas.len() == 1`" shape
    /// is UNCHANGED from Milestone B7a (an existing test asserts it);
    /// `area_cells` is new — every area's cells, always populated, feeding
    /// item 5's multi-area Paste.
    fn copy_areas_to_clipboard(&mut self, sheet: String, areas: Vec<Rect>, source_addr: String) {
        let empty: HashMap<(u32, u32), CellContent> = HashMap::new();
        let sheet_cells = self.sheets.get(&sheet).unwrap_or(&empty);
        let get = |row: u32, col: u32| -> Variant {
            sheet_cells
                .get(&(row, col))
                .map(|c| c.value.clone())
                .unwrap_or(Variant::Empty)
        };
        let first = areas[0];
        let cells: Vec<Vec<Variant>> = if areas.len() == 1 {
            (first.start_row..=first.end_row)
                .map(|r| {
                    (first.start_col..=first.end_col)
                        .map(|c| get(r, c))
                        .collect()
                })
                .collect()
        } else {
            Vec::new()
        };
        let area_cells: Vec<Vec<Vec<Variant>>> = areas
            .iter()
            .map(|a| {
                (a.start_row..=a.end_row)
                    .map(|r| (a.start_col..=a.end_col).map(|c| get(r, c)).collect())
                    .collect()
            })
            .collect();
        self.clipboard = Some(ClipboardState {
            source_addr,
            src_sheet: sheet,
            rows: first.rows(),
            cols: first.cols(),
            cells,
            span: self.current_span.unwrap_or(SourceSpan { start: 0, end: 0 }),
            areas,
            area_cells,
        });
    }

    /// Checks the 3 merged-cell-conflict cases for a Paste (Milestone
    /// B6c2), called once from `do_paste` after the destination
    /// anchor/fill dimensions are known — common to both its range- and
    /// single-cell-destination branches. Unconditional in every mode that
    /// executes the macro, same precedent as the shape-mismatch/
    /// protection checks above it.
    #[allow(clippy::too_many_arguments)]
    fn check_merge_conflicts(
        &mut self,
        dest_sheet: &str,
        dest_addr: &str,
        anchor_row: u32,
        anchor_col: u32,
        fill_rows: u32,
        fill_cols: u32,
        clip: &ClipboardState,
        transpose: bool,
    ) -> Result<(), String> {
        let dest_rect = (
            (anchor_row, anchor_col),
            (anchor_row + fill_rows - 1, anchor_col + fill_cols - 1),
        );
        let dest_merges: Vec<MergeRect> = self
            .merged_ranges
            .get(dest_sheet)
            .cloned()
            .unwrap_or_default();

        // 1. Anchor cell is a covered (non-top-left) cell of an existing
        // merge — applies regardless of destination shape.
        let anchor_point = ((anchor_row, anchor_col), (anchor_row, anchor_col));
        for &m in &dest_merges {
            if rect_contains(m, anchor_point) && m.0 != (anchor_row, anchor_col) {
                self.last_resolution_failure =
                    Some(ResolutionFailureKind::PasteIntoNonAnchorMergedCell {
                        dest_addr: dest_addr.to_string(),
                        dest_sheet: dest_sheet.to_string(),
                        merged_range: m,
                        copy_span: Some(clip.span),
                    });
                return Err(
                    "Paste error: destination cell is inside a merged range but isn't its top-left cell"
                        .to_string(),
                );
            }
        }

        // 2. Destination range partially overlaps one or more merges — a
        // single-cell destination can't "partially" overlap anything (case
        // 1 above already covers landing inside one), so this only applies
        // to a genuinely multi-cell destination.
        if fill_rows > 1 || fill_cols > 1 {
            let conflicts: Vec<_> = dest_merges
                .iter()
                .copied()
                .filter(|&m| rects_overlap(dest_rect, m) && !rect_contains(dest_rect, m))
                .collect();
            if !conflicts.is_empty() {
                self.last_resolution_failure =
                    Some(ResolutionFailureKind::PastePartialMergedRange {
                        dest_addr: dest_addr.to_string(),
                        dest_sheet: dest_sheet.to_string(),
                        conflicts,
                        copy_span: Some(clip.span),
                    });
                return Err(
                    "Paste error: destination partially overlaps a merged range".to_string()
                );
            }
        }

        // 3. Source and destination merge layouts differ — only meaningful
        // when the source itself has a real shape to compare against.
        let single_cell_source = clip.rows == 1 && clip.cols == 1;
        if !single_cell_source
            && let Some(src_rect @ ((sr1, sc1), _)) = self.resolve_range_addr(&clip.source_addr)
        {
            let src_merges: Vec<MergeRect> = self
                .merged_ranges
                .get(&clip.src_sheet)
                .cloned()
                .unwrap_or_default();

            let mut src_rel: Vec<MergeRect> = src_merges
                .iter()
                .filter(|&&m| rect_contains(src_rect, m))
                .map(|&((r1, c1), (r2, c2))| ((r1 - sr1, c1 - sc1), (r2 - sr1, c2 - sc1)))
                .collect();
            src_rel.sort();

            let mut dest_rel: Vec<MergeRect> = dest_merges
                .iter()
                .filter(|&&m| rect_contains(dest_rect, m))
                .map(|&((r1, c1), (r2, c2))| {
                    let (rr1, rc1) = (r1 - anchor_row, c1 - anchor_col);
                    let (rr2, rc2) = (r2 - anchor_row, c2 - anchor_col);
                    if transpose {
                        ((rc1, rr1), (rc2, rr2))
                    } else {
                        ((rr1, rc1), (rr2, rc2))
                    }
                })
                .collect();
            dest_rel.sort();

            if src_rel != dest_rel {
                let conflicts: Vec<MergeRect> = dest_merges
                    .iter()
                    .copied()
                    .filter(|&m| rect_contains(dest_rect, m))
                    .collect();
                self.last_resolution_failure =
                    Some(ResolutionFailureKind::PasteMergeLayoutMismatch {
                        source_addr: clip.source_addr.clone(),
                        source_sheet: clip.src_sheet.clone(),
                        dest_addr: dest_addr.to_string(),
                        dest_sheet: dest_sheet.to_string(),
                        conflicts,
                        copy_span: Some(clip.span),
                    });
                return Err("Paste error: source and destination merge layouts differ".to_string());
            }
        }

        Ok(())
    }

    /// Pastes the current clipboard into `dest_addr` — shared by
    /// `Stmt::RangePaste`, `Stmt::SheetRangePaste`, and `Stmt::RangeCopy`'s
    /// immediate `Destination:=` form (Milestone B6b). A missing clipboard
    /// (no prior `.Copy`, or `Application.CutCopyMode` cleared since) and a
    /// destination range whose shape doesn't match the clipboard's (after
    /// accounting for `transpose`) are both unconditional hard errors, in
    /// every mode — real Excel raises Error 1004 for these regardless of
    /// any error-handling state, so this is a fidelity improvement, not a
    /// gated diagnostic-only behavior (see the B6b plan's "Key decision").
    /// Two cases are never shape-checked, matching real Excel: a single
    /// destination *cell* (no `:`) auto-expands from the anchor, and a
    /// single-*cell source* fills an explicit destination range of any size
    /// (real Excel's well-known "paste one value into many cells" fill
    /// behavior — a destination range that's an exact multiple of a
    /// multi-cell source, i.e. tiling, is a rarer sibling left unmodeled).
    /// Once the anchor/fill dimensions are settled, `check_merge_conflicts`
    /// (Milestone B6c2) applies the same unconditional-hard-error treatment
    /// to merged-cell conflicts — same reasoning, same precedent.
    fn do_paste(&mut self, dest_addr: &str, transpose: bool) -> Result<(), String> {
        let active = self.active_sheet.clone();
        self.check_sheet_not_protected(&active, &active)?;
        let clip = match &self.clipboard {
            Some(c) => c.clone(),
            None => {
                self.last_resolution_failure = Some(ResolutionFailureKind::PasteWithoutCopy {
                    dest_addr: dest_addr.to_string(),
                });
                return Err("Paste error: Clipboard is empty".to_string());
            }
        };

        // Milestone B7a: multi-area foundation. Only taken when source or
        // destination unambiguously parses to more than one area; an
        // unparseable destination or a single-area-only paste falls
        // through untouched to the existing logic below.
        let dest_areas_probe = self.resolve_multi_area_addr(dest_addr);
        let is_multi_area =
            clip.areas.len() > 1 || dest_areas_probe.as_ref().is_some_and(|d| d.len() > 1);
        if is_multi_area {
            // Milestone B7c item 5: the one multi-area paste shape that's
            // now actually executed rather than only diagnosed — source
            // and destination both multi-area, with the same `Areas.Count`
            // and matching per-area shapes (pairwise, in order). Every
            // other multi-area shape (count/shape mismatch, or either side
            // single-area) stays diagnose-only exactly as in B7a — see
            // `multi_area_paste_failure`'s doc comment.
            //
            // `transpose` isn't modeled for this shape (real Excel's own
            // per-area transpose semantics aren't obviously well-defined
            // either) — `&& !transpose` below is load-bearing, not
            // cosmetic: without it, a `Transpose:=True` multi-area paste
            // would silently write UN-transposed data instead of either
            // transposing or erroring, trading a loud pre-existing failure
            // for a silently wrong answer. With it, that case still falls
            // through to `multi_area_paste_failure` below, unchanged.
            //
            // This path also skips `check_merge_conflicts` (unlike the
            // single-area path below it) — merged-cell conflicts aren't
            // checked for a multi-area destination. `check_sheet_not_
            // protected` above still applies.
            if let Some(dest_areas) = dest_areas_probe.clone()
                && clip.areas.len() > 1
                && !transpose
                && dest_areas.len() == clip.areas.len()
                && clip
                    .areas
                    .iter()
                    .zip(dest_areas.iter())
                    .all(|(s, d)| s.rows() == d.rows() && s.cols() == d.cols())
            {
                for (i, dst_area) in dest_areas.iter().enumerate() {
                    let src_area = clip.areas[i];
                    for r in 0..src_area.rows() {
                        for c in 0..src_area.cols() {
                            let v = clip.area_cells[i][r as usize][c as usize].clone();
                            self.cells_mut().insert(
                                (dst_area.start_row + r, dst_area.start_col + c),
                                CellContent {
                                    formula: None,
                                    value: v,
                                },
                            );
                        }
                    }
                }
                return Ok(());
            }
            return Err(self.multi_area_paste_failure(&clip, dest_addr, dest_areas_probe));
        }

        let single_cell_source = clip.rows == 1 && clip.cols == 1;
        let (expected_rows, expected_cols) = if transpose {
            (clip.cols, clip.rows)
        } else {
            (clip.rows, clip.cols)
        };
        let (anchor_row, anchor_col, fill_rows, fill_cols) = if dest_addr.contains(':') {
            let ((r1, c1), (r2, c2)) = self
                .resolve_range_addr(dest_addr)
                .ok_or_else(|| format!("Paste error: invalid destination range '{}'", dest_addr))?;
            let dest_rows = r2 - r1 + 1;
            let dest_cols = c2 - c1 + 1;
            if !single_cell_source && (dest_rows != expected_rows || dest_cols != expected_cols) {
                self.last_resolution_failure = Some(ResolutionFailureKind::PasteShapeMismatch {
                    source_addr: clip.source_addr.clone(),
                    source_rows: clip.rows,
                    source_cols: clip.cols,
                    dest_addr: dest_addr.to_string(),
                    dest_rows,
                    dest_cols,
                    dest_row1: r1,
                    dest_col1: c1,
                    transpose,
                    copy_span: Some(clip.span),
                });
                return Err(format!(
                    "Paste error: shape mismatch (source {}x{}, destination {}x{})",
                    expected_rows, expected_cols, dest_rows, dest_cols
                ));
            }
            (r1, c1, dest_rows, dest_cols)
        } else {
            let (r, c) = parse_cell_addr(dest_addr).ok_or_else(|| {
                format!("Paste error: invalid destination address '{}'", dest_addr)
            })?;
            (r, c, expected_rows, expected_cols)
        };
        self.check_merge_conflicts(
            &active, dest_addr, anchor_row, anchor_col, fill_rows, fill_cols, &clip, transpose,
        )?;
        for r in 0..fill_rows {
            for c in 0..fill_cols {
                let v = if single_cell_source {
                    clip.cells[0][0].clone()
                } else if transpose {
                    clip.cells[c as usize][r as usize].clone()
                } else {
                    clip.cells[r as usize][c as usize].clone()
                };
                self.cells_mut().insert(
                    (anchor_row + r, anchor_col + c),
                    CellContent {
                        formula: None,
                        value: v,
                    },
                );
            }
        }
        Ok(())
    }

    /// Classifies a multi-area Copy/Paste and returns the matching error
    /// message — never returns success. Called only once `do_paste` has
    /// already established that source or destination has more than one
    /// area *and* (Milestone B7c item 5) that the shape isn't the one
    /// matching-`Areas.Count`-and-per-area-shape case `do_paste` now
    /// executes directly instead of calling this.
    ///
    /// `dest_areas` is `None` when the destination address itself didn't
    /// resolve (malformed, or a named-range miss) — that's a plain invalid-
    /// address error, same message pre-B7a callers already saw, not a
    /// multi-area classification.
    fn multi_area_paste_failure(
        &mut self,
        clip: &ClipboardState,
        dest_addr: &str,
        dest_areas: Option<Vec<Rect>>,
    ) -> String {
        let source_areas = clip.areas.clone();
        let dest_areas = match dest_areas {
            Some(d) => d,
            None => {
                return format!("Paste error: invalid destination range '{}'", dest_addr);
            }
        };

        if source_areas.len() > 1 && dest_areas.len() == 1 {
            let count = source_areas.len();
            self.last_resolution_failure =
                Some(ResolutionFailureKind::MultiAreaToSingleAreaPaste {
                    source_areas,
                    destination_areas: dest_areas,
                });
            return format!(
                "Paste error: source has {} disjoint areas but the destination is a single area",
                count
            );
        }

        if source_areas.len() > 1 && dest_areas.len() > 1 {
            if source_areas.len() != dest_areas.len() {
                let (src_count, dst_count) = (source_areas.len(), dest_areas.len());
                self.last_resolution_failure =
                    Some(ResolutionFailureKind::MultiAreaCountMismatch {
                        source_areas,
                        destination_areas: dest_areas,
                    });
                return format!(
                    "Paste error: source has {} areas but destination has {} areas",
                    src_count, dst_count
                );
            }
            for (i, (s, d)) in source_areas.iter().zip(dest_areas.iter()).enumerate() {
                if s.rows() != d.rows() || s.cols() != d.cols() {
                    let (area_index, source_area, destination_area) = (i + 1, *s, *d);
                    let (sr, sc, dr, dc) = (s.rows(), s.cols(), d.rows(), d.cols());
                    self.last_resolution_failure =
                        Some(ResolutionFailureKind::MultiAreaShapeMismatch {
                            area_index,
                            source_area,
                            destination_area,
                        });
                    return format!(
                        "Paste error: area {} shape mismatch (source {}x{}, destination {}x{})",
                        area_index, sr, sc, dr, dc
                    );
                }
            }
        }

        // Nothing structurally wrong, but not the one shape `do_paste`
        // executes directly (a single-area source into a multi-area
        // destination, or the reverse) — see `MultiAreaPasteUnsupported`'s
        // doc comment.
        self.last_resolution_failure = Some(ResolutionFailureKind::MultiAreaPasteUnsupported {
            source_areas,
            destination_areas: dest_areas,
        });
        "Paste error: multi-area paste is not yet supported by elixcee (diagnosed, not executed)"
            .to_string()
    }

    /// Computes the `RANGE_CONTAINS_HIDDEN_CELLS` observation (Milestone
    /// B7b) for the range last `.Copy`'d, intersected with the copied
    /// sheet's hidden row/column metadata. Read-only and idempotent — not
    /// a "drain" like `take_resolution_failure` — callable any time after
    /// a run regardless of success/failure. `None` when: nothing has been
    /// copied (or `Application.CutCopyMode = False` cleared it since — the
    /// same "last surviving Copy" limitation `PASTE_WITHOUT_COPY` already
    /// has on the failure side); the copy spanned more than one area
    /// (multi-area sources are deferred, see the B7b plan's design
    /// decisions); the copied sheet has no registered hidden rows/columns;
    /// or none of those hidden rows/columns actually overlap the copied
    /// range.
    pub fn hidden_cells_observation(&self) -> Option<HiddenCellsObservation> {
        let clip = self.clipboard.as_ref()?;
        if clip.areas.len() != 1 {
            return None;
        }
        let rect = clip.areas[0];
        let visibility = self.sheet_visibility.get(&clip.src_sheet)?;

        let hidden_rows: Vec<Interval> = visibility
            .hidden_rows
            .iter()
            .filter_map(|iv| iv.clip(rect.start_row, rect.end_row))
            .collect();
        let hidden_columns: Vec<Interval> = visibility
            .hidden_columns
            .iter()
            .filter_map(|iv| iv.clip(rect.start_col, rect.end_col))
            .collect();
        if hidden_rows.is_empty() && hidden_columns.is_empty() {
            return None;
        }

        // Assumes hidden row/column intervals for a sheet don't overlap
        // each other (true for any real XLSX: rows are scanned in
        // increasing order, columns come from non-overlapping <col>
        // spans) — `saturating_sub` is just arithmetic hygiene against a
        // malformed input, not validation of a case this milestone models.
        let hidden_row_count: u64 = hidden_rows
            .iter()
            .map(|iv| (iv.end - iv.start + 1) as u64)
            .sum();
        let hidden_col_count: u64 = hidden_columns
            .iter()
            .map(|iv| (iv.end - iv.start + 1) as u64)
            .sum();
        let rows = rect.rows() as u64;
        let cols = rect.cols() as u64;
        let total_cells = rows * cols;
        let visible_cells =
            rows.saturating_sub(hidden_row_count) * cols.saturating_sub(hidden_col_count);

        Some(HiddenCellsObservation {
            sheet: clip.src_sheet.clone(),
            address: clip.source_addr.clone(),
            rows: rect.rows(),
            columns: rect.cols(),
            hidden_rows,
            hidden_columns,
            total_cells,
            visible_cells,
        })
    }

    /// Loads a `.xlsx`/`.xlsm`/`.ods` file's sheets and cells into this `Vm`
    /// and sets the active sheet to the first one loaded. Returns the
    /// loaded sheet names (lowercase, in file order) on success. Extracted
    /// from `main.rs`'s run-mode `--file` handling (Milestone B5a) so the
    /// new `test-workbook` subcommand can reuse it instead of duplicating
    /// the loop.
    ///
    /// Two failure messages are preserved exactly so CLI callers can keep
    /// classifying them the way `--file` already does (`E3001`/`io_error`
    /// vs `E3002`/`sheet_setup_error`): a literal `"workbook has no sheets"`
    /// for an empty workbook, or `"cannot read '<path>': <reader error>"`
    /// for anything else.
    pub fn load_workbook_file(&mut self, path: &str) -> Result<Vec<String>, String> {
        self.loaded_workbook_name = std::path::Path::new(path)
            .file_name()
            .map(|n| n.to_string_lossy().to_string());
        self.loaded_workbook_path = Some(path.to_string());
        let sheets =
            reader::read_workbook(path).map_err(|e| format!("cannot read '{}': {}", path, e))?;
        if sheets.is_empty() {
            return Err("workbook has no sheets".to_string());
        }
        Ok(self.populate_from_sheets(sheets))
    }

    /// Populates this `Vm` from already-read sheet data and sets the active
    /// sheet to the first one. Split out from `load_workbook_file` so the
    /// mixed-case-sheet-name fix (see below) is unit-testable without going
    /// through a real file, since `save_workbook`-built fixtures always
    /// lowercase sheet names and would never exercise it.
    ///
    /// Clears `self.sheets` first: `Vm::new()` pre-seeds an empty `"sheet1"`
    /// so a macro that writes cells before any workbook is loaded has a
    /// valid `active_sheet` to land in. Without this clear, that default
    /// survives untouched into any loaded workbook whose sheets are never
    /// named `"Sheet1"`, and `sheet_names()`/the writer would carry it
    /// through as a genuine extra empty sheet on save (found via a
    /// synthetic two-sheet fixture named "First"/"Second" — the output
    /// gained a third, unrequested "sheet1"). `populate_from_sheets` is
    /// only ever called once per `Vm` (right after `Vm::new()`, before any
    /// macro runs — see `load_workbook_file` and `lib.rs`'s `load_workbook`
    /// PyO3 binding), so this is a full replace, not a data-losing merge.
    pub(crate) fn populate_from_sheets(&mut self, sheets: Vec<WorkbookSheet>) -> Vec<String> {
        self.sheets.clear();
        self.sheet_order.clear();
        let mut names = Vec::with_capacity(sheets.len());
        for sheet_data in &sheets {
            self.ensure_sheet(&sheet_data.name);
            let prev = self.active_sheet.clone();
            // Lowercased, matching `active_sheet`'s documented invariant —
            // `ensure_sheet` already lowercases the stored key, so leaving
            // this un-lowercased (as the pre-extraction code did) meant
            // `cells_mut()` couldn't find the sheet for any file with a
            // non-lowercase sheet name (found and fixed during extraction:
            // confirmed via a hand-crafted .xlsx with a sheet named "Input"
            // that panicked with "active sheet must exist" before this fix).
            self.active_sheet = sheet_data.name.to_lowercase();
            for (&(row, col), cell) in &sheet_data.cells {
                let value = match cell {
                    SheetCell::Integer(n) => Variant::Integer(*n),
                    SheetCell::Float(f) => Variant::Float(*f),
                    SheetCell::Str(s) => Variant::Str(s.clone()),
                    SheetCell::Bool(b) => Variant::Boolean(*b),
                    SheetCell::Error(e) => Variant::Error(e.clone()),
                };
                self.cells_mut().insert(
                    (row, col),
                    CellContent {
                        formula: sheet_data.formulas.get(&(row, col)).cloned(),
                        value,
                    },
                );
            }
            self.active_sheet = prev;
            let key = sheet_data.name.to_lowercase();
            if !sheet_data.merged_ranges.is_empty() {
                self.merged_ranges
                    .insert(key.clone(), sheet_data.merged_ranges.clone());
            }
            if !sheet_data.raw_style_indices.is_empty() {
                self.cell_style_indices
                    .insert(key.clone(), sheet_data.raw_style_indices.clone());
            }
            if !sheet_data.cell_number_formats.is_empty() {
                self.cell_number_formats
                    .insert(key.clone(), sheet_data.cell_number_formats.clone());
            }
            if !sheet_data.hidden_rows.is_empty() || !sheet_data.hidden_columns.is_empty() {
                self.sheet_visibility.insert(
                    key.clone(),
                    SheetVisibility {
                        hidden_rows: sheet_data
                            .hidden_rows
                            .iter()
                            .map(|&(start, end)| Interval { start, end })
                            .collect(),
                        hidden_columns: sheet_data
                            .hidden_columns
                            .iter()
                            .map(|&(start, end)| Interval { start, end })
                            .collect(),
                    },
                );
            }
            self.worksheet_origins.insert(
                key.clone(),
                WorksheetOrigin {
                    original_sheet_id: sheet_data.sheet_id.clone(),
                    original_workbook_rel_id: sheet_data.workbook_rel_id.clone(),
                    original_part_name: sheet_data.source_part_name.clone(),
                    original_display_name: Some(sheet_data.name.clone()),
                },
            );
            names.push(key);
        }
        let first = names[0].clone();
        self.set_active_sheet(&first)
            .expect("just-inserted sheet must exist");
        names
    }

    fn sheet_cells_mut(&mut self, name: &str) -> Option<&mut HashMap<(u32, u32), CellContent>> {
        self.cell_index_dirty = true;
        self.sheets.get_mut(&name.to_lowercase())
    }

    /// Drain and return every MsgBox message recorded since the last call
    /// (or since `run_sub` started, since `run_sub` clears the log first).
    pub fn take_messages(&mut self) -> Vec<String> {
        std::mem::take(&mut self.msgbox_log)
    }

    /// Span of the statement that was executing the last time `exec_stmt`
    /// ran — i.e. where a runtime error happened, if `run_sub` just
    /// returned one. `None` if no statement has executed yet.
    pub fn current_span(&self) -> Option<SourceSpan> {
        self.current_span
    }

    pub fn run_sub(&mut self, program: &Program, sub_name: &str) -> Result<(), String> {
        // Each run starts with a clean message log — otherwise a Vm reused
        // across multiple run_sub calls (e.g. from the Python bindings)
        // would leak the previous run's MsgBox text into this run's result.
        self.msgbox_log.clear();
        self.last_resolution_failure = None;
        self.err_number = 0;
        self.err_description.clear();
        self.err_source.clear();
        self.err_help_file.clear();
        self.err_help_context = 0;
        self.pending_raised_error = None;
        // A Vm reused across multiple run_sub calls must not carry the
        // previous run's call frames (or their On Error state) into this
        // one — call_sub_def pushes the entrypoint's own frame below, so
        // this only needs to be empty, not seeded.
        self.call_stack.clear();
        self.option_base = program.option_base;
        // Cache user-defined functions, subs, and type definitions.
        self.user_funcs = program
            .funcs
            .iter()
            .map(|f| (f.name.clone(), f.clone()))
            .collect();
        self.user_subs = program
            .subs
            .iter()
            .map(|s| (s.name.clone(), s.clone()))
            .collect();
        for td in &program.type_defs {
            self.type_defs.insert(td.name.clone(), td.fields.clone());
        }

        // Pre-flight compile-time check, run before any statement (including
        // the entrypoint's own first line) executes — see
        // `check::compile_check_errors`'s own doc comment for exactly what
        // this catches. Running it here, ahead of `call_sub_def`, is what
        // makes these errors uncatchable by `On Error` for free: no `On
        // Error` statement has had a chance to take effect yet, matching
        // real VBA (these are compile errors, not runtime ones).
        if let Some((msg, span)) = check::compile_check_errors(program, &HashSet::new()) {
            self.current_span = Some(span);
            return Err(msg);
        }

        let name = sub_name.to_lowercase();
        let sub = self
            .user_subs
            .get(&name)
            .ok_or_else(|| format!("Sub '{}' not found", sub_name))?
            .clone();
        self.call_sub_def(&sub, &[])
    }

    /// Multi-module entrypoint (Milestone B2): `modules` is a list of
    /// (module_name, Program) pairs. Rejects the run at load time if any
    /// bare Sub or Function name collides across modules — the flat merge
    /// used for in-body calls can't express VBA's own-module-first/Private
    /// scoping, so a colliding name is refused rather than resolved
    /// silently (see `parser::find_cross_module_sub_collisions`). Otherwise
    /// behaves like `run_sub`, generalized to N modules; `entrypoint` may be
    /// a bare name or a `Module.Sub`-qualified one.
    pub fn run_sub_multi(
        &mut self,
        modules: &[(String, Program)],
        entrypoint: &str,
    ) -> Result<(), String> {
        let sub_collisions = parser::find_cross_module_sub_collisions(modules);
        if let Some((name, mods)) = sub_collisions.first() {
            return Err(format!(
                "duplicate Sub '{}' across modules '{}' — cross-module name collisions aren't supported yet; own-module-first/Private scoping isn't modeled — rename one of them",
                name,
                mods.join("', '")
            ));
        }
        let func_collisions = parser::find_cross_module_func_collisions(modules);
        if let Some((name, mods)) = func_collisions.first() {
            return Err(format!(
                "duplicate Function '{}' across modules '{}' — cross-module name collisions aren't supported yet; own-module-first/Private scoping isn't modeled — rename one of them",
                name,
                mods.join("', '")
            ));
        }

        self.msgbox_log.clear();
        self.last_resolution_failure = None;
        self.err_number = 0;
        self.err_description.clear();
        self.err_source.clear();
        self.err_help_file.clear();
        self.err_help_context = 0;
        self.pending_raised_error = None;
        self.call_stack.clear();
        // Real VBA scopes `Option Base` per module; this codebase's `Vm`
        // is a single flat namespace across every loaded module (same
        // simplification `user_funcs`/`user_subs`/`type_defs` already
        // make), so this takes the first module that declares one rather
        // than modeling true per-module scoping.
        self.option_base = modules
            .iter()
            .map(|(_, p)| p.option_base)
            .find(|&b| b != 0)
            .unwrap_or(0);
        self.user_funcs.clear();
        self.user_subs.clear();
        for (_, program) in modules {
            for f in &program.funcs {
                self.user_funcs.insert(f.name.clone(), f.clone());
            }
            for s in &program.subs {
                self.user_subs.insert(s.name.clone(), s.clone());
            }
            for td in &program.type_defs {
                self.type_defs.insert(td.name.clone(), td.fields.clone());
            }
        }

        // Same pre-flight compile-time check as `run_sub`, run once per
        // module — each module only sees its own `Program`, so
        // `other_module_names` (every bare Sub/Function name declared in
        // every *other* module) is built the same way
        // `main.rs`'s own multi-module `elixcee check` path already does,
        // or a legitimate unqualified cross-module call would be
        // misreported as undefined.
        for (name, program) in modules {
            let mut other_module_names: HashSet<String> = HashSet::new();
            for (other_name, other_program) in modules {
                if other_name != name {
                    other_module_names.extend(other_program.subs.iter().map(|s| s.name.clone()));
                    other_module_names.extend(other_program.funcs.iter().map(|f| f.name.clone()));
                }
            }
            if let Some((msg, span)) = check::compile_check_errors(program, &other_module_names) {
                self.current_span = Some(span);
                return Err(msg);
            }
        }

        let sub = match parser::resolve_entrypoint(modules, entrypoint) {
            EntrypointResolution::Found(sub) => sub.clone(),
            EntrypointResolution::NotFound => {
                return Err(format!("Sub '{}' not found", entrypoint));
            }
        };
        self.call_sub_def(&sub, &[])
    }

    fn call_sub_def(&mut self, sub: &SubDef, args: &[Variant]) -> Result<(), String> {
        let saved: Vec<(String, Option<Variant>)> = sub
            .params
            .iter()
            .enumerate()
            .map(|(i, p)| {
                let old = self.variables.get(p).cloned();
                if let Some(v) = args.get(i) {
                    self.variables.insert(p.clone(), v.clone());
                }
                (p.clone(), old)
            })
            .collect();
        let body = sub.body.clone();
        // A fresh frame per call — real VBA's `On Error` scope is the
        // procedure, not the call site, so this callee starts with no
        // active handler regardless of what the caller's own frame has set
        // (see `exec_body`'s doc comment). Popped before propagating any
        // error, so the caller's own frame is what a failure that escapes
        // this call is actually checked against.
        self.call_stack.push(CallFrame {
            procedure_name: sub.name.clone(),
            error_mode: ErrorMode::Disabled,
        });
        let result = self.exec_body(&body, |f| matches!(f, ExitKind::Sub));
        self.call_stack.pop();
        result?;
        for (p, old) in saved {
            match old {
                Some(v) => {
                    self.variables.insert(p, v);
                }
                None => {
                    self.variables.remove(&p);
                }
            }
        }
        Ok(())
    }

    fn call_func_def(&mut self, func: &FuncDef, args: &[Variant]) -> Result<Variant, String> {
        let saved: Vec<(String, Option<Variant>)> = func
            .params
            .iter()
            .enumerate()
            .map(|(i, p)| {
                let old = self.variables.get(p).cloned();
                if let Some(v) = args.get(i) {
                    self.variables.insert(p.clone(), v.clone());
                }
                (p.clone(), old)
            })
            .collect();
        let ret_name = func.name.clone();
        let old_ret = self.variables.remove(&ret_name);
        let body = func.body.clone();
        self.call_stack.push(CallFrame {
            procedure_name: func.name.clone(),
            error_mode: ErrorMode::Disabled,
        });
        let result = self.exec_body(&body, |f| matches!(f, ExitKind::Function | ExitKind::Sub));
        self.call_stack.pop();
        result?;
        let ret_val = self.variables.remove(&ret_name).unwrap_or(Variant::Empty);
        for (p, old) in saved {
            match old {
                Some(v) => {
                    self.variables.insert(p, v);
                }
                None => {
                    self.variables.remove(&p);
                }
            }
        }
        if let Some(v) = old_ret {
            self.variables.insert(ret_name, v);
        }
        Ok(ret_val)
    }

    /// Execute a body slice with label-jump support (for GoTo and On Error GoTo).
    /// The existing per-statement `exec_stmt` error catch (resume_next) is preserved.
    fn exec_body<F>(&mut self, stmts: &[SpannedStmt], is_exit: F) -> Result<(), String>
    where
        F: Fn(&ExitKind) -> bool,
    {
        let mut i = 0;
        while i < stmts.len() {
            // Handle pending unconditional GoTo
            if let Some(label) = self.pending_goto.take() {
                match stmts
                    .iter()
                    .position(|s| matches!(&s.stmt, Stmt::Label(l) if l == &label))
                {
                    Some(pos) => {
                        i = pos;
                        continue;
                    }
                    None => return Err(format!("GoTo: label '{}' not found", label)),
                }
            }

            if let Some(ref f) = self.exit_flag {
                if is_exit(f) {
                    self.exit_flag = None;
                    break;
                }
                break; // other exit kinds bubble up
            }

            let result = self.exec_stmt(&stmts[i]); // per-stmt catch preserves resume_next behavior
            match result {
                Ok(()) => {}
                Err(e) => {
                    // On Error GoTo: jump to handler label — skipped in
                    // strict-resolution mode (`diagnose`) so the first
                    // failure always propagates instead of being redirected
                    // to a handler that would mask it. Reads/consumes only
                    // the *current* frame's mode (the procedure whose body
                    // this `stmts` slice belongs to) — a callee's own frame
                    // never sees or resolves labels that only exist in a
                    // caller's body; see `call_sub_def`/`call_func_def`.
                    if !self.strict_resolution
                        && let Some(ErrorMode::GoTo(label)) = self.current_error_mode()
                    {
                        self.set_current_error_mode(ErrorMode::Disabled);
                        self.record_error(&e);
                        match stmts
                            .iter()
                            .position(|s| matches!(&s.stmt, Stmt::Label(l) if l == &label))
                        {
                            Some(pos) => {
                                i = pos;
                                continue;
                            }
                            None => {
                                return Err(format!("On Error GoTo: label '{}' not found", label));
                            }
                        }
                    }
                    return Err(e);
                }
            }
            i += 1;
        }
        Ok(())
    }

    /// The innermost (currently-executing) call frame's `ErrorMode` — the
    /// procedure this statement is actually running inside, not any
    /// caller's. `None` only if called with an empty `call_stack`, which
    /// shouldn't happen while any statement is executing (`call_sub_def`/
    /// `call_func_def` always push a frame first).
    fn current_error_mode(&self) -> Option<ErrorMode> {
        self.call_stack.last().map(|f| f.error_mode.clone())
    }

    fn set_current_error_mode(&mut self, mode: ErrorMode) {
        if let Some(frame) = self.call_stack.last_mut() {
            frame.error_mode = mode;
        }
    }

    fn exec_stmt(&mut self, spanned: &SpannedStmt) -> Result<(), String> {
        if self.exit_flag.is_some() {
            return Ok(());
        }
        self.current_span = Some(spanned.span);
        let result = self.exec_stmt_inner(&spanned.stmt);
        match result {
            Ok(()) => Ok(()),
            // `On Error Resume Next` is not honored in strict-resolution
            // mode (`diagnose`) — see the field doc on `strict_resolution`.
            Err(e)
                if !self.strict_resolution
                    && matches!(self.current_error_mode(), Some(ErrorMode::ResumeNext)) =>
            {
                self.record_error(&e);
                Ok(())
            }
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
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                let v = self.eval_expr(value)?;
                self.cells_mut().insert(
                    (r, c),
                    CellContent {
                        formula: None,
                        value: v,
                    },
                );
            }
            Stmt::SetCalcMode(mode) => {
                let m = match mode {
                    CalcModeValue::Automatic => CalculationMode::Automatic,
                    CalcModeValue::Manual => CalculationMode::Manual,
                };
                self.set_calc_mode(m)?;
            }
            Stmt::For {
                var,
                from,
                to,
                step,
                body,
            } => {
                let mut i = to_f64(&self.eval_expr(from)?)?;
                let to_f = to_f64(&self.eval_expr(to)?)?;
                let step_f = match step {
                    Some(s) => to_f64(&self.eval_expr(s)?)?,
                    None => 1.0,
                };
                if step_f == 0.0 {
                    return Err("For loop: step cannot be zero".into());
                }
                'for_loop: while (step_f > 0.0 && i <= to_f) || (step_f < 0.0 && i >= to_f) {
                    self.check_deadline()?;
                    self.variables.insert(var.clone(), as_int_if_whole(i));
                    for s in body {
                        self.exec_stmt(s)?;
                        if matches!(self.exit_flag, Some(ExitKind::For)) {
                            self.exit_flag = None;
                            break 'for_loop;
                        }
                        if self.exit_flag.is_some() {
                            return Ok(());
                        }
                    }
                    i += step_f;
                }
            }
            Stmt::ForEach {
                var,
                range_addr,
                body,
            } => {
                let ((r1, c1), (r2, c2)) = self
                    .resolve_range_addr(range_addr)
                    .ok_or_else(|| format!("ForEach: invalid range '{}'", range_addr))?;
                'fe_outer: for r in r1..=r2 {
                    for c in c1..=c2 {
                        self.check_deadline()?;
                        let v = self.get_cell(r, c);
                        self.variables.insert(var.clone(), v);
                        // `For Each c In Range(...)` binds `c` to a real
                        // single-cell Range *object*, not just the cell's
                        // value — so `c.Value` reads that cell (it used to
                        // fall through to the UDT path and silently yield
                        // Empty) and a `Dim c As Range` loop variable stops
                        // reading as the never-Set Nothing its declaration
                        // registered. The plain value stays in `variables`
                        // too, so a bare `c` in an arithmetic context keeps
                        // working exactly as before (VBA's own default-
                        // property behavior).
                        self.object_variables.insert(
                            var.clone(),
                            ObjectRef::Range(RangeRef::single(
                                self.active_sheet.clone(),
                                Rect {
                                    start_row: r,
                                    start_col: c,
                                    end_row: r,
                                    end_col: c,
                                },
                            )),
                        );
                        for s in body {
                            self.exec_stmt(s)?;
                            if matches!(self.exit_flag, Some(ExitKind::For)) {
                                self.exit_flag = None;
                                break 'fe_outer;
                            }
                            if self.exit_flag.is_some() {
                                return Ok(());
                            }
                        }
                    }
                }
            }
            Stmt::If {
                condition,
                then_body,
                else_body,
            } => {
                let branch = if is_truthy(&self.eval_expr(condition)?) {
                    then_body
                } else {
                    else_body
                };
                for s in branch {
                    self.exec_stmt(s)?;
                    if self.exit_flag.is_some() {
                        return Ok(());
                    }
                }
            }
            Stmt::DoLoop {
                pre_cond,
                post_cond,
                body,
            } => {
                let check = |vm: &mut Vm, cond: &Option<(bool, Expr)>| -> Result<bool, String> {
                    match cond {
                        None => Ok(true),
                        Some((is_until, expr)) => {
                            let v = vm.eval_expr(expr)?;
                            Ok(if *is_until {
                                !is_truthy(&v)
                            } else {
                                is_truthy(&v)
                            })
                        }
                    }
                };
                'do_loop: while check(self, pre_cond)? {
                    self.check_deadline()?;
                    for s in body.clone() {
                        self.exec_stmt(&s)?;
                        if matches!(self.exit_flag, Some(ExitKind::Do)) {
                            self.exit_flag = None;
                            break 'do_loop;
                        }
                        if self.exit_flag.is_some() {
                            return Ok(());
                        }
                    }
                    if !check(self, post_cond)? {
                        break 'do_loop;
                    }
                }
            }
            Stmt::SelectCase {
                expr,
                cases,
                else_body,
            } => {
                let val = self.eval_expr(expr)?;
                let mut matched = false;
                'outer: for (matchers, body) in cases {
                    for m in matchers {
                        let hit = match m {
                            CaseMatch::Value(v) => vba_eq(&val, &self.eval_expr(v)?),
                            CaseMatch::Range(lo, hi) => {
                                let l = self.eval_expr(lo)?;
                                let h = self.eval_expr(hi)?;
                                vba_cmp(&val, &l)? != std::cmp::Ordering::Less
                                    && vba_cmp(&val, &h)? != std::cmp::Ordering::Greater
                            }
                            CaseMatch::IsOp(op, rhs) => {
                                let r = self.eval_expr(rhs)?;
                                match op {
                                    VbaBinOp::Eq => vba_eq(&val, &r),
                                    VbaBinOp::Ne => !vba_eq(&val, &r),
                                    VbaBinOp::Lt => vba_cmp(&val, &r)? == std::cmp::Ordering::Less,
                                    VbaBinOp::Le => {
                                        vba_cmp(&val, &r)? != std::cmp::Ordering::Greater
                                    }
                                    VbaBinOp::Gt => {
                                        vba_cmp(&val, &r)? == std::cmp::Ordering::Greater
                                    }
                                    VbaBinOp::Ge => vba_cmp(&val, &r)? != std::cmp::Ordering::Less,
                                    _ => false,
                                }
                            }
                        };
                        if hit {
                            for s in body {
                                self.exec_stmt(s)?;
                            }
                            matched = true;
                            break 'outer;
                        }
                    }
                }
                if !matched {
                    for s in else_body {
                        self.exec_stmt(s)?;
                    }
                }
            }
            Stmt::ExitFor => self.exit_flag = Some(ExitKind::For),
            Stmt::ExitDo => self.exit_flag = Some(ExitKind::Do),
            Stmt::ExitSub => self.exit_flag = Some(ExitKind::Sub),
            Stmt::ExitFunction => self.exit_flag = Some(ExitKind::Function),
            Stmt::OnError { resume_next } => {
                // `On Error Resume Next` (true) / `On Error GoTo 0` (false)
                // — both scoped to the current frame only.
                self.set_current_error_mode(if *resume_next {
                    ErrorMode::ResumeNext
                } else {
                    ErrorMode::Disabled
                });
            }
            Stmt::OnErrorGoTo(label) => {
                self.set_current_error_mode(ErrorMode::GoTo(label.clone()));
            }
            Stmt::ErrClear => {
                self.err_number = 0;
                self.err_description.clear();
                self.err_source.clear();
                self.err_help_file.clear();
                self.err_help_context = 0;
            }
            Stmt::ErrRaise {
                number,
                source,
                description,
                help_file,
                help_context,
            } => {
                let number = to_i64_rounded(&self.eval_expr(number)?)?;
                let source = match source {
                    Some(s) => self.eval_expr(s)?.to_string(),
                    None => String::new(),
                };
                let description = match description {
                    Some(d) => self.eval_expr(d)?.to_string(),
                    None => default_description_for_vba_error_number(number).to_string(),
                };
                let help_file = match help_file {
                    Some(h) => self.eval_expr(h)?.to_string(),
                    None => String::new(),
                };
                let help_context = match help_context {
                    Some(h) => to_i64_rounded(&self.eval_expr(h)?)?,
                    None => 0,
                };
                self.pending_raised_error = Some(RaisedError {
                    number,
                    description: description.clone(),
                    source,
                    help_file,
                    help_context,
                });
                return Err(description);
            }
            Stmt::Label(_) => {} // no-op during normal execution
            Stmt::GoTo(label) => {
                self.pending_goto = Some(label.clone());
            }
            Stmt::Resume { .. } => {
                // After error handler runs: clear error state, continue.
                // Already `Disabled` by the time a `GoTo` handler's own
                // body reaches a `Resume` (consumed on the jump itself —
                // see `exec_body`), so this only matters if the handler
                // ran under `Resume Next` instead.
                self.set_current_error_mode(ErrorMode::Disabled);
            }
            Stmt::CallSub { name, args } => {
                let arg_vals: Vec<Variant> = args
                    .iter()
                    .map(|a| self.eval_expr(a))
                    .collect::<Result<_, _>>()?;
                if let Some(func) = self.user_funcs.get(name).cloned() {
                    let _ = self.call_func_def(&func, &arg_vals)?;
                } else if let Some(sub) = self.user_subs.get(name).cloned() {
                    self.call_sub_def(&sub, &arg_vals)?;
                } else {
                    return Err(format!("Sub/Function '{}' not found", name));
                }
            }
            Stmt::SetAppProp { prop, value } => {
                let v = self.eval_expr(value);
                if prop == "cutcopymode"
                    && let Ok(v) = &v
                    && !is_truthy(v)
                {
                    self.clipboard = None;
                }
            }
            Stmt::RangeName { addr, name } => {
                self.named_ranges.insert(name.to_lowercase(), addr.clone());
            }
            Stmt::RangeWrite {
                addr,
                is_formula,
                value,
            } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let v = self.eval_expr(value)?;
                let ((r1, c1), (r2, c2)) = self
                    .resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeWrite: invalid address '{}'", addr))?;
                if *is_formula {
                    let s = vba_to_str(&v);
                    for r in r1..=r2 {
                        for c in c1..=c2 {
                            self.set_cell_formula(r, c, &s)?;
                        }
                    }
                } else {
                    // Batch writes: access sheet directly to avoid N dirty-flag sets
                    let sheet = self.active_sheet.clone();
                    if let Some(cells) = self.sheets.get_mut(&sheet) {
                        for r in r1..=r2 {
                            for c in c1..=c2 {
                                cells.insert(
                                    (r, c),
                                    CellContent {
                                        formula: None,
                                        value: v.clone(),
                                    },
                                );
                            }
                        }
                    }
                    self.cell_index_dirty = true;
                }
            }
            Stmt::RangeClear { addr, .. } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let ((r1, c1), (r2, c2)) = self
                    .resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeClear: invalid address '{}'", addr))?;
                let sheet = self.active_sheet.clone();
                if let Some(cells) = self.sheets.get_mut(&sheet) {
                    for r in r1..=r2 {
                        for c in c1..=c2 {
                            cells.remove(&(r, c));
                        }
                    }
                }
                self.cell_index_dirty = true;
            }
            Stmt::RangeOffsetWrite {
                addr,
                row_off,
                col_off,
                value,
            } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let v = self.eval_expr(value)?;
                let (base_r, base_c) = parse_cell_addr(addr)
                    .ok_or_else(|| format!("RangeOffsetWrite: invalid address '{}'", addr))?;
                let ro = to_f64(&self.eval_expr(row_off)?)? as i64;
                let co = to_f64(&self.eval_expr(col_off)?)? as i64;
                let row = (base_r as i64 + ro) as u32;
                let col = (base_c as i64 + co) as u32;
                self.cells_mut().insert(
                    (row, col),
                    CellContent {
                        formula: None,
                        value: v,
                    },
                );
            }
            Stmt::RangeDelete { addr, axis } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let ((r1, c1), (r2, c2)) = self
                    .resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeDelete: invalid address '{}'", addr))?;
                match axis {
                    Axis::Row => self.delete_rows(r1, r2 - r1 + 1),
                    Axis::Column => self.delete_cols(c1, c2 - c1 + 1),
                }
            }
            Stmt::RangeInsert { addr, axis } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let ((r1, c1), (r2, c2)) = self
                    .resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeInsert: invalid address '{}'", addr))?;
                match axis {
                    Axis::Row => self.insert_rows(r1, r2 - r1 + 1),
                    Axis::Column => self.insert_cols(c1, c2 - c1 + 1),
                }
            }
            Stmt::RowColDelete { axis, index } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let idx = to_f64(&self.eval_expr(index)?)? as u32;
                match axis {
                    Axis::Row => self.delete_rows(idx, 1),
                    Axis::Column => self.delete_cols(idx, 1),
                }
            }
            Stmt::RowColInsert { axis, index } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let idx = to_f64(&self.eval_expr(index)?)? as u32;
                match axis {
                    Axis::Row => self.insert_rows(idx, 1),
                    Axis::Column => self.insert_cols(idx, 1),
                }
            }
            Stmt::RangeSort {
                addr,
                key_col,
                descending,
                header,
            } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let ((r1, c1), (r2, c2)) = self
                    .resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeSort: invalid address '{}'", addr))?;
                self.sort_range_on_sheet(&active, r1, c1, r2, c2, *key_col, *descending, *header);
            }
            Stmt::RangeAutoFilter {
                addr,
                field,
                criteria1,
            } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let ((r1, c1), (r2, _)) = self
                    .resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeAutoFilter: invalid address '{}'", addr))?;
                // A bare AutoFilter (no Field/Criteria1) is a real no-op here -- see
                // RangeAutoFilter's own doc comment for why (no <autoFilter> element is
                // persisted, so there'd be nothing to visibly turn on either way).
                if let (Some(field_expr), Some(criteria_expr)) = (field, criteria1) {
                    // Field is 1-based, relative to addr's own left edge (real VBA's
                    // convention -- "the leftmost field is 1"), not an absolute column.
                    let field_off = to_f64(&self.eval_expr(field_expr)?)? as u32;
                    let filter_col = c1 + field_off.saturating_sub(1);
                    let criteria = vba_to_str(&self.eval_expr(criteria_expr)?);
                    // Row r1 is always the header -- AutoFilter never hides it.
                    let mut newly_hidden = Vec::new();
                    for r in (r1 + 1)..=r2 {
                        if vba_to_str(&self.get_cell(r, filter_col)) != criteria {
                            newly_hidden.push(Interval { start: r, end: r });
                        }
                    }
                    if !newly_hidden.is_empty() {
                        self.sheet_visibility
                            .entry(active)
                            .or_default()
                            .hidden_rows
                            .extend(newly_hidden);
                    }
                }
            }
            Stmt::RangeCopy { src, dst } => {
                let areas = self
                    .resolve_multi_area_addr(src)
                    .ok_or_else(|| format!("RangeCopy: invalid source range '{}'", src))?;
                // `resolve_multi_area_addr` never returns `Some(vec![])` — every
                // comma-separated piece (at least 1) must itself parse.
                let sheet = self.active_sheet.clone();
                self.copy_areas_to_clipboard(sheet, areas, src.clone());
                if let Some(dst_addr) = dst {
                    self.do_paste(dst_addr, false)?;
                }
            }
            Stmt::RangeObjectCopy { var, dst } => {
                self.require_live_object(var)?;
                let obj = self
                    .object_variables
                    .get(var)
                    .cloned()
                    .ok_or_else(|| format!("'{}' is Nothing — Set was never called", var))?;
                let r = expect_range_ref(obj, "Copy")?;
                let display = format!("<{}>", var);
                self.copy_areas_to_clipboard(r.sheet, r.areas, display);
                if let Some(dst_addr) = dst {
                    self.do_paste(dst_addr, false)?;
                }
            }
            Stmt::Set { var, value } => {
                if let ObjectExpr::Var(name) = value {
                    // `Set ws = ActiveSheet` / `Set wb = ThisWorkbook` /
                    // `Set wb = ActiveWorkbook` (Phase 2C items 7/8) — these
                    // three parse as a bare `ObjectExpr::Var` the same as
                    // any other identifier in object position (the parser
                    // can't distinguish them from a real object-variable
                    // reference at parse time — see `ast::ObjectExpr::Var`'s
                    // doc), so they're recognized here by name, ahead of the
                    // generic "not a live object variable" no-op below.
                    // Previously all three fell into that no-op — a bare
                    // `Set ws = ActiveSheet` silently did nothing.
                    match name.as_str() {
                        "activesheet" => {
                            self.object_variables.insert(
                                var.clone(),
                                ObjectRef::Worksheet(self.active_sheet.clone()),
                            );
                            return Ok(());
                        }
                        "thisworkbook" | "activeworkbook" => {
                            self.object_variables
                                .insert(var.clone(), ObjectRef::Workbook);
                            return Ok(());
                        }
                        // `Set r = Nothing` clears ONLY this variable's own
                        // reference. Any variable previously assigned *from*
                        // it (`Set r2 = r`) keeps pointing at the real object
                        // — `ObjectRef` is copied by value into each
                        // variable's own slot, so there is no shared cell to
                        // clear. Matched here by name, ahead of the generic
                        // "not a live object variable" no-op below, which is
                        // where this used to land (silently doing nothing).
                        "nothing" => {
                            self.object_variables
                                .insert(var.clone(), ObjectRef::Nothing);
                            return Ok(());
                        }
                        _ => {}
                    }
                }
                if let ObjectExpr::Var(name) = value
                    && !self.object_variables.contains_key(name)
                {
                    // A bare identifier in object position that isn't a
                    // live object variable and isn't one of the three names
                    // handled above — either a genuinely unset `Set b = a`
                    // or another unmodeled VBA object keyword the parser
                    // can't distinguish from a variable reference at parse
                    // time (`Selection`, `Nothing`, ... — `Nothing` in
                    // particular needs this: `Set rng = Nothing` must stay a
                    // no-op, never a hard error). No-op, same precedent as
                    // `Stmt::Dim`/`Stmt::Unsupported` for any other
                    // unmodeled construct — safer than guessing wrong and
                    // raising a confusing runtime error. Errors from a
                    // *resolvable* object expression (an out-of-range
                    // `Areas(n)`, a cross-sheet `Union`, an invalid
                    // `Range(...)` address) are unaffected and still
                    // propagate below.
                    return Ok(());
                }
                let obj = self.eval_object_expr(value)?;
                self.object_variables.insert(var.clone(), obj);
            }
            Stmt::RangePaste {
                dest_addr,
                transpose,
            } => {
                let t = match transpose {
                    Some(e) => is_truthy(&self.eval_expr(e)?),
                    None => false,
                };
                self.do_paste(dest_addr, t)?;
            }
            Stmt::SheetRangePaste { sheet, dest_addr } => {
                let (key, display) = self.resolve_sheet_expr(sheet)?;
                self.check_strict_sheet_exists(&display, &key)?;
                let prev = self.active_sheet.clone();
                if !self.strict_resolution {
                    self.ensure_sheet(&key);
                }
                self.active_sheet = key;
                self.cell_index_dirty = true;
                let result = self.do_paste(dest_addr, false);
                self.active_sheet = prev;
                self.cell_index_dirty = true;
                result?;
            }
            Stmt::SheetCellWrite {
                sheet,
                row,
                col,
                value,
            } => {
                let (key, display) = self.resolve_sheet_expr(sheet)?;
                self.check_strict_sheet_exists(&display, &key)?;
                self.check_sheet_not_protected(&key, &display)?;
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                let v = self.eval_expr(value)?;
                if !self.strict_resolution {
                    self.ensure_sheet(&key);
                }
                self.sheet_cells_mut(&key).unwrap().insert(
                    (r, c),
                    CellContent {
                        formula: None,
                        value: v,
                    },
                );
            }
            Stmt::SheetRangeWrite {
                sheet,
                addr,
                is_formula,
                value,
            } => {
                let (key, display) = self.resolve_sheet_expr(sheet)?;
                self.check_strict_sheet_exists(&display, &key)?;
                self.check_sheet_not_protected(&key, &display)?;
                let ((r1, c1), (r2, c2)) = parse_range_addr(addr)
                    .ok_or_else(|| format!("SheetRangeWrite: invalid address '{}'", addr))?;
                let v = self.eval_expr(value)?;
                if !self.strict_resolution {
                    self.ensure_sheet(&key);
                }
                if *is_formula {
                    let s = vba_to_str(&v);
                    let prev = self.active_sheet.clone();
                    self.active_sheet = key.clone();
                    for r in r1..=r2 {
                        for c in c1..=c2 {
                            self.set_cell_formula(r, c, &s)?;
                        }
                    }
                    self.active_sheet = prev;
                } else if let Some(cells) = self.sheet_cells_mut(&key) {
                    for r in r1..=r2 {
                        for c in c1..=c2 {
                            cells.insert(
                                (r, c),
                                CellContent {
                                    formula: None,
                                    value: v.clone(),
                                },
                            );
                        }
                    }
                    self.cell_index_dirty = true;
                }
            }
            Stmt::WithSheet { sheet_name, body } => {
                self.check_strict_sheet_exists(sheet_name, &sheet_name.to_lowercase())?;
                let prev = self.active_sheet.clone();
                if !self.strict_resolution {
                    self.ensure_sheet(sheet_name);
                }
                let key = sheet_name.to_lowercase();
                self.active_sheet = key.clone();
                self.cell_index_dirty = true;
                // Pushes onto the runtime With stack *as well as* swapping
                // the active sheet: the swap is what makes an ordinary
                // `Range(...)`/`Cells(...)` statement in the body land on
                // this sheet, the stack entry is what makes a bare
                // `.member` resolve against it (bare `.member` is no longer
                // rewritten at parse time, so without this push it would
                // find an empty stack).
                let result = self.run_with_body(WithValue::Sheet(key), body);
                self.active_sheet = prev;
                self.cell_index_dirty = true;
                result?;
            }
            Stmt::SheetsAdd => {
                // `self.sheets.len() + 1` alone collides whenever the sheet
                // set has a gap (e.g. Sheet2 was deleted, leaving Sheet1/
                // Sheet3 -- len()==2 would compute "sheet3", which already
                // exists). ensure_sheet() no-ops on an existing key, so an
                // unguarded collision silently drops the Add entirely.
                // Probe upward from the same starting point until a free
                // name is found.
                let mut n = self.sheets.len() + 1;
                let mut new_name = format!("sheet{n}");
                while self.sheets.contains_key(&new_name) {
                    n += 1;
                    new_name = format!("sheet{n}");
                }
                self.ensure_sheet(&new_name);
            }
            Stmt::SheetsDelete { sheet } => {
                let (key, display) = self.resolve_sheet_expr(sheet)?;
                self.remove_sheet(&key, &display)?;
            }
            Stmt::SheetProtection {
                sheet,
                protect,
                ui_only,
            } => {
                let (key, display) = self.resolve_sheet_expr(sheet)?;
                self.require_sheet_exists(&display, &key)?;
                if *protect {
                    // UserInterfaceOnly:=True means real Excel blocks manual
                    // UI edits but not macro writes — since elixcee has no
                    // UI to block, that leaves the sheet macro-writable.
                    let ui_only = match ui_only {
                        Some(e) => is_truthy(&self.eval_expr(e)?),
                        None => false,
                    };
                    if !ui_only {
                        self.protected_sheets.insert(key);
                    }
                } else {
                    self.protected_sheets.remove(&key);
                }
            }
            Stmt::Dim => {}
            Stmt::DimBare { var } => {
                // Registers the name as a real, Empty-valued variable —
                // `IsEmpty(x)` must be True right after `Dim x` runs, not
                // "Undefined variable" (the old bare `Stmt::Dim` never
                // recorded the name at all). `entry(...).or_insert(...)`
                // rather than an unconditional `insert`: if this statement
                // somehow re-executes (e.g. inside a loop body — unusual
                // VBA style, but not illegal), it must not reset an
                // already-assigned value back to Empty, matching real
                // VBA's own Dim-is-a-declaration-not-a-reset semantics.
                self.variables.entry(var.clone()).or_insert(Variant::Empty);
            }
            Stmt::DimMulti(decls) => {
                for s in decls {
                    self.exec_stmt_inner(s)?;
                }
            }
            Stmt::Unsupported { .. } => {}
            Stmt::DimArray { name, sizes } => {
                if sizes.is_empty() {
                    // `Dim arr()` — dynamic array, unsized until a later
                    // `ReDim`. Rank-0 rather than rank-1-empty: it doesn't
                    // yet have a shape at all, so any subscript access on it
                    // (rank always mismatches) correctly errors with
                    // "Subscript out of range", same as real VBA.
                    self.variables.insert(
                        name.clone(),
                        Variant::VbaArray(VbaArray {
                            bounds: vec![],
                            elements: vec![],
                        }),
                    );
                } else {
                    let bounds = self.eval_array_bounds(sizes)?;
                    let arr = VbaArray::new_zeroed(bounds)?;
                    self.variables.insert(name.clone(), Variant::VbaArray(arr));
                }
            }
            Stmt::ReDim {
                name,
                sizes,
                preserve,
            } => {
                let bounds = self.eval_array_bounds(sizes)?;
                let new_arr = if *preserve {
                    match self.variables.get(name) {
                        Some(Variant::VbaArray(old)) if old.rank() == bounds.len() => {
                            redim_preserve(old, &bounds)?
                        }
                        _ => VbaArray::new_zeroed(bounds)?,
                    }
                } else {
                    VbaArray::new_zeroed(bounds)?
                };
                self.variables
                    .insert(name.clone(), Variant::VbaArray(new_arr));
            }
            Stmt::Erase { name } => {
                if let Some(Variant::VbaArray(arr)) = self.variables.get_mut(name) {
                    for v in arr.elements.iter_mut() {
                        *v = Variant::Empty;
                    }
                }
            }
            Stmt::ArrayWrite {
                name,
                indices,
                value,
            } => {
                let v = self.eval_expr(value)?;
                let idx = self.eval_array_indices(indices)?;
                let bounds = match self.variables.get(name) {
                    Some(Variant::VbaArray(arr)) => arr.bounds.clone(),
                    _ => return Err(format!("'{}' is not an array", name)),
                };
                match VbaArray::linear_index_for(&bounds, &idx) {
                    Ok(i) => {
                        if let Some(Variant::VbaArray(arr)) = self.variables.get_mut(name) {
                            arr.elements[i] = v;
                        }
                    }
                    Err(_) => return Err(self.vba_array_oob_error_for(name, &idx, &bounds)),
                }
            }
            Stmt::With { target, body } => {
                // Resolve the target ONCE, here, on block entry — not per
                // `.member` access and not at parse time.
                let value = self.eval_with_target(target)?;
                self.run_with_body(value, body)?;
            }
            Stmt::WithDot { member, value } => {
                let v = self.eval_expr(value)?;
                self.write_with_member(member, v)?;
            }
            Stmt::MsgBox { message } => {
                let msg = self.eval_expr(message)?;
                // Record before checking error_on_msgbox: `messages` should
                // reflect every MsgBox the macro attempted to show, even
                // ones that are then treated as a blocking error.
                self.msgbox_log.push(msg.to_string());
                if self.error_on_msgbox {
                    return Err(format!("MsgBox: {}", msg));
                }
                if self.print_msgbox {
                    println!("{}", msg);
                }
            }
            Stmt::DimRecord { var, type_name } => {
                if let Some(fields) = self.type_defs.get(type_name).cloned() {
                    let record = make_record_default(&fields, &self.type_defs);
                    self.variables.insert(var.clone(), record);
                } else if matches!(type_name.as_str(), "range" | "worksheet" | "workbook") {
                    // `Dim r As Range` declares an *object* variable holding
                    // no reference yet — real VBA's `r Is Nothing` is True
                    // here, and any member access through it raises error 91
                    // until a `Set` gives it one. Registering it (rather than
                    // the previous silent no-op) is what makes both of those
                    // observable. Guarded on there being no user-defined
                    // `Type` of the same name, so a `Type Range ... End Type`
                    // module still wins — VBA's own name resolution prefers
                    // the user type too.
                    self.object_variables
                        .entry(var.clone())
                        .or_insert(ObjectRef::Nothing);
                }
                // Any other unknown type name → no-op (built-in type)
            }
            Stmt::DimArrayRecord {
                name,
                sizes,
                type_name,
            } => {
                let upper = to_f64(&self.eval_expr(&sizes[0])?)? as usize;
                let element = if let Some(fields) = self.type_defs.get(type_name).cloned() {
                    make_record_default(&fields, &self.type_defs)
                } else {
                    Variant::Empty
                };
                self.variables
                    .insert(name.clone(), Variant::Array(vec![element; upper + 1]));
            }
            Stmt::RecordSetNested { var, fields, value } => {
                let v = self.eval_expr(value)?;
                let target = self
                    .variables
                    .entry(var.clone())
                    .or_insert_with(|| Variant::Record(HashMap::new()));
                nested_set(target, fields, v);
            }
            Stmt::ArrayRecordSet {
                name,
                indices,
                field,
                value,
            } => {
                let v = self.eval_expr(value)?;
                let idx = to_f64(&self.eval_expr(&indices[0])?)? as usize;
                let oob_len = match self.variables.get(name) {
                    Some(Variant::Array(arr)) if idx >= arr.len() => Some(arr.len()),
                    Some(Variant::Array(_)) => None,
                    _ => return Err(format!("'{}' is not an array", name)),
                };
                if let Some(len) = oob_len {
                    // DimArrayRecord/ArrayRecordSet don't track a lower
                    // bound (see `eval_array_dim0`'s doc) — always 0.
                    return Err(self.array_oob_error(name, idx as i64, 0, len));
                }
                if let Some(Variant::Array(arr)) = self.variables.get_mut(name) {
                    match &mut arr[idx] {
                        Variant::Record(m) => {
                            m.insert(field.clone(), v);
                        }
                        slot => {
                            let mut m = HashMap::new();
                            m.insert(field.clone(), v);
                            *slot = Variant::Record(m);
                        }
                    }
                }
            }
            Stmt::RecordSet { var, field, value } => {
                // `<var>.Value = ...` / `<var>.Formula = ...` where `var`
                // is a `Set`-assigned object variable (Milestone B7c) —
                // shares this statement's `<var>.<field> = <value>`
                // grammar with plain record-field assignment (no separate
                // AST node needed), disambiguated here at runtime by which
                // namespace actually holds `var`. `object_variables` is
                // only ever populated by `Set`, so this can't misfire
                // against a genuine `Type`-based record that happens to
                // have a field literally named "value"/"formula".
                self.require_live_object(var)?;
                if let Some(ObjectRef::Range(r)) = self.object_variables.get(var).cloned() {
                    if field == "value" || field == "formula" {
                        let v = self.eval_expr(value)?;
                        self.write_range_ref_value(&r, field == "formula", &v)?;
                    }
                    // Any other property write on an object variable is a
                    // harmless no-op, same leniency as the generic
                    // `.Method`-without-assignment fallback in
                    // `parse_ident_stmt`.
                    return Ok(());
                }
                let v = self.eval_expr(value)?;
                let entry = self
                    .variables
                    .entry(var.clone())
                    .or_insert(Variant::Record(std::collections::HashMap::new()));
                if let Variant::Record(m) = entry {
                    m.insert(field.clone(), v);
                } else {
                    self.variables.insert(
                        var.clone(),
                        Variant::Record({
                            let mut m = std::collections::HashMap::new();
                            m.insert(field.clone(), v);
                            m
                        }),
                    );
                }
            }
        }
        Ok(())
    }

    pub fn eval_expr(&mut self, expr: &Expr) -> Result<Variant, String> {
        match expr {
            Expr::Integer(n) => Ok(Variant::Integer(*n)),
            Expr::Float(f) => Ok(Variant::Float(*f)),
            Expr::Str(s) => Ok(Variant::Str(s.clone())),
            Expr::Bool(b) => Ok(Variant::Boolean(*b)),
            Expr::ErrNumber => Ok(Variant::Integer(self.err_number)),
            Expr::ErrDescription => Ok(Variant::Str(self.err_description.clone())),
            Expr::ErrSource => Ok(Variant::Str(self.err_source.clone())),
            Expr::ErrHelpFile => Ok(Variant::Str(self.err_help_file.clone())),
            Expr::ErrHelpContext => Ok(Variant::Integer(self.err_help_context)),
            Expr::Var(name) => {
                if let Some(v) = self.variables.get(name) {
                    return Ok(v.clone());
                }
                // Excel built-in constants
                Ok(match name.as_str() {
                    // Calculation mode
                    "xlcalculationmanual" => Variant::Integer(-4135),
                    "xlcalculationautomatic" => Variant::Integer(-4105),
                    "xlcalculationsemiautomatic" => Variant::Integer(2),
                    // Direction
                    "xlup" => Variant::Integer(-4162),
                    "xldown" => Variant::Integer(-4121),
                    "xltoleft" => Variant::Integer(-4159),
                    "xltoright" => Variant::Integer(-4161),
                    // Cursor
                    "xlwait" => Variant::Integer(2),
                    "xldefault" => Variant::Integer(1),
                    "xlibeam" => Variant::Integer(3),
                    "xlnorthwestarrow" => Variant::Integer(4),
                    // VB string constants
                    "vbcrlf" => Variant::Str("\r\n".into()),
                    "vblf" => Variant::Str("\n".into()),
                    "vbcr" => Variant::Str("\r".into()),
                    "vbtab" => Variant::Str("\t".into()),
                    "vbnullstring" => Variant::Str(String::new()),
                    "vbnullchar" => Variant::Str("\0".into()),
                    // VB boolean constants (in addition to True/False literals)
                    "vbtrue" => Variant::Boolean(true),
                    "vbfalse" => Variant::Boolean(false),
                    // VB MsgBox return values
                    "vbok" => Variant::Integer(1),
                    "vbcancel" => Variant::Integer(2),
                    "vbyes" => Variant::Integer(6),
                    "vbno" => Variant::Integer(7),
                    // `Null` is its own value, not Empty: "no valid data"
                    // versus "uninitialized". Folding the two together (as
                    // this line used to) made every documented Null rule
                    // unobservable. `Nothing` stays Empty — it's the null
                    // *object* reference, and object state lives in
                    // `object_variables`, not in a `Variant` (see
                    // `Expr::IsNothing`).
                    "null" => return Ok(Variant::Null),
                    "empty" | "nothing" => return Ok(Variant::Empty),
                    // Real VBA allows omitting `()` on these three zero-arg
                    // functions (`Date`, not just `Date()`) — every other
                    // `eval_vba_func` entry needs at least one argument, so
                    // this doesn't generalize to "any unrecognized bare
                    // identifier might be a function call" (that would risk
                    // masking a genuine variable-name typo as a function
                    // call instead of the clearer "Undefined variable").
                    "date" | "now" | "time" => return self.eval_vba_func(name, &[]),
                    _ => return Err(format!("Undefined variable: '{}'", name)),
                })
            }
            Expr::UnaryMinus(inner) => match self.eval_expr(inner)? {
                Variant::Integer(n) => Ok(Variant::Integer(-n)),
                Variant::Float(f) => Ok(Variant::Float(-f)),
                // Same documented rule as binary `-`: "If one or both
                // expressions are Null expressions, result is Null."
                Variant::Null => Ok(Variant::Null),
                other => Err(format!("Unary minus on non-numeric: {}", other)),
            },
            Expr::UnaryNot(inner) => {
                // Mirrors And/Or/Xor's own logical-vs-bitwise split (see
                // `eval_binop`): a genuine Boolean gets logical negation: a
                // numeric operand gets a real bitwise complement (`Not 5` is
                // `-6`, not `False`) — VBA's own distinction, not a
                // truthy/falsy coercion.
                let v = self.eval_expr(inner)?;
                match v {
                    Variant::Boolean(b) => Ok(Variant::Boolean(!b)),
                    // Documented on the Not operator page: `Not Null` is
                    // Null — the third row of its own truth table.
                    Variant::Null => Ok(Variant::Null),
                    other => Ok(Variant::Integer(!to_i64_bitwise(&other)?)),
                }
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
                    let arg_vals: Vec<Variant> = args
                        .iter()
                        .map(|a| self.eval_expr(a))
                        .collect::<Result<_, _>>()?;
                    return self.call_func_def(&func, &arg_vals);
                }
                // Array subscript access on a plain (Range-value-read /
                // formula-array-result / record-array) `Variant::Array` —
                // always 0-based, unchanged from before this round (a real
                // `Dim`-declared array is `Variant::VbaArray`, below).
                if matches!(self.variables.get(name.as_str()), Some(Variant::Array(_))) {
                    let vba_idx = to_f64(
                        &self.eval_expr(
                            args.first()
                                .ok_or_else(|| format!("Array '{}' requires index", name))?,
                        )?,
                    )? as i64;
                    let lower = 0;
                    let internal = vba_idx - lower;
                    let (found, len) = match self.variables.get(name.as_str()) {
                        Some(Variant::Array(arr)) => {
                            let v = if internal >= 0 {
                                arr.get(internal as usize).cloned()
                            } else {
                                None
                            };
                            (v, arr.len())
                        }
                        _ => return Err(format!("'{}' is not an array", name)),
                    };
                    return match found {
                        Some(v) => Ok(v),
                        None => Err(self.array_oob_error(name, vba_idx, lower, len)),
                    };
                }
                // Array subscript access on a real (possibly multi-dim)
                // VBA-declared array: arr(i, j, ...).
                if matches!(
                    self.variables.get(name.as_str()),
                    Some(Variant::VbaArray(_))
                ) {
                    let idx = self.eval_array_indices(args)?;
                    let bounds = match self.variables.get(name.as_str()) {
                        Some(Variant::VbaArray(arr)) => arr.bounds.clone(),
                        _ => return Err(format!("'{}' is not an array", name)),
                    };
                    return match VbaArray::linear_index_for(&bounds, &idx) {
                        Ok(i) => match self.variables.get(name.as_str()) {
                            Some(Variant::VbaArray(arr)) => Ok(arr.elements[i].clone()),
                            _ => Err(format!("'{}' is not an array", name)),
                        },
                        Err(_) => Err(self.vba_array_oob_error_for(name, &idx, &bounds)),
                    };
                }
                self.eval_vba_func(name, args)
            }
            Expr::RangeRead { addr } => {
                let ((r1, c1), (r2, c2)) = self
                    .resolve_range_addr(addr)
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
            Expr::RangeOffsetRead {
                addr,
                row_off,
                col_off,
            } => {
                let (base_r, base_c) = parse_cell_addr(addr)
                    .ok_or_else(|| format!("RangeOffsetRead: invalid address '{}'", addr))?;
                let ro = to_f64(&self.eval_expr(row_off)?)? as i64;
                let co = to_f64(&self.eval_expr(col_off)?)? as i64;
                Ok(self.get_cell((base_r as i64 + ro) as u32, (base_c as i64 + co) as u32))
            }
            Expr::SheetCellRead { sheet, row, col } => {
                let (key, display) = self.resolve_sheet_expr(sheet)?;
                self.check_strict_sheet_exists(&display, &key)?;
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                Ok(self
                    .sheets
                    .get(&key)
                    .and_then(|s| s.get(&(r, c)))
                    .map(|cell| cell.value.clone())
                    .unwrap_or(Variant::Empty))
            }
            Expr::SheetRangeRead { sheet, addr } => {
                let (key, display) = self.resolve_sheet_expr(sheet)?;
                self.check_strict_sheet_exists(&display, &key)?;
                let ((r1, c1), (r2, c2)) = parse_range_addr(addr)
                    .ok_or_else(|| format!("SheetRangeRead: invalid address '{}'", addr))?;
                let cells = self.sheets.get(&key);
                let get = |r: u32, c: u32| {
                    cells
                        .and_then(|s| s.get(&(r, c)))
                        .map(|cell| cell.value.clone())
                        .unwrap_or(Variant::Empty)
                };
                if r1 == r2 && c1 == c2 {
                    Ok(get(r1, c1))
                } else {
                    let arr = (r1..=r2)
                        .flat_map(|r| (c1..=c2).map(move |c| (r, c)))
                        .map(|(r, c)| get(r, c))
                        .collect();
                    Ok(Variant::Array(arr))
                }
            }
            Expr::WorkbookQualifiedSheet { .. } => {
                // Only meaningful as the `sheet` field wrapped inside another
                // sheet-access node (see `resolve_sheet_expr`) — never
                // evaluated as a standalone expression by the parser.
                Err(
                    "Workbooks(...).Worksheets(...) is only valid as part of a Cells/Range access"
                        .to_string(),
                )
            }
            // Same "only meaningful wrapped inside a sheet-access node" story
            // as `WorkbookQualifiedSheet` above — `resolve_sheet_expr`
            // intercepts `ActiveSheetRef` before it ever reaches here, for
            // every use the parser actually produces (Milestone B7c item 6).
            Expr::ActiveSheetRef => Ok(Variant::Str(self.active_sheet.clone())),
            // Same "only meaningful wrapped inside a sheet-access node"
            // story as `ActiveSheetRef` above — `resolve_sheet_expr`
            // intercepts `ObjectVarSheet` before it ever reaches here, for
            // every use the parser actually produces (Phase 2C items 7/8).
            Expr::ObjectVarSheet(name) => match self.object_variables.get(name).cloned() {
                Some(ObjectRef::Worksheet(key)) => Ok(Variant::Str(key)),
                _ => Err(format!("'{}' is not a Worksheet object variable", name)),
            },
            Expr::CellsFind { what, find_row } => {
                let target = self.eval_expr(what)?;
                let mut keys: Vec<(u32, u32)> = self.cells().keys().cloned().collect();
                keys.sort(); // 行優先スキャン
                for (r, c) in keys {
                    if vba_eq(&self.get_cell(r, c), &target) {
                        return Ok(Variant::Integer(if *find_row {
                            r as i64
                        } else {
                            c as i64
                        }));
                    }
                }
                Ok(Variant::Integer(0)) // not found
            }
            Expr::RowsCount => Ok(Variant::Integer(1_048_576)),
            Expr::ColsCount => Ok(Variant::Integer(16_384)),
            Expr::CellsEndProp {
                row,
                col,
                dir,
                prop,
            } => {
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                let result = match (dir, prop) {
                    (XlDir::Up, XlEndProp::Row) => self.last_nonempty_row(c, r),
                    (XlDir::Down, XlEndProp::Row) => self.first_empty_row(c, r).saturating_sub(1),
                    (XlDir::Left, XlEndProp::Column) => self.last_nonempty_col(r, c),
                    (XlDir::Right, XlEndProp::Column) => {
                        self.first_empty_col(r, c).saturating_sub(1)
                    }
                    (XlDir::Up, XlEndProp::Column) | (XlDir::Down, XlEndProp::Column) => c,
                    (XlDir::Left, XlEndProp::Row) | (XlDir::Right, XlEndProp::Row) => r,
                };
                Ok(Variant::Integer(result as i64))
            }
            // A bare `.member` read — e.g. the right-hand side of
            // `.Value = .Value + 1`. Resolved against the innermost active
            // With target, at whatever depth in the AST it appears.
            Expr::WithDot(fields) => match self.current_with()? {
                WithValue::Range(r) => {
                    if fields.first().map(String::as_str) == Some("value") {
                        self.read_range_ref_value(&r)
                    } else {
                        Ok(Variant::Empty)
                    }
                }
                WithValue::Record(var) => {
                    let mut cur = self.variables.get(&var).cloned().unwrap_or(Variant::Empty);
                    for f in fields {
                        cur = match cur {
                            Variant::Record(m) => m.get(f).cloned().unwrap_or(Variant::Empty),
                            _ => Variant::Empty,
                        };
                    }
                    Ok(cur)
                }
                // Worksheet/unmodeled property reads aren't modeled — the
                // same `Empty` an unmodeled `<var>.<field>` read already
                // gives, rather than a new error condition.
                WithValue::Sheet(_) | WithValue::Unmodeled => Ok(Variant::Empty),
            },
            Expr::IsNothing(name) => {
                // True for a declared-but-unset object variable and for one
                // explicitly `Set` to Nothing — real VBA can't tell those
                // apart either. A name absent from `object_variables`
                // entirely is reported as Nothing too: in real VBA a
                // non-object operand here is a compile-time type error, and
                // "no live object reference" is the closest true answer
                // elixcee can give without inventing a type error the rest
                // of the VM has no way to raise.
                Ok(Variant::Boolean(!matches!(
                    self.object_variables.get(name),
                    Some(ObjectRef::Range(_))
                        | Some(ObjectRef::Worksheet(_))
                        | Some(ObjectRef::Workbook)
                )))
            }
            Expr::RecordGet { var, field } => {
                // `x = <var>.Value` where `var` is a `Set`-assigned object
                // variable (Milestone B7c) — see `Stmt::RecordSet`'s
                // matching write-side comment for why this is safe to
                // disambiguate purely by which namespace holds `var`.
                self.require_live_object(var)?;
                if let Some(ObjectRef::Range(r)) = self.object_variables.get(var).cloned() {
                    return if field == "value" {
                        self.read_range_ref_value(&r)
                    } else {
                        Ok(Variant::Empty)
                    };
                }
                match self.variables.get(var) {
                    Some(Variant::Record(m)) => Ok(m.get(field).cloned().unwrap_or(Variant::Empty)),
                    _ => Ok(Variant::Empty),
                }
            }
            Expr::RecordGetNested { var, fields } => {
                // `x = <var>.Areas.Count` (Milestone B7c item 3) — same
                // object-variable disambiguation as above.
                self.require_live_object(var)?;
                if let Some(ObjectRef::Range(r)) = self.object_variables.get(var).cloned() {
                    return if fields.len() == 2 && fields[0] == "areas" && fields[1] == "count" {
                        Ok(Variant::Integer(r.areas.len() as i64))
                    } else {
                        Ok(Variant::Empty)
                    };
                }
                let mut cur = self.variables.get(var).cloned().unwrap_or(Variant::Empty);
                for f in fields {
                    cur = match cur {
                        Variant::Record(m) => m.get(f).cloned().unwrap_or(Variant::Empty),
                        _ => Variant::Empty,
                    };
                }
                Ok(cur)
            }
            Expr::ArrayRecordGet {
                name,
                indices,
                field,
            } => {
                let idx = to_f64(&self.eval_expr(&indices[0])?)? as usize;
                let (found, len) = match self.variables.get(name) {
                    Some(Variant::Array(arr)) => (arr.get(idx).cloned(), arr.len()),
                    _ => return Err(format!("'{}' is not an array", name)),
                };
                match found {
                    Some(Variant::Record(m)) => Ok(m.get(field).cloned().unwrap_or(Variant::Empty)),
                    Some(other) => Ok(other),
                    None => Err(self.array_oob_error(name, idx as i64, 0, len)),
                }
            }
        }
    }

    fn eval_vba_func(&mut self, name: &str, args: &[Expr]) -> Result<Variant, String> {
        let vals: Vec<Variant> = args
            .iter()
            .map(|a| self.eval_expr(a))
            .collect::<Result<_, _>>()?;
        match name {
            "int" => {
                let f = to_f64(vals.first().ok_or("INT requires 1 argument")?)?;
                Ok(as_int_if_whole(f.floor()))
            }
            "clng" | "cint" => {
                // Real VBA's CInt/CLng use banker's rounding (round-half-
                // to-even), same as Round() — `to_i64_rounded`'s own doc
                // comment already claims this ("the same round-half-to-
                // even ... that CLng/Round use"), but this arm used Rust's
                // default round-half-away-from-zero until now: `CInt(0.5)`
                // was `1`, not real VBA's `0`.
                let v = vals.first().ok_or("CInt/CLng requires 1 argument")?;
                Ok(Variant::Integer(to_i64_rounded(v)?))
            }
            "cbool" => {
                let v = vals.first().ok_or("CBool requires 1 argument")?;
                let b = match v {
                    Variant::Boolean(b) => *b,
                    // A string is only ever a literal "True"/"False" or a
                    // numeric-string in real VBA — never routed through the
                    // shared CLng/CInt numeric-coercion path (that's the
                    // bug this arm used to have: CBool("True") tried to
                    // parse "True" as a number and errored).
                    Variant::Str(s) => match s.trim().to_lowercase().as_str() {
                        "true" => true,
                        "false" => false,
                        _ => to_f64(v)? != 0.0,
                    },
                    other => to_f64(other)? != 0.0,
                };
                Ok(Variant::Boolean(b))
            }
            "fix" => {
                // Truncates toward zero — unlike Int(), which floors toward
                // negative infinity. `Fix(-3.9)` is `-3`, not `-4`.
                let f = to_f64(vals.first().ok_or("Fix requires 1 argument")?)?;
                Ok(as_int_if_whole(f.trunc()))
            }
            "sgn" => {
                let f = to_f64(vals.first().ok_or("Sgn requires 1 argument")?)?;
                let n: i64 = if f > 0.0 {
                    1
                } else if f < 0.0 {
                    -1
                } else {
                    0
                };
                Ok(Variant::Integer(n))
            }
            "round" => {
                // VBA's own Round() uses banker's rounding (round-half-to-
                // even) — the same convention `to_i64_rounded` documents
                // for `\`/`Mod` operand coercion — which is NOT what
                // WorksheetFunction.Round (Excel's ROUND() formula, round-
                // half-away-from-zero) does. They're genuinely different
                // functions in real VBA/Excel, not aliases of each other,
                // so this doesn't share `eval_wsf`'s "round" arm.
                let f = to_f64(vals.first().ok_or("Round requires 1 argument")?)?;
                let digits = if vals.len() >= 2 {
                    to_f64(&vals[1])? as i32
                } else {
                    0
                };
                // Unlike WorksheetFunction.Round/Excel's ROUND(), which both
                // accept a negative NumDigitsAfterDecimal to round left of
                // the decimal point, real VBA's Round() raises "Invalid
                // procedure call or argument" for a negative digit count.
                if digits < 0 {
                    return Err("Invalid procedure call or argument".into());
                }
                let factor = 10f64.powi(digits);
                Ok(as_int_if_whole((f * factor).round_ties_even() / factor))
            }
            "cdbl" | "csng" => {
                let f = to_f64(vals.first().ok_or("CDbl requires 1 argument")?)?;
                Ok(Variant::Float(f))
            }
            "cstr" => {
                let s = vals.first().ok_or("CStr requires 1 argument")?.to_string();
                Ok(Variant::Str(s))
            }
            "str" => {
                // Unlike CStr, real VBA's Str() reserves a leading space
                // for the sign position on a non-negative number
                // (`Str(459)` is `" 459"`, not `"459"`) — a well-known VBA
                // quirk, and a real behavior difference from CStr, not an
                // alias of it. Scoped to numeric inputs, the only case
                // Str() is documented for; anything else falls back to the
                // same plain-Display formatting CStr uses.
                let v = vals.first().ok_or("Str requires 1 argument")?;
                let non_negative_number = matches!(v, Variant::Integer(n) if *n >= 0)
                    || matches!(v, Variant::Float(f) if *f >= 0.0);
                let s = if non_negative_number {
                    format!(" {}", v)
                } else {
                    v.to_string()
                };
                Ok(Variant::Str(s))
            }
            "val" => {
                // Real VBA's Val() parses a leading numeric prefix and
                // stops at the first character that doesn't fit, returning
                // 0 only if there's no valid numeric prefix at all —
                // Val("123abc") is 123, not 0. A strict whole-string parse
                // (what this used to do) makes any trailing non-numeric
                // character silently zero out the entire value.
                let s = match vals.first().ok_or("Val requires 1 argument")? {
                    Variant::Str(s) => s.clone(),
                    v => v.to_string(),
                };
                Ok(as_int_if_whole(parse_vba_val_prefix(&s)))
            }
            "len" => {
                let s = match vals.first().ok_or("Len requires 1 argument")? {
                    Variant::Str(s) => s.chars().count() as i64,
                    Variant::Empty => 0,
                    v => v.to_string().chars().count() as i64,
                };
                Ok(Variant::Integer(s))
            }
            "left" => {
                let s = vba_to_str(vals.first().ok_or("Left requires 2 arguments")?);
                let n = to_f64(vals.get(1).ok_or("Left requires 2 arguments")?)? as usize;
                Ok(Variant::Str(s.chars().take(n).collect()))
            }
            "right" => {
                let s = vba_to_str(vals.first().ok_or("Right requires 2 arguments")?);
                let n = to_f64(vals.get(1).ok_or("Right requires 2 arguments")?)? as usize;
                let chars: Vec<char> = s.chars().collect();
                Ok(Variant::Str(
                    chars[chars.len().saturating_sub(n)..].iter().collect(),
                ))
            }
            "mid" => {
                if vals.len() < 2 {
                    return Err("Mid requires at least 2 arguments".into());
                }
                let s = vba_to_str(&vals[0]);
                let start = (to_f64(&vals[1])? as usize).saturating_sub(1);
                let len = if vals.len() >= 3 {
                    to_f64(&vals[2])? as usize
                } else {
                    usize::MAX
                };
                Ok(Variant::Str(s.chars().skip(start).take(len).collect()))
            }
            "ucase" => Ok(Variant::Str(
                vba_to_str(vals.first().ok_or("UCase requires 1 argument")?).to_uppercase(),
            )),
            "lcase" => Ok(Variant::Str(
                vba_to_str(vals.first().ok_or("LCase requires 1 argument")?).to_lowercase(),
            )),
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
            // Genuinely different questions, so no longer one shared arm:
            // IsNull asks "is this the Null value", IsEmpty asks "is this an
            // uninitialized Variant". IsNull(Empty) and IsEmpty(Null) are
            // both False in real VBA.
            "isnull" => Ok(Variant::Boolean(matches!(
                vals.first(),
                Some(Variant::Null)
            ))),
            "isempty" => Ok(Variant::Boolean(matches!(
                vals.first(),
                Some(Variant::Empty) | None
            ))),
            "isnumeric" => {
                // Real VBA's IsNumeric also accepts a string that parses as
                // a number (`IsNumeric("123")` is True) and Empty (an
                // uninitialized variable coerces to 0 in a numeric
                // context) — not just an already-numeric Variant. Scoped
                // to plain decimal/scientific-notation strings (Rust's own
                // f64 parser, after trimming whitespace); real VBA's fuller
                // numeric-string grammar (currency symbols, locale-specific
                // decimal separators, parenthesized negatives, ...) isn't
                // attempted here — no evidence (corpus or otherwise) it's
                // needed, and guessing at locale-specific parsing rules
                // isn't this project's style.
                let is_numeric = match vals.first() {
                    Some(Variant::Integer(_)) | Some(Variant::Float(_)) | Some(Variant::Empty) => {
                        true
                    }
                    Some(Variant::Str(s)) => s.trim().parse::<f64>().is_ok(),
                    _ => false,
                };
                Ok(Variant::Boolean(is_numeric))
            }
            "chr" => {
                let n = to_f64(vals.first().ok_or("Chr requires 1 argument")?)? as u32;
                char::from_u32(n)
                    .map(|c| Variant::Str(c.to_string()))
                    .ok_or_else(|| format!("Chr: invalid code {}", n))
            }
            "asc" => {
                let s = vba_to_str(vals.first().ok_or("Asc requires 1 argument")?);
                s.chars()
                    .next()
                    .map(|c| Variant::Integer(c as i64))
                    .ok_or_else(|| "Asc: empty string".into())
            }
            "instr" => {
                // InStr([start,] string1, string2 [, compare])
                let (start, s1, s2) = if vals.len() >= 3 {
                    (
                        to_f64(&vals[0])? as usize,
                        vba_to_str(&vals[1]),
                        vba_to_str(&vals[2]),
                    )
                } else {
                    (
                        1,
                        vba_to_str(vals.first().ok_or("InStr requires at least 2 arguments")?),
                        vba_to_str(vals.get(1).ok_or("InStr requires at least 2 arguments")?),
                    )
                };
                let h: Vec<char> = s1.chars().collect();
                let n: Vec<char> = s2.chars().collect();
                // VBA: empty needle → return start position; start > len → return 0
                if n.is_empty() {
                    return Ok(Variant::Integer(start as i64));
                }
                let begin = start.saturating_sub(1);
                if begin >= h.len() {
                    return Ok(Variant::Integer(0));
                }
                let pos = h[begin..]
                    .windows(n.len())
                    .position(|w| {
                        w.iter()
                            .map(|c| c.to_uppercase().next().unwrap_or(*c))
                            .eq(n.iter().map(|c| c.to_uppercase().next().unwrap_or(*c)))
                    })
                    .map(|p| p + start)
                    .unwrap_or(0);
                Ok(Variant::Integer(pos as i64))
            }
            "replace" => {
                if vals.len() < 3 {
                    return Err("Replace requires at least 3 arguments".into());
                }
                let s = vba_to_str(&vals[0]);
                let old = vba_to_str(&vals[1]);
                let new = vba_to_str(&vals[2]);
                Ok(Variant::Str(if old.is_empty() {
                    s
                } else {
                    s.replace(&old as &str, &new as &str)
                }))
            }
            // Real VBA's `Now`/`Date`/`Time` all return a Date-typed value —
            // Excel's own epoch-serial number, split into a whole-day part
            // (the date) and a 0.0-1.0 fractional part (the time of day).
            // `date_to_serial`'s 25569 offset (Excel serial of 1970-01-01)
            // matches the same constant the formula engine's own NOW()
            // (`formula::eval::func_now`) already uses — kept independent
            // rather than shared to avoid a new formula<->vm cross-module
            // dependency for one constant.
            "date" => {
                let unix_days = unix_epoch_days();
                Ok(Variant::Date(unix_days as i64 + 25569))
            }
            "time" => {
                // Time-of-day only, as a fraction — `Variant::Date` is a
                // whole-day-only `i64` in this codebase (see its doc in
                // `elixcee-types`), so a sub-day value can't round-trip
                // through it; `Variant::Float` at least carries the
                // numerically correct value, same as real VBA's own
                // internal Double representation for a time-only value.
                // `TypeName(Time)` will report "Double" here, not real
                // VBA's "Date" — a known, disclosed gap (see ROADMAP.md),
                // not something silently wrong.
                Ok(Variant::Float(unix_seconds_of_day() as f64 / 86400.0))
            }
            "now" => {
                let unix_days = unix_epoch_days();
                let frac = unix_seconds_of_day() as f64 / 86400.0;
                // Same TypeName caveat as "time" above.
                Ok(Variant::Float(unix_days as f64 + 25569.0 + frac))
            }
            // ── Inline conditional ───────────────────────────────────────────
            "iif" => {
                if vals.len() < 3 {
                    return Err("IIf requires 3 arguments".into());
                }
                Ok(if is_truthy(&vals[0]) {
                    vals[1].clone()
                } else {
                    vals[2].clone()
                })
            }
            // ── Format ───────────────────────────────────────────────────────
            "format" => {
                if vals.is_empty() {
                    return Err("Format requires at least 1 argument".into());
                }
                let v = &vals[0];
                let fmt = if vals.len() >= 2 {
                    vba_to_str(&vals[1])
                } else {
                    String::new()
                };
                Ok(Variant::Str(format_vba(v, &fmt)))
            }
            // ── Type inspection ──────────────────────────────────────────────
            "typename" => {
                let name = match vals.first().ok_or("TypeName requires 1 argument")? {
                    Variant::Integer(_) => "Long",
                    Variant::Float(_) => "Double",
                    Variant::Str(_) => "String",
                    Variant::Boolean(_) => "Boolean",
                    Variant::Date(_) => "Date",
                    Variant::Error(_) => "Error",
                    Variant::Array(_) | Variant::VbaArray(_) => "Variant()",
                    Variant::Empty => "Empty",
                    Variant::Null => "Null",
                    Variant::Record(_) => "Object",
                };
                Ok(Variant::Str(name.into()))
            }
            "vartype" => {
                let n: i64 = match vals.first().ok_or("VarType requires 1 argument")? {
                    Variant::Empty => 0,
                    Variant::Null => 1,                               // vbNull
                    Variant::Integer(_) => 3,                         // vbLong
                    Variant::Float(_) => 5,                           // vbDouble
                    Variant::Str(_) => 8,                             // vbString
                    Variant::Boolean(_) => 11,                        // vbBoolean
                    Variant::Date(_) => 7,                            // vbDate
                    Variant::Array(_) | Variant::VbaArray(_) => 8204, // vbArray + vbVariant
                    Variant::Error(_) => 10,                          // vbError
                    Variant::Record(_) => 0,                          // vbEmpty as fallback
                };
                Ok(Variant::Integer(n))
            }
            // ── Array functions ──────────────────────────────────────────────
            // `Array(a, b, c)` builds a zero-based, rank-1 VbaArray from its
            // arguments. `Array()` with no arguments is a legal empty array.
            "array" => Ok(Variant::VbaArray(VbaArray::from_vec(vals.to_vec()))),
            "split" => {
                if vals.is_empty() {
                    return Err("Split requires at least 1 argument".into());
                }
                let s = vba_to_str(&vals[0]);
                let delim = if vals.len() >= 2 {
                    vba_to_str(&vals[1])
                } else {
                    " ".to_string()
                };
                let parts = s
                    .split(delim.as_str())
                    .map(|p| Variant::Str(p.to_string()))
                    .collect();
                Ok(Variant::VbaArray(VbaArray::from_vec(parts)))
            }
            "join" => {
                if vals.is_empty() {
                    return Err("Join requires at least 1 argument".into());
                }
                let parts = match &vals[0] {
                    Variant::Array(a) => a.iter().map(vba_to_str).collect::<Vec<_>>(),
                    Variant::VbaArray(a) => a.elements.iter().map(vba_to_str).collect::<Vec<_>>(),
                    v => vec![vba_to_str(v)],
                };
                let delim = if vals.len() >= 2 {
                    vba_to_str(&vals[1])
                } else {
                    " ".to_string()
                };
                Ok(Variant::Str(parts.join(&delim)))
            }
            "ubound" => {
                match vals.first().ok_or("UBound requires 1 argument")? {
                    // Not a VBA `Dim`-declared array (a Range-value read, a
                    // formula-array result, a record array, …) — always
                    // 0-based and always answers for its one and only
                    // dimension, ignoring the optional dimension argument,
                    // unchanged from before this round (this type has no
                    // per-dimension bounds to consult in the first place).
                    Variant::Array(a) => Ok(Variant::Integer(a.len() as i64 - 1)),
                    Variant::VbaArray(a) => {
                        let dim = self.array_func_dimension(args.get(1))?;
                        a.ubound(dim).map(Variant::Integer)
                    }
                    _ => Err("UBound: argument is not an array".into()),
                }
            }
            "lbound" => match vals.first().ok_or("LBound requires 1 argument")? {
                Variant::Array(_) => Ok(Variant::Integer(0)),
                Variant::VbaArray(a) => {
                    let dim = self.array_func_dimension(args.get(1))?;
                    a.lbound(dim).map(Variant::Integer)
                }
                _ => Err("LBound: argument is not an array".into()),
            },
            "isarray" => Ok(Variant::Boolean(matches!(
                vals.first(),
                Some(Variant::Array(_) | Variant::VbaArray(_))
            ))),
            // ── Range object (used as WSF arg) ───────────────────────────────
            "range" => {
                if let Some(Variant::Str(addr)) = vals.first() {
                    let ((r1, c1), (r2, c2)) = self
                        .resolve_range_addr(addr)
                        .ok_or_else(|| format!("Range: invalid address '{}'", addr))?;
                    let arr = (r1..=r2)
                        .flat_map(|r| (c1..=c2).map(move |c| (r, c)))
                        .map(|(r, c)| self.get_cell(r, c))
                        .collect();
                    Ok(Variant::Array(arr))
                } else {
                    Err("Range: requires a string address argument".into())
                }
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
        self.cells()
            .get(&(row, col))
            .map(|c| c.value.clone())
            .unwrap_or(Variant::Empty)
    }

    /// The active sheet's resolved number-format code for a cell (GitHub #4), e.g.
    /// `"m/d/yyyy"` for a date-formatted cell -- `None` for a cell with no format, the
    /// General format, or a sheet built purely in-VBA/loaded from `.ods`. Letting a
    /// caller do the serial-number-to-date conversion itself (rather than this VM
    /// guessing and changing `get_cell`'s return type) matches the reporter's own
    /// stated preference and avoids a breaking change to `get_cell`.
    pub fn get_cell_number_format(&self, row: u32, col: u32) -> Option<&str> {
        self.cell_number_formats
            .get(&self.active_sheet)?
            .get(&(row, col))
            .map(String::as_str)
    }

    pub fn set_cell_formula(&mut self, row: u32, col: u32, formula: &str) -> Result<(), String> {
        let expr = formula::parse(formula)?;
        let value = formula::evaluate(&expr, self.cells())?;
        self.cells_mut().insert(
            (row, col),
            CellContent {
                formula: Some(formula.to_string()),
                value,
            },
        );
        Ok(())
    }

    pub fn recalculate_all(&mut self) -> Result<(), String> {
        // Collect all formula cells and parse them
        let formula_cells: Vec<(u32, u32, formula::FormulaExpr)> = {
            self.cells()
                .iter()
                .filter_map(|((r, c), cell)| {
                    cell.formula
                        .as_ref()
                        .and_then(|f| formula::parse(f).ok().map(|e| (*r, *c, e)))
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
            if let Some(cell) = self
                .sheets
                .get_mut(&active)
                .and_then(|m| m.get_mut(&(row, col)))
            {
                cell.value = value;
            }
        }
        // Mark index dirty once (formula values changed, End queries may be stale)
        if !formula_cells.is_empty() {
            self.cell_index_dirty = true;
        }
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
        if self.cell_index_dirty {
            self.rebuild_cell_index();
        }
        self.col_rows
            .get(&col)
            .and_then(|rows| rows.range(..=max_row).next_back().copied())
            .unwrap_or(1)
    }

    /// Find the first empty row in `col` at or below `start_row` (xlDown helper).
    pub fn first_empty_row(&mut self, col: u32, start_row: u32) -> u32 {
        if self.cell_index_dirty {
            self.rebuild_cell_index();
        }
        if let Some(rows) = self.col_rows.get(&col) {
            let mut prev = start_row.saturating_sub(1);
            for &r in rows.range(start_row..) {
                if r != prev + 1 {
                    return prev + 1;
                } // gap found
                prev = r;
            }
            prev + 1
        } else {
            start_row // column is entirely empty
        }
    }

    /// Find the last non-empty column in `row` at or left of `max_col` (xlToLeft).
    pub fn last_nonempty_col(&mut self, row: u32, max_col: u32) -> u32 {
        if self.cell_index_dirty {
            self.rebuild_cell_index();
        }
        self.row_cols
            .get(&row)
            .and_then(|cols| cols.range(..=max_col).next_back().copied())
            .unwrap_or(1)
    }

    /// Find the first empty column in `row` at or right of `start_col` (xlToRight helper).
    pub fn first_empty_col(&mut self, row: u32, start_col: u32) -> u32 {
        if self.cell_index_dirty {
            self.rebuild_cell_index();
        }
        if let Some(cols) = self.row_cols.get(&row) {
            let mut prev = start_col.saturating_sub(1);
            for &c in cols.range(start_col..) {
                if c != prev + 1 {
                    return prev + 1;
                }
                prev = c;
            }
            prev + 1
        } else {
            start_col
        }
    }
}

impl Default for Vm {
    fn default() -> Self {
        Self::new()
    }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

// ── Range address helpers ─────────────────────────────────────────────────────
// col_letters_to_num_vm/parse_cell_addr/parse_range_addr moved to
// elixcee-types (Phase 2A); re-exported near the top of this file.

/// Splits `addr` on top-level commas and parses each piece with
/// `parse_range_addr` (Milestone B7a) — `"A1:A3,C1:C3"` becomes 2 `Rect`s;
/// a plain `"A1:C10"` (no comma) still returns a 1-element `Vec` so callers
/// can treat single- and multi-area addresses uniformly. `parse_range_addr`
/// itself is untouched — every other caller keeps its current signature.
pub fn parse_multi_area_addr(addr: &str) -> Option<Vec<Rect>> {
    addr.split(',')
        .map(|piece| {
            let ((start_row, start_col), (end_row, end_col)) = parse_range_addr(piece)?;
            Some(Rect {
                start_row,
                start_col,
                end_row,
                end_col,
            })
        })
        .collect()
}

/// Does rect `a` overlap rect `b` at all? (Milestone B6c2) — no existing
/// helper for this anywhere in the codebase before now.
fn rects_overlap(a: MergeRect, b: MergeRect) -> bool {
    let ((ar1, ac1), (ar2, ac2)) = a;
    let ((br1, bc1), (br2, bc2)) = b;
    ar1 <= br2 && br1 <= ar2 && ac1 <= bc2 && bc1 <= ac2
}

/// Does rect `outer` fully contain rect `inner`? (Milestone B6c2)
fn rect_contains(outer: MergeRect, inner: MergeRect) -> bool {
    let ((or1, oc1), (or2, oc2)) = outer;
    let ((ir1, ic1), (ir2, ic2)) = inner;
    or1 <= ir1 && ir2 <= or2 && oc1 <= ic1 && ic2 <= oc2
}

/// `(sheet_name_lowercase, (r1,c1), (r2,c2))`.
pub type SheetRange = (String, (u32, u32), (u32, u32));

/// Parses `"Sheet!A1:B10"` (or bare `"A1:B10"`, defaulting to `active_sheet`)
/// into a `SheetRange` — for CLI/fixture-facing range strings outside VBA
/// syntax (e.g. Milestone B5a's `test-workbook` TOML), not used anywhere
/// inside VBA statement execution itself.
pub fn parse_sheet_range_addr(s: &str, active_sheet: &str) -> Option<SheetRange> {
    let s = s.trim();
    let (sheet, range_part) = match s.find('!') {
        Some(i) => (s[..i].trim().to_lowercase(), &s[i + 1..]),
        None => (active_sheet.to_lowercase(), s),
    };
    let range = parse_range_addr(range_part)?;
    Some((sheet, range.0, range.1))
}

// ── B6a: resolution-failure evidence helpers ────────────────────────────────

/// Levenshtein edit distance, hand-rolled to avoid a new dependency for a
/// single "did you mean" suggestion (same zero-new-runtime-dependency
/// rationale as B5a's hand-rolled TOML parser). Operates on `char`s (not
/// bytes) so CJK names are compared correctly.
fn levenshtein(a: &str, b: &str) -> usize {
    let a: Vec<char> = a.chars().collect();
    let b: Vec<char> = b.chars().collect();
    let mut prev: Vec<usize> = (0..=b.len()).collect();
    let mut curr = vec![0usize; b.len() + 1];
    for (i, &ca) in a.iter().enumerate() {
        curr[0] = i + 1;
        for (j, &cb) in b.iter().enumerate() {
            let cost = if ca == cb { 0 } else { 1 };
            curr[j + 1] = (prev[j + 1] + 1).min(curr[j] + 1).min(prev[j] + cost);
        }
        std::mem::swap(&mut prev, &mut curr);
    }
    prev[b.len()]
}

/// The closest name to `requested` among `candidates` by edit distance —
/// only returned if the distance is small relative to the requested name's
/// length, so an unrelated name is never suggested (e.g. a 1-character typo
/// in a 4-character name is worth suggesting; a completely different name
/// of similar length is not).
fn closest_match(requested: &str, candidates: &[String]) -> Option<String> {
    let requested_lower = requested.to_lowercase();
    let bound = (requested_lower.chars().count() / 2).max(2);
    candidates
        .iter()
        .map(|c| (c, levenshtein(&requested_lower, &c.to_lowercase())))
        .filter(|(_, d)| *d <= bound)
        .min_by_key(|(_, d)| *d)
        .map(|(c, _)| c.clone())
}

// ── UDT helpers ──────────────────────────────────────────────────────────────

/// Build a `Variant::Record` with type-appropriate defaults, supporting nested UDTs.
fn make_record_default(
    fields: &[(String, String)],
    type_defs: &HashMap<String, Vec<(String, String)>>,
) -> Variant {
    let map: HashMap<String, Variant> = fields
        .iter()
        .map(|(name, vba_type)| {
            let default = match vba_type.as_str() {
                "integer" | "long" | "longlong" | "byte" => Variant::Integer(0),
                "single" | "double" | "currency" | "decimal" => Variant::Float(0.0),
                "boolean" => Variant::Boolean(false),
                "string" => Variant::Str(String::new()),
                other => {
                    if let Some(nested) = type_defs.get(other) {
                        make_record_default(nested, type_defs)
                    } else {
                        Variant::Empty
                    }
                }
            };
            (name.clone(), default)
        })
        .collect();
    Variant::Record(map)
}

/// Recursively set a value at the path given by `fields` inside a `Variant::Record` tree.
fn nested_set(target: &mut Variant, fields: &[String], value: Variant) {
    if fields.is_empty() {
        *target = value;
        return;
    }
    let field = &fields[0];
    let rest = &fields[1..];
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
        CellRef { col, row } => [(*row, *col)].into(),
        Range { c1, r1, c2, r2 } => {
            let mut s = HashSet::new();
            for r in *r1..=*r2 {
                for c in *c1..=*c2 {
                    s.insert((r, c));
                }
            }
            s
        }
        BinOp { lhs, rhs, .. } => {
            let mut s = extract_cell_refs(lhs);
            s.extend(extract_cell_refs(rhs));
            s
        }
        UnaryMinus(inner) => extract_cell_refs(inner),
        FuncCall { args, .. } => args.iter().flat_map(extract_cell_refs).collect(),
        Number(_) | Str(_) | Bool(_) => HashSet::new(),
    }
}

/// Topological sort of formula cells by dependency order.
/// Returns indices into `cells` in safe evaluation order.
/// Cells with no inter-formula dependencies appear first.
/// Returns `Err` if a circular reference is detected.
fn topo_sort_formulas(cells: &[(u32, u32, formula::FormulaExpr)]) -> Result<Vec<usize>, String> {
    let n = cells.len();
    // map (row, col) → index in cells slice
    let pos: HashMap<(u32, u32), usize> = cells
        .iter()
        .enumerate()
        .map(|(i, (r, c, _))| ((*r, *c), i))
        .collect();

    // in_degree[i] = number of formula cells that i depends on
    let mut in_degree = vec![0usize; n];
    // adj[j] = list of formula cells that depend on j
    let mut adj: Vec<Vec<usize>> = vec![vec![]; n];

    for (i, (_, _, expr)) in cells.iter().enumerate() {
        for dep in extract_cell_refs(expr) {
            if let Some(&j) = pos.get(&dep)
                && j != i
            {
                // skip self-reference
                adj[j].push(i);
                in_degree[i] += 1;
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
            if in_degree[j] == 0 {
                queue.push_back(j);
            }
        }
    }

    if order.len() != n {
        // Circular reference detected — evaluate remaining cells in original order with a warning
        let visited: HashSet<usize> = order.iter().copied().collect();
        for i in 0..n {
            if !visited.contains(&i) {
                order.push(i);
            }
        }
        // Return Ok with best-effort order rather than hard-erroring; circular refs will show stale values
    }
    Ok(order)
}

/// `ReDim Preserve arr(...)` on an array of the same rank as `new_bounds`.
/// Real VBA only allows the *last* dimension's *upper* bound to change under
/// `Preserve`; every other dimension (and the last one's lower bound) keeps
/// its existing value exactly, or there's no well-defined way to keep an
/// existing element at the same subscript it had before — that case is
/// Error 9, same family as every other array-shape mismatch this module
/// reports (not independently confirmed against a live Excel, but
/// internally consistent with this codebase's own established convention
/// for array-shape errors).
fn redim_preserve(old: &VbaArray, new_bounds: &[ArrayBound]) -> Result<VbaArray, String> {
    if new_bounds.is_empty() {
        // No real dimension to preserve into — surfaces the same
        // "at least one dimension" error `new_zeroed` itself would.
        return VbaArray::new_zeroed(new_bounds.to_vec());
    }
    let last = new_bounds.len() - 1;
    if old.bounds[last].lower != new_bounds[last].lower {
        return Err("Subscript out of range".to_string());
    }
    if old.bounds[..last] != new_bounds[..last] {
        return Err("Subscript out of range".to_string());
    }
    let mut new_arr = VbaArray::new_zeroed(new_bounds.to_vec())?;
    // Walk every index in `old`'s own shape (an odometer: the last
    // dimension advances fastest) and copy each element to the same
    // subscript in `new_arr`, if that subscript still exists there.
    let mut idx: Vec<i64> = old.bounds.iter().map(|b| b.lower).collect();
    if old.elements.is_empty() {
        return Ok(new_arr);
    }
    loop {
        if let (Ok(old_i), Ok(new_i)) = (old.linear_index(&idx), new_arr.linear_index(&idx)) {
            new_arr.elements[new_i] = old.elements[old_i].clone();
        }
        let mut d = idx.len();
        loop {
            if d == 0 {
                return Ok(new_arr);
            }
            d -= 1;
            idx[d] += 1;
            if idx[d] <= old.bounds[d].upper {
                break;
            }
            idx[d] = old.bounds[d].lower;
        }
    }
}

fn vba_to_str(v: &Variant) -> String {
    match v {
        Variant::Str(s) => s.clone(),
        Variant::Integer(n) => n.to_string(),
        Variant::Float(f) => f.to_string(),
        Variant::Boolean(b) => {
            if *b {
                "True".into()
            } else {
                "False".into()
            }
        }
        Variant::Date(s) => serial_to_display(*s),
        Variant::Error(e) => e.as_str().to_string(),
        // "Any expression that is Empty is also treated as a zero-length
        // string"; a lone Null concatenates the same way (the both-Null
        // case is decided earlier, in `null_rule`).
        Variant::Empty | Variant::Null => String::new(),
        Variant::Array(a) => a.iter().map(vba_to_str).collect::<Vec<_>>().join(", "),
        Variant::VbaArray(a) => a
            .elements
            .iter()
            .map(vba_to_str)
            .collect::<Vec<_>>()
            .join(", "),
        Variant::Record(_) => "[Record]".into(),
    }
}

fn cmp_variants(a: &Variant, b: &Variant) -> std::cmp::Ordering {
    use std::cmp::Ordering::*;
    match (a, b) {
        (Variant::Integer(x), Variant::Integer(y)) => x.cmp(y),
        (Variant::Float(x), Variant::Float(y)) => x.partial_cmp(y).unwrap_or(Equal),
        (Variant::Integer(x), Variant::Float(y)) => (*x as f64).partial_cmp(y).unwrap_or(Equal),
        (Variant::Float(x), Variant::Integer(y)) => x.partial_cmp(&(*y as f64)).unwrap_or(Equal),
        (Variant::Str(x), Variant::Str(y)) => x.to_lowercase().cmp(&y.to_lowercase()),
        (Variant::Empty, Variant::Empty) => Equal,
        (Variant::Empty, _) => Less,
        (_, Variant::Empty) => Greater,
        _ => vba_to_str(a)
            .to_lowercase()
            .cmp(&vba_to_str(b).to_lowercase()),
    }
}

fn flat_nums(vals: &[Variant]) -> Vec<f64> {
    let mut out = vec![];
    for v in vals {
        match v {
            Variant::Array(a) => out.extend(a.iter().filter_map(|x| to_f64_excel(x).ok())),
            Variant::VbaArray(a) => {
                out.extend(a.elements.iter().filter_map(|x| to_f64_excel(x).ok()))
            }
            _ => {
                if let Ok(f) = to_f64_excel(v) {
                    out.push(f);
                }
            }
        }
    }
    out
}

fn flat_all(vals: &[Variant]) -> Vec<Variant> {
    vals.iter()
        .flat_map(|v| match v {
            Variant::Array(a) => a.clone(),
            Variant::VbaArray(a) => a.elements.clone(),
            other => vec![other.clone()],
        })
        .collect()
}

fn eval_wsf(func: &str, vals: &[Variant]) -> Result<Variant, String> {
    match func {
        "sum" => {
            let nums = flat_nums(vals);
            Ok(as_int_if_whole(nums.iter().sum::<f64>()))
        }
        "max" => {
            let nums = flat_nums(vals);
            nums.iter()
                .cloned()
                .reduce(f64::max)
                .map(as_int_if_whole)
                .ok_or_else(|| "WorksheetFunction.Max: no values".into())
        }
        "min" => {
            let nums = flat_nums(vals);
            nums.iter()
                .cloned()
                .reduce(f64::min)
                .map(as_int_if_whole)
                .ok_or_else(|| "WorksheetFunction.Min: no values".into())
        }
        "average" => {
            let nums = flat_nums(vals);
            if nums.is_empty() {
                return Err("WorksheetFunction.Average: no values".into());
            }
            Ok(as_int_if_whole(
                nums.iter().sum::<f64>() / nums.len() as f64,
            ))
        }
        "count" => {
            let n = flat_all(vals)
                .iter()
                .filter(|v| matches!(v, Variant::Integer(_) | Variant::Float(_)))
                .count();
            Ok(Variant::Integer(n as i64))
        }
        "counta" => {
            let n = flat_all(vals)
                .iter()
                .filter(|v| !matches!(v, Variant::Empty))
                .count();
            Ok(Variant::Integer(n as i64))
        }
        "countblank" => {
            let n = flat_all(vals)
                .iter()
                .filter(|v| matches!(v, Variant::Empty))
                .count();
            Ok(Variant::Integer(n as i64))
        }
        "countif" => {
            if vals.len() < 2 {
                return Err("WorksheetFunction.CountIf requires 2 arguments".into());
            }
            let range = flat_all(&vals[..1]);
            let criteria = &vals[1];
            let n = range
                .iter()
                .filter(|v| wsf_criteria_match(v, criteria))
                .count();
            Ok(Variant::Integer(n as i64))
        }
        "sumif" => {
            // SumIf(range, criteria [, sum_range])
            if vals.len() < 2 {
                return Err("WorksheetFunction.SumIf requires at least 2 arguments".into());
            }
            let crit_range = flat_all(&vals[..1]);
            let criteria = &vals[1];
            let sum_range = if vals.len() >= 3 {
                flat_all(&vals[2..3])
            } else {
                crit_range.clone()
            };
            let total: f64 = crit_range
                .iter()
                .zip(sum_range.iter())
                .filter(|(cv, _)| wsf_criteria_match(cv, criteria))
                .filter_map(|(_, sv)| to_f64_excel(sv).ok())
                .sum();
            Ok(as_int_if_whole(total))
        }
        "round" => {
            if vals.is_empty() {
                return Err("WorksheetFunction.Round requires arguments".into());
            }
            let f = to_f64_excel(&vals[0])?;
            let digits = if vals.len() >= 2 {
                to_f64_excel(&vals[1])? as i32
            } else {
                0
            };
            let factor = 10f64.powi(digits);
            Ok(as_int_if_whole((f * factor).round() / factor))
        }
        "abs" => {
            let f = to_f64_excel(vals.first().ok_or("Abs: no arg")?)?;
            Ok(as_int_if_whole(f.abs()))
        }
        "sqrt" => {
            let f = to_f64_excel(vals.first().ok_or("Sqrt: no arg")?)?;
            Ok(Variant::Float(f.sqrt()))
        }
        "power" => {
            if vals.len() < 2 {
                return Err("Power requires 2 arguments".into());
            }
            Ok(as_int_if_whole(
                to_f64_excel(&vals[0])?.powf(to_f64_excel(&vals[1])?),
            ))
        }
        "log" => {
            let x = to_f64_excel(vals.first().ok_or("Log: no arg")?)?;
            let base = if vals.len() >= 2 {
                to_f64_excel(&vals[1])?
            } else {
                std::f64::consts::E
            };
            Ok(Variant::Float(x.log(base)))
        }
        "match" => {
            // Match(lookup_val, lookup_array, [match_type]) — returns 1-based position
            if vals.len() < 2 {
                return Err("Match: requires at least 2 arguments".into());
            }
            let target = &vals[0];
            let arr = flat_all(&vals[1..2]);
            let pos = arr
                .iter()
                .position(|v| vba_eq(v, target))
                .map(|i| Variant::Integer(i as i64 + 1))
                .unwrap_or(Variant::Error(ExcelError::NA));
            Ok(pos)
        }
        "index" => {
            // Index(array, row_num [, col_num])
            if vals.len() < 2 {
                return Err("Index: requires at least 2 arguments".into());
            }
            let arr = flat_all(&vals[0..1]);
            let idx = (to_f64_excel(&vals[1])? as usize).saturating_sub(1);
            Ok(arr
                .get(idx)
                .cloned()
                .unwrap_or(Variant::Error(ExcelError::Ref)))
        }
        _ => Err(format!("WorksheetFunction.{} is not implemented", func)),
    }
}

/// `true` iff `name` is a recognized built-in VBA function or
/// `WorksheetFunction.*` method (via the `wsf_` prefix). Used by the
/// `check` subcommand to consult the *real* dispatch table instead of a
/// hand-maintained mirror that would drift as functions are added — a
/// throwaway `Vm` + zero-arg probe call is cheap and has no second source
/// of truth to go stale.
pub fn is_known_builtin_function(name: &str) -> bool {
    let mut vm = Vm::new();
    match vm.eval_vba_func(name, &[]) {
        Ok(_) => true,
        Err(msg) => {
            !msg.starts_with("Unknown VBA function: '") && !msg.ends_with("is not implemented")
        }
    }
}

/// The exact error `eval_vba_func` produces for a zero-arg call to `name`,
/// or `None` if it's actually a known, working builtin. Same zero-arg-probe
/// technique as `is_known_builtin_function` (safe here for the same reason:
/// every "not known" arm — the generic fallback and `eval_wsf`'s own
/// catch-all — matches purely on `name`, never on the argument list, so the
/// probed message is identical to what a real call with any other argument
/// count would produce).
///
/// Used by `check::compile_check_errors` so its pre-flight rejection of an
/// unresolvable `Expr::FuncCall` reports the same wording running it would
/// have produced, instead of inventing separate text that could drift from
/// a dispatch arm's own message — `wsf_textjoin`, for instance, fails
/// inside `eval_wsf` with "WorksheetFunction.textjoin is not implemented",
/// not the generic "Unknown VBA function" `eval_vba_func`'s own top-level
/// fallback arm uses.
pub fn builtin_call_error(name: &str) -> Option<String> {
    let mut vm = Vm::new();
    vm.eval_vba_func(name, &[]).err()
}

fn wsf_criteria_match(v: &Variant, criteria: &Variant) -> bool {
    match criteria {
        Variant::Str(s) => {
            let s = s.trim();
            // Comparison criteria like ">5", "<>0", ">=10"
            if let Some(rest) = s.strip_prefix(">=") {
                if let Ok(n) = rest.parse::<f64>() {
                    return to_f64(v).is_ok_and(|f| f >= n);
                }
            } else if let Some(rest) = s.strip_prefix("<=") {
                if let Ok(n) = rest.parse::<f64>() {
                    return to_f64(v).is_ok_and(|f| f <= n);
                }
            } else if let Some(rest) = s.strip_prefix("<>") {
                if let Ok(n) = rest.parse::<f64>() {
                    return to_f64(v).is_ok_and(|f| f != n);
                }
                return vba_to_str(v).to_lowercase() != rest.to_lowercase();
            } else if let Some(rest) = s.strip_prefix('>') {
                if let Ok(n) = rest.parse::<f64>() {
                    return to_f64(v).is_ok_and(|f| f > n);
                }
            } else if let Some(rest) = s.strip_prefix('<')
                && let Ok(n) = rest.parse::<f64>()
            {
                return to_f64(v).is_ok_and(|f| f < n);
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
    let dec_places = fmt
        .find('.')
        .map(|i| {
            fmt[i + 1..]
                .chars()
                .filter(|c| *c == '0' || *c == '#')
                .count()
        })
        .unwrap_or(0);
    match v {
        Variant::Integer(n) => {
            let f = *n as f64;
            if thousands {
                // Simple thousands separator
                let int_part = format!("{}", n.abs());
                let grouped: String = int_part
                    .chars()
                    .rev()
                    .enumerate()
                    .flat_map(|(i, c)| {
                        if i > 0 && i % 3 == 0 {
                            vec![',', c]
                        } else {
                            vec![c]
                        }
                    })
                    .collect::<String>()
                    .chars()
                    .rev()
                    .collect();
                let signed = if *n < 0 {
                    format!("-{}", grouped)
                } else {
                    grouped
                };
                if dec_places > 0 {
                    format!("{}.{}", signed, "0".repeat(dec_places))
                } else {
                    signed
                }
            } else if dec_places > 0 {
                format!("{:.prec$}", f, prec = dec_places)
            } else {
                format!("{}", n)
            }
        }
        Variant::Float(f) => {
            if thousands {
                let int_part = format!("{}", (*f as i64).abs());
                let grouped: String = int_part
                    .chars()
                    .rev()
                    .enumerate()
                    .flat_map(|(i, c)| {
                        if i > 0 && i % 3 == 0 {
                            vec![',', c]
                        } else {
                            vec![c]
                        }
                    })
                    .collect::<String>()
                    .chars()
                    .rev()
                    .collect();
                let signed = if *f < 0.0 {
                    format!("-{}", grouped)
                } else {
                    grouped
                };
                if dec_places > 0 {
                    format!("{}.{:.prec$}", signed, f.fract().abs(), prec = dec_places)
                } else {
                    signed
                }
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
        Variant::Float(f) => Ok(*f),
        // VBA represents True as -1 internally (CInt(True) = -1), unlike Excel worksheet
        // formulas where TRUE arithmetic-coerces to 1 -- see formula::eval's separate
        // to_float, which is correct as 1.0 for that different language. Found via the
        // vba-semantics suite's operator-coercion matrix: True + 5 was returning 6 instead
        // of VBA's documented 4.
        Variant::Boolean(b) => Ok(if *b { -1.0 } else { 0.0 }),
        Variant::Date(s) => Ok(*s as f64),
        Variant::Error(e) => Err(e.to_string()),
        Variant::Empty => Ok(0.0),
        // Real VBA error 94. Unlike Empty (documented as 0 in a numeric
        // context), Null has no numeric value at all. Every documented
        // Null-propagating operator short-circuits before reaching here, so
        // this only fires where a Null genuinely can't propagate (e.g. a
        // function argument that must be a number).
        Variant::Null => Err("Invalid use of Null".into()),
        Variant::Str(s) => s
            .parse::<f64>()
            .map_err(|_| format!("Cannot convert '{}' to number", s)),
        Variant::Array(_) | Variant::VbaArray(_) => Err("Cannot convert array to number".into()),
        Variant::Record(_) => Err("Cannot convert record to number".into()),
    }
}

/// Same as to_f64, but for WorksheetFunction.* (eval_wsf/flat_nums) call sites only:
/// Application.WorksheetFunction bridges into Excel's own calculation engine, so its
/// Boolean coercion matches a worksheet formula (TRUE=1), not VBA's own True=-1 --
/// confirmed live (WorksheetFunction.Sum(True, True) is 2 in real VBA, not -2). Found when
/// the True=-1 fix to plain to_f64 silently changed this too, since flat_nums used to share
/// it; every other Variant kind coerces identically either way.
fn to_f64_excel(v: &Variant) -> Result<f64, String> {
    match v {
        Variant::Boolean(b) => Ok(if *b { 1.0 } else { 0.0 }),
        other => to_f64(other),
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
        Variant::Empty => false,
        // Documented on the If...Then...Else statement page: "If condition
        // is Null, condition is treated as False." Not an error — this is
        // the one place VBA gives Null a Boolean reading.
        Variant::Null => false,
        Variant::Array(a) => !a.is_empty(),
        Variant::VbaArray(a) => !a.elements.is_empty(),
        Variant::Record(_) => true,
    }
}

fn vba_eq(a: &Variant, b: &Variant) -> bool {
    match (a, b) {
        (Variant::Integer(x), Variant::Integer(y)) => x == y,
        (Variant::Float(x), Variant::Float(y)) => x == y,
        (Variant::Integer(x), Variant::Float(y)) => (*x as f64) == *y,
        (Variant::Float(x), Variant::Integer(y)) => *x == (*y as f64),
        (Variant::Date(x), Variant::Date(y)) => x == y,
        (Variant::Date(x), Variant::Integer(y)) => x == y,
        (Variant::Integer(x), Variant::Date(y)) => x == y,
        (Variant::Str(x), Variant::Str(y)) => x.to_uppercase() == y.to_uppercase(),
        (Variant::Boolean(x), Variant::Boolean(y)) => x == y,
        (Variant::Empty, Variant::Empty) => true,
        // Documented VBA comparison rules: Empty numeric-compares as 0, string-compares as
        // "" -- vba_cmp (used for </>) already applies this via to_f64's Empty=>0.0 arm, but
        // vba_eq's old catch-all fell through to `false` for e.g. `0 = Empty`, an internal
        // inconsistency between = and < on the same operand pair. Found via the
        // vba-semantics suite's comparison-coercion matrix.
        (Variant::Empty, Variant::Str(s)) | (Variant::Str(s), Variant::Empty) => s.is_empty(),
        (Variant::Empty, other) | (other, Variant::Empty) => {
            to_f64(other).map(|f| f == 0.0).unwrap_or(false)
        }
        (Variant::Error(_), _) | (_, Variant::Error(_)) => false,
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
        (Ok(fa), Ok(fb)) => fa
            .partial_cmp(&fb)
            .ok_or_else(|| "Cannot compare NaN values".into()),
        _ => {
            let sa = match a {
                Variant::Str(s) => s.clone(),
                _ => format!("{:?}", a),
            };
            let sb = match b {
                Variant::Str(s) => s.clone(),
                _ => format!("{:?}", b),
            };
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

/// Parses `Val()`'s leading numeric prefix: optional leading whitespace,
/// optional sign, digits, optional `.digits` — stopping at (not erroring
/// on) the first character that doesn't fit, same as real VBA. Returns 0.0
/// if no valid numeric prefix exists at all. Scoped to this core grammar;
/// real VBA's documented embedded-whitespace-stripping inside the numeric
/// prefix (`Val("1 2 3")` == 123) isn't attempted — no evidence it's
/// needed, and this project doesn't guess at rarely-exercised VBA quirks.
fn parse_vba_val_prefix(s: &str) -> f64 {
    let trimmed = s.trim_start();
    let bytes = trimmed.as_bytes();
    let mut i = 0;
    if i < bytes.len() && (bytes[i] == b'+' || bytes[i] == b'-') {
        i += 1;
    }
    let mut end = i;
    let mut seen_digit = false;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
        seen_digit = true;
        end = i;
    }
    if i < bytes.len() && bytes[i] == b'.' {
        let after_dot = i + 1;
        let mut j = after_dot;
        while j < bytes.len() && bytes[j].is_ascii_digit() {
            j += 1;
        }
        if j > after_dot {
            end = j;
            seen_digit = true;
        }
    }
    if !seen_digit {
        return 0.0;
    }
    trimmed[..end].parse::<f64>().unwrap_or(0.0)
}

/// Whole days since the Unix epoch, for `Date`/`Now` — same
/// `unix_days + 25569` (Excel serial of 1970-01-01) convention the formula
/// engine's own `func_now` uses.
fn unix_epoch_days() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs()
        / 86400
}

/// Seconds elapsed since local midnight UTC, for `Time`/`Now`'s fractional
/// time-of-day component.
fn unix_seconds_of_day() -> u64 {
    std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default()
        .as_secs()
        % 86400
}

fn to_cell_index(v: Variant, label: &str) -> Result<u32, String> {
    let f = to_f64(&v)?;
    if f < 1.0 || f.fract() != 0.0 {
        return Err(format!(
            "Cell {} must be a positive integer, got {}",
            label, f
        ));
    }
    Ok(f as u32)
}

/// VBA's `\`/`Mod` operators round each operand to a whole number first
/// (real VBA coerces to Long via the same round-half-to-even — "banker's
/// rounding" — that `CLng`/`Round` use, not truncation and not Rust's
/// default round-half-away-from-zero `f64::round()`) before doing integer
/// division/modulus. E.g. `5 \ 0.5`: 0.5 rounds to 0 (nearest even), giving
/// `5 \ 0` — a division-by-zero error, not `5 \ 1`.
fn to_i64_rounded(v: &Variant) -> Result<i64, String> {
    Ok(to_f64(v)?.round_ties_even() as i64)
}

/// `And`/`Or`/`Xor`/(unary) `Not`'s numeric-context coercion: same
/// round-to-Long as `to_i64_rounded`, except a `Boolean` converts to VBA's
/// actual internal bit pattern (`True` = -1 i.e. all-ones, `False` = 0),
/// not `to_f64`'s `1.0`/`0.0` — needed so a mixed `<number> And <boolean>`
/// bitwise-ANDs against the Boolean's real all-ones/all-zeros pattern (e.g.
/// `5 And True` = 5, not 1).
fn to_i64_bitwise(v: &Variant) -> Result<i64, String> {
    match v {
        Variant::Boolean(b) => Ok(if *b { -1 } else { 0 }),
        other => to_i64_rounded(other),
    }
}

/// Applies every documented Null rule for `op` before the ordinary
/// coercion path ever sees the operands, returning `Some(result)` when Null
/// decides the answer. All sourced from Microsoft's own VBA language
/// reference, fetched live rather than recalled:
///
/// - `+`/`-` (and the rest of arithmetic): "If one or both expressions are
///   Null expressions, result is Null."
/// - `&`: "If both expressions are Null, result is Null. However, if only
///   one expression is Null, that expression is treated as a zero-length
///   string." — the one operator where a single Null does *not* propagate.
/// - comparison operators: each of `<`, `<=`, `>`, `>=`, `=`, `<>` lists
///   "Null if expression1 or expression2 = Null" as a third outcome
///   alongside True/False.
/// - `And`/`Or`/`Xor`/`Not`: three-valued truth tables in which Null does
///   *not* uniformly propagate — `False And Null` is False and
///   `True Or Null` is True, because those two answers are already
///   determined without knowing the missing operand.
///
/// Returning `Option` rather than short-circuiting inside `eval_binop`'s own
/// arms keeps each rule stated once, next to its citation, instead of spread
/// across seven operator branches.
fn null_rule(op: &VbaBinOp, l: &Variant, r: &Variant) -> Option<Variant> {
    let l_null = matches!(l, Variant::Null);
    let r_null = matches!(r, Variant::Null);
    if !l_null && !r_null {
        return None;
    }
    match op {
        VbaBinOp::Add
        | VbaBinOp::Sub
        | VbaBinOp::Mul
        | VbaBinOp::Div
        | VbaBinOp::Pow
        | VbaBinOp::IntDiv
        | VbaBinOp::Mod
        | VbaBinOp::Eq
        | VbaBinOp::Ne
        | VbaBinOp::Lt
        | VbaBinOp::Le
        | VbaBinOp::Gt
        | VbaBinOp::Ge => Some(Variant::Null),
        // Only both-Null concatenates to Null; a single Null is "".
        VbaBinOp::Concat => {
            if l_null && r_null {
                Some(Variant::Null)
            } else {
                None
            }
        }
        // `False And Null` -> False; `Null And False` -> False; everything
        // else involving Null -> Null.
        VbaBinOp::And => {
            if matches!(l, Variant::Boolean(false)) || matches!(r, Variant::Boolean(false)) {
                Some(Variant::Boolean(false))
            } else {
                Some(Variant::Null)
            }
        }
        // `True Or Null` -> True; `Null Or True` -> True; everything else
        // involving Null -> Null.
        VbaBinOp::Or => {
            if matches!(l, Variant::Boolean(true)) || matches!(r, Variant::Boolean(true)) {
                Some(Variant::Boolean(true))
            } else {
                Some(Variant::Null)
            }
        }
        // "However, if either expression is Null, result is also Null."
        VbaBinOp::Xor => Some(Variant::Null),
    }
}

/// `to_f64` for an *arithmetic operator's* operand: a string that can't
/// convert to a number gets real VBA's own wording for that situation
/// ("One expression is a numeric data type and the other is a String |
/// A `Type mismatch` error occurs" — the + operator reference), instead of
/// `to_f64`'s internal coercion-failure text.
///
/// Deliberately a wrapper rather than a change to `to_f64` itself: that
/// helper has ~54 call sites across the VM (loop bounds, array indices,
/// function arguments, …), each with its own correct real-VBA wording for
/// its own failure, and only the arithmetic operators are documented to say
/// "Type mismatch" here. Every other `to_f64` message is untouched.
fn arith_to_f64(v: &Variant) -> Result<f64, String> {
    match v {
        Variant::Str(s) if s.parse::<f64>().is_err() => Err("Type mismatch".into()),
        other => to_f64(other),
    }
}

fn eval_binop(op: &VbaBinOp, l: Variant, r: Variant) -> Result<Variant, String> {
    if let Some(v) = null_rule(op, &l, &r) {
        return Ok(v);
    }
    match op {
        VbaBinOp::Add | VbaBinOp::Sub | VbaBinOp::Mul | VbaBinOp::Div | VbaBinOp::Pow => {
            let lf = arith_to_f64(&l)?;
            let rf = arith_to_f64(&r)?;
            let result = match op {
                VbaBinOp::Add => lf + rf,
                VbaBinOp::Sub => lf - rf,
                VbaBinOp::Mul => lf * rf,
                VbaBinOp::Div => {
                    if rf == 0.0 {
                        return Err("Division by zero".into());
                    }
                    lf / rf
                }
                VbaBinOp::Pow => lf.powf(rf),
                _ => unreachable!(),
            };
            Ok(as_int_if_whole(result))
        }
        VbaBinOp::IntDiv | VbaBinOp::Mod => {
            let li = to_i64_rounded(&l)?;
            let ri = to_i64_rounded(&r)?;
            let result = match op {
                VbaBinOp::IntDiv => li.checked_div(ri),
                VbaBinOp::Mod => li.checked_rem(ri),
                _ => unreachable!(),
            };
            match result {
                Some(v) => Ok(Variant::Integer(v)),
                None if ri == 0 => Err("Division by zero".into()),
                None => Err("Integer division overflow".into()),
            }
        }
        VbaBinOp::And | VbaBinOp::Or | VbaBinOp::Xor => {
            // Both operands genuinely Boolean → logical op, Boolean result
            // (VBA's own distinction between the "logical" and "bitwise"
            // reading of the same operator). Otherwise, numeric bitwise op.
            if let (Variant::Boolean(lb), Variant::Boolean(rb)) = (&l, &r) {
                let result = match op {
                    VbaBinOp::And => *lb && *rb,
                    VbaBinOp::Or => *lb || *rb,
                    VbaBinOp::Xor => *lb != *rb,
                    _ => unreachable!(),
                };
                Ok(Variant::Boolean(result))
            } else {
                let li = to_i64_bitwise(&l)?;
                let ri = to_i64_bitwise(&r)?;
                let result = match op {
                    VbaBinOp::And => li & ri,
                    VbaBinOp::Or => li | ri,
                    VbaBinOp::Xor => li ^ ri,
                    _ => unreachable!(),
                };
                Ok(Variant::Integer(result))
            }
        }
        VbaBinOp::Concat => {
            // A single Null operand is documented to concatenate as a
            // zero-length string, exactly like Empty (the both-Null case
            // already returned Null in `null_rule`).
            let l = if matches!(l, Variant::Null) {
                Variant::Empty
            } else {
                l
            };
            let r = if matches!(r, Variant::Null) {
                Variant::Empty
            } else {
                r
            };
            Ok(Variant::Str(format!("{}{}", l, r)))
        }
        VbaBinOp::Eq => Ok(Variant::Boolean(vba_eq(&l, &r))),
        VbaBinOp::Ne => Ok(Variant::Boolean(!vba_eq(&l, &r))),
        VbaBinOp::Lt => Ok(Variant::Boolean(
            vba_cmp(&l, &r)? == std::cmp::Ordering::Less,
        )),
        VbaBinOp::Le => Ok(Variant::Boolean(
            vba_cmp(&l, &r)? != std::cmp::Ordering::Greater,
        )),
        VbaBinOp::Gt => Ok(Variant::Boolean(
            vba_cmp(&l, &r)? == std::cmp::Ordering::Greater,
        )),
        VbaBinOp::Ge => Ok(Variant::Boolean(
            vba_cmp(&l, &r)? != std::cmp::Ordering::Less,
        )),
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
        assert_eq!(
            run("Sub MySub()\n    a = 42\nEnd Sub\n").variables["a"],
            Variant::Integer(42)
        );
    }

    #[test]
    fn test_variable_assignment_float() {
        assert_eq!(
            run("Sub MySub()\n    x = 1.5\nEnd Sub\n").variables["x"],
            Variant::Float(1.5)
        );
    }

    #[test]
    fn test_variable_assignment_string() {
        assert_eq!(
            run("Sub MySub()\n    s = \"hello\"\nEnd Sub\n").variables["s"],
            Variant::Str("hello".into())
        );
    }

    #[test]
    fn test_cell_write_literal() {
        assert_eq!(
            run("Sub MySub()\n    Cells(1, 1).Value = 100\nEnd Sub\n").get_cell(1, 1),
            Variant::Integer(100)
        );
    }

    #[test]
    fn test_cell_write_from_variable() {
        assert_eq!(
            run("Sub MySub()\n    x = 99\n    Cells(2, 3).Value = x\nEnd Sub\n").get_cell(2, 3),
            Variant::Integer(99)
        );
    }

    #[test]
    fn test_cell_write_string() {
        assert_eq!(
            run("Sub MySub()\n    Cells(1, 2).Value = \"world\"\nEnd Sub\n").get_cell(1, 2),
            Variant::Str("world".into())
        );
    }

    #[test]
    fn test_cell_empty_by_default() {
        assert_eq!(Vm::new().get_cell(1, 1), Variant::Empty);
    }

    #[test]
    fn test_multiple_cells() {
        let vm = run(
            "Sub MySub()\n    Cells(1, 1).Value = 1\n    Cells(1, 2).Value = 2\n    Cells(2, 1).Value = 3\nEnd Sub\n",
        );
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
    // 3.14 is an arbitrary decimal test value for Variant::Float's Display impl, not π.
    #[allow(clippy::approx_constant)]
    fn test_variant_display() {
        assert_eq!(Variant::Integer(42).to_string(), "42");
        assert_eq!(Variant::Float(3.14).to_string(), "3.14");
        assert_eq!(Variant::Boolean(true).to_string(), "True");
        assert_eq!(Variant::Empty.to_string(), "");
    }

    // ── Arithmetic ────────────────────────────────────────────────────────────

    #[test]
    fn test_arithmetic_assignment() {
        let vm = run(
            "Sub MySub()\n    a = 3 + 4\n    b = 10 - 3\n    c = 2 * 5\n    d = 10 / 4\nEnd Sub\n",
        );
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
        )
        .unwrap();
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
        let vm = run(
            "Sub MySub()\n    x = 0\n    Do While x < 5\n        x = x + 1\n    Loop\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test]
    fn test_do_until_loop() {
        let vm = run(
            "Sub MySub()\n    x = 0\n    Do Until x >= 5\n        x = x + 1\n    Loop\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test]
    fn test_do_loop_while_post() {
        // Post-check: body runs at least once even if condition is already false
        let vm = run(
            "Sub MySub()\n    x = 99\n    Do\n        x = x + 1\n    Loop While x < 5\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(100));
    }

    // ── Select Case ───────────────────────────────────────────────────────────

    #[test]
    fn test_select_case_value() {
        let vm = run(
            "Sub MySub()\n    x = 2\n    Select Case x\n        Case 1\n            r = \"one\"\n        Case 2\n            r = \"two\"\n        Case Else\n            r = \"other\"\n    End Select\nEnd Sub\n",
        );
        assert_eq!(vm.variables["r"], Variant::Str("two".into()));
    }

    #[test]
    fn test_select_case_else() {
        let vm = run(
            "Sub MySub()\n    x = 99\n    Select Case x\n        Case 1\n            r = 1\n        Case Else\n            r = 0\n    End Select\nEnd Sub\n",
        );
        assert_eq!(vm.variables["r"], Variant::Integer(0));
    }

    #[test]
    fn test_select_case_multi_value() {
        let vm = run(
            "Sub MySub()\n    x = 3\n    Select Case x\n        Case 1, 2\n            r = 12\n        Case 3, 4\n            r = 34\n    End Select\nEnd Sub\n",
        );
        assert_eq!(vm.variables["r"], Variant::Integer(34));
    }

    #[test]
    fn test_select_case_is_op() {
        let vm = run(
            "Sub MySub()\n    x = 10\n    Select Case x\n        Case Is > 5\n            r = 1\n        Case Else\n            r = 0\n    End Select\nEnd Sub\n",
        );
        assert_eq!(vm.variables["r"], Variant::Integer(1));
    }

    #[test]
    fn test_select_case_range() {
        let vm = run(
            "Sub MySub()\n    x = 3\n    Select Case x\n        Case 1 To 5\n            r = 1\n        Case Else\n            r = 0\n    End Select\nEnd Sub\n",
        );
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
        let vm = run(
            "Sub MySub()\n    With Sheet1\n        .Cells(1, 1).Value = 100\n        .Cells(2, 1).Value = 200\n    End With\nEnd Sub\n",
        );
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
        // -2.5's nearest even integer is -2, not -3 (Rust's own default
        // f64::round() would give -3, away from zero — see
        // test_vba_cint_clng_use_banker_rounding for the full tie-case
        // coverage this single value used to lack an assertion for).
        assert_eq!(vm.variables["b"], Variant::Integer(-2));
    }

    #[test]
    fn test_vba_cint_clng_use_banker_rounding() {
        // Real VBA's CInt/CLng round half-to-even, like Round() — not
        // Rust's own default f64::round() (half-away-from-zero). CInt(0.5)
        // used to be 1, not real VBA's 0.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    a = CInt(0.5)\n",
            "    b = CInt(1.5)\n",
            "    c = CLng(2.5)\n",
            "    d = CLng(-2.5)\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["a"], Variant::Integer(0));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
        assert_eq!(vm.variables["c"], Variant::Integer(2));
        assert_eq!(vm.variables["d"], Variant::Integer(-2));
    }

    #[test]
    fn test_vba_fix_truncates_toward_zero() {
        // Fix() truncates toward zero, unlike Int() which floors toward
        // negative infinity — Fix(-3.9) is -3, not -4 (see test_vba_int's
        // Int(-3.1) == -4 for the contrast).
        let vm = run("Sub MySub()\n    a = Fix(3.9)\n    b = Fix(-3.9)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(3));
        assert_eq!(vm.variables["b"], Variant::Integer(-3));
    }

    #[test]
    fn test_vba_sgn() {
        let vm = run("Sub MySub()\n    a = Sgn(-5)\n    b = Sgn(5)\n    c = Sgn(0)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(-1));
        assert_eq!(vm.variables["b"], Variant::Integer(1));
        assert_eq!(vm.variables["c"], Variant::Integer(0));
    }

    #[test]
    fn test_vba_round_uses_banker_rounding_unlike_worksheetfunction_round() {
        // Real VBA's own Round() rounds half-to-even; WorksheetFunction.Round
        // (Excel's ROUND() formula) rounds half-away-from-zero — genuinely
        // different functions, not aliases. `Round(2.5)` == 2 (nearest even),
        // not 3 (which is what WorksheetFunction.Round(2.5) gives).
        let vm = run(concat!(
            "Sub MySub()\n",
            "    a = Round(3.14159)\n",
            "    b = Round(-1.5)\n",
            "    c = Round(2.5)\n",
            "    d = Round(0.125, 2)\n",
            "    e = Application.WorksheetFunction.Round(2.5)\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["a"], Variant::Integer(3));
        assert_eq!(vm.variables["b"], Variant::Integer(-2));
        assert_eq!(vm.variables["c"], Variant::Integer(2));
        assert_eq!(vm.variables["d"], Variant::Float(0.12));
        assert_eq!(vm.variables["e"], Variant::Integer(3));
    }

    #[test]
    fn test_vba_round_rejects_negative_digits() {
        // Unlike WorksheetFunction.Round/Excel's ROUND(), real VBA's own
        // Round() errors on a negative digit count rather than rounding
        // left of the decimal point.
        let prog = parser::parse("Sub MySub()\n    a = Round(1234.5, -2)\nEnd Sub\n").unwrap();
        let err = Vm::new().run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Invalid procedure call or argument");
    }

    #[test]
    fn test_vba_date_now_time_return_real_values_not_a_debug_string() {
        // Used to return a Rust debug-formatted `SystemTime { tv_sec: ...
        // }` string regardless of which of the three was called — visibly
        // wrong, not just imprecise. `Date()` must round-trip through the
        // same Excel-serial epoch math the formula engine's own NOW() uses
        // (25569 == Excel serial of 1970-01-01), matching the real system
        // clock; `Time()`/`Now()` must at least be numerically-plausible
        // Doubles (0.0..1.0 for Time, an Excel-serial-plus-fraction value
        // for Now), not a debug string, even though `Variant::Date` being
        // whole-day-only means their `TypeName` can't be "Date" here (a
        // disclosed, separate gap — see ROADMAP.md).
        let vm = run(concat!(
            "Sub MySub()\n",
            "    d = Date()\n",
            "    dt = TypeName(Date())\n",
            "    t = Time()\n",
            "    n = Now()\n",
            "End Sub\n",
        ));
        let expected_serial = unix_epoch_days() as i64 + 25569;
        assert_eq!(vm.variables["d"], Variant::Date(expected_serial));
        assert_eq!(vm.variables["dt"], Variant::Str("Date".to_string()));
        match vm.variables["t"] {
            Variant::Float(f) => assert!((0.0..1.0).contains(&f), "Time() out of range: {f}"),
            ref other => panic!("expected Variant::Float, got {other:?}"),
        }
        match vm.variables["n"] {
            Variant::Float(f) => assert!(f > 25569.0, "Now() looks wrong: {f}"),
            ref other => panic!("expected Variant::Float, got {other:?}"),
        }
    }

    #[test]
    fn test_vba_date_now_time_work_without_parens() {
        // Real VBA allows omitting `()` on these three zero-arg functions.
        // `Expr::Var("date")` used to always hit "Undefined variable" — a
        // bare identifier only recognized these three as a fallback after
        // failing the ordinary variable/constant lookups, so this doesn't
        // risk masking a genuine variable-name typo as a function call.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    d = Date\n",
            "    dt = TypeName(Date)\n",
            "    t = Time\n",
            "    n = Now\n",
            "End Sub\n",
        ));
        let expected_serial = unix_epoch_days() as i64 + 25569;
        assert_eq!(vm.variables["d"], Variant::Date(expected_serial));
        assert_eq!(vm.variables["dt"], Variant::Str("Date".to_string()));
        assert!(matches!(vm.variables["t"], Variant::Float(_)));
        assert!(matches!(vm.variables["n"], Variant::Float(_)));
    }

    #[test]
    fn test_undefined_variable_typo_still_errors_after_date_now_time_fallback() {
        // The new "date"/"now"/"time" bare-identifier fallback must not
        // swallow a genuine undefined-variable typo into some other
        // behavior — anything else still hits the same error as before.
        let prog = parser::parse("Sub MySub()\n    x = someUndefinedVar\nEnd Sub\n").unwrap();
        let err = Vm::new().run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Undefined variable: 'someundefinedvar'");
    }

    #[test]
    fn test_vba_cbool() {
        // CBool used to be grouped with CLng/CInt and return a
        // Variant::Integer via numeric coercion — CBool("True") then tried
        // to parse "True" as a number and errored. It must return a real
        // Variant::Boolean, and a literal "True"/"False" string must not
        // go through numeric parsing at all.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    a = CBool(\"True\")\n",
            "    b = CBool(\"False\")\n",
            "    c = CBool(5)\n",
            "    d = CBool(0)\n",
            "    t = TypeName(CBool(5))\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Boolean(false));
        assert_eq!(vm.variables["c"], Variant::Boolean(true));
        assert_eq!(vm.variables["d"], Variant::Boolean(false));
        assert_eq!(vm.variables["t"], Variant::Str("Boolean".to_string()));
    }

    #[test]
    fn test_vba_len() {
        let vm = run("Sub MySub()\n    a = Len(\"Hello\")\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(5));
    }

    #[test]
    fn test_vba_str_reserves_a_leading_space_unlike_cstr() {
        // Real VBA's Str() reserves a leading space for the sign position
        // on a non-negative number (Str(459) is " 459") -- a real behavior
        // difference from CStr(459) == "459", not an alias of it. Both used
        // to share one arm with no space.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    a = Str(459)\n",
            "    b = Str(-459)\n",
            "    c = CStr(459)\n",
            "    d = Str(0)\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["a"], Variant::Str(" 459".to_string()));
        assert_eq!(vm.variables["b"], Variant::Str("-459".to_string()));
        assert_eq!(vm.variables["c"], Variant::Str("459".to_string()));
        assert_eq!(vm.variables["d"], Variant::Str(" 0".to_string()));
    }

    #[test]
    fn test_vba_mid_left_right() {
        let vm = run(
            "Sub MySub()\n    a = Mid(\"Hello\", 2, 3)\n    b = Left(\"Hello\", 3)\n    c = Right(\"Hello\", 2)\nEnd Sub\n",
        );
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

    #[test]
    fn test_vba_not_is_bitwise_on_numbers() {
        // Real VBA: `Not 5` is `-6` (bitwise complement), not `False` from a
        // truthy coercion — and combining it with the already-bitwise `And`
        // must round-trip to a consistent result: `Not 5 And 3` == 2.
        let vm = run("Sub MySub()\n    a = Not 5\n    b = Not 5 And 3\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(-6));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
    }

    // ── ElseIf ───────────────────────────────────────────────────────────────

    #[test]
    fn test_elseif_chain() {
        let vm = run(
            "Sub MySub()\n    x = 7\n    If x > 10 Then\n        r = 1\n    ElseIf x > 5 Then\n        r = 2\n    Else\n        r = 3\n    End If\nEnd Sub\n",
        );
        assert_eq!(vm.variables["r"], Variant::Integer(2));
    }

    #[test]
    fn test_single_line_if_end_to_end() {
        // The exact shape found unparseable in the 581-scenario VBA corpus
        // (arrays_0007..0010) while verifying the comma-Dim fix: a
        // single-line `If cond Then stmt` with no `End If`, used as a
        // running-max accumulator inside a loop.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    Dim arr(3) As Double\n",
            "    arr(0) = 2\n",
            "    arr(1) = 9\n",
            "    arr(2) = 4\n",
            "    arr(3) = 1\n",
            "    Dim i As Integer, m As Double\n",
            "    m = arr(0)\n",
            "    For i = 1 To 3\n",
            "        If arr(i) > m Then m = arr(i)\n",
            "    Next i\n",
            "    result = m\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["result"], Variant::Integer(9));
    }

    #[test]
    fn test_single_line_if_else_end_to_end() {
        let vm = run("Sub MySub()\n    x = 3\n    If x > 10 Then y = 1 Else y = 2\nEnd Sub\n");
        assert_eq!(vm.variables["y"], Variant::Integer(2));
    }

    #[test]
    fn test_single_line_if_exit_sub_actually_exits() {
        // Routing `Exit Sub` through the generic identifier-statement parser
        // (as if it were a bare, unrecognized identifier) would silently
        // no-op it instead of exiting — the statement after the single-line
        // If would then wrongly still run. `y` must stay unset.
        let vm = run("Sub MySub()\n    x = 1\n    If x > 0 Then Exit Sub\n    y = 99\nEnd Sub\n");
        assert!(!vm.variables.contains_key("y"));
    }

    #[test]
    fn test_single_line_if_exit_sub_then_else_on_one_line() {
        // `If cond Then Exit Sub Else stmt` — makes sure parse_exit's own
        // fixed-arity `consume_ident()` (for the Exit target) doesn't
        // swallow a following `Else` the way GoTo's label parse could.
        let exits = run(
            "Sub MySub()\n    x = 5\n    If x > 0 Then Exit Sub Else y = 1\n    y = 2\nEnd Sub\n",
        );
        assert!(!exits.variables.contains_key("y"));

        let takes_else = run(
            "Sub MySub()\n    x = -1\n    If x > 0 Then Exit Sub Else y = 1\n    z = y\nEnd Sub\n",
        );
        assert_eq!(takes_else.variables["y"], Variant::Integer(1));
        assert_eq!(takes_else.variables["z"], Variant::Integer(1));
    }

    #[test]
    fn test_single_line_if_supports_range_cells_and_msgbox_branches() {
        // Before parse_single_line_if_branch was refactored to share
        // parse_simple_stmt_no_eol with block-form parse_stmt, only
        // identifier-led statements were recognized inline -- `Range(...)`/
        // `Cells(...)`/`MsgBox`/etc. are their own dedicated keyword arms in
        // block-form VBA, not covered by parse_ident_stmt's generic
        // name(args)/name(args)=value dispatch. `Range("A1").Value = 1`
        // specifically mis-parsed as an array write to a variable literally
        // named "range" with index "A1", which then failed trying to
        // convert the string "A1" to a number -- found by
        // compat/vba-semantics/ on exactly this shape, not by source audit.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    x = -1\n",
            "    If x > 0 Then Exit Sub Else Range(\"A1\").Value = 1\n",
            "    If x < 0 Then Cells(1, 2).Value = 42\n",
            "End Sub\n",
        ));
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(1)); // A1
        assert_eq!(vm.get_cell(1, 2), Variant::Integer(42)); // B1
    }

    #[test]
    fn test_single_line_if_supports_with_dot_branches() {
        // parse_stmt gained a Tok::Dot arm (parse_with_dot_stmt) when the runtime With
        // stack replaced the old With-body-only parse-time special case, but
        // parse_single_line_if_branch's own dispatch only checked Tok::Ident and was never
        // updated to match -- so a bare `.member` branch inside a single-line If nested in
        // a With body (`If cond Then .Value = x`) silently degraded to Stmt::Unsupported:
        // no parse error, but the assignment never ran. Same bug *class* as the
        // Range()/Cells() gap above (a single-line-If branch dispatch lagging behind
        // block-form parse_stmt's own statement coverage), found by manual testing while
        // integrating the With-stack work, not by either subagent's own test suite.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    With Range(\"A1\")\n",
            "        .Value = 5\n",
            "        If .Value > 0 Then .Value = .Value + 1\n",
            "    End With\n",
            "End Sub\n",
        ));
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(6)); // A1
    }

    #[test]
    fn test_single_line_if_assigns_variable_named_after_a_block_keyword() {
        // parse_stmt's "bare `name = ...` is always assignment" override must fire before
        // the block-construct keyword dispatch even when reached via
        // parse_single_line_if_branch -> parse_simple_stmt_no_eol, not just via block-form
        // parse_stmt directly (the shape prop_vba_assignment_parses covers). `do`/`select`/
        // etc. as a variable name is unusual, but the single-line-If path shares the same
        // dispatch table, so a fix that only covers one caller and not the other would be
        // incomplete.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    x = 1\n",
            "    If x > 0 Then do = 5\n",
            "    Range(\"A1\").Value = do\n",
            "End Sub\n",
        ));
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(5)); // A1
    }

    #[test]
    fn test_boolean_true_is_negative_one_in_arithmetic() {
        // VBA represents True as -1 internally (CInt(True) = -1), distinct from Excel
        // worksheet formula semantics (TRUE = 1 in a =TRUE+1 cell formula, unrelated code
        // path in formula::eval, intentionally left untouched). Found via the
        // vba-semantics suite's operator-coercion matrix: to_f64 previously coerced
        // Variant::Boolean(true) to 1.0, giving True + 5 = 6 instead of VBA's documented 4.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    Range(\"A1\").Value = True + 5\n",
            "    Range(\"A2\").Value = CInt(True)\n",
            "End Sub\n",
        ));
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(4));
        assert_eq!(vm.get_cell(2, 1), Variant::Integer(-1));
    }

    #[test]
    fn test_worksheet_function_boolean_coercion_uses_excel_not_vba_semantics() {
        // WorksheetFunction.* bridges into Excel's own calculation engine, so its Boolean
        // coercion matches a worksheet formula (TRUE=1), not VBA's own True=-1 -- fixing
        // to_f64 for VBA's own arithmetic (see test_boolean_true_is_negative_one_in_arithmetic)
        // silently changed this too, since flat_nums/eval_wsf used to share that same
        // function. WorksheetFunction.Sum(True, True) must stay 2, not become -2.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    Range(\"A1\").Value = WorksheetFunction.Sum(True, True)\n",
            "    Range(\"A2\").Value = True + True\n",
            "End Sub\n",
        ));
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(2));
        assert_eq!(vm.get_cell(2, 1), Variant::Integer(-2));
    }

    #[test]
    fn test_empty_equals_zero_and_empty_string() {
        // Documented VBA equality rule: Empty numeric-compares as 0, string-compares as ""
        // -- vba_cmp (used for </>) already applied this via to_f64's Empty=>0.0 arm, but
        // vba_eq (used for =/<>) fell through to its catch-all `false` for e.g. `0 = Empty`,
        // an inconsistency between = and < on the exact same operand pair. Found via the
        // vba-semantics suite's comparison-coercion matrix.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    Dim r1, r2, r3\n",
            "    Range(\"A1\").Value = (0 = r1)\n",
            "    Range(\"A2\").Value = (r2 = \"\")\n",
            "    Range(\"A3\").Value = (1 = r3)\n",
            "End Sub\n",
        ));
        assert_eq!(vm.get_cell(1, 1), Variant::Boolean(true));
        assert_eq!(vm.get_cell(2, 1), Variant::Boolean(true));
        assert_eq!(vm.get_cell(3, 1), Variant::Boolean(false));
    }

    #[test]
    fn test_single_line_if_goto_actually_jumps() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    x = 1\n",
            "    If x > 0 Then GoTo Skip\n",
            "    y = 99\n",
            "Skip:\n",
            "    z = 1\n",
            "End Sub\n",
        ));
        assert!(!vm.variables.contains_key("y"));
        assert_eq!(vm.variables["z"], Variant::Integer(1));
    }

    // ── Exit ─────────────────────────────────────────────────────────────────

    #[test]
    fn test_exit_for() {
        let vm = run(
            "Sub MySub()\n    s = 0\n    For i = 1 To 10\n        If i > 3 Then\n            Exit For\n        End If\n        s = s + i\n    Next i\nEnd Sub\n",
        );
        assert_eq!(vm.variables["s"], Variant::Integer(6)); // 1+2+3
    }

    #[test]
    fn test_exit_do() {
        let vm = run(
            "Sub MySub()\n    x = 0\n    Do\n        x = x + 1\n        If x >= 5 Then\n            Exit Do\n        End If\n    Loop While x < 100\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    // ── On Error Resume Next ──────────────────────────────────────────────────

    #[test]
    fn test_on_error_resume_next() {
        let vm = run(
            "Sub MySub()\n    On Error Resume Next\n    a = 1\n    b = 2\n    a = 1\nEnd Sub\n",
        );
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
            "    x = 99\n",                // should be skipped
            "    Exit Sub\n",
            "ErrH:\n",
            "    handled = 1\n",
            "End Sub\n",
        );
        let vm = run(code);
        assert_eq!(vm.variables["x"], Variant::Integer(1)); // set before error
        assert!(!vm.variables.contains_key("x") || vm.variables["x"] != Variant::Integer(99)); // not 99
        assert_eq!(vm.variables["handled"], Variant::Integer(1)); // handler ran
    }

    // ── Err.Number / Err.Description / Err.Clear / Err.Raise ────────────────

    #[test]
    fn err_number_is_zero_before_any_error() {
        let vm = run("Sub MySub()\n    n = Err.Number\n    d = Err.Description\nEnd Sub\n");
        assert_eq!(vm.variables["n"], Variant::Integer(0));
        assert_eq!(vm.variables["d"], Variant::Str(String::new()));
    }

    #[test]
    fn on_error_resume_next_records_a_confirmed_vba_error_number() {
        // Division by zero is real VBA error 11 — see classify_vba_error_number.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    x = 1 / 0\n",
            "    n = Err.Number\n",
            "    d = Err.Description\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Integer(11));
        assert_eq!(vm.variables["d"], Variant::Str("Division by zero".into()));
    }

    #[test]
    fn err_clear_resets_every_err_property() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    Err.Raise 513, \"MySource\", \"custom text\", \"help.chm\", 100\n",
            "    Err.Clear\n",
            "    n = Err.Number\n",
            "    d = Err.Description\n",
            "    s = Err.Source\n",
            "    h = Err.HelpFile\n",
            "    c = Err.HelpContext\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Integer(0));
        assert_eq!(vm.variables["d"], Variant::Str(String::new()));
        assert_eq!(vm.variables["s"], Variant::Str(String::new()));
        assert_eq!(vm.variables["h"], Variant::Str(String::new()));
        assert_eq!(vm.variables["c"], Variant::Integer(0));
    }

    #[test]
    fn err_source_help_file_help_context_are_empty_zero_for_an_internally_raised_error() {
        // This project doesn't model a VBA project/class name, so an
        // error the VM itself raises (not via Err.Raise) never populates
        // Source/HelpFile/HelpContext with anything but their zero values.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    x = 1 / 0\n",
            "    s = Err.Source\n",
            "    h = Err.HelpFile\n",
            "    c = Err.HelpContext\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["s"], Variant::Str(String::new()));
        assert_eq!(vm.variables["h"], Variant::Str(String::new()));
        assert_eq!(vm.variables["c"], Variant::Integer(0));
    }

    #[test]
    fn err_raise_with_only_a_number_fills_in_the_default_description() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    Err.Raise 5\n",
            "    n = Err.Number\n",
            "    d = Err.Description\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Integer(5));
        assert_eq!(
            vm.variables["d"],
            Variant::Str("Invalid procedure call or argument".into())
        );
    }

    #[test]
    fn err_raise_skipped_source_does_not_get_read_as_description() {
        // `Err.Raise Number, , Description` — the idiomatic way to skip
        // Source. A positional-args-without-skip-support implementation
        // would misread "custom text" as Source, not Description.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    Err.Raise 513, , \"custom text\"\n",
            "    n = Err.Number\n",
            "    d = Err.Description\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Integer(513));
        assert_eq!(vm.variables["d"], Variant::Str("custom text".into()));
    }

    #[test]
    fn err_raise_with_explicit_source_and_description() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    Err.Raise 513, \"MySource\", \"custom text\"\n",
            "    n = Err.Number\n",
            "    d = Err.Description\n",
            "    s = Err.Source\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Integer(513));
        assert_eq!(vm.variables["d"], Variant::Str("custom text".into()));
        assert_eq!(vm.variables["s"], Variant::Str("MySource".into()));
    }

    #[test]
    fn err_raise_with_all_five_arguments_fills_in_every_err_property() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    Err.Raise 513, \"MySource\", \"custom text\", \"help.chm\", 100\n",
            "    n = Err.Number\n",
            "    d = Err.Description\n",
            "    s = Err.Source\n",
            "    h = Err.HelpFile\n",
            "    c = Err.HelpContext\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Integer(513));
        assert_eq!(vm.variables["d"], Variant::Str("custom text".into()));
        assert_eq!(vm.variables["s"], Variant::Str("MySource".into()));
        assert_eq!(vm.variables["h"], Variant::Str("help.chm".into()));
        assert_eq!(vm.variables["c"], Variant::Integer(100));
    }

    #[test]
    fn err_raise_skipped_help_file_does_not_get_read_as_help_context() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    Err.Raise 513, \"MySource\", \"custom text\", , 100\n",
            "    h = Err.HelpFile\n",
            "    c = Err.HelpContext\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["h"], Variant::Str(String::new()));
        assert_eq!(vm.variables["c"], Variant::Integer(100));
    }

    #[test]
    fn on_error_goto_handler_sees_err_number_from_the_error_that_triggered_it() {
        let code = concat!(
            "Sub MySub()\n",
            "    On Error GoTo ErrH\n",
            "    x = 1 / 0\n",
            "    Exit Sub\n",
            "ErrH:\n",
            "    n = Err.Number\n",
            "    d = Err.Description\n",
            "End Sub\n",
        );
        let vm = run(code);
        assert_eq!(vm.variables["n"], Variant::Integer(11));
        assert_eq!(vm.variables["d"], Variant::Str("Division by zero".into()));
    }

    #[test]
    fn err_raise_propagates_when_no_on_error_handler_is_active() {
        let prog = parser::parse("Sub MySub()\n    Err.Raise 513, , \"boom\"\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "boom");
    }

    #[test]
    fn a_raise_caught_inside_a_called_sub_does_not_leak_its_number_into_a_later_unrelated_error() {
        // `pending_raised_error` is a single Vm-wide slot (not scoped per
        // call frame, unlike `On Error`'s own mode — see `CallFrame`).
        // Helper gets a fresh, `Disabled` frame of its own (real VBA: error
        // handling doesn't inherit into a callee), so `Err.Raise 9` isn't
        // caught inside Helper's own body — it propagates out and is caught
        // by MySub's still-active `On Error Resume Next` at the `Call
        // Helper()` statement itself, in MySub's own frame. This confirms
        // `pending_raised_error` doesn't survive past the `Err.Raise` it
        // belongs to either way: the later, unrelated division error must
        // still report its own number (11), not 9.
        let prog = parser::parse(concat!(
            "Sub Helper()\n",
            "    Err.Raise 9\n",
            "End Sub\n",
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    Call Helper()\n",
            "    x = 1 / 0\n",
            "    n = Err.Number\n",
            "End Sub\n",
        ))
        .unwrap();
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.variables["n"], Variant::Integer(11));
    }

    #[test]
    fn on_error_resume_next_in_the_caller_does_not_resume_the_rest_of_a_failed_callees_body() {
        // Deliberate behavior change from the old Vm-wide `on_error_resume_next`
        // flag: previously the catch fired inside `exec_stmt`, *inside Child's
        // own body*, so Child's remaining statements kept running after the
        // error. Real VBA does not work that way — `On Error Resume Next` set
        // in MySub does not extend into Child's frame (Child has no handler of
        // its own), so the error propagates out of Child entirely and is only
        // caught at the `Call Child()` statement in MySub. Child's own
        // remaining statements after the failing line must NOT run.
        let vm = run(concat!(
            "Sub Child()\n",
            "    x = 1 / 0\n",
            "    aftertheerror = \"did it run\"\n",
            "End Sub\n",
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    Call Child()\n",
            "    n = Err.Number\n",
            "End Sub\n",
        ));
        assert!(!vm.variables.contains_key("aftertheerror"));
        assert_eq!(vm.variables["n"], Variant::Integer(11));
    }

    #[test]
    fn err_number_is_a_real_variable_named_err_is_still_a_plain_variable() {
        // `err` with no `.Number`/`.Description`/`.Clear`/`.Raise` suffix is
        // an ordinary user variable — the Err-object parsing only guards on
        // those exact member names, never a bare `err`.
        let vm = run("Sub MySub()\n    err = 42\nEnd Sub\n");
        assert_eq!(vm.variables["err"], Variant::Integer(42));
    }

    // ── Call-frame On Error scoping ──────────────────────────────────────
    // The bug this phase's call-frame rework fixes: `on_error_goto_label`
    // used to be a single Vm-wide field, so a callee's own `exec_body`
    // could see a caller's still-set label and try (and fail) to resolve
    // it against the callee's own body instead of letting the error
    // propagate to where the label actually lives.

    #[test]
    fn on_error_goto_in_a_caller_catches_an_error_raised_inside_a_called_sub() {
        let vm = run(concat!(
            "Sub Child()\n",
            "    x = 1 / 0\n",
            "End Sub\n",
            "Sub MySub()\n",
            "    On Error GoTo Handler\n",
            "    Call Child()\n",
            "    result = \"not reached\"\n",
            "    Exit Sub\n",
            "Handler:\n",
            "    n = Err.Number\n",
            "    d = Err.Description\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Integer(11));
        assert_eq!(vm.variables["d"], Variant::Str("Division by zero".into()));
        assert!(!vm.variables.contains_key("result"));
    }

    #[test]
    fn a_callees_own_handler_catches_its_own_error_without_involving_the_caller() {
        let vm = run(concat!(
            "Sub Child()\n",
            "    On Error GoTo ChildHandler\n",
            "    y = 1 / 0\n",
            "    Exit Sub\n",
            "ChildHandler:\n",
            "    childcaught = Err.Number\n",
            "End Sub\n",
            "Sub MySub()\n",
            "    On Error GoTo Handler\n",
            "    Call Child()\n",
            "    parentreached = True\n",
            "    Exit Sub\n",
            "Handler:\n",
            "    parentcaught = True\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["childcaught"], Variant::Integer(11));
        assert_eq!(vm.variables["parentreached"], Variant::Boolean(true));
        assert!(!vm.variables.contains_key("parentcaught"));
    }

    #[test]
    fn a_second_failure_inside_a_callees_own_handler_propagates_to_the_caller() {
        // A GoTo handler is consumed the moment it fires (real VBA: without
        // a fresh On Error inside the handler itself, a second failure
        // there isn't caught again by the same handler).
        let vm = run(concat!(
            "Sub Child()\n",
            "    On Error GoTo ChildHandler\n",
            "    y = 1 / 0\n",
            "    Exit Sub\n",
            "ChildHandler:\n",
            "    z = 1 / 0\n",
            "    afterseconderror = \"not reached\"\n",
            "End Sub\n",
            "Sub MySub()\n",
            "    On Error GoTo Handler\n",
            "    Call Child()\n",
            "    result = \"not reached\"\n",
            "    Exit Sub\n",
            "Handler:\n",
            "    n = Err.Number\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Integer(11));
        assert!(!vm.variables.contains_key("afterseconderror"));
        assert!(!vm.variables.contains_key("result"));
    }

    #[test]
    fn on_error_goto_0_inside_a_callee_disables_only_that_callees_own_frame() {
        let vm = run(concat!(
            "Sub Child()\n",
            "    On Error GoTo ChildHandler\n",
            "    On Error GoTo 0\n",
            "    y = 1 / 0\n",
            "    Exit Sub\n",
            "ChildHandler:\n",
            "    childcaught = True\n",
            "End Sub\n",
            "Sub MySub()\n",
            "    On Error GoTo Handler\n",
            "    Call Child()\n",
            "    result = \"not reached\"\n",
            "    Exit Sub\n",
            "Handler:\n",
            "    n = Err.Number\n",
            "End Sub\n",
        ));
        assert!(!vm.variables.contains_key("childcaught"));
        assert_eq!(vm.variables["n"], Variant::Integer(11));
    }

    #[test]
    fn on_error_resume_next_does_not_leak_into_a_sibling_procedure() {
        let prog = parser::parse(concat!(
            "Sub SiblingA()\n",
            "    On Error Resume Next\n",
            "    a = 1 / 0\n",
            "    acaught = Err.Number\n",
            "End Sub\n",
            "Sub SiblingB()\n",
            "    b = 1 / 0\n",
            "End Sub\n",
            "Sub MySub()\n",
            "    Call SiblingA()\n",
            "    Call SiblingB()\n",
            "End Sub\n",
        ))
        .unwrap();
        let mut vm = Vm::new();
        // If SiblingA's Resume Next had leaked into SiblingB, this would
        // succeed instead (SiblingB's division silently swallowed too).
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Division by zero");
    }

    #[test]
    fn recursive_calls_do_not_share_an_on_error_mode_across_frames() {
        let prog = parser::parse(concat!(
            "Sub Recur(n)\n",
            "    If n = 0 Then\n",
            "        On Error Resume Next\n",
            "        z = 1 / 0\n",
            "        zcaught = Err.Number\n",
            "    Else\n",
            "        Call Recur(n - 1)\n",
            "        y = 1 / 0\n",
            "    End If\n",
            "End Sub\n",
            "Sub MySub()\n",
            "    Call Recur(2)\n",
            "End Sub\n",
        ))
        .unwrap();
        let mut vm = Vm::new();
        // If n=0's Resume Next had leaked into n=1's own frame, n=1's own
        // `y = 1 / 0` would be silently swallowed too and the whole call
        // would succeed instead of propagating.
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Division by zero");
    }

    #[test]
    fn a_raise_with_all_five_arguments_inside_a_called_sub_reaches_the_callers_handler_intact() {
        let vm = run(concat!(
            "Sub Child()\n",
            "    Err.Raise 513, \"MySource\", \"custom text\", \"help.chm\", 100\n",
            "End Sub\n",
            "Sub MySub()\n",
            "    On Error GoTo Handler\n",
            "    Call Child()\n",
            "    Exit Sub\n",
            "Handler:\n",
            "    n = Err.Number\n",
            "    d = Err.Description\n",
            "    s = Err.Source\n",
            "    h = Err.HelpFile\n",
            "    c = Err.HelpContext\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Integer(513));
        assert_eq!(vm.variables["d"], Variant::Str("custom text".into()));
        assert_eq!(vm.variables["s"], Variant::Str("MySource".into()));
        assert_eq!(vm.variables["h"], Variant::Str("help.chm".into()));
        assert_eq!(vm.variables["c"], Variant::Integer(100));
    }

    #[test]
    fn a_function_call_propagates_its_error_to_the_caller_the_same_way_a_sub_call_does() {
        let vm = run(concat!(
            "Function Child()\n",
            "    x = 1 / 0\n",
            "End Function\n",
            "Sub MySub()\n",
            "    On Error GoTo Handler\n",
            "    y = Child()\n",
            "    result = \"not reached\"\n",
            "    Exit Sub\n",
            "Handler:\n",
            "    n = Err.Number\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["n"], Variant::Integer(11));
        assert!(!vm.variables.contains_key("result"));
    }

    #[test]
    fn test_goto_unconditional() {
        let code = concat!(
            "Sub MySub()\n",
            "    a = 1\n",
            "    GoTo Skip\n",
            "    a = 99\n", // should be skipped
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
            "    result = p.y\n", // p.y not set → Empty → 0 in arithmetic
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
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    total = 0\n    For Each cell In Range(\"A1:A3\")\n        total = total + cell\n    Next cell\nEnd Sub\n",
        );
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
        let vm = run(
            "Function Square(n)\n    Square = n * n\nEnd Function\nSub MySub()\n    result = Square(7)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["result"], Variant::Integer(49));
    }

    #[test]
    fn test_function_return_value_nested() {
        let vm = run(
            "Function Add(a, b)\n    Add = a + b\nEnd Function\nSub MySub()\n    x = Add(3, 4) + Add(1, 2)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(10));
    }

    #[test]
    fn test_function_recursive() {
        // Factorial: 5! = 120
        let vm = run(
            "Function Fact(n)\n    If n <= 1 Then\n        Fact = 1\n    Else\n        Fact = n * Fact(n - 1)\n    End If\nEnd Function\nSub MySub()\n    result = Fact(5)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["result"], Variant::Integer(120));
    }

    #[test]
    fn test_function_in_cell_write() {
        let vm = run(
            "Function Double(x)\n    Double = x * 2\nEnd Function\nSub MySub()\n    Cells(1, 1).Value = Double(21)\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
    }

    #[test]
    fn test_call_sub_with_args() {
        let prog = parser::parse("Sub FillRow(rowNum, val)\n    Cells(rowNum, 1).Value = val\nEnd Sub\nSub MySub()\n    Call FillRow(3, 99)\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(99));
    }

    // ── Phase 2C item 6: typed Function parameters and return type ───────────

    #[test]
    fn typed_function_params_and_return_type_parse_and_execute() {
        // `Function DoubleIt(x As Integer) As Integer` — previously failed
        // to parse ("expected newline, got Ident(\"as\")") at the return-
        // type annotation; typed params alone already worked.
        let prog = parser::parse(
            "Function DoubleIt(x As Integer) As Integer\n    DoubleIt = x * 2\nEnd Function\nSub MySub()\n    result = DoubleIt(21)\nEnd Sub\n",
        )
        .unwrap();
        assert_eq!(prog.funcs[0].params, vec!["x"]);
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.variables["result"], Variant::Integer(42));
    }

    #[test]
    fn typed_function_multiple_params_and_double_return_type() {
        let vm = run(
            "Function Helper(x As Double, y As Double) As Double\n    Helper = x * x + y\nEnd Function\nSub MySub()\n    Range(\"A1\").Value = Helper(2, 3)\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(7));
    }

    // ── vb constants ─────────────────────────────────────────────────────────

    #[test]
    fn test_vb_string_constants() {
        let vm = run(
            "Sub MySub()\n    a = \"Hello\" & vbCrLf & \"World\"\n    b = \"tab\" & vbTab & \"here\"\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Str("Hello\r\nWorld".into()));
        assert_eq!(vm.variables["b"], Variant::Str("tab\there".into()));
    }

    // ── While ... Wend ───────────────────────────────────────────────────────

    #[test]
    fn test_while_wend() {
        let vm =
            run("Sub MySub()\n    x = 0\n    While x < 5\n        x = x + 1\n    Wend\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test]
    fn test_while_wend_no_iteration() {
        let vm = run(
            "Sub MySub()\n    x = 10\n    While x < 5\n        x = x + 1\n    Wend\n    y = 99\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(10));
        assert_eq!(vm.variables["y"], Variant::Integer(99));
    }

    // ── Const ─────────────────────────────────────────────────────────────────

    #[test]
    fn test_const_declaration() {
        let vm = run("Sub MySub()\n    Const MAX_ROW As Long = 100\n    x = MAX_ROW\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(100));
    }

    #[test]
    fn test_const_string() {
        let vm =
            run("Sub MySub()\n    Const PREFIX = \"ID_\"\n    s = PREFIX & \"001\"\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("ID_001".into()));
    }

    // ── Empty / Null / Nothing ────────────────────────────────────────────────

    #[test]
    fn test_empty_literal() {
        // `Null` is no longer folded into `Empty` — they are genuinely
        // different VBA values (see `Variant::Null`), which is what makes
        // every documented Null-propagation rule expressible. `Nothing`
        // stays Empty: it's the null *object* reference, tracked in
        // `object_variables`, not a `Variant`.
        let vm = run("Sub MySub()\n    a = Empty\n    b = Null\n    c = Nothing\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Empty);
        assert_eq!(vm.variables["b"], Variant::Null);
        assert_eq!(vm.variables["c"], Variant::Empty);
    }

    // ── IIf ──────────────────────────────────────────────────────────────────

    #[test]
    fn test_iif_true() {
        let vm = run("Sub MySub()\n    x = IIf(1 > 0, \"yes\", \"no\")\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Str("yes".into()));
    }

    #[test]
    fn test_iif_false() {
        let vm = run("Sub MySub()\n    x = IIf(0 > 1, 10, 20)\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(20));
    }

    // ── Format ───────────────────────────────────────────────────────────────

    #[test]
    fn test_format_decimal() {
        let vm = run("Sub MySub()\n    s = Format(3.14159, \"0.00\")\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("3.14".into()));
    }

    #[test]
    fn test_format_integer_no_dec() {
        let vm = run("Sub MySub()\n    s = Format(42, \"0\")\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("42".into()));
    }

    // ── TypeName / VarType ───────────────────────────────────────────────────

    #[test]
    fn test_typename() {
        let vm = run(
            "Sub MySub()\n    a = TypeName(42)\n    b = TypeName(\"hi\")\n    c = TypeName(True)\n    d = TypeName(Empty)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Str("Long".into()));
        assert_eq!(vm.variables["b"], Variant::Str("String".into()));
        assert_eq!(vm.variables["c"], Variant::Str("Boolean".into()));
        assert_eq!(vm.variables["d"], Variant::Str("Empty".into()));
    }

    // ── Arrays ───────────────────────────────────────────────────────────────

    #[test]
    fn test_dim_array_write_read() {
        let vm = run(
            "Sub MySub()\n    Dim arr(5)\n    arr(0) = 10\n    arr(3) = 99\n    a = arr(0)\n    b = arr(3)\n    c = arr(1)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(10));
        assert_eq!(vm.variables["b"], Variant::Integer(99));
        assert_eq!(vm.variables["c"], Variant::Empty);
    }

    #[test]
    fn test_dim_array_loop() {
        let vm = run(
            "Sub MySub()\n    Dim arr(4)\n    For i = 0 To 4\n        arr(i) = i * 2\n    Next i\n    s = 0\n    For i = 0 To 4\n        s = s + arr(i)\n    Next i\nEnd Sub\n",
        );
        assert_eq!(vm.variables["s"], Variant::Integer(20)); // 0+2+4+6+8
    }

    #[test]
    fn test_redim_preserve() {
        let vm = run(
            "Sub MySub()\n    Dim arr(2)\n    arr(0) = 1\n    arr(1) = 2\n    ReDim Preserve arr(4)\n    arr(3) = 99\n    a = arr(0)\n    b = arr(3)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(99));
    }

    #[test]
    fn test_ubound_lbound() {
        let vm =
            run("Sub MySub()\n    Dim arr(9)\n    u = UBound(arr)\n    l = LBound(arr)\nEnd Sub\n");
        assert_eq!(vm.variables["u"], Variant::Integer(9));
        assert_eq!(vm.variables["l"], Variant::Integer(0));
    }

    #[test]
    fn dim_array_explicit_lower_bound_is_honored() {
        let vm = run(
            "Sub MySub()\n    Dim arr(2 To 8)\n    l = LBound(arr)\n    u = UBound(arr)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["l"], Variant::Integer(2));
        assert_eq!(vm.variables["u"], Variant::Integer(8));
    }

    #[test]
    fn dim_array_explicit_lower_bound_indices_read_and_write_at_their_real_positions() {
        let vm = run(
            "Sub MySub()\n    Dim arr(2 To 4)\n    arr(2) = 10\n    arr(4) = 30\n    a = arr(2)\n    b = arr(4)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(10));
        assert_eq!(vm.variables["b"], Variant::Integer(30));
    }

    #[test]
    fn dim_array_index_below_an_explicit_lower_bound_is_out_of_range() {
        let prog =
            parser::parse("Sub MySub()\n    Dim arr(2 To 4)\n    arr(1) = 1\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");
    }

    #[test]
    fn option_base_one_shifts_the_default_lower_bound() {
        let vm = run(
            "Option Base 1\nSub MySub()\n    Dim arr(5)\n    l = LBound(arr)\n    u = UBound(arr)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["l"], Variant::Integer(1));
        assert_eq!(vm.variables["u"], Variant::Integer(5));
    }

    #[test]
    fn option_base_one_does_not_override_an_explicit_lower_bound() {
        let vm =
            run("Option Base 1\nSub MySub()\n    Dim arr(0 To 5)\n    l = LBound(arr)\nEnd Sub\n");
        assert_eq!(vm.variables["l"], Variant::Integer(0));
    }

    #[test]
    fn dim_array_empty_parens_declares_a_dynamic_array_sizable_by_redim() {
        let vm = run(
            "Sub MySub()\n    Dim arr()\n    ReDim arr(5)\n    arr(3) = 99\n    u = UBound(arr)\n    v = arr(3)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["u"], Variant::Integer(5));
        assert_eq!(vm.variables["v"], Variant::Integer(99));
    }

    #[test]
    fn erase_on_a_fixed_array_resets_elements_to_empty_but_keeps_its_bounds() {
        let vm = run(
            "Sub MySub()\n    Dim arr(3)\n    arr(0) = 5\n    arr(1) = 10\n    Erase arr\n    e = IsEmpty(arr(0))\n    u = UBound(arr)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["e"], Variant::Boolean(true));
        assert_eq!(vm.variables["u"], Variant::Integer(3));
    }

    #[test]
    fn erase_on_an_explicit_lower_bound_array_preserves_that_bound_too() {
        let vm = run(
            "Sub MySub()\n    Dim arr(2 To 4)\n    arr(2) = 5\n    Erase arr\n    l = LBound(arr)\n    e = IsEmpty(arr(2))\nEnd Sub\n",
        );
        assert_eq!(vm.variables["l"], Variant::Integer(2));
        assert_eq!(vm.variables["e"], Variant::Boolean(true));
    }

    #[test]
    fn test_split_join() {
        let vm = run(
            "Sub MySub()\n    arr = Split(\"a,b,c\", \",\")\n    s = Join(arr, \"-\")\n    n = UBound(arr)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["s"], Variant::Str("a-b-c".into()));
        assert_eq!(vm.variables["n"], Variant::Integer(2));
    }

    #[test]
    fn test_isarray() {
        let vm = run(
            "Sub MySub()\n    Dim arr(3)\n    a = IsArray(arr)\n    b = IsArray(42)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Boolean(false));
    }

    // ── Multi-dimensional VBA arrays ────────────────────────────────────

    #[test]
    fn two_dimensional_array_second_index_selects_a_genuinely_distinct_element() {
        // The bug this phase fixes: elixcee's array storage used to be
        // genuinely 1-D, so `arr(2,0)` and `arr(2,1)` silently collapsed
        // onto the same internal slot.
        let vm = run(
            "Sub MySub()\n    Dim arr(3, 2)\n    arr(2, 0) = 111\n    arr(2, 1) = 222\n\
             a = arr(2, 0)\n    b = arr(2, 1)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(111));
        assert_eq!(vm.variables["b"], Variant::Integer(222));
    }

    #[test]
    fn lbound_ubound_report_each_dimension_of_a_two_dimensional_array_independently() {
        let vm = run("Sub MySub()\n    Dim arr(1 To 3, -2 To 2)\n\
             a = LBound(arr, 1)\n    b = UBound(arr, 1)\n\
             c = LBound(arr, 2)\n    d = UBound(arr, 2)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(3));
        assert_eq!(vm.variables["c"], Variant::Integer(-2));
        assert_eq!(vm.variables["d"], Variant::Integer(2));
    }

    #[test]
    fn ubound_dimension_zero_is_subscript_out_of_range() {
        let prog =
            parser::parse("Sub MySub()\n    Dim arr(3, 2)\n    u = UBound(arr, 0)\nEnd Sub\n")
                .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");
    }

    #[test]
    fn ubound_of_a_dimension_beyond_the_arrays_rank_is_subscript_out_of_range() {
        let prog =
            parser::parse("Sub MySub()\n    Dim arr(3, 2)\n    u = UBound(arr, 3)\nEnd Sub\n")
                .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");
    }

    #[test]
    fn too_few_subscripts_on_a_two_dimensional_array_is_not_silently_accepted() {
        let prog =
            parser::parse("Sub MySub()\n    Dim arr(3, 2)\n    a = arr(1)\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");
    }

    #[test]
    fn too_many_subscripts_on_a_two_dimensional_array_is_not_silently_accepted() {
        let prog = parser::parse("Sub MySub()\n    Dim arr(3, 2)\n    a = arr(1, 1, 1)\nEnd Sub\n")
            .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");
    }

    #[test]
    fn array_out_of_bounds_evidence_on_a_two_dimensional_array_reports_the_failing_dimensions_own_bounds()
     {
        // Regression for a zip/enumerate mixup: the evidence must name
        // dimension 2's bounds (0..2), not dimension 1's (0..3), since
        // dimension 2's index (-1) is the one that's actually out of range.
        let prog =
            parser::parse("Sub MySub()\n    Dim arr(3, 2)\n    arr(0, -1) = 1\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::ArrayIndexOutOfBounds {
                name,
                index,
                lower,
                upper,
            }) => {
                assert_eq!(name, "arr");
                assert_eq!(index, -1);
                assert_eq!(lower, 0);
                assert_eq!(upper, 2);
            }
            other => panic!("expected ArrayIndexOutOfBounds, got {:?}", other),
        }
    }

    #[test]
    fn an_out_of_range_index_on_either_dimension_of_a_two_dimensional_array_errors() {
        let prog =
            parser::parse("Sub MySub()\n    Dim arr(3, 2)\n    arr(4, 0) = 1\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");

        let prog2 =
            parser::parse("Sub MySub()\n    Dim arr(3, 2)\n    arr(0, -1) = 1\nEnd Sub\n").unwrap();
        let mut vm2 = Vm::new();
        let err2 = vm2.run_sub(&prog2, "mysub").unwrap_err();
        assert_eq!(err2, "Subscript out of range");
    }

    #[test]
    fn option_base_one_shifts_every_dimensions_default_lower_bound() {
        let vm = run("Option Base 1\nSub MySub()\n    Dim arr(3, 2)\n\
             a = LBound(arr, 1)\n    b = LBound(arr, 2)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(1));
    }

    #[test]
    fn three_dimensional_array_reads_and_writes_at_their_real_positions() {
        let vm = run(
            "Sub MySub()\n    Dim arr(1, 1, 1)\n    arr(0, 0, 0) = 1\n    arr(1, 1, 1) = 8\n\
             a = arr(0, 0, 0)\n    b = arr(1, 1, 1)\n    c = arr(0, 1, 0)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(8));
        // Never written — real VBA's default Empty, not a collision with
        // one of the two writes above.
        assert_eq!(vm.variables["c"], Variant::Empty);
    }

    #[test]
    fn redim_preserve_on_a_two_dimensional_array_keeps_every_element_at_its_own_subscript() {
        let vm = run(
            "Sub MySub()\n    Dim arr(1, 1)\n    arr(0, 0) = 1\n    arr(0, 1) = 2\n\
             arr(1, 0) = 3\n    arr(1, 1) = 4\n    ReDim Preserve arr(1, 3)\n\
             a = arr(0, 0)\n    b = arr(1, 1)\n    c = arr(1, 3)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(4));
        assert_eq!(vm.variables["c"], Variant::Empty);
    }

    #[test]
    fn redim_preserve_changing_a_non_last_dimension_of_a_two_dimensional_array_errors() {
        let prog = parser::parse(
            "Sub MySub()\n    Dim arr(1, 1)\n    ReDim Preserve arr(3, 1)\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");
    }

    #[test]
    fn redim_preserve_changing_the_last_dimensions_lower_bound_errors_even_at_rank_1() {
        // Real VBA's `Preserve` only grows/shrinks the last dimension's
        // *upper* bound — its lower bound (and every other dimension's
        // bounds entirely) must stay exactly what they were, or an
        // existing element has no well-defined subscript to land at.
        let prog = parser::parse(
            "Sub MySub()\n    Dim arr(0 To 5)\n    ReDim Preserve arr(2 To 8)\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");
    }

    #[test]
    fn erase_on_a_two_dimensional_array_resets_elements_but_keeps_every_dimensions_bounds() {
        let vm = run(
            "Sub MySub()\n    Dim arr(1 To 2, 1 To 2)\n    arr(1, 1) = 9\n    Erase arr\n\
             e = IsEmpty(arr(1, 1))\n    l1 = LBound(arr, 1)\n    l2 = LBound(arr, 2)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["e"], Variant::Boolean(true));
        assert_eq!(vm.variables["l1"], Variant::Integer(1));
        assert_eq!(vm.variables["l2"], Variant::Integer(1));
    }

    #[test]
    fn a_dynamic_array_before_its_first_redim_errors_on_any_subscript_access() {
        let prog = parser::parse("Sub MySub()\n    Dim arr()\n    a = arr(0)\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");
    }

    #[test]
    fn an_absurdly_huge_two_dimensional_dim_is_rejected_as_out_of_memory_not_a_hang_or_crash() {
        let prog =
            parser::parse("Sub MySub()\n    Dim arr(2000000000, 2000000000)\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Out of memory");
    }

    #[test]
    fn a_dimension_product_that_overflows_i64_is_also_rejected_as_out_of_memory() {
        let prog = parser::parse(&format!(
            "Sub MySub()\n    Dim arr({}, {})\nEnd Sub\n",
            i64::MAX / 2,
            i64::MAX / 2
        ))
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Out of memory");
    }

    #[test]
    fn array_function_and_split_still_produce_a_plain_one_dimensional_zero_based_array() {
        // Regression: `Array()`/`Split()` migrated to `Variant::VbaArray`
        // alongside real `Dim` arrays — this pins their externally
        // observable shape (0-based, rank 1) stays exactly what it was.
        let vm = run(
            "Sub MySub()\n    a = Array(10, 20, 30)\n    l = LBound(a)\n    u = UBound(a)\n\
             v = a(1)\n    s = Split(\"x,y\", \",\")\n    su = UBound(s)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["l"], Variant::Integer(0));
        assert_eq!(vm.variables["u"], Variant::Integer(2));
        assert_eq!(vm.variables["v"], Variant::Integer(20));
        assert_eq!(vm.variables["su"], Variant::Integer(1));
    }

    #[test]
    fn a_two_dimensional_array_assigned_to_another_variable_keeps_its_shape() {
        let vm = run(
            "Sub MySub()\n    Dim arr(1, 1)\n    arr(1, 1) = 42\n    other = arr\n\
             v = other(1, 1)\n    u = UBound(other, 2)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["v"], Variant::Integer(42));
        assert_eq!(vm.variables["u"], Variant::Integer(1));
    }

    #[test]
    fn a_two_dimensional_array_passed_to_a_function_and_returned_keeps_its_shape() {
        let vm = run("Function Echo(a)\n    Echo = a\nEnd Function\n\
             Sub MySub()\n    Dim arr(1, 1)\n    arr(0, 1) = 7\n    r = Echo(arr)\n\
             v = r(0, 1)\n    u2 = UBound(r, 2)\nEnd Sub\n");
        assert_eq!(vm.variables["v"], Variant::Integer(7));
        assert_eq!(vm.variables["u2"], Variant::Integer(1));
    }

    #[test]
    fn test_isnumeric_accepts_numeric_strings_not_just_numeric_variants() {
        // Used to only check the Variant's own type (Integer/Float),
        // missing real VBA's IsNumeric("123") == True entirely.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    a = IsNumeric(\"123\")\n",
            "    b = IsNumeric(\"12.5\")\n",
            "    c = IsNumeric(\" 42 \")\n",
            "    d = IsNumeric(\"abc\")\n",
            "    e = IsNumeric(\"12abc\")\n",
            "    f = IsNumeric(123)\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Boolean(true));
        assert_eq!(vm.variables["c"], Variant::Boolean(true));
        assert_eq!(vm.variables["d"], Variant::Boolean(false));
        assert_eq!(vm.variables["e"], Variant::Boolean(false));
        assert_eq!(vm.variables["f"], Variant::Boolean(true));
    }

    #[test]
    fn test_vba_val_parses_leading_numeric_prefix() {
        // Real VBA's Val() parses a leading numeric prefix and stops at
        // the first character that doesn't fit -- Val("123abc") is 123,
        // not 0. Used to require the entire string to parse as f64, so
        // any trailing non-numeric character silently zeroed the result.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    a = Val(\"123abc\")\n",
            "    b = Val(\"  42.5xyz\")\n",
            "    c = Val(\"abc\")\n",
            "    d = Val(\"\")\n",
            "    e = Val(\"-5.5xyz\")\n",
            "    f = Val(\".5\")\n",
            "    g = Val(\"5.\")\n",
            "    h = Val(\"5\")\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["a"], Variant::Integer(123));
        assert_eq!(vm.variables["b"], Variant::Float(42.5));
        assert_eq!(vm.variables["c"], Variant::Integer(0));
        assert_eq!(vm.variables["d"], Variant::Integer(0));
        assert_eq!(vm.variables["e"], Variant::Float(-5.5));
        assert_eq!(vm.variables["f"], Variant::Float(0.5));
        assert_eq!(vm.variables["g"], Variant::Integer(5));
        assert_eq!(vm.variables["h"], Variant::Integer(5));
    }

    // ── Application properties ────────────────────────────────────────────────

    #[test]
    fn test_app_prop_noop() {
        let vm = run(
            "Sub MySub()\n    Application.ScreenUpdating = False\n    Application.EnableEvents = False\n    x = 1\n    Application.ScreenUpdating = True\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(1));
    }

    #[test]
    fn test_xl_constants() {
        let vm = run(
            "Sub MySub()\n    a = xlUp\n    b = xlDown\n    c = xlCalculationManual\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(-4162));
        assert_eq!(vm.variables["b"], Variant::Integer(-4121));
        assert_eq!(vm.variables["c"], Variant::Integer(-4135));
    }

    // ── Range write / copy / read ─────────────────────────────────────────────

    #[test]
    fn test_range_write_value() {
        let vm = run(
            "Sub MySub()\n    Range(\"A1\").Value = 42\n    Range(\"B2\").Value = \"hello\"\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
        assert_eq!(vm.get_cell(2, 2), Variant::Str("hello".into()));
    }

    #[test]
    fn test_range_write_formula() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(1,2).Value = 20\n    Range(\"C1\").Formula = \"=SUM(A1:B1)\"\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 3), Variant::Integer(30));
    }

    #[test]
    fn test_range_copy() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    Range(\"A1:A3\").Copy Destination:=Range(\"B1\")\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 2), Variant::Integer(10));
        assert_eq!(vm.get_cell(2, 2), Variant::Integer(20));
        assert_eq!(vm.get_cell(3, 2), Variant::Integer(30));
    }

    #[test]
    fn test_range_read_expr() {
        let vm =
            run("Sub MySub()\n    Cells(5,1).Value = 99\n    x = Range(\"A5\").Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(99));
    }

    // ── WorksheetFunction ────────────────────────────────────────────────────

    #[test]
    fn test_wsf_sum() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    s = WorksheetFunction.Sum(Range(\"A1:A3\"))\nEnd Sub\n",
        );
        assert_eq!(vm.variables["s"], Variant::Integer(60));
    }

    #[test]
    fn test_wsf_max_min() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 5\n    Cells(2,1).Value = 3\n    Cells(3,1).Value = 8\n    mx = WorksheetFunction.Max(Range(\"A1:A3\"))\n    mn = WorksheetFunction.Min(Range(\"A1:A3\"))\nEnd Sub\n",
        );
        assert_eq!(vm.variables["mx"], Variant::Integer(8));
        assert_eq!(vm.variables["mn"], Variant::Integer(3));
    }

    #[test]
    fn test_wsf_average() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    av = WorksheetFunction.Average(Range(\"A1:A2\"))\nEnd Sub\n",
        );
        assert_eq!(vm.variables["av"], Variant::Integer(15));
    }

    #[test]
    fn test_wsf_countif() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 5\n    Cells(2,1).Value = 10\n    Cells(3,1).Value = 3\n    n = WorksheetFunction.CountIf(Range(\"A1:A3\"), \">4\")\nEnd Sub\n",
        );
        assert_eq!(vm.variables["n"], Variant::Integer(2));
    }

    #[test]
    fn test_wsf_sumif() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 5\n    Cells(2,1).Value = 10\n    Cells(3,1).Value = 3\n    s = WorksheetFunction.SumIf(Range(\"A1:A3\"), \">4\")\nEnd Sub\n",
        );
        assert_eq!(vm.variables["s"], Variant::Integer(15));
    }

    #[test]
    fn test_wsf_application_prefix() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 7\n    Cells(2,1).Value = 3\n    s = Application.WorksheetFunction.Sum(Range(\"A1:A2\"))\nEnd Sub\n",
        );
        assert_eq!(vm.variables["s"], Variant::Integer(10));
    }

    #[test]
    fn test_wsf_match() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = \"a\"\n    Cells(2,1).Value = \"b\"\n    Cells(3,1).Value = \"c\"\n    pos = WorksheetFunction.Match(\"b\", Range(\"A1:A3\"), 0)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["pos"], Variant::Integer(2));
    }

    // ── Range("A1:A10").Value 多セル読み取り ─────────────────────────────────

    #[test]
    fn test_range_read_multi_cell() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    x = Range(\"A1:A3\").Value\nEnd Sub\n",
        );
        assert_eq!(
            vm.variables["x"],
            Variant::Array(vec![
                Variant::Integer(1),
                Variant::Integer(2),
                Variant::Integer(3)
            ])
        );
    }

    #[test]
    fn test_range_read_single_cell_backward_compat() {
        let vm =
            run("Sub MySub()\n    Cells(5,1).Value = 99\n    x = Range(\"A5\").Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(99));
    }

    #[test]
    fn test_range_read_2d_row_major() {
        // A1=1, B1=2, A2=3, B2=4 → Range("A1:B2").Value → [1,2,3,4]
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(1,2).Value = 2\n    Cells(2,1).Value = 3\n    Cells(2,2).Value = 4\n    arr = Range(\"A1:B2\").Value\n    a = arr(0)\n    b = arr(1)\n    c = arr(2)\n    d = arr(3)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
        assert_eq!(vm.variables["c"], Variant::Integer(3));
        assert_eq!(vm.variables["d"], Variant::Integer(4));
    }

    // ── Cells.Find ───────────────────────────────────────────────────────────

    #[test]
    fn test_cells_find_row() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = \"apple\"\n    Cells(2,1).Value = \"banana\"\n    Cells(3,1).Value = \"cherry\"\n    r = Cells.Find(What:=\"banana\").Row\nEnd Sub\n",
        );
        assert_eq!(vm.variables["r"], Variant::Integer(2));
    }

    #[test]
    fn test_cells_find_column() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = \"x\"\n    Cells(1,2).Value = \"y\"\n    Cells(1,3).Value = \"z\"\n    c = Cells.Find(What:=\"y\").Column\nEnd Sub\n",
        );
        assert_eq!(vm.variables["c"], Variant::Integer(2));
    }

    #[test]
    fn test_cells_find_not_found() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = \"a\"\n    r = Cells.Find(What:=\"missing\").Row\nEnd Sub\n",
        );
        assert_eq!(vm.variables["r"], Variant::Integer(0));
    }

    #[test]
    fn test_cells_find_extra_kwargs() {
        let vm = run(
            "Sub MySub()\n    Cells(2,1).Value = 42\n    r = Cells.Find(What:=42, LookIn:=xlValues, SearchDirection:=xlPrevious).Row\nEnd Sub\n",
        );
        assert_eq!(vm.variables["r"], Variant::Integer(2));
    }

    // ── EntireRow / EntireColumn ──────────────────────────────────────────────

    #[test]
    fn test_entirerow_delete() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    Range(\"A2:A2\").EntireRow.Delete\n    x = Cells(2,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(3)); // 3 が行2に移動
    }

    #[test]
    fn test_entirerow_clear() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 99\n    Range(\"A1\").EntireRow.Clear\n    x = Cells(1,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Empty);
    }

    #[test]
    fn test_entirerow_clear_contents() {
        let vm = run(
            "Sub MySub()\n    Cells(2,1).Value = 55\n    Range(\"A2\").EntireRow.ClearContents\n    x = Cells(2,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Empty);
    }

    #[test]
    fn test_entirecolumn_delete() {
        // GitHub #8: EntireColumn.Delete used to delete the ROW the reference cell was
        // in (RangeDelete ignored which axis EntireRow/EntireColumn meant) -- column B
        // must be removed here, column A and C's former contents (now shifted into B)
        // must survive, and row 1 (which B1 is in) must NOT be the thing that vanishes.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(1,2).Value = 20\n    Cells(1,3).Value = 30\n    Range(\"B1\").EntireColumn.Delete\n    a = Cells(1,1).Value\n    b = Cells(1,2).Value\n    c = Cells(1,3).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(10)); // column A untouched
        assert_eq!(vm.variables["b"], Variant::Integer(30)); // column C shifted into B
        assert_eq!(vm.variables["c"], Variant::Empty); // nothing shifted into C
    }

    #[test]
    fn test_entirecolumn_delete_multi_column_range() {
        // Range("A1:B1").EntireColumn spans 2 columns (A-B, not just the reference
        // cell's own column) -- mirrors test_entirerow_delete's multi-row coverage.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(1,2).Value = 2\n    Cells(1,3).Value = 3\n    Range(\"A1:B1\").EntireColumn.Delete\n    a = Cells(1,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(3)); // old C1 shifted into A1
    }

    #[test]
    fn test_entirerow_insert() {
        // GitHub #7: EntireRow.Insert was a silent no-op.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Range(\"A2\").EntireRow.Insert\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Empty); // new blank row
        assert_eq!(vm.variables["c"], Variant::Integer(2)); // shifted down
    }

    #[test]
    fn test_entirecolumn_insert() {
        // GitHub #7: EntireColumn.Insert was a silent no-op.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(1,2).Value = 2\n    Range(\"B1\").EntireColumn.Insert\n    a = Cells(1,1).Value\n    b = Cells(1,2).Value\n    c = Cells(1,3).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Empty); // new blank column
        assert_eq!(vm.variables["c"], Variant::Integer(2)); // shifted right
    }

    // ── Rows(n) / Columns(n) ─────────────────────────────────────────────────

    #[test]
    fn test_rows_delete() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    Rows(2).Delete\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(3)); // old row 3 shifted up
    }

    #[test]
    fn test_rows_insert() {
        // GitHub #7: Rows(n).Insert was a silent no-op.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Rows(2).Insert\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Empty);
        assert_eq!(vm.variables["c"], Variant::Integer(2));
    }

    #[test]
    fn test_columns_delete() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(1,2).Value = 2\n    Cells(1,3).Value = 3\n    Columns(2).Delete\n    a = Cells(1,1).Value\n    b = Cells(1,2).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(3)); // old column C shifted left
    }

    #[test]
    fn test_columns_insert() {
        // GitHub #7: Columns(n).Insert was a silent no-op.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(1,2).Value = 2\n    Columns(2).Insert\n    a = Cells(1,1).Value\n    b = Cells(1,2).Value\n    c = Cells(1,3).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Empty);
        assert_eq!(vm.variables["c"], Variant::Integer(2));
    }

    #[test]
    fn test_rows_delete_accepts_a_variable_index() {
        // Rows(n)/Columns(n) take an Expr (like Cells(row, col)), not a parse-time
        // string literal like Range(...) -- confirms a variable index actually works.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    n = 2\n    Rows(n).Delete\n    a = Cells(2,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(3));
    }

    #[test]
    fn test_range_noop_hidden() {
        let vm = run("Sub MySub()\n    Range(\"A1\").Hidden = True\n    x = 1\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1));
    }

    #[test]
    fn test_range_noop_interior_color() {
        let vm = run("Sub MySub()\n    Range(\"A1:B2\").Interior.Color = 3\n    x = 2\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(2));
    }

    #[test]
    fn test_range_noop_numberformat() {
        let vm =
            run("Sub MySub()\n    Range(\"A1\").NumberFormat = \"0.00\"\n    x = 3\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(3));
    }

    // ── Range.Delete / Range.Insert ──────────────────────────────────────────

    #[test]
    fn test_range_delete_shifts_up() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    Cells(4,1).Value = 4\n    Range(\"A2:A2\").Delete\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(3)); // 3 shifted up
        assert_eq!(vm.variables["c"], Variant::Integer(4)); // 4 shifted up
    }

    #[test]
    fn test_range_insert_shifts_down() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Range(\"A2:A2\").Insert\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Empty); // new empty row
        assert_eq!(vm.variables["c"], Variant::Integer(2)); // shifted down
    }

    // ── Range.Sort ───────────────────────────────────────────────────────────

    #[test]
    fn test_range_sort_ascending() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 3\n    Cells(2,1).Value = 1\n    Cells(3,1).Value = 2\n    Range(\"A1:A3\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
        assert_eq!(vm.variables["c"], Variant::Integer(3));
    }

    #[test]
    fn test_range_sort_descending() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 3\n    Cells(2,1).Value = 1\n    Cells(3,1).Value = 2\n    Range(\"A1:A3\").Sort Key1:=Range(\"A1\"), Order1:=xlDescending\n    a = Cells(1,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(3));
    }

    #[test]
    fn test_range_sort_header_yes_excludes_the_first_row() {
        // GitHub #6: Header:=xlYes was ignored -- the header row got swept into the sort.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = \"Name\"\n    Cells(2,1).Value = 3\n    Cells(3,1).Value = 1\n    Cells(4,1).Value = 2\n    Range(\"A1:A4\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending, Header:=xlYes\n    header = Cells(1,1).Value\n    a = Cells(2,1).Value\n    b = Cells(3,1).Value\n    c = Cells(4,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["header"], Variant::Str("Name".into())); // untouched
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
        assert_eq!(vm.variables["c"], Variant::Integer(3));
    }

    #[test]
    fn test_range_sort_header_no_is_unchanged_from_the_default() {
        // Header:=xlNo (or omitted) must keep sorting the whole given range, same as
        // before Header:=xlYes existed -- no regression on the already-working path.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 3\n    Cells(2,1).Value = 1\n    Cells(3,1).Value = 2\n    Range(\"A1:A3\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending, Header:=xlNo\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
        assert_eq!(vm.variables["c"], Variant::Integer(3));
    }

    // ── GitHub #5: Range.AutoFilter ────────────────────────────────────────

    #[test]
    fn test_range_autofilter_hides_rows_not_matching_the_criteria() {
        // Name/Age, matching the exact GitHub #5 repro: Charlie/25, Alice/40, Bob/10,
        // Dan/25 -- Field:=2 (the 2nd column of the range, "Age") Criteria1:="25" must
        // hide rows 3 (Alice/40) and 4 (Bob/10), keep the header (row 1) and the two
        // matching rows (2, 5) visible.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = \"Name\"\n    Cells(1,2).Value = \"Age\"\n    \
             Cells(2,1).Value = \"Charlie\"\n    Cells(2,2).Value = 25\n    \
             Cells(3,1).Value = \"Alice\"\n    Cells(3,2).Value = 40\n    \
             Cells(4,1).Value = \"Bob\"\n    Cells(4,2).Value = 10\n    \
             Cells(5,1).Value = \"Dan\"\n    Cells(5,2).Value = 25\n    \
             Range(\"A1:B5\").AutoFilter Field:=2, Criteria1:=\"25\"\nEnd Sub\n",
        );
        let hidden = &vm.sheet_visibility.get("sheet1").unwrap().hidden_rows;
        assert!(
            hidden.contains(&Interval { start: 3, end: 3 }),
            "row 3 (Alice/40) must be hidden: {hidden:?}"
        );
        assert!(
            hidden.contains(&Interval { start: 4, end: 4 }),
            "row 4 (Bob/10) must be hidden: {hidden:?}"
        );
        assert!(
            !hidden.iter().any(|iv| iv.start == 1),
            "the header row must never be hidden by AutoFilter: {hidden:?}"
        );
        assert!(
            !hidden.iter().any(|iv| iv.start == 2 || iv.start == 5),
            "matching rows (Charlie/25, Dan/25) must stay visible: {hidden:?}"
        );
    }

    #[test]
    fn test_range_autofilter_bare_form_hides_nothing() {
        // A bare .AutoFilter (no Field/Criteria1) just turns on the dropdown arrows in
        // real Excel -- no rows get hidden, matching GitHub #5's own description.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = \"Name\"\n    Cells(2,1).Value = \"Charlie\"\n    \
             Range(\"A1:A2\").AutoFilter\nEnd Sub\n",
        );
        assert!(
            vm.sheet_visibility
                .get("sheet1")
                .map(|v| v.hidden_rows.is_empty())
                .unwrap_or(true),
            "bare AutoFilter must not hide any rows"
        );
    }

    // ── Range clear / offset / multi-cell write / Sheets() ───────────────────

    #[test]
    fn test_range_clear_contents() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 99\n    Range(\"A1\").ClearContents\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Empty);
    }

    #[test]
    fn test_range_clear() {
        let vm = run(
            "Sub MySub()\n    Range(\"A1:A3\").Value = 5\n    Range(\"A1:A3\").Clear\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Empty);
        assert_eq!(vm.get_cell(2, 1), Variant::Empty);
        assert_eq!(vm.get_cell(3, 1), Variant::Empty);
    }

    #[test]
    fn test_range_write_multi_cell() {
        let vm = run("Sub MySub()\n    Range(\"A1:A3\").Value = 7\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(7));
        assert_eq!(vm.get_cell(2, 1), Variant::Integer(7));
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(7));
    }

    #[test]
    fn test_range_offset_read() {
        let vm = run(
            "Sub MySub()\n    Cells(2,2).Value = 42\n    x = Range(\"A1\").Offset(1,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    #[test]
    fn test_range_offset_write() {
        let vm = run("Sub MySub()\n    Range(\"A1\").Offset(2,0).Value = 99\nEnd Sub\n");
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(99));
    }

    #[test]
    fn test_sheets_cell_write() {
        let vm = run("Sub MySub()\n    Sheets(\"Sheet1\").Cells(1,1).Value = 123\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(123));
    }

    #[test]
    fn test_sheets_cell_read() {
        let vm = run(
            "Sub MySub()\n    Sheets(\"Sheet1\").Cells(1,1).Value = 55\n    x = Sheets(\"Sheet1\").Cells(1,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(55));
    }

    #[test]
    fn test_worksheets_cell_write() {
        let vm = run("Sub MySub()\n    Worksheets(\"Data\").Cells(2,3).Value = 77\nEnd Sub\n");
        // Now routes to "data" sheet, not the active "sheet1"
        let cell = vm
            .get_sheet_cells("data")
            .and_then(|s| s.get(&(2, 3)))
            .map(|c| c.value.clone());
        assert_eq!(cell, Some(Variant::Integer(77)));
    }

    // ── Multi-sheet (Phase 9) ────────────────────────────────────────────────

    #[test]
    fn test_multisheet_write_read_different_sheets() {
        let vm = run(
            "Sub MySub()\n    Sheets(\"Sheet1\").Cells(1,1).Value = 10\n    Sheets(\"Sheet2\").Cells(1,1).Value = 20\nEnd Sub\n",
        );
        let s1 = vm
            .get_sheet_cells("sheet1")
            .and_then(|s| s.get(&(1, 1)))
            .map(|c| c.value.clone());
        let s2 = vm
            .get_sheet_cells("sheet2")
            .and_then(|s| s.get(&(1, 1)))
            .map(|c| c.value.clone());
        assert_eq!(s1, Some(Variant::Integer(10)));
        assert_eq!(s2, Some(Variant::Integer(20)));
    }

    #[test]
    fn test_multisheet_cross_sheet_read() {
        let vm = run(
            "Sub MySub()\n    Sheets(\"Data\").Cells(1,1).Value = 42\n    x = Sheets(\"Data\").Cells(1,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    #[test]
    fn test_with_sheets_block() {
        let vm = run(
            "Sub MySub()\n    With Sheets(\"Sheet2\")\n        .Cells(1,1).Value = 99\n    End With\n    x = Sheets(\"Sheet2\").Cells(1,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(99));
    }

    #[test]
    fn test_with_sheets_restores_active() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    With Sheets(\"Sheet2\")\n        .Cells(1,1).Value = 2\n    End With\n    x = Cells(1,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(1)); // active sheet unchanged
    }

    // ── Milestone B7c item 6: ThisWorkbook / ActiveWorkbook / ActiveSheet ────

    #[test]
    fn activesheet_cells_write_and_read_target_the_active_sheet() {
        let vm = run(
            "Sub MySub()\n    ActiveSheet.Cells(1,1).Value = 5\n    x = ActiveSheet.Cells(1,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(5));
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test]
    fn activesheet_range_write_and_read() {
        let vm = run(
            "Sub MySub()\n    ActiveSheet.Range(\"B2\").Value = 9\n    x = ActiveSheet.Range(\"B2\").Value\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(2, 2), Variant::Integer(9));
        assert_eq!(vm.variables["x"], Variant::Integer(9));
    }

    #[test]
    fn activesheet_tracks_the_active_sheet_after_it_changes() {
        // Real VBA's ActiveSheet is dynamic — it reflects whichever sheet
        // is active *at the moment it's evaluated*, unlike a Range object
        // (which fixes its parent sheet at creation time).
        let vm = run(
            "Sub MySub()\n    With Sheets(\"Sheet2\")\n        ActiveSheet.Cells(1,1).Value = 42\n    End With\nEnd Sub\n",
        );
        let cell = vm
            .get_sheet_cells("sheet2")
            .and_then(|s| s.get(&(1, 1)))
            .map(|c| c.value.clone());
        assert_eq!(cell, Some(Variant::Integer(42)));
    }

    #[test]
    fn thisworkbook_worksheets_cell_write_targets_the_named_sheet() {
        let vm = run(
            "Sub MySub()\n    ThisWorkbook.Worksheets(\"Data\").Cells(2,3).Value = 77\nEnd Sub\n",
        );
        let cell = vm
            .get_sheet_cells("data")
            .and_then(|s| s.get(&(2, 3)))
            .map(|c| c.value.clone());
        assert_eq!(cell, Some(Variant::Integer(77)));
    }

    #[test]
    fn activeworkbook_worksheets_range_read_targets_the_named_sheet() {
        let vm = run(
            "Sub MySub()\n    Sheets(\"Data\").Cells(1,1).Value = 42\n    \
             x = ActiveWorkbook.Worksheets(\"Data\").Range(\"A1\").Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    #[test]
    fn test_sheets_add() {
        let vm = run("Sub MySub()\n    Sheets.Add\n    n = 1\nEnd Sub\n");
        assert!(vm.sheet_names().len() >= 2);
        // `sheet_order` (the writer's real-order source, distinct from
        // `sheet_names()`'s alphabetical sort) must have grown by exactly
        // one entry too, and the new sheet must land at the end -- it
        // wasn't present at `Vm::new()` time.
        assert_eq!(vm.sheet_order.len(), 2);
        assert!(!vm.sheet_order[..1].contains(&vm.sheet_order[1]));
    }

    #[test]
    fn sheets_add_after_deleting_a_middle_sheet_does_not_collide_with_a_later_survivor() {
        // Regression for a real bug: naming a new sheet purely from
        // `self.sheets.len() + 1` collides whenever the sheet set has a
        // gap. sheet1/sheet2/sheet3 -> delete sheet2 -> len()==2 -> the old
        // code computed "sheet3", which still exists as a survivor, and
        // ensure_sheet() silently no-ops on a collision -- the Add produced
        // nothing at all, with no error.
        let vm = run("Sub MySub()\n    \
             Sheets.Add\n    \
             Sheets.Add\n    \
             Sheets(\"Sheet3\").Cells(1,1).Value = 99\n    \
             Sheets(\"Sheet2\").Delete\n    \
             Sheets.Add\n\
             End Sub\n");
        assert_eq!(
            vm.sheet_order,
            vec!["sheet1", "sheet3", "sheet4"],
            "the post-delete Add must land on a fresh name, not silently no-op on \
             the sheet3 collision"
        );
        assert_eq!(
            vm.get_sheet_cells("sheet3")
                .unwrap()
                .get(&(1, 1))
                .map(|c| &c.value),
            Some(&Variant::Integer(99)),
            "sheet3's own data must survive untouched by the later Add"
        );
        assert!(vm.get_sheet_cells("sheet4").unwrap().is_empty());
    }

    #[test]
    fn test_sheets_delete() {
        let vm = run(
            "Sub MySub()\n    Sheets(\"Sheet2\").Cells(1,1).Value = 5\n    Sheets(\"Sheet2\").Delete\n    n = 1\nEnd Sub\n",
        );
        assert!(!vm.sheet_names().contains(&"sheet2".to_string()));
        assert!(
            !vm.sheet_order.contains(&"sheet2".to_string()),
            "sheet_order must drop a deleted sheet too, or the writer would still emit it"
        );
    }

    // ── GitHub #3: Vm::delete_sheet / Vm::ensure_sheet_at ────────────────────

    #[test]
    fn delete_sheet_removes_a_non_active_sheet() {
        let mut vm = Vm::new();
        vm.ensure_sheet("Extra");
        vm.delete_sheet("Extra").unwrap();
        assert!(!vm.sheet_names().contains(&"extra".to_string()));
        assert!(!vm.sheet_order.contains(&"extra".to_string()));
    }

    #[test]
    fn delete_sheet_errors_on_an_unknown_name_instead_of_silently_no_opping() {
        // Unlike Sheets("Typo").Delete (which resolves via resolve_sheet_expr, never
        // validates existence, and silently no-ops), the direct API is a caller-facing
        // structural operation and should say so on a typo.
        let mut vm = Vm::new();
        let err = vm.delete_sheet("Typo").unwrap_err();
        assert!(err.contains("not found"), "unexpected error: {err}");
    }

    #[test]
    fn ensure_sheet_at_inserts_at_the_requested_position() {
        let mut vm = Vm::new(); // starts with just "sheet1"
        vm.ensure_sheet("Second");
        vm.ensure_sheet_at("First", Some(0));
        assert_eq!(vm.sheet_order, vec!["first", "sheet1", "second"]);
    }

    #[test]
    fn ensure_sheet_at_clamps_an_out_of_range_index_instead_of_panicking() {
        let mut vm = Vm::new();
        vm.ensure_sheet_at("Later", Some(999));
        assert_eq!(vm.sheet_order, vec!["sheet1", "later"]);
    }

    #[test]
    fn ensure_sheet_at_ignores_the_index_when_the_sheet_already_exists() {
        let mut vm = Vm::new();
        vm.ensure_sheet("Second");
        vm.ensure_sheet_at("Second", Some(0)); // already exists -- must not move it
        assert_eq!(vm.sheet_order, vec!["sheet1", "second"]);
    }

    // ── P1 core 3: Vm::rename_sheet / Vm::move_sheet ─────────────────────────

    #[test]
    fn rename_sheet_updates_all_eight_per_sheet_maps() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(42),
            },
        );
        vm.merged_ranges
            .insert("sheet1".to_string(), vec![((1, 2), (1, 4))]);
        vm.sheet_visibility.insert(
            "sheet1".to_string(),
            SheetVisibility {
                hidden_rows: vec![Interval { start: 5, end: 5 }],
                hidden_columns: vec![],
            },
        );
        vm.cell_style_indices
            .insert("sheet1".to_string(), HashMap::from([((1, 1), 3u32)]));
        vm.cell_number_formats.insert(
            "sheet1".to_string(),
            HashMap::from([((1, 1), "0.00".to_string())]),
        );

        vm.rename_sheet("Sheet1", "Renamed").unwrap();

        assert!(vm.sheets.contains_key("renamed"));
        assert!(!vm.sheets.contains_key("sheet1"));
        assert!(vm.merged_ranges.contains_key("renamed"));
        assert!(!vm.merged_ranges.contains_key("sheet1"));
        assert!(vm.sheet_visibility.contains_key("renamed"));
        assert!(!vm.sheet_visibility.contains_key("sheet1"));
        assert!(vm.cell_style_indices.contains_key("renamed"));
        assert!(!vm.cell_style_indices.contains_key("sheet1"));
        assert!(vm.cell_number_formats.contains_key("renamed"));
        assert!(!vm.cell_number_formats.contains_key("sheet1"));
        assert!(vm.worksheet_origins.contains_key("renamed"));
        assert!(!vm.worksheet_origins.contains_key("sheet1"));
        assert!(vm.sheet_order.contains(&"renamed".to_string()));
        assert!(!vm.sheet_order.contains(&"sheet1".to_string()));
    }

    #[test]
    fn rename_sheet_preserves_tab_position_in_sheet_order() {
        let mut vm = Vm::new();
        vm.ensure_sheet("B");
        vm.ensure_sheet("C");
        assert_eq!(vm.sheet_order, vec!["sheet1", "b", "c"]);
        vm.rename_sheet("B", "Renamed").unwrap();
        assert_eq!(vm.sheet_order, vec!["sheet1", "renamed", "c"]);
    }

    #[test]
    fn rename_sheet_updates_active_sheet_when_renaming_the_active_sheet() {
        let mut vm = Vm::new(); // active sheet is "sheet1"
        vm.rename_sheet("Sheet1", "Renamed").unwrap();
        assert_eq!(vm.active_sheet, "renamed");
        // Must not panic -- `cells()` unwraps on `self.sheets[active_sheet]`.
        assert_eq!(vm.cells().len(), 0);
    }

    #[test]
    fn rename_sheet_does_not_touch_active_sheet_when_renaming_a_different_sheet() {
        let mut vm = Vm::new();
        vm.ensure_sheet("Other");
        vm.rename_sheet("Other", "Renamed").unwrap();
        assert_eq!(vm.active_sheet, "sheet1");
    }

    #[test]
    fn rename_sheet_sets_worksheet_origins_display_name_to_the_new_name() {
        let mut vm = Vm::new();
        vm.rename_sheet("Sheet1", "NewName").unwrap();
        assert_eq!(
            vm.worksheet_origins
                .get("newname")
                .and_then(|o| o.original_display_name.clone()),
            Some("NewName".to_string())
        );
    }

    #[test]
    fn rename_sheet_allows_case_only_rename_of_the_same_sheet() {
        let mut vm = Vm::new();
        vm.rename_sheet("Sheet1", "SHEET1").unwrap();
        assert_eq!(
            vm.worksheet_origins
                .get("sheet1")
                .and_then(|o| o.original_display_name.clone()),
            Some("SHEET1".to_string())
        );
    }

    #[test]
    fn rename_sheet_errors_if_old_name_not_found() {
        let mut vm = Vm::new();
        let err = vm.rename_sheet("Typo", "New").unwrap_err();
        assert!(err.contains("not found"), "unexpected error: {err}");
    }

    #[test]
    fn rename_sheet_errors_if_new_name_collides_with_a_different_existing_sheet() {
        let mut vm = Vm::new();
        vm.ensure_sheet("Other");
        let err = vm.rename_sheet("Sheet1", "Other").unwrap_err();
        assert!(err.contains("already exists"), "unexpected error: {err}");
    }

    #[test]
    fn rename_sheet_errors_on_empty_or_whitespace_new_name() {
        let mut vm = Vm::new();
        assert!(vm.rename_sheet("Sheet1", "").is_err());
        assert!(vm.rename_sheet("Sheet1", "   ").is_err());
    }

    #[test]
    fn rename_sheet_errors_if_sheet_is_protected() {
        let mut vm = Vm::new();
        vm.protected_sheets.insert("sheet1".to_string());
        let err = vm.rename_sheet("Sheet1", "New").unwrap_err();
        assert!(err.contains("protected"), "unexpected error: {err}");
    }

    #[test]
    fn move_sheet_reorders_sheet_order_without_touching_other_maps() {
        let mut vm = Vm::new();
        vm.ensure_sheet("B");
        vm.ensure_sheet("C");
        vm.merged_ranges
            .insert("sheet1".to_string(), vec![((1, 1), (1, 1))]);
        let active_before = vm.active_sheet.clone();

        vm.move_sheet("C", 0).unwrap();

        assert_eq!(vm.sheet_order, vec!["c", "sheet1", "b"]);
        assert_eq!(vm.active_sheet, active_before);
        assert!(vm.merged_ranges.contains_key("sheet1"));
    }

    #[test]
    fn move_sheet_clamps_out_of_range_index_to_the_end() {
        let mut vm = Vm::new();
        vm.ensure_sheet("B");
        vm.move_sheet("Sheet1", 999).unwrap();
        assert_eq!(vm.sheet_order, vec!["b", "sheet1"]);
    }

    #[test]
    fn move_sheet_errors_if_sheet_not_found() {
        let mut vm = Vm::new();
        let err = vm.move_sheet("Typo", 0).unwrap_err();
        assert!(err.contains("not found"), "unexpected error: {err}");
    }

    #[test]
    fn move_sheet_does_not_check_protected_sheets() {
        let mut vm = Vm::new();
        vm.ensure_sheet("B");
        vm.protected_sheets.insert("sheet1".to_string());
        vm.move_sheet("Sheet1", 1).unwrap();
        assert_eq!(vm.sheet_order, vec!["b", "sheet1"]);
    }

    #[test]
    fn rename_sheet_and_move_sheet_both_flag_defined_names_as_possibly_stale() {
        // A rename can leave a <definedName>'s TEXT dangling (e.g. "Sheet1!$F$5"
        // after "Sheet1" is renamed); a reorder can invalidate a positional
        // localSheetId. Both set the same flag -- caught in review after an
        // earlier version of this fix only covered move_sheet, missing that
        // rename_sheet needed to set it too.
        let mut vm = Vm::new();
        assert!(!vm.defined_names_may_be_stale);
        vm.rename_sheet("Sheet1", "Renamed").unwrap();
        assert!(vm.defined_names_may_be_stale);

        let mut vm2 = Vm::new();
        vm2.ensure_sheet("B");
        assert!(!vm2.defined_names_may_be_stale);
        vm2.move_sheet("B", 0).unwrap();
        assert!(vm2.defined_names_may_be_stale);
    }

    // ── P1 core 3: row/col insert-delete sheet-parameterized siblings ───────

    #[test]
    fn insert_rows_on_sheet_does_not_affect_a_different_sheet_or_change_active_sheet() {
        let mut vm = Vm::new(); // active sheet is "sheet1"
        vm.ensure_sheet("Other");
        vm.sheet_cells_mut("other").unwrap().insert(
            (5, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(99),
            },
        );
        vm.cells_mut().insert(
            (5, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );

        vm.insert_rows_on_sheet("other", 1, 2);

        assert_eq!(vm.active_sheet, "sheet1");
        assert_eq!(vm.get_cell(5, 1), Variant::Integer(1)); // active sheet untouched
        assert_eq!(
            vm.get_sheet_cells("other")
                .unwrap()
                .get(&(7, 1))
                .unwrap()
                .value,
            Variant::Integer(99)
        ); // shifted down by 2 on the target sheet
    }

    #[test]
    fn delete_rows_on_sheet_shifts_only_the_target_sheet() {
        let mut vm = Vm::new();
        vm.ensure_sheet("Other");
        vm.sheet_cells_mut("other").unwrap().insert(
            (5, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(99),
            },
        );
        vm.cells_mut().insert(
            (5, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );

        vm.delete_rows_on_sheet("other", 1, 2);

        assert_eq!(vm.active_sheet, "sheet1");
        assert_eq!(vm.get_cell(5, 1), Variant::Integer(1)); // active sheet untouched
        assert_eq!(
            vm.get_sheet_cells("other")
                .unwrap()
                .get(&(3, 1))
                .unwrap()
                .value,
            Variant::Integer(99)
        ); // shifted up by 2 on the target sheet
    }

    #[test]
    fn delete_rows_on_sheet_removes_the_stale_entry_at_the_pre_shift_row() {
        // Regression test for a bug caught in review: an earlier draft of
        // delete_rows_on_sheet used `retain(row < first || row > last)`, which
        // kept the pre-shift entry AND inserted a second copy at the shifted
        // position -- silent stale duplicate data. The correct predicate is
        // `retain(row < first)`, verified here.
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (10, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(7),
            },
        );
        vm.delete_rows_on_sheet("sheet1", 1, 2);
        assert_eq!(vm.get_cell(10, 1), Variant::Empty); // stale pre-shift position is gone
        assert_eq!(vm.get_cell(8, 1), Variant::Integer(7)); // correct shifted position
    }

    #[test]
    fn insert_cols_on_sheet_does_not_affect_a_different_sheet_or_change_active_sheet() {
        let mut vm = Vm::new();
        vm.ensure_sheet("Other");
        vm.sheet_cells_mut("other").unwrap().insert(
            (1, 5),
            CellContent {
                formula: None,
                value: Variant::Integer(99),
            },
        );
        vm.cells_mut().insert(
            (1, 5),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );

        vm.insert_cols_on_sheet("other", 1, 2);

        assert_eq!(vm.active_sheet, "sheet1");
        assert_eq!(vm.get_cell(1, 5), Variant::Integer(1));
        assert_eq!(
            vm.get_sheet_cells("other")
                .unwrap()
                .get(&(1, 7))
                .unwrap()
                .value,
            Variant::Integer(99)
        );
    }

    #[test]
    fn delete_cols_on_sheet_removes_the_stale_entry_at_the_pre_shift_col() {
        // Column-axis mirror of the row-axis stale-position regression test above.
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 10),
            CellContent {
                formula: None,
                value: Variant::Integer(7),
            },
        );
        vm.delete_cols_on_sheet("sheet1", 1, 2);
        assert_eq!(vm.get_cell(1, 10), Variant::Empty);
        assert_eq!(vm.get_cell(1, 8), Variant::Integer(7));
    }

    // ── P1 remainder: sort_range_on_sheet (extracted from Stmt::RangeSort) ──

    #[test]
    fn sort_range_on_sheet_does_not_affect_a_different_sheet_or_change_active_sheet() {
        let mut vm = Vm::new(); // active sheet is "sheet1"
        vm.ensure_sheet("Other");
        vm.write_rect(
            "other",
            (1, 1),
            &[
                vec![Variant::Integer(3)],
                vec![Variant::Integer(1)],
                vec![Variant::Integer(2)],
            ],
        );
        vm.write_rect("sheet1", (1, 1), &[vec![Variant::Integer(99)]]);

        vm.sort_range_on_sheet("other", 1, 1, 3, 1, 1, false, false);

        assert_eq!(vm.active_sheet, "sheet1");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(99)); // active sheet untouched
        assert_eq!(
            vm.iter_rows_values("other", 1, Some(3), 1, Some(1)),
            vec![
                vec![Variant::Integer(1)],
                vec![Variant::Integer(2)],
                vec![Variant::Integer(3)],
            ]
        );
    }

    #[test]
    fn sort_range_on_sheet_sorts_ascending_and_descending() {
        let mut vm = Vm::new();
        vm.write_rect(
            "sheet1",
            (1, 1),
            &[
                vec![Variant::Integer(3)],
                vec![Variant::Integer(1)],
                vec![Variant::Integer(2)],
            ],
        );
        vm.sort_range_on_sheet("sheet1", 1, 1, 3, 1, 1, false, false);
        assert_eq!(
            vm.iter_rows_values("sheet1", 1, Some(3), 1, Some(1)),
            vec![
                vec![Variant::Integer(1)],
                vec![Variant::Integer(2)],
                vec![Variant::Integer(3)],
            ]
        );

        let mut vm2 = Vm::new();
        vm2.write_rect(
            "sheet1",
            (1, 1),
            &[
                vec![Variant::Integer(3)],
                vec![Variant::Integer(1)],
                vec![Variant::Integer(2)],
            ],
        );
        vm2.sort_range_on_sheet("sheet1", 1, 1, 3, 1, 1, true, false);
        assert_eq!(
            vm2.iter_rows_values("sheet1", 1, Some(3), 1, Some(1)),
            vec![
                vec![Variant::Integer(3)],
                vec![Variant::Integer(2)],
                vec![Variant::Integer(1)],
            ]
        );
    }

    #[test]
    fn sort_range_on_sheet_excludes_the_header_row() {
        let mut vm = Vm::new();
        vm.write_rect(
            "sheet1",
            (1, 1),
            &[
                vec![Variant::Integer(9)], // header -- must stay at row 1
                vec![Variant::Integer(3)],
                vec![Variant::Integer(1)],
            ],
        );
        vm.sort_range_on_sheet("sheet1", 1, 1, 3, 1, 1, false, true);
        assert_eq!(
            vm.iter_rows_values("sheet1", 1, Some(3), 1, Some(1)),
            vec![
                vec![Variant::Integer(9)],
                vec![Variant::Integer(1)],
                vec![Variant::Integer(3)],
            ]
        );
    }

    #[test]
    fn sort_range_on_sheet_with_an_out_of_range_key_col_clamps_via_saturating_sub() {
        // Pins the preserved (not fixed) VBA-path behavior: a key_col below
        // the range's own c1 saturates to offset 0 instead of erroring, so
        // it silently sorts by the range's first column. This must not
        // change without a conscious decision -- PyVm::sort_range validates
        // key_col explicitly instead of inheriting this clamp.
        let mut vm = Vm::new();
        vm.write_rect(
            "sheet1",
            (1, 2), // range starts at column 2 (B)
            &[
                vec![Variant::Integer(3)],
                vec![Variant::Integer(1)],
                vec![Variant::Integer(2)],
            ],
        );
        vm.sort_range_on_sheet("sheet1", 1, 2, 3, 2, 1, false, false); // key_col=1 < c1=2
        assert_eq!(
            vm.iter_rows_values("sheet1", 1, Some(3), 2, Some(2)),
            vec![
                vec![Variant::Integer(1)],
                vec![Variant::Integer(2)],
                vec![Variant::Integer(3)],
            ]
        );
    }

    // ── Milestone B6a: strict_resolution + resolution-failure evidence ──────

    #[test]
    fn sheet_range_write_and_read_round_trip() {
        // New Milestone B6a construct: Sheets(name).Range(addr) — previously
        // only .Cells(r,c) was supported off a sheet name.
        let vm = run(
            "Sub MySub()\n    Sheets(\"Sheet2\").Range(\"B2\").Value = 123\n    x = Sheets(\"Sheet2\").Range(\"B2\").Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(123));
    }

    #[test]
    fn strict_mode_write_to_a_missing_sheet_is_a_resolution_failure_not_auto_vivify() {
        let prog = parser::parse(
            "Sub MySub()\n    Worksheets(\"NoSuchSheet\").Cells(1,1).Value = 1\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        vm.strict_resolution = true;
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Sheet 'NoSuchSheet' not found");
        assert!(
            !vm.sheet_names().contains(&"nosuchsheet".to_string()),
            "strict mode must not auto-vivify"
        );
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::WorksheetNotFound(e)) => {
                assert_eq!(e.requested, "NoSuchSheet");
                assert!(e.available.contains(&"sheet1".to_string()));
            }
            other => panic!("expected WorksheetNotFound, got {:?}", other),
        }
    }

    #[test]
    fn non_strict_mode_write_to_a_missing_sheet_still_auto_vivifies() {
        // Confirms strict_resolution is opt-in only — every existing caller
        // (default: false) keeps today's convenience behavior unchanged.
        let prog = parser::parse(
            "Sub MySub()\n    Worksheets(\"NewSheet\").Cells(1,1).Value = 42\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(
            vm.get_sheet_cells("newsheet")
                .and_then(|s| s.get(&(1, 1)))
                .map(|c| c.value.clone()),
            Some(Variant::Integer(42))
        );
    }

    #[test]
    fn strict_mode_read_from_a_missing_sheet_is_a_resolution_failure_not_empty() {
        let prog = parser::parse(
            "Sub MySub()\n    x = Worksheets(\"NoSuchSheet\").Cells(1,1).Value\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        vm.strict_resolution = true;
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Sheet 'NoSuchSheet' not found");
    }

    #[test]
    fn non_strict_mode_read_from_a_missing_sheet_is_still_silently_empty() {
        let prog = parser::parse(
            "Sub MySub()\n    x = Worksheets(\"NoSuchSheet\").Cells(1,1).Value\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.variables["x"], Variant::Empty);
    }

    #[test]
    fn strict_mode_with_sheets_on_a_missing_sheet_is_a_resolution_failure() {
        // `With Sheets("...")` parses its sheet name to a plain lowercased
        // String (unlike the Expr-based Sheets(...)/Worksheets(...) forms
        // above) — left untouched by B6a, so the evidence shows the
        // already-lowercased name here, not the as-written case.
        let prog = parser::parse("Sub MySub()\n    With Sheets(\"NoSuchSheet\")\n        .Cells(1,1).Value = 1\n    End With\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        vm.strict_resolution = true;
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Sheet 'nosuchsheet' not found");
    }

    #[test]
    fn numeric_sheet_index_selects_by_alphabetical_position_in_both_modes() {
        // elixcee has no real tab-order tracking, so a numeric index resolves
        // against `sheet_names()`'s alphabetical order — documented as an
        // honest fidelity gap, not real Excel tab order.
        let vm = run(
            "Sub MySub()\n    Sheets(\"Alpha\").Cells(1,1).Value = 1\n    Sheets(\"Beta\").Cells(1,1).Value = 2\n    Worksheets(2).Cells(2,2).Value = 99\nEnd Sub\n",
        );
        // sheet_names() alphabetical: alpha, beta, sheet1 -> index 2 = "beta"
        assert_eq!(
            vm.get_sheet_cells("beta")
                .and_then(|s| s.get(&(2, 2)))
                .map(|c| c.value.clone()),
            Some(Variant::Integer(99))
        );
    }

    #[test]
    fn numeric_sheet_index_out_of_range_is_a_hard_error_even_without_strict_mode() {
        // Numeric indexing is new in B6a — there's no pre-B6a lenient
        // behavior to preserve for it, so it's always a hard error.
        let prog = parser::parse("Sub MySub()\n    x = Worksheets(99).Cells(1,1).Value\nEnd Sub\n")
            .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Sheet index 99 not found");
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::WorksheetNotFound(e)) => assert_eq!(e.requested, "99"),
            other => panic!("expected WorksheetNotFound, got {:?}", other),
        }
    }

    #[test]
    fn workbooks_qualified_sheet_access_matches_the_loaded_workbook_by_name() {
        let out_path = std::env::temp_dir().join("elixcee_vm_workbooks_match_test.xlsx");
        crate::save_workbook(&Vm::new(), out_path.to_str().unwrap()).unwrap();
        let file_name = std::path::Path::new(&out_path)
            .file_name()
            .unwrap()
            .to_str()
            .unwrap()
            .to_string();

        let src = format!(
            "Sub MySub()\n    Workbooks(\"{}\").Worksheets(\"Sheet1\").Cells(1,1).Value = 7\nEnd Sub\n",
            file_name
        );
        let prog = parser::parse(&src).unwrap();
        let mut vm = Vm::new();
        vm.load_workbook_file(out_path.to_str().unwrap()).unwrap();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(
            vm.get_sheet_cells("sheet1")
                .and_then(|s| s.get(&(1, 1)))
                .map(|c| c.value.clone()),
            Some(Variant::Integer(7))
        );
    }

    #[test]
    fn workbooks_qualified_sheet_access_reports_a_mismatch_unconditionally() {
        // A workbook mismatch is always a hard error — not gated behind
        // strict_resolution, since Workbooks(...) is a brand-new B6a
        // construct with no pre-B6a lenient behavior to preserve.
        let out_path = std::env::temp_dir().join("elixcee_vm_workbooks_mismatch_test.xlsx");
        crate::save_workbook(&Vm::new(), out_path.to_str().unwrap()).unwrap();

        let prog = parser::parse(
            "Sub MySub()\n    Workbooks(\"other.xlsx\").Worksheets(1).Cells(1,1).Value = 1\nEnd Sub\n",
        ).unwrap();
        let mut vm = Vm::new();
        vm.load_workbook_file(out_path.to_str().unwrap()).unwrap();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Workbook 'other.xlsx' not found");
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::WorkbookNotFound(e)) => {
                assert_eq!(e.requested, "other.xlsx");
                assert!(!e.available.is_empty());
            }
            other => panic!("expected WorkbookNotFound, got {:?}", other),
        }
    }

    #[test]
    fn array_out_of_bounds_evidence_reports_zero_based_bounds() {
        let prog = parser::parse("Sub MySub()\n    Dim arr(3)\n    arr(9) = 1\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Subscript out of range");
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::ArrayIndexOutOfBounds {
                name,
                index,
                lower,
                upper,
            }) => {
                assert_eq!(name, "arr");
                assert_eq!(index, 9);
                assert_eq!(lower, 0);
                assert_eq!(upper, 3);
            }
            other => panic!("expected ArrayIndexOutOfBounds, got {:?}", other),
        }
    }

    #[test]
    fn last_resolution_failure_does_not_leak_across_separate_run_sub_calls() {
        let bad = parser::parse("Sub Bad()\n    Dim arr(1)\n    arr(9) = 1\nEnd Sub\n").unwrap();
        let good = parser::parse("Sub Good()\n    x = 1\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        assert!(vm.run_sub(&bad, "bad").is_err());
        assert!(vm.take_resolution_failure().is_some());
        vm.run_sub(&good, "good").unwrap();
        assert!(
            vm.take_resolution_failure().is_none(),
            "stale evidence from a prior failed run must not leak into a later successful run"
        );
    }

    #[test]
    fn levenshtein_distance_matches_hand_counted_edits() {
        assert_eq!(levenshtein("", ""), 0);
        assert_eq!(levenshtein("abc", "abc"), 0);
        assert_eq!(levenshtein("abc", "abd"), 1);
        assert_eq!(levenshtein("kitten", "sitting"), 3);
        assert_eq!(levenshtein("売上2025", "売上2026"), 1);
    }

    #[test]
    fn closest_match_suggests_only_within_a_bounded_distance() {
        let candidates = vec![
            "Sales2026".to_string(),
            "Summary".to_string(),
            "Input".to_string(),
        ];
        assert_eq!(
            closest_match("Sales2025", &candidates),
            Some("Sales2026".to_string())
        );
        // Nothing here is meaningfully close to "ZzzUnrelated" — no suggestion.
        assert_eq!(closest_match("ZzzUnrelated", &candidates), None);
    }

    #[test]
    fn test_sheet_names() {
        let vm = run(
            "Sub MySub()\n    Sheets(\"Alpha\").Cells(1,1).Value = 1\n    Sheets(\"Beta\").Cells(1,1).Value = 2\nEnd Sub\n",
        );
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
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    lastRow = Cells(Rows.Count,1).End(xlUp).Row\nEnd Sub\n",
        );
        assert_eq!(vm.variables["lastrow"], Variant::Integer(3));
    }

    #[test]
    fn test_cells_end_col_left() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = \"a\"\n    Cells(1,2).Value = \"b\"\n    Cells(1,3).Value = \"c\"\n    lastCol = Cells(1,Columns.Count).End(xlToLeft).Column\nEnd Sub\n",
        );
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
    fn test_dim_multi_declarator_end_to_end() {
        // `Dim a As Integer, b As Person` — a comma-separated multi-declarator
        // Dim mixing a built-in type with a user-defined type. Previously
        // unparseable (see CHANGELOG.md's Known limitations / ROADMAP.md):
        // `b As Person` returned early from `parse_dim` without consuming
        // the rest of the line, so `eat_eol()` hit the trailing comma and
        // failed the whole macro at parse time.
        let vm = run(concat!(
            "Type Person\n",
            "    Name As String\n",
            "End Type\n",
            "\n",
            "Sub MySub()\n",
            "    Dim a As Integer, b As Person\n",
            "    a = 7\n",
            "    b.Name = \"Alice\"\n",
            "    x = a\n",
            "    y = b.Name\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["x"], Variant::Integer(7));
        assert_eq!(vm.variables["y"], Variant::Str("Alice".into()));
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
        assert_eq!(vm.variables["xi"], Variant::Integer(0));
        assert_eq!(vm.variables["yi"], Variant::Integer(0));
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
            "    s = o.Val\n",  // default "" for String
            "    x = o.Pt.X\n", // default 0 for nested Integer
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
            "    d = items(0).Value\n", // default 0
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
        assert_eq!(vm.variables["rx"], Variant::Integer(5));
        assert_eq!(vm.variables["ry"], Variant::Integer(10));
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
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(10),
            },
        );
        // Force index rebuild
        let _ = vm.last_nonempty_row(1, 1_048_576);
        assert!(!vm.cell_index_dirty);
        // Write via sheet_cells_mut (simulated by inserting through the public method)
        vm.cells_mut().insert(
            (5, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(50),
            },
        );
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

    // ── msgbox_log lifecycle ───────────────────────────────────────────────────

    #[test]
    fn test_msgbox_log_does_not_leak_across_runs() {
        let prog1 = parser::parse("Sub First()\n    MsgBox \"from first\"\nEnd Sub\n").unwrap();
        let prog2 = parser::parse("Sub Second()\n    MsgBox \"from second\"\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();

        vm.run_sub(&prog1, "First").unwrap();
        assert_eq!(vm.take_messages(), vec!["from first".to_string()]);

        // Reusing the same Vm for a second run must not carry over the first
        // run's messages, even if take_messages() wasn't called in between.
        vm.run_sub(&prog2, "Second").unwrap();
        assert_eq!(vm.take_messages(), vec!["from second".to_string()]);
    }

    #[test]
    fn test_msgbox_log_survives_a_later_runtime_error() {
        let prog = parser::parse(
            "Sub MySub()\n    MsgBox \"seen before failure\"\n    x = totla + 1\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let result = vm.run_sub(&prog, "MySub");
        assert!(
            result.is_err(),
            "expected the undefined-variable error to propagate"
        );
        assert_eq!(vm.take_messages(), vec!["seen before failure".to_string()]);
    }

    #[test]
    fn test_msgbox_blocked_is_recorded_before_failing() {
        let prog = parser::parse("Sub MySub()\n    MsgBox \"blocked\"\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        vm.error_on_msgbox = true;
        let result = vm.run_sub(&prog, "MySub");
        assert!(
            result.is_err(),
            "MsgBox must still fail when error_on_msgbox is set"
        );
        // Spec: messages reflects every MsgBox the macro attempted to show,
        // even ones that are then treated as a blocking error.
        assert_eq!(vm.take_messages(), vec!["blocked".to_string()]);
    }

    #[test]
    fn test_take_messages_drains_the_log() {
        let prog = parser::parse("Sub MySub()\n    MsgBox \"once\"\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        vm.run_sub(&prog, "MySub").unwrap();
        assert_eq!(vm.take_messages(), vec!["once".to_string()]);
        // A second drain with no new MsgBox calls must come back empty.
        assert!(vm.take_messages().is_empty());
    }

    // ── Dim registers a real Empty-valued variable ──────────────────────────

    #[test]
    fn test_dim_without_type_registers_an_empty_variable() {
        // Dim x used to be a complete no-op — the variable name was never
        // recorded at all, so IsEmpty(x)/x + 5 hit "Undefined variable"
        // instead of real VBA's Empty. Found by compat/vba-semantics/, a
        // new value-correctness suite, on its very first run.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    Dim x\n",
            "    a = IsEmpty(x)\n",
            "    b = x + 5\n",
            "    x = 10\n",
            "    c = IsEmpty(x)\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Integer(5));
        assert_eq!(vm.variables["c"], Variant::Boolean(false));
    }

    #[test]
    fn test_dim_as_builtin_type_also_registers_an_empty_variable() {
        let vm = run("Sub MySub()\n    Dim x As Integer\n    a = IsEmpty(x)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
    }

    #[test]
    fn test_dim_multi_declarator_registers_every_bare_name() {
        let vm = run(concat!(
            "Sub MySub()\n",
            "    Dim x As Integer, y\n",
            "    a = IsEmpty(x)\n",
            "    b = IsEmpty(y)\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Boolean(true));
    }

    #[test]
    fn test_dim_does_not_reset_an_already_assigned_variable_if_rerun() {
        // Dim is a declaration, not a runtime reset -- if the statement
        // somehow re-executes (a loop containing a Dim, unusual but legal
        // VBA), an already-assigned value must survive.
        let vm = run(concat!(
            "Sub MySub()\n",
            "    For i = 1 To 2\n",
            "        Dim x\n",
            "        If i = 1 Then x = 42\n",
            "    Next i\n",
            "    a = x\n",
            "End Sub\n",
        ));
        assert_eq!(vm.variables["a"], Variant::Integer(42));
    }

    // ── Stmt::Unsupported executes as a true no-op ──────────────────────────

    #[test]
    fn test_unsupported_stmt_is_a_true_noop() {
        // Range("A1").NumberFormat isn't a recognized Range property, so it
        // parses to Stmt::Unsupported — confirm it doesn't error and later
        // statements still run normally, exactly like Stmt::Dim.
        let vm =
            run("Sub MySub()\n    Range(\"A1\").NumberFormat = \"0.00\"\n    x = 3\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(3));
    }

    // ── is_known_builtin_function ───────────────────────────────────────────

    #[test]
    fn known_builtin_vba_functions_are_recognized() {
        assert!(is_known_builtin_function("len"));
        assert!(is_known_builtin_function("iif"));
        assert!(is_known_builtin_function("range"));
    }

    #[test]
    fn known_worksheet_functions_are_recognized_via_wsf_prefix() {
        assert!(is_known_builtin_function("wsf_sum"));
        assert!(is_known_builtin_function("wsf_countif"));
    }

    #[test]
    fn unknown_names_are_not_recognized() {
        assert!(!is_known_builtin_function("totallyfake"));
        assert!(!is_known_builtin_function("wsf_totallyfake"));
    }

    // ── run_sub_multi (Milestone B2) ────────────────────────────────────────

    fn module(name: &str, src: &str) -> (String, Program) {
        (name.to_string(), parser::parse(src).unwrap())
    }

    #[test]
    fn run_sub_multi_single_module_behaves_like_run_sub() {
        let modules = vec![module("module1", "Sub Main()\n    x = 42\nEnd Sub\n")];
        let mut vm = Vm::new();
        vm.run_sub_multi(&modules, "Main").unwrap();
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    #[test]
    fn run_sub_multi_resolves_unique_bare_name_across_modules() {
        let modules = vec![
            module("module1", "Sub Helper()\n    y = 1\nEnd Sub\n"),
            module(
                "module2",
                "Sub Main()\n    Call Helper()\n    x = 42\nEnd Sub\n",
            ),
        ];
        let mut vm = Vm::new();
        vm.run_sub_multi(&modules, "Main").unwrap();
        assert_eq!(vm.variables["x"], Variant::Integer(42));
        assert_eq!(vm.variables["y"], Variant::Integer(1));
    }

    #[test]
    fn run_sub_multi_resolves_qualified_entrypoint() {
        // Qualification works even without a collision forcing it — useful
        // for explicit scripting even when the bare name would resolve
        // fine on its own. (Disambiguating a *genuine* same-name collision
        // via qualification is not supported: the flat cross-module merge
        // used for in-body calls can't safely coexist with it, so any
        // collision is rejected at load regardless of qualification — see
        // `run_sub_multi_rejects_a_genuine_sub_collision_before_executing_anything`.)
        let modules = vec![
            module("module1", "Sub Other()\n    x = 1\nEnd Sub\n"),
            module("module2", "Sub Main()\n    x = 2\nEnd Sub\n"),
        ];
        let mut vm = Vm::new();
        vm.run_sub_multi(&modules, "Module2.Main").unwrap();
        assert_eq!(vm.variables["x"], Variant::Integer(2));
    }

    #[test]
    fn run_sub_multi_rejects_a_genuine_sub_collision_before_executing_anything() {
        let modules = vec![
            module("module1", "Sub Main()\n    x = 1\nEnd Sub\n"),
            module("module2", "Sub Main()\n    x = 2\nEnd Sub\n"),
        ];
        let mut vm = Vm::new();
        let err = vm.run_sub_multi(&modules, "Main").unwrap_err();
        assert!(err.contains("duplicate Sub 'main'"), "{:?}", err);
        assert!(
            !vm.variables.contains_key("x"),
            "no execution should have happened"
        );
    }

    #[test]
    fn run_sub_multi_rejects_a_genuine_func_collision() {
        let modules = vec![
            module(
                "module1",
                "Function Foo()\n    Foo = 1\nEnd Function\nSub Main()\n    x = 1\nEnd Sub\n",
            ),
            module("module2", "Function Foo()\n    Foo = 2\nEnd Function\n"),
        ];
        let mut vm = Vm::new();
        let err = vm.run_sub_multi(&modules, "Module1.Main").unwrap_err();
        assert!(err.contains("duplicate Function 'foo'"), "{:?}", err);
    }

    #[test]
    fn run_sub_multi_entrypoint_not_found() {
        let modules = vec![module("module1", "Sub Main()\n    x = 1\nEnd Sub\n")];
        let mut vm = Vm::new();
        let err = vm.run_sub_multi(&modules, "Bogus").unwrap_err();
        assert!(err.contains("not found"), "{:?}", err);
    }

    // ── Milestone B5a: parse_sheet_range_addr / load_workbook_file / deadline ──

    #[test]
    fn parse_sheet_range_addr_with_sheet_prefix() {
        let (sheet, from, to) = parse_sheet_range_addr("Input!B2:B10", "sheet1").unwrap();
        assert_eq!(sheet, "input");
        assert_eq!(from, (2, 2));
        assert_eq!(to, (10, 2));
    }

    #[test]
    fn parse_sheet_range_addr_without_sheet_prefix_uses_active_sheet() {
        let (sheet, from, to) = parse_sheet_range_addr("A1:B3", "sheet1").unwrap();
        assert_eq!(sheet, "sheet1");
        assert_eq!(from, (1, 1));
        assert_eq!(to, (3, 2));
    }

    #[test]
    fn parse_sheet_range_addr_rejects_invalid_range() {
        assert!(parse_sheet_range_addr("Input!not_a_range", "sheet1").is_none());
    }

    #[test]
    fn load_workbook_file_populates_cells_and_sets_active_sheet() {
        // Build a real .xlsx in-process via the existing writer (same
        // technique as lib.rs's diff_reader_tests) rather than shelling out
        // to the CLI binary — CARGO_BIN_EXE_* isn't available inside a
        // `cargo test --lib` unit test.
        let out_path = std::env::temp_dir().join("elixcee_vm_load_workbook_test.xlsx");
        let mut source_vm = Vm::new();
        source_vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(42),
            },
        );
        crate::save_workbook(&source_vm, out_path.to_str().unwrap()).unwrap();

        let mut vm = Vm::new();
        let names = vm.load_workbook_file(out_path.to_str().unwrap()).unwrap();
        assert_eq!(names, vec!["sheet1".to_string()]);
        assert_eq!(vm.active_sheet, "sheet1");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
    }

    #[test]
    fn populate_from_sheets_lowercases_a_mixed_case_sheet_name() {
        // Regression test for the bug found while extracting
        // `load_workbook_file` out of main.rs: real Excel files commonly
        // default to a sheet named "Sheet1" (capital S), and `save_workbook`
        // always lowercases names on write — so a fixture built via
        // `save_workbook` (as in the test above) can never exercise a
        // mixed-case name and would pass identically with or without the
        // lowercasing fix. Constructing a `WorkbookSheet` directly, as a
        // real XLSX reader would produce, closes that hole.
        let mut cells = std::collections::HashMap::new();
        cells.insert((1, 1), SheetCell::Integer(42));
        let sheets = vec![WorkbookSheet {
            name: "Input".to_string(),
            cells,
            sheet_id: None,
            workbook_rel_id: None,
            source_part_name: None,
            merged_ranges: Vec::new(),
            hidden_rows: Vec::new(),
            hidden_columns: Vec::new(),
            raw_style_indices: HashMap::new(),
            formulas: HashMap::new(),
            cell_number_formats: HashMap::new(),
        }];

        let mut vm = Vm::new();
        let names = vm.populate_from_sheets(sheets);

        assert_eq!(names, vec!["input".to_string()]);
        assert_eq!(vm.active_sheet, "input");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
        // Regression: `Vm::new()`'s default "sheet1" (pre-seeded so a
        // macro can write cells before any workbook is loaded) must not
        // survive into a loaded workbook whose real sheets are never named
        // "Sheet1" -- else the writer would carry it through as a genuine
        // extra empty sheet on save.
        assert!(!vm.sheets.contains_key("sheet1"));
    }

    #[test]
    fn populate_from_sheets_does_not_leak_the_default_sheet1_when_absent() {
        // Found via a real synthetic .xlsx round-trip: a two-sheet workbook
        // named "First"/"Second" (neither "Sheet1") gained an unrequested
        // third, empty "sheet1" on save. Root cause: `populate_from_sheets`
        // used to only ever *add* sheets via `ensure_sheet`, never clearing
        // `Vm::new()`'s pre-seeded default "sheet1" first.
        let sheets = vec![
            WorkbookSheet {
                name: "First".to_string(),
                cells: HashMap::new(),
                sheet_id: Some("5".to_string()),
                workbook_rel_id: None,
                source_part_name: None,
                merged_ranges: Vec::new(),
                hidden_rows: Vec::new(),
                hidden_columns: Vec::new(),
                raw_style_indices: HashMap::new(),
                formulas: HashMap::new(),
                cell_number_formats: HashMap::new(),
            },
            WorkbookSheet {
                name: "Second".to_string(),
                cells: HashMap::new(),
                sheet_id: Some("9".to_string()),
                workbook_rel_id: None,
                source_part_name: None,
                merged_ranges: Vec::new(),
                hidden_rows: Vec::new(),
                hidden_columns: Vec::new(),
                raw_style_indices: HashMap::new(),
                formulas: HashMap::new(),
                cell_number_formats: HashMap::new(),
            },
        ];

        let mut vm = Vm::new();
        assert!(
            vm.sheets.contains_key("sheet1"),
            "sanity: default present before load"
        );
        let names = vm.populate_from_sheets(sheets);

        assert_eq!(names, vec!["first".to_string(), "second".to_string()]);
        assert_eq!(
            vm.sheet_names(),
            vec!["first".to_string(), "second".to_string()]
        );
        assert!(!vm.sheets.contains_key("sheet1"));
    }

    #[test]
    fn populate_from_sheets_keeps_a_real_sheet1_when_present() {
        // A workbook whose first real sheet actually is "Sheet1" must keep
        // behaving exactly as before -- `ensure_sheet` re-inserting the same
        // key is a no-op, not a duplicate.
        let sheets = vec![WorkbookSheet {
            name: "Sheet1".to_string(),
            cells: HashMap::new(),
            sheet_id: Some("1".to_string()),
            workbook_rel_id: None,
            source_part_name: None,
            merged_ranges: Vec::new(),
            hidden_rows: Vec::new(),
            hidden_columns: Vec::new(),
            raw_style_indices: HashMap::new(),
            formulas: HashMap::new(),
            cell_number_formats: HashMap::new(),
        }];

        let mut vm = Vm::new();
        let names = vm.populate_from_sheets(sheets);

        assert_eq!(names, vec!["sheet1".to_string()]);
        assert_eq!(vm.sheet_names(), vec!["sheet1".to_string()]);
    }

    #[test]
    fn load_workbook_file_reports_a_clear_error_for_a_missing_file() {
        let mut vm = Vm::new();
        let err = vm
            .load_workbook_file("/nonexistent/path/does_not_exist.xlsx")
            .unwrap_err();
        assert!(err.starts_with("cannot read"), "{:?}", err);
    }

    #[test]
    fn deadline_none_means_unlimited_loop_execution() {
        // A loop well past the 256-iteration check gate must still run to
        // completion with no deadline set — the default, zero-overhead path.
        let mut vm = Vm::new();
        assert!(vm.deadline.is_none());
        let prog = parser::parse(
            "Sub MySub()\n    n = 0\n    For i = 1 To 2000\n        n = n + 1\n    Next i\nEnd Sub\n",
        )
        .unwrap();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.variables["n"], Variant::Integer(2000));
    }

    #[test]
    fn deadline_exceeded_stops_a_tight_for_loop_with_a_timeout_error() {
        let mut vm = Vm::new();
        vm.deadline = Some(std::time::Instant::now()); // already past
        let prog = parser::parse(
            "Sub MySub()\n    For i = 1 To 100000000\n        n = i\n    Next i\nEnd Sub\n",
        )
        .unwrap();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.starts_with("TIMEOUT:"), "{:?}", err);
    }

    #[test]
    fn deadline_exceeded_stops_a_tight_do_loop_with_a_timeout_error() {
        let mut vm = Vm::new();
        vm.deadline = Some(std::time::Instant::now());
        let prog = parser::parse(
            "Sub MySub()\n    i = 0\n    Do While i < 100000000\n        i = i + 1\n    Loop\nEnd Sub\n",
        )
        .unwrap();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.starts_with("TIMEOUT:"), "{:?}", err);
    }

    // ── Milestone B6b: Copy/Paste shape diagnosis + Clipboard state ─────────

    #[test]
    fn bare_copy_then_paste_special_round_trips_matching_shapes() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(1,2).Value = 20\n    \
             Cells(2,1).Value = 30\n    Cells(2,2).Value = 40\n    \
             Range(\"A1:B2\").Copy\n    Range(\"E1:F2\").PasteSpecial\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 5), Variant::Integer(10));
        assert_eq!(vm.get_cell(1, 6), Variant::Integer(20));
        assert_eq!(vm.get_cell(2, 5), Variant::Integer(30));
        assert_eq!(vm.get_cell(2, 6), Variant::Integer(40));
    }

    #[test]
    fn transpose_true_swaps_rows_and_columns_on_paste() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(1,2).Value = 20\n    \
             Range(\"A1:B1\").Copy\n    Range(\"E1:E2\").PasteSpecial Transpose:=True\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 5), Variant::Integer(10));
        assert_eq!(vm.get_cell(2, 5), Variant::Integer(20));
    }

    #[test]
    fn paste_shape_mismatch_is_a_hard_error_with_evidence() {
        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1:C10\").Copy\n    Range(\"E1:F10\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("shape mismatch"), "{:?}", err);
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::PasteShapeMismatch {
                source_addr,
                source_rows,
                source_cols,
                dest_addr,
                dest_rows,
                dest_cols,
                dest_row1,
                dest_col1,
                transpose,
                copy_span,
            }) => {
                assert_eq!(source_addr, "A1:C10");
                assert_eq!((source_rows, source_cols), (10, 3));
                assert_eq!(dest_addr, "E1:F10");
                assert_eq!((dest_rows, dest_cols), (10, 2));
                assert_eq!((dest_row1, dest_col1), (1, 5));
                assert!(!transpose);
                assert!(copy_span.is_some());
            }
            other => panic!("expected PasteShapeMismatch, got {:?}", other),
        }
    }

    #[test]
    fn paste_without_a_prior_copy_is_a_hard_error() {
        let prog = parser::parse("Sub MySub()\n    Range(\"A1\").PasteSpecial\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("Clipboard is empty"), "{:?}", err);
        assert_eq!(
            vm.take_resolution_failure(),
            Some(ResolutionFailureKind::PasteWithoutCopy {
                dest_addr: "A1".to_string()
            })
        );
    }

    // ── Milestone B7a: multi-area Range foundation ───────────────────────────

    #[test]
    fn rect_and_range_ref_helpers_report_dimensions_and_area_count() {
        let a = Rect {
            start_row: 1,
            start_col: 1,
            end_row: 10,
            end_col: 3,
        };
        assert_eq!((a.rows(), a.cols()), (10, 3));

        let single = RangeRef::single("sheet1".to_string(), a);
        assert!(single.is_single_area());
        assert_eq!(single.single_rect(), Some(&a));
        assert_eq!(single.cell_count(), 30);

        let b = Rect {
            start_row: 1,
            start_col: 5,
            end_row: 4,
            end_col: 5,
        };
        let multi = RangeRef {
            sheet: "sheet1".to_string(),
            areas: vec![a, b],
        };
        assert!(!multi.is_single_area());
        assert_eq!(multi.single_rect(), None);
        assert_eq!(multi.cell_count(), 30 + 4);
    }

    #[test]
    fn parse_multi_area_addr_returns_one_rect_for_a_plain_range() {
        assert_eq!(
            parse_multi_area_addr("A1:C10"),
            Some(vec![Rect {
                start_row: 1,
                start_col: 1,
                end_row: 10,
                end_col: 3
            }])
        );
    }

    #[test]
    fn parse_multi_area_addr_splits_multiple_comma_separated_pieces() {
        assert_eq!(
            parse_multi_area_addr("A1:A10,C1:C10"),
            Some(vec![
                Rect {
                    start_row: 1,
                    start_col: 1,
                    end_row: 10,
                    end_col: 1
                },
                Rect {
                    start_row: 1,
                    start_col: 3,
                    end_row: 10,
                    end_col: 3
                },
            ])
        );
    }

    #[test]
    fn parse_multi_area_addr_rejects_a_malformed_piece() {
        assert_eq!(parse_multi_area_addr("A1:A10,bogus"), None);
    }

    #[test]
    fn range_copy_of_a_multi_area_source_populates_clipboard_areas_without_snapshotting_cells() {
        let prog =
            parser::parse("Sub MySub()\n    Range(\"A1:A10,C1:C10\").Copy\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
        let clip = vm
            .clipboard
            .as_ref()
            .expect("Copy should have populated the clipboard");
        assert_eq!(
            clip.areas,
            vec![
                Rect {
                    start_row: 1,
                    start_col: 1,
                    end_row: 10,
                    end_col: 1
                },
                Rect {
                    start_row: 1,
                    start_col: 3,
                    end_row: 10,
                    end_col: 3
                },
            ]
        );
        assert!(
            clip.cells.is_empty(),
            "multi-area Copy shouldn't snapshot per-area cell values in v1: {:?}",
            clip.cells
        );
    }

    #[test]
    fn multi_area_source_pasted_into_single_area_destination_reports_the_completion_condition() {
        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1:A10,C1:C10\").Copy\n    Range(\"E1:F10\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("disjoint areas"), "{:?}", err);
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::MultiAreaToSingleAreaPaste {
                source_areas,
                destination_areas,
            }) => {
                assert_eq!(
                    source_areas,
                    vec![
                        Rect {
                            start_row: 1,
                            start_col: 1,
                            end_row: 10,
                            end_col: 1
                        },
                        Rect {
                            start_row: 1,
                            start_col: 3,
                            end_row: 10,
                            end_col: 3
                        },
                    ]
                );
                assert_eq!(
                    destination_areas,
                    vec![Rect {
                        start_row: 1,
                        start_col: 5,
                        end_row: 10,
                        end_col: 6
                    }]
                );
            }
            other => panic!("expected MultiAreaToSingleAreaPaste, got {:?}", other),
        }
    }

    #[test]
    fn multi_area_source_and_destination_with_differing_area_counts_reports_count_mismatch() {
        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1:A10,C1:C10,E1:E10\").Copy\n    \
             Range(\"G1:G10,I1:I10\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("3 areas"), "{:?}", err);
        assert!(err.contains("2 areas"), "{:?}", err);
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::MultiAreaCountMismatch {
                source_areas,
                destination_areas,
            }) => {
                assert_eq!(source_areas.len(), 3);
                assert_eq!(destination_areas.len(), 2);
            }
            other => panic!("expected MultiAreaCountMismatch, got {:?}", other),
        }
    }

    #[test]
    fn multi_area_source_and_destination_with_matching_counts_but_differing_shapes_reports_shape_mismatch()
     {
        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1:A10,C1:C10\").Copy\n    \
             Range(\"G1:G10,I1:J10\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("area 2"), "{:?}", err);
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::MultiAreaShapeMismatch {
                area_index,
                source_area,
                destination_area,
            }) => {
                assert_eq!(area_index, 2);
                assert_eq!(
                    source_area,
                    Rect {
                        start_row: 1,
                        start_col: 3,
                        end_row: 10,
                        end_col: 3
                    }
                );
                assert_eq!(
                    destination_area,
                    Rect {
                        start_row: 1,
                        start_col: 9,
                        end_row: 10,
                        end_col: 10
                    }
                );
            }
            other => panic!("expected MultiAreaShapeMismatch, got {:?}", other),
        }
    }

    #[test]
    fn single_area_source_pasted_into_multi_area_destination_reports_paste_unsupported() {
        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1:B10\").Copy\n    Range(\"E1:E10,G1:G10\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("not yet supported"), "{:?}", err);
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::MultiAreaPasteUnsupported {
                source_areas,
                destination_areas,
            }) => {
                assert_eq!(source_areas.len(), 1);
                assert_eq!(destination_areas.len(), 2);
            }
            other => panic!("expected MultiAreaPasteUnsupported, got {:?}", other),
        }
    }

    // ── Milestone B7c item 1: Range object variables ────────────────────────

    #[test]
    fn set_and_dot_value_read_write_a_range_object_through_to_the_sheet() {
        let vm = run("Sub MySub()\n    Set rng = Range(\"B2\")\n    rng.Value = 42\nEnd Sub\n");
        assert_eq!(vm.get_cell(2, 2), Variant::Integer(42));
    }

    #[test]
    fn set_reference_semantics_a_third_party_write_is_visible_through_every_alias() {
        // The discriminating case: `Set b = a` must NOT snapshot A1's
        // value at `Set`-time — it must alias the same coordinates, so a
        // write through a completely different path (a plain
        // `Range("A1").Value =`, not through `a` or `b`) is still visible
        // when reading `b.Value` afterwards. A value-copy implementation
        // would fail this (it would see whatever A1 held at `Set` time,
        // i.e. Empty).
        let vm = run(
            "Sub MySub()\n    Set a = Range(\"A1\")\n    Set b = a\n    \
             Range(\"A1\").Value = 9\n    x = b.Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(9));
    }

    #[test]
    fn range_object_value_read_for_a_multi_cell_area_returns_an_array() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(1,2).Value = 2\n    \
             Set rng = Range(\"A1:B1\")\n    x = rng.Value\nEnd Sub\n",
        );
        assert_eq!(
            vm.variables["x"],
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(2)])
        );
    }

    #[test]
    fn set_with_an_unresolved_bare_identifier_is_a_no_op_not_a_hard_error() {
        // `Set rng = Nothing` (and any other unmodeled bare object
        // keyword, e.g. a future `ActiveSheet` before item 6 lands) must
        // degrade gracefully — see the doc comment on `Stmt::Set`'s VM
        // handler.
        let vm = run("Sub MySub()\n    Set rng = Nothing\n    x = 1\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1));
    }

    #[test]
    fn range_object_copy_and_paste_round_trips_through_an_object_variable_source() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 7\n    Set rng = Range(\"A1\")\n    \
             rng.Copy\n    Range(\"C3\").PasteSpecial\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(3, 3), Variant::Integer(7));
    }

    // ── Milestone B7c item 2: Union ──────────────────────────────────────────

    #[test]
    fn union_combines_two_ranges_into_one_multi_area_object() {
        let vm = run(
            "Sub MySub()\n    Set u = Union(Range(\"A1:A2\"), Range(\"C1:C2\"))\n    \
             n = u.Areas.Count\nEnd Sub\n",
        );
        assert_eq!(vm.variables["n"], Variant::Integer(2));
    }

    #[test]
    fn union_accepts_object_variables_not_only_range_literals() {
        let vm = run(
            "Sub MySub()\n    Set a = Range(\"A1:A2\")\n    Set b = Range(\"C1:C2\")\n    \
             Set u = Union(a, b)\n    n = u.Areas.Count\nEnd Sub\n",
        );
        assert_eq!(vm.variables["n"], Variant::Integer(2));
    }

    // ── Milestone B7c item 3: .Areas ─────────────────────────────────────────

    #[test]
    fn areas_count_is_1_for_a_single_area_range() {
        let vm =
            run("Sub MySub()\n    Set rng = Range(\"A1:B2\")\n    n = rng.Areas.Count\nEnd Sub\n");
        assert_eq!(vm.variables["n"], Variant::Integer(1));
    }

    #[test]
    fn areas_index_returns_the_nth_single_area_range_1_based() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 11\n    Cells(1,3).Value = 33\n    \
             Set u = Range(\"A1,C1\")\n    Set second = u.Areas(2)\n    x = second.Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(33));
    }

    #[test]
    fn areas_index_out_of_range_is_a_runtime_error() {
        let prog = parser::parse(
            "Sub MySub()\n    Set u = Range(\"A1,C1\")\n    Set bad = u.Areas(3)\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("out of range"), "{:?}", err);
    }

    // ── Milestone B7c item 4: SpecialCells(xlCellTypeVisible) ────────────────

    #[test]
    fn specialcells_visible_excludes_a_hidden_row() {
        let mut vm = Vm::new();
        vm.sheet_visibility.insert(
            "sheet1".to_string(),
            SheetVisibility {
                hidden_rows: vec![Interval { start: 2, end: 2 }],
                hidden_columns: vec![],
            },
        );
        let prog = parser::parse(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    \
             Set rng = Range(\"A1:A3\")\n    Set vis = rng.SpecialCells(xlCellTypeVisible)\n    \
             n = vis.Areas.Count\nEnd Sub\n",
        )
        .unwrap();
        vm.run_sub(&prog, "mysub").unwrap();
        // Row 2 is hidden, splitting A1:A3 into two visible areas: A1 and A3.
        assert_eq!(vm.variables["n"], Variant::Integer(2));
    }

    #[test]
    fn specialcells_visible_is_the_whole_range_when_nothing_is_hidden() {
        let vm = run(
            "Sub MySub()\n    Set rng = Range(\"A1:A3\")\n    Set vis = rng.SpecialCells(xlCellTypeVisible)\n    \
             n = vis.Areas.Count\nEnd Sub\n",
        );
        assert_eq!(vm.variables["n"], Variant::Integer(1));
    }

    // ── Milestone B7c item 5: multi-area Copy/Paste ──────────────────────────

    #[test]
    fn matching_shape_multi_area_paste_actually_completes() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(1,3).Value = 3\n    \
             Range(\"A1:A1,C1:C1\").Copy\n    Range(\"E1:E1,G1:G1\").PasteSpecial\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 5), Variant::Integer(1));
        assert_eq!(vm.get_cell(1, 7), Variant::Integer(3));
    }

    #[test]
    fn matching_shape_multi_area_paste_from_a_union_object_variable() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    \
             Cells(1,3).Value = 30\n    Cells(2,3).Value = 40\n    \
             Set u = Union(Range(\"A1:A2\"), Range(\"C1:C2\"))\n    u.Copy\n    \
             Range(\"E1:E2,G1:G2\").PasteSpecial\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 5), Variant::Integer(10));
        assert_eq!(vm.get_cell(2, 5), Variant::Integer(20));
        assert_eq!(vm.get_cell(1, 7), Variant::Integer(30));
        assert_eq!(vm.get_cell(2, 7), Variant::Integer(40));
    }

    #[test]
    fn matching_shape_multi_area_paste_with_transpose_still_errors_instead_of_silently_mis_pasting()
    {
        // `transpose` isn't modeled for the multi-area execution path — it
        // must fall through to the pre-existing diagnose-only error rather
        // than silently writing UN-transposed data while claiming success.
        let prog = parser::parse(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(1,3).Value = 3\n    \
             Range(\"A1:A1,C1:C1\").Copy\n    \
             Range(\"E1:E1,G1:G1\").PasteSpecial Transpose:=True\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("not yet supported"), "{:?}", err);
        assert_eq!(vm.get_cell(1, 5), Variant::Empty);
        assert_eq!(vm.get_cell(1, 7), Variant::Empty);
    }

    #[test]
    fn cutcopymode_false_clears_the_clipboard_so_a_later_paste_fails() {
        let prog = parser::parse(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Range(\"A1\").Copy\n    \
             Application.CutCopyMode = False\n    Range(\"B1\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("Clipboard is empty"), "{:?}", err);
    }

    #[test]
    fn copy_destination_to_a_matching_shape_range_writes_there_correctly() {
        // Closes a latent bug: the old `RangeCopy` execution parsed `dst`
        // via `parse_cell_addr` (single-cell only) and silently fell back
        // to the source's own top-left cell for any real range address —
        // never noticed because no prior test exercised a range Destination.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    \
             Cells(3,1).Value = 3\n    Range(\"A1:A3\").Copy Destination:=Range(\"B1:B3\")\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 2), Variant::Integer(1));
        assert_eq!(vm.get_cell(2, 2), Variant::Integer(2));
        assert_eq!(vm.get_cell(3, 2), Variant::Integer(3));
    }

    #[test]
    fn copy_destination_shape_mismatch_is_also_a_hard_error() {
        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1:C10\").Copy Destination:=Range(\"E1:F10\")\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("shape mismatch"), "{:?}", err);
    }

    #[test]
    fn single_cell_source_fills_a_larger_destination_range_without_a_shape_error() {
        // Real Excel's well-known "paste one value into many cells" fill
        // behavior — not a shape mismatch, even though 1x1 != 10x1.
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 42\n    Range(\"A1\").Copy\n    \
             Range(\"B1:B10\").PasteSpecial\nEnd Sub\n",
        );
        for row in 1..=10 {
            assert_eq!(vm.get_cell(row, 2), Variant::Integer(42), "row {}", row);
        }
    }

    #[test]
    fn worksheet_paste_destination_writes_into_the_named_sheet() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 7\n    Range(\"A1\").Copy\n    \
             Worksheets(\"Sheet1\").Paste Destination:=Range(\"C1\")\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 3), Variant::Integer(7));
    }

    // ── Milestone B6c: sheet protection diagnosis ───────────────────────────

    #[test]
    fn protecting_a_sheet_blocks_a_later_cell_write() {
        let prog = parser::parse(
            "Sub MySub()\n    Worksheets(\"Sheet1\").Protect\n    Cells(1,1).Value = 1\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("protected"), "{:?}", err);
        assert_eq!(
            vm.take_resolution_failure(),
            Some(ResolutionFailureKind::SheetProtected {
                sheet: "sheet1".to_string()
            })
        );
    }

    #[test]
    fn unprotecting_a_sheet_restores_write_access() {
        let vm = run("Sub MySub()\n    Worksheets(\"Sheet1\").Protect\n    \
             Worksheets(\"Sheet1\").Unprotect\n    Cells(1,1).Value = 42\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
    }

    #[test]
    fn protection_does_not_block_reads() {
        let vm = run(
            "Sub MySub()\n    Cells(1,1).Value = 5\n    Worksheets(\"Sheet1\").Protect\n    \
             x = Cells(1,1).Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test]
    fn protecting_a_nonexistent_sheet_is_a_hard_error_unconditionally() {
        // Unconditional (not gated behind strict_resolution) — brand-new
        // construct, same precedent as `WorkbookQualifiedSheet`.
        let prog = parser::parse("Sub MySub()\n    Worksheets(\"NoSuchSheet\").Protect\nEnd Sub\n")
            .unwrap();
        let mut vm = Vm::new();
        assert!(!vm.strict_resolution);
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("not found"), "{:?}", err);
        assert!(matches!(
            vm.take_resolution_failure(),
            Some(ResolutionFailureKind::WorksheetNotFound(_))
        ));
    }

    #[test]
    fn protect_accepts_and_discards_a_password_kwarg() {
        let prog = parser::parse(
            "Sub MySub()\n    Worksheets(\"Sheet1\").Protect Password:=\"secret\"\n    \
             Cells(1,1).Value = 1\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("protected"), "{:?}", err);
    }

    #[test]
    fn protect_user_interface_only_true_does_not_block_macro_writes() {
        // Real Excel's UserInterfaceOnly:=True blocks manual UI edits but
        // not macro writes — this is the standard idiom for a sheet a
        // macro must keep writing to while the user can't touch it by hand.
        let vm = run(
            "Sub MySub()\n    Worksheets(\"Sheet1\").Protect UserInterfaceOnly:=True\n    \
             Cells(1,1).Value = 42\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
    }

    #[test]
    fn protect_user_interface_only_false_still_blocks_macro_writes() {
        let prog = parser::parse(
            "Sub MySub()\n    Worksheets(\"Sheet1\").Protect UserInterfaceOnly:=False\n    \
             Cells(1,1).Value = 1\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("protected"), "{:?}", err);
    }

    #[test]
    fn range_write_range_clear_and_copy_paste_are_all_blocked_by_protection() {
        let cases = [
            "Range(\"A1\").Value = 1",
            "Range(\"A1\").ClearContents",
            "Range(\"A1\").Copy Destination:=Range(\"B1\")",
        ];
        for stmt in cases {
            let src = format!(
                "Sub MySub()\n    Worksheets(\"Sheet1\").Protect\n    {}\nEnd Sub\n",
                stmt
            );
            let prog = parser::parse(&src).unwrap();
            let mut vm = Vm::new();
            let err = vm.run_sub(&prog, "mysub").unwrap_err();
            assert!(err.contains("protected"), "stmt {:?}: {:?}", stmt, err);
        }
    }

    #[test]
    fn sheets_delete_is_blocked_on_a_protected_sheet() {
        let prog = parser::parse(
            "Sub MySub()\n    Worksheets(\"Extra\").Cells(1,1).Value = 1\n    \
             Worksheets(\"Extra\").Protect\n    Sheets(\"Extra\").Delete\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("protected"), "{:?}", err);
    }

    // ── Milestone B6c2: merged-cell-aware Paste diagnosis ───────────────────

    /// There's no VBA construct to create a merge (out of scope — see the
    /// plan's non-goals), so tests inject `merged_ranges` directly, the same
    /// way a real `.xlsx`/`.ods` reader would via `populate_from_sheets`.
    fn vm_with_merge(sheet: &str, rect: MergeRect) -> Vm {
        let mut vm = Vm::new();
        vm.merged_ranges.insert(sheet.to_string(), vec![rect]);
        vm
    }

    #[test]
    fn paste_into_non_anchor_merged_cell_is_a_hard_error() {
        // Sheet1 has a merge B1:D1 (cols 2-4, row 1); pasting into C1 (a
        // covered cell of that merge, not its top-left B1) must fail.
        let mut vm = vm_with_merge("sheet1", ((1, 2), (1, 4)));
        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1\").Copy\n    Range(\"C1\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("top-left"), "{:?}", err);
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::PasteIntoNonAnchorMergedCell {
                dest_addr,
                dest_sheet,
                merged_range,
                ..
            }) => {
                assert_eq!(dest_addr, "C1");
                assert_eq!(dest_sheet, "sheet1");
                assert_eq!(merged_range, ((1, 2), (1, 4)));
            }
            other => panic!("expected PasteIntoNonAnchorMergedCell, got {:?}", other),
        }
    }

    #[test]
    fn paste_into_a_merges_own_anchor_cell_succeeds() {
        // Pasting into B1 itself (the merge's top-left) is the normal way
        // to write to a merged cell in real Excel — must not error.
        let mut vm = vm_with_merge("sheet1", ((1, 2), (1, 4)));
        let prog = parser::parse(
            "Sub MySub()\n    Cells(1,1).Value = 42\n    Range(\"A1\").Copy\n    \
             Range(\"B1\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.get_cell(1, 2), Variant::Integer(42));
    }

    #[test]
    fn paste_partial_merged_range_is_a_hard_error() {
        // Merge B1:D1 spans cols 2-4; a paste destination of only B1:C1
        // (cols 2-3) crosses part of it without covering it fully.
        let mut vm = vm_with_merge("sheet1", ((1, 2), (1, 4)));
        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1:B1\").Copy\n    Range(\"B1:C1\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("partially overlaps"), "{:?}", err);
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::PastePartialMergedRange {
                dest_addr,
                dest_sheet,
                conflicts,
                ..
            }) => {
                assert_eq!(dest_addr, "B1:C1");
                assert_eq!(dest_sheet, "sheet1");
                assert_eq!(conflicts, vec![((1, 2), (1, 4))]);
            }
            other => panic!("expected PastePartialMergedRange, got {:?}", other),
        }
    }

    #[test]
    fn paste_merge_layout_mismatch_reports_the_correct_root_cause() {
        // The user's own motivating example: same shape (10x3), but the
        // destination has a merged first row (E1:G1) the source doesn't.
        let mut vm = vm_with_merge("sheet1", ((1, 5), (1, 7)));
        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1:C10\").Copy\n    Range(\"E1:G10\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("layouts differ"), "{:?}", err);
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::PasteMergeLayoutMismatch {
                source_addr,
                source_sheet,
                dest_addr,
                dest_sheet,
                conflicts,
                ..
            }) => {
                assert_eq!(source_addr, "A1:C10");
                assert_eq!(source_sheet, "sheet1");
                assert_eq!(dest_addr, "E1:G10");
                assert_eq!(dest_sheet, "sheet1");
                assert_eq!(conflicts, vec![((1, 5), (1, 7))]);
            }
            other => panic!("expected PasteMergeLayoutMismatch, got {:?}", other),
        }
    }

    #[test]
    fn matching_merge_layouts_on_both_sides_paste_cleanly() {
        // Source A1:C10 has a merged first row (A1:C1); destination E1:G10
        // has a merged first row at the same relative position (E1:G1) —
        // identical layouts must not be flagged as a mismatch.
        let mut vm = Vm::new();
        vm.merged_ranges.insert(
            "sheet1".to_string(),
            vec![((1, 1), (1, 3)), ((1, 5), (1, 7))],
        );
        let prog = parser::parse(
            "Sub MySub()\n    Cells(2,1).Value = 42\n    Range(\"A1:C10\").Copy\n    \
             Range(\"E1:G10\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.get_cell(2, 5), Variant::Integer(42));
    }

    #[test]
    fn matching_merge_layouts_paste_cleanly_with_transpose() {
        // Source A1:C2 (2 rows x 3 cols) has a merge in its top row spanning
        // its last 2 cols (B1:C1, relative (0,1)-(0,2)). Transpose:=True
        // swaps rows/cols, so the matching destination shape is 3 rows x 2
        // cols (E1:F3); the merge that lines up after transposing is a
        // 2-row x 1-col merge in the *second column-worth-of-source's-rows*
        // position — E2:E3 (rows 2-3, col E only), not E1:E2. This position
        // (not just matching dimensions) only lines up if `do_paste`
        // actually applies the `(rr,rc) -> (rc,rr)` swap to the destination
        // merge's relative coordinates before comparing — a naive
        // non-transposed comparison would flag this as a mismatch instead.
        let mut vm = Vm::new();
        vm.merged_ranges.insert(
            "sheet1".to_string(),
            vec![((1, 2), (1, 3)), ((2, 5), (3, 5))],
        );
        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1:C2\").Copy\n    \
             Range(\"E1:F3\").PasteSpecial Transpose:=True\nEnd Sub\n",
        )
        .unwrap();
        vm.run_sub(&prog, "mysub").unwrap();
    }

    #[test]
    fn populate_from_sheets_threads_merged_ranges_into_the_vm() {
        // Every other B6c2 test injects `vm.merged_ranges` directly
        // (`vm_with_merge`), bypassing `populate_from_sheets` entirely — so
        // the actual reader -> VM wiring added to `populate_from_sheets`
        // (the `if !sheet_data.merged_ranges.is_empty() { .. }` block) was
        // never exercised by any test. Mixed-case "Input" also confirms the
        // merge map is keyed lowercase, same as `active_sheet`/`sheets`.
        let mut cells = std::collections::HashMap::new();
        cells.insert((1, 1), SheetCell::Integer(1));
        let sheets = vec![WorkbookSheet {
            name: "Input".to_string(),
            cells,
            sheet_id: None,
            workbook_rel_id: None,
            source_part_name: None,
            merged_ranges: vec![((1, 2), (1, 4))],
            hidden_rows: Vec::new(),
            hidden_columns: Vec::new(),
            raw_style_indices: HashMap::new(),
            formulas: HashMap::new(),
            cell_number_formats: HashMap::new(),
        }];
        let mut vm = Vm::new();
        vm.populate_from_sheets(sheets);

        let prog = parser::parse(
            "Sub MySub()\n    Range(\"A1\").Copy\n    Range(\"C1\").PasteSpecial\nEnd Sub\n",
        )
        .unwrap();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(err.contains("top-left"), "{:?}", err);
        match vm.take_resolution_failure() {
            Some(ResolutionFailureKind::PasteIntoNonAnchorMergedCell { dest_sheet, .. }) => {
                assert_eq!(dest_sheet, "input");
            }
            other => panic!("expected PasteIntoNonAnchorMergedCell, got {:?}", other),
        }
    }

    // ── P1 remainder: merge_cells / unmerge_cells ────────────────────────────

    #[test]
    fn merge_cells_rejects_a_single_cell_range() {
        let mut vm = Vm::new();
        let err = vm.merge_cells("sheet1", 1, 1, 1, 1).unwrap_err();
        assert!(err.contains("at least 2 cells"), "{:?}", err);
    }

    #[test]
    fn merge_cells_rejects_a_range_that_overlaps_an_existing_merge() {
        let mut vm = vm_with_merge("sheet1", ((1, 2), (1, 4))); // B1:D1
        let err = vm.merge_cells("sheet1", 1, 3, 2, 5).unwrap_err(); // C1:E2 overlaps
        assert!(err.contains("overlap"), "{:?}", err);
        assert_eq!(vm.merged_ranges.get("sheet1").unwrap().len(), 1); // rejected merge not added
    }

    #[test]
    fn merge_cells_allows_a_non_overlapping_second_merge_on_the_same_sheet() {
        let mut vm = vm_with_merge("sheet1", ((1, 2), (1, 4))); // B1:D1
        vm.merge_cells("sheet1", 3, 1, 3, 2).unwrap(); // A3:B3, no overlap
        assert_eq!(
            vm.merged_ranges.get("sheet1").unwrap(),
            &vec![((1, 2), (1, 4)), ((3, 1), (3, 2))]
        );
    }

    #[test]
    fn merge_cells_does_not_touch_existing_cell_values_in_the_covered_range() {
        let mut vm = Vm::new();
        vm.write_rect(
            "sheet1",
            (1, 1),
            &[vec![Variant::Integer(1), Variant::Integer(2)]],
        );
        vm.merge_cells("sheet1", 1, 1, 1, 2).unwrap();
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(1));
        assert_eq!(vm.get_cell(1, 2), Variant::Integer(2));
    }

    #[test]
    fn unmerge_cells_removes_an_exact_match() {
        let mut vm = vm_with_merge("sheet1", ((1, 2), (1, 4))); // B1:D1
        vm.unmerge_cells("sheet1", 1, 2, 1, 4).unwrap();
        assert!(vm.merged_ranges.get("sheet1").unwrap().is_empty());
    }

    #[test]
    fn unmerge_cells_errors_on_no_match_instead_of_silently_no_opping() {
        let mut vm = Vm::new();
        let err = vm.unmerge_cells("sheet1", 1, 1, 1, 2).unwrap_err();
        assert!(err.contains("no merge found"), "{:?}", err);
    }

    #[test]
    fn unmerge_cells_errors_on_a_partial_overlap_that_is_not_an_exact_match() {
        let mut vm = vm_with_merge("sheet1", ((1, 2), (1, 4))); // B1:D1
        let err = vm.unmerge_cells("sheet1", 1, 2, 1, 3).unwrap_err(); // B1:C1, not exact
        assert!(err.contains("no merge found"), "{:?}", err);
        assert_eq!(vm.merged_ranges.get("sheet1").unwrap().len(), 1); // original merge untouched
    }

    // ── Milestone B7b: hidden row/column metadata foundation ────────────────

    #[test]
    fn populate_from_sheets_threads_hidden_rows_and_columns_into_the_vm() {
        // Mixed-case "Input" confirms `sheet_visibility` is keyed lowercase,
        // same as `merged_ranges`/`active_sheet`/`sheets`.
        let mut cells = std::collections::HashMap::new();
        cells.insert((1, 1), SheetCell::Integer(1));
        let sheets = vec![WorkbookSheet {
            name: "Input".to_string(),
            cells,
            sheet_id: None,
            workbook_rel_id: None,
            source_part_name: None,
            merged_ranges: Vec::new(),
            hidden_rows: vec![(3, 5)],
            hidden_columns: vec![(2, 2)],
            raw_style_indices: HashMap::new(),
            formulas: HashMap::new(),
            cell_number_formats: HashMap::new(),
        }];
        let mut vm = Vm::new();
        vm.populate_from_sheets(sheets);

        let prog = parser::parse("Sub MySub()\n    Range(\"A1:C10\").Copy\nEnd Sub\n").unwrap();
        vm.run_sub(&prog, "mysub").unwrap();

        let obs = vm
            .hidden_cells_observation()
            .expect("should observe hidden rows/columns overlapping the copy");
        assert_eq!(obs.sheet, "input");
        assert_eq!(obs.address, "A1:C10");
        assert_eq!(obs.rows, 10);
        assert_eq!(obs.columns, 3);
        assert_eq!(obs.hidden_rows, vec![Interval { start: 3, end: 5 }]);
        assert_eq!(obs.hidden_columns, vec![Interval { start: 2, end: 2 }]);
        assert_eq!(obs.total_cells, 30);
        // (10 rows - 3 hidden) * (3 cols - 1 hidden) = 7 * 2 = 14.
        assert_eq!(obs.visible_cells, 14);
    }

    #[test]
    fn hidden_cells_observation_is_none_for_a_multi_area_copy() {
        let mut cells = std::collections::HashMap::new();
        cells.insert((1, 1), SheetCell::Integer(1));
        let sheets = vec![WorkbookSheet {
            name: "Sheet1".to_string(),
            cells,
            sheet_id: None,
            workbook_rel_id: None,
            source_part_name: None,
            merged_ranges: Vec::new(),
            hidden_rows: vec![(3, 5)],
            hidden_columns: Vec::new(),
            raw_style_indices: HashMap::new(),
            formulas: HashMap::new(),
            cell_number_formats: HashMap::new(),
        }];
        let mut vm = Vm::new();
        vm.populate_from_sheets(sheets);

        let prog =
            parser::parse("Sub MySub()\n    Range(\"A1:A10,C1:C10\").Copy\nEnd Sub\n").unwrap();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.hidden_cells_observation(), None);
    }

    #[test]
    fn hidden_cells_observation_is_none_when_no_hidden_interval_overlaps_the_range() {
        let mut cells = std::collections::HashMap::new();
        cells.insert((1, 1), SheetCell::Integer(1));
        let sheets = vec![WorkbookSheet {
            name: "Sheet1".to_string(),
            cells,
            sheet_id: None,
            workbook_rel_id: None,
            source_part_name: None,
            merged_ranges: Vec::new(),
            hidden_rows: vec![(50, 60)],
            hidden_columns: Vec::new(),
            raw_style_indices: HashMap::new(),
            formulas: HashMap::new(),
            cell_number_formats: HashMap::new(),
        }];
        let mut vm = Vm::new();
        vm.populate_from_sheets(sheets);

        let prog = parser::parse("Sub MySub()\n    Range(\"A1:C10\").Copy\nEnd Sub\n").unwrap();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.hidden_cells_observation(), None);
    }

    #[test]
    fn hidden_cells_observation_is_none_when_nothing_was_copied() {
        assert_eq!(Vm::new().hidden_cells_observation(), None);
    }

    // ── Phase 2C: Mod / \ / ^ / infix And Or Xor Not ─────────────────────────

    #[test]
    fn mod_operator_computes_modulus() {
        let vm = run("Sub MySub()\n    a = 7 Mod 3\n    b = -7 Mod 3\n    c = 7 Mod -3\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        // Result sign follows the dividend (left operand), same as Rust's `%`.
        assert_eq!(vm.variables["b"], Variant::Integer(-1));
        assert_eq!(vm.variables["c"], Variant::Integer(1));
    }

    #[test]
    fn intdiv_operator_truncates_toward_zero() {
        let vm = run("Sub MySub()\n    a = 7 \\ 2\n    b = -7 \\ 2\n    c = 2 \\ 3\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(3));
        assert_eq!(vm.variables["b"], Variant::Integer(-3)); // truncated, not floored (-4)
        assert_eq!(vm.variables["c"], Variant::Integer(0));
    }

    #[test]
    fn intdiv_and_mod_round_fractional_operands_half_to_even_first() {
        // 0.5 rounds to 0 (nearest even), not 1 — so `5 \ 0.5` is `5 \ 0`.
        let prog =
            parser::parse("Sub MySub()\n    Cells(1, 1).Value = 5 \\ 0.5\nEnd Sub\n").unwrap();
        assert!(Vm::new().run_sub(&prog, "mysub").is_err());
        // 2.5 rounds to 2 (nearest even), not 3.
        let vm = run("Sub MySub()\n    a = 2.5 \\ 4\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(0));
    }

    #[test]
    fn pow_operator_computes_exponentiation() {
        let vm = run("Sub MySub()\n    a = 2 ^ 3\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(8));
    }

    #[test]
    fn pow_binds_tighter_than_unary_minus() {
        // -2 ^ 2 is -(2 ^ 2) = -4, not (-2) ^ 2 = 4.
        let vm = run("Sub MySub()\n    a = -2 ^ 2\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(-4));
    }

    #[test]
    fn pow_accepts_a_negative_exponent() {
        let vm = run("Sub MySub()\n    a = 2 ^ -1\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Float(0.5));
    }

    #[test]
    fn pow_is_left_associative() {
        // 2 ^ 3 ^ 2 is (2 ^ 3) ^ 2 = 64, not 2 ^ (3 ^ 2) = 512.
        let vm = run("Sub MySub()\n    a = 2 ^ 3 ^ 2\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(64));
    }

    #[test]
    fn full_precedence_chain_matches_real_vba() {
        // 2 + 3 * 2 ^ 2 = 2 + 3*4 = 2 + 12 = 14 (not 20, not 100).
        let vm = run("Sub MySub()\n    a = 2 + 3 * 2 ^ 2\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(14));
    }

    #[test]
    fn infix_and_or_xor_compute_boolean_logic() {
        let vm = run(
            "Sub MySub()\n    a = True And False\n    b = True Or False\n    c = True Xor True\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Boolean(false));
        assert_eq!(vm.variables["b"], Variant::Boolean(true));
        assert_eq!(vm.variables["c"], Variant::Boolean(false));
    }

    #[test]
    fn and_or_xor_do_numeric_bitwise_math_on_non_boolean_operands() {
        let vm = run("Sub MySub()\n    a = 6 And 3\n    b = 6 Or 1\n    c = 5 Xor 1\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(2));
        assert_eq!(vm.variables["b"], Variant::Integer(7));
        assert_eq!(vm.variables["c"], Variant::Integer(4));
    }

    #[test]
    fn not_binds_looser_than_comparison_and_tighter_than_and() {
        // `Not a And b` is `(Not a) And b`, not `Not (a And b)`.
        let vm =
            run("Sub MySub()\n    a = Not False And True\n    b = Not (False And True)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Boolean(true));
        // `Not x = y` is `Not (x = y)`, not `(Not x) = y`.
        let vm2 = run("Sub MySub()\n    a = Not 1 = 2\nEnd Sub\n");
        assert_eq!(vm2.variables["a"], Variant::Boolean(true)); // Not(1=2) = Not(False) = True
    }

    #[test]
    fn if_not_condition_works() {
        let vm = run(
            "Sub MySub()\n    x = False\n    If Not x Then\n        a = 1\n    Else\n        a = 2\n    End If\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Integer(1));
    }

    #[test]
    fn or_binds_looser_than_and() {
        // a Or b And c is a Or (b And c), not (a Or b) And c.
        let vm = run("Sub MySub()\n    a = True Or False And False\nEnd Sub\n");
        // True Or (False And False) = True Or False = True.
        // (True Or False) And False would be False — this distinguishes them.
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
    }

    #[test]
    fn parenthesized_expression_supports_and_or_xor() {
        let vm = run("Sub MySub()\n    a = (True And False) Or True\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
    }

    #[test]
    fn concat_binds_looser_than_plus_minus() {
        // "x" & 1 + 2 is "x" & (1 + 2) = "x3", not ("x" & 1) + 2. Previously
        // `&` was folded into the same precedence tier as `+`/`-`
        // (equal precedence, left-to-right), which would have given the
        // latter (and likely a runtime type error, since "x1" + 2 isn't
        // valid VBA arithmetic).
        let vm = run("Sub MySub()\n    a = \"x\" & 1 + 2\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Str("x3".into()));
    }

    #[test]
    fn range_value_write_with_and_or_mod_intdiv_pow_all_parse_and_execute() {
        // The exact constructs the integration review found broken at parse
        // time — confirms they now reach the VM and produce real values, not
        // just "doesn't error".
        let vm = run(
            "Sub MySub()\n    Range(\"A1\").Value = 2 Mod 3\n    Range(\"A2\").Value = 7 \\ 3\n    Range(\"A3\").Value = 2 ^ 3\n    Range(\"A4\").Value = True And False\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(2));
        assert_eq!(vm.get_cell(2, 1), Variant::Integer(2));
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(8));
        assert_eq!(vm.get_cell(4, 1), Variant::Boolean(false));
    }

    // ── Phase 2C: With Range(...) ─────────────────────────────────────────────

    #[test]
    fn with_range_bare_value_reads_and_writes_the_with_target() {
        let vm = run(
            "Sub MySub()\n    Cells(1, 1).Value = 5\n    With Range(\"A1\")\n        .Value = .Value + 1000\n    End With\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(1005));
    }

    #[test]
    fn with_range_bare_value_write_alone() {
        let vm = run(
            "Sub MySub()\n    With Range(\"B2\")\n        .Value = 42\n    End With\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(2, 2), Variant::Integer(42));
    }

    #[test]
    fn with_range_nested_range_reference_still_works() {
        // A nested `.Range(...)`/`.Cells(...)` inside a `With Range(...)`
        // body is its own independent reference, not the With's own target
        // — same convention `parse_with_dot_stmt` already used for
        // `With Sheets(...)`.
        let vm = run(
            "Sub MySub()\n    With Range(\"A1\")\n        .Value = 1\n        .Range(\"B1\").Value = 2\n    End With\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(1));
        assert_eq!(vm.get_cell(1, 2), Variant::Integer(2));
    }

    // ── Phase 2C items 7/8: Set ws = ActiveSheet / Set wb = ThisWorkbook ─────

    #[test]
    fn set_ws_activesheet_then_range_and_cells_write_and_read() {
        // Previously a silent no-op (see `Stmt::Set`'s old comment) — `ws`
        // would stay unset, and `ws.Cells(...)`/`ws.Range(...)` wouldn't
        // even parse to anything meaningful. Both write and read now
        // actually reach the sheet.
        let vm = run(
            "Sub MySub()\n    Set ws = ActiveSheet\n    ws.Cells(1, 1).Value = 5\n    ws.Range(\"B2\").Value = 9\n    x = ws.Cells(1, 1).Value\n    y = ws.Range(\"B2\").Value\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(5));
        assert_eq!(vm.get_cell(2, 2), Variant::Integer(9));
        assert_eq!(vm.variables["x"], Variant::Integer(5));
        assert_eq!(vm.variables["y"], Variant::Integer(9));
    }

    #[test]
    fn set_ws_activesheet_captures_a_snapshot_not_a_dynamic_reference() {
        // Real VBA's `Set` fixes a Worksheet reference's identity at
        // assignment time — unlike the bare `ActiveSheet` keyword itself
        // (dynamic; see `activesheet_tracks_the_active_sheet_after_it_
        // changes`), `ws` must keep pointing at Sheet2 even after the
        // active sheet reverts to Sheet1 when the `With` block ends.
        let vm = run(
            "Sub MySub()\n    With Sheets(\"Sheet2\")\n        Set ws = ActiveSheet\n    End With\n    ws.Cells(1, 1).Value = 42\nEnd Sub\n",
        );
        let cell = vm
            .get_sheet_cells("sheet2")
            .and_then(|s| s.get(&(1, 1)))
            .map(|c| c.value.clone());
        assert_eq!(cell, Some(Variant::Integer(42)));
        // And Sheet1 (still the active sheet) is untouched.
        assert_eq!(vm.get_cell(1, 1), Variant::Empty);
    }

    #[test]
    fn set_wb_thisworkbook_then_worksheets_write_targets_the_named_sheet() {
        let vm = run(
            "Sub MySub()\n    Set wb = ThisWorkbook\n    wb.Worksheets(\"Data\").Cells(2, 3).Value = 77\nEnd Sub\n",
        );
        let cell = vm
            .get_sheet_cells("data")
            .and_then(|s| s.get(&(2, 3)))
            .map(|c| c.value.clone());
        assert_eq!(cell, Some(Variant::Integer(77)));
    }

    #[test]
    fn set_wb_activeworkbook_then_sheets_read() {
        let vm = run(
            "Sub MySub()\n    Sheets(\"Data\").Cells(1, 1).Value = 42\n    Set wb = ActiveWorkbook\n    x = wb.Sheets(\"Data\").Range(\"A1\").Value\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    // ── ObjectRef::Nothing — unset / cleared object variables ───────────────

    fn run_err(code: &str) -> String {
        let prog = parser::parse(code).unwrap();
        Vm::new().run_sub(&prog, "mysub").unwrap_err()
    }

    #[test]
    fn dim_as_range_registers_an_unset_object_variable() {
        let vm = run("Sub MySub()\n    Dim r As Range\nEnd Sub\n");
        assert_eq!(vm.object_variables.get("r"), Some(&ObjectRef::Nothing));
    }

    #[test]
    fn member_write_through_a_never_set_object_variable_raises_error_91() {
        assert_eq!(
            run_err("Sub MySub()\n    Dim r As Range\n    r.Value = 5\nEnd Sub\n"),
            OBJECT_NOT_SET
        );
    }

    #[test]
    fn member_read_through_a_never_set_object_variable_raises_error_91() {
        assert_eq!(
            run_err("Sub MySub()\n    Dim r As Range\n    x = r.Value\nEnd Sub\n"),
            OBJECT_NOT_SET
        );
    }

    #[test]
    fn set_nothing_clears_the_reference() {
        assert_eq!(
            run_err(
                "Sub MySub()\n    Dim r As Range\n    Set r = Range(\"A1\")\n    \
                 Set r = Nothing\n    r.Value = 5\nEnd Sub\n"
            ),
            OBJECT_NOT_SET
        );
    }

    #[test]
    fn set_nothing_clears_only_that_variable_not_an_alias() {
        // The alias case: `Set r2 = r` copies the reference into r2's own
        // slot, so clearing r afterwards must leave r2 fully live.
        let vm = run(
            "Sub MySub()\n    Range(\"B1\").Value = 42\n    Dim r As Range\n    \
             Dim r2 As Range\n    Set r = Range(\"B1\")\n    Set r2 = r\n    \
             Set r = Nothing\n    x = r2.Value\nEnd Sub\n",
        );
        assert_eq!(vm.object_variables.get("r"), Some(&ObjectRef::Nothing));
        assert!(matches!(
            vm.object_variables.get("r2"),
            Some(ObjectRef::Range(_))
        ));
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    #[test]
    fn is_nothing_reflects_each_variables_own_state() {
        let vm = run("Sub MySub()\n    Dim r As Range\n    Dim r2 As Range\n    \
             a = (r Is Nothing)\n    Set r = Range(\"A1\")\n    b = (r Is Nothing)\n    \
             Set r2 = r\n    Set r = Nothing\n    c = (r Is Nothing)\n    \
             d = (r2 Is Nothing)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Boolean(false));
        assert_eq!(vm.variables["c"], Variant::Boolean(true));
        assert_eq!(vm.variables["d"], Variant::Boolean(false));
    }

    #[test]
    fn a_scalar_variable_never_acquires_object_state() {
        // Scalar (`x = 5`) and object (`Set r = ...`) assignment are
        // genuinely different in VBA and live in different namespaces —
        // Nothing-tracking must not leak into the scalar path.
        let vm = run("Sub MySub()\n    Dim x As Long\n    x = 5\n    x = x + 1\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(6));
        assert!(!vm.object_variables.contains_key("x"));
    }

    #[test]
    fn a_non_object_name_still_auto_creates_a_record_on_field_write() {
        // Only a *registered* object variable may raise error 91. A name
        // that isn't one at all keeps its pre-existing record behavior.
        let vm = run("Sub MySub()\n    p.x = 3\n    y = p.x\nEnd Sub\n");
        assert_eq!(vm.variables["y"], Variant::Integer(3));
    }

    // ── Runtime With stack ──────────────────────────────────────────────────
    // With-target resolution is a runtime mechanism now, not a parse-time
    // textual rewrite. These pin the properties only a real runtime stack
    // has: a computed target, resolution at arbitrary AST depth, evaluate-
    // once-on-entry, and push/pop discipline that survives early exits.

    #[test]
    fn with_computed_cells_target_resolves_at_runtime() {
        let vm = run(
            "Sub MySub()\n    r = 2\n    c = 3\n    With Cells(r, c)\n        .Value = 42\n    \
             End With\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(2, 3), Variant::Integer(42));
    }

    #[test]
    fn with_target_is_evaluated_once_on_block_entry() {
        // Reassigning the index variable inside the body must not retarget
        // the block — the object expression is evaluated once, on entry.
        let vm = run(
            "Sub MySub()\n    i = 1\n    With Cells(i, 1)\n        i = 5\n        .Value = 3\n    \
             End With\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(3));
        assert_eq!(vm.get_cell(5, 1), Variant::Empty);
    }

    #[test]
    fn bare_dot_member_resolves_inside_nested_block_constructs() {
        let vm = run(
            "Sub MySub()\n    With Range(\"A1\")\n        For i = 1 To 2\n            \
             If i = 2 Then\n                .Value = 4\n            End If\n        Next i\n    \
             End With\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(4));
    }

    #[test]
    fn with_object_variable_target_resolves_against_its_reference() {
        let vm = run(
            "Sub MySub()\n    Dim rng As Range\n    Set rng = Range(\"C3\")\n    With rng\n        \
             .Value = 5\n    End With\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(3, 3), Variant::Integer(5));
    }

    #[test]
    fn with_worksheet_variable_target_qualifies_cells_to_that_sheet() {
        let vm = run(
            "Sub MySub()\n    Sheets(\"Data\").Range(\"A1\").Value = 0\n    \
             Set ws = ActiveSheet\n    Sheets(\"Data\").Select\n    With ws\n        \
             .Cells(1, 1).Value = 42\n    End With\nEnd Sub\n",
        );
        // `ws` captured Sheet1 (the active sheet at Set time), so the write
        // lands there regardless of anything else.
        assert_eq!(
            vm.get_sheet_cells("sheet1")
                .and_then(|s| s.get(&(1, 1)))
                .map(|c| c.value.clone()),
            Some(Variant::Integer(42))
        );
    }

    #[test]
    fn three_levels_of_nested_with_restore_each_outer_target() {
        let vm = run(
            "Sub MySub()\n    With Range(\"A1\")\n        .Value = 1\n        \
             With Range(\"B1\")\n            .Value = 2\n            With Range(\"C1\")\n                \
             .Value = 3\n            End With\n            .Value = .Value + 20\n        End With\n        \
             .Value = .Value + 10\n    End With\nEnd Sub\n",
        );
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(11));
        assert_eq!(vm.get_cell(1, 2), Variant::Integer(22));
        assert_eq!(vm.get_cell(1, 3), Variant::Integer(3));
    }

    #[test]
    fn the_with_stack_is_empty_again_after_a_block_completes() {
        let vm =
            run("Sub MySub()\n    With Range(\"A1\")\n        .Value = 1\n    End With\nEnd Sub\n");
        assert!(vm.with_stack.is_empty());
    }

    #[test]
    fn the_with_stack_does_not_leak_after_an_exit_sub_inside_a_with_body() {
        // The pop must happen on the early-exit path too — a leaked entry
        // would silently mis-resolve whatever `.member` ran next.
        let vm = run(
            "Sub MySub()\n    With Range(\"A1\")\n        .Value = 1\n        Exit Sub\n    \
             End With\nEnd Sub\n",
        );
        assert!(vm.with_stack.is_empty());
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(1));
    }

    #[test]
    fn the_with_stack_does_not_leak_after_a_runtime_error_inside_a_with_body() {
        let prog = parser::parse(
            "Sub MySub()\n    With Range(\"A1\")\n        .Value = 1 / 0\n    End With\nEnd Sub\n",
        )
        .unwrap();
        let mut vm = Vm::new();
        assert!(vm.run_sub(&prog, "mysub").is_err());
        assert!(vm.with_stack.is_empty());
    }

    #[test]
    fn a_bare_dot_member_with_no_enclosing_with_block_is_a_runtime_error() {
        // It used to be a *parse* error; a bare `.member` is now a general
        // statement form, so the "no target" case moves to runtime rather
        // than disappearing.
        assert_eq!(
            run_err("Sub MySub()\n    .Value = 1\nEnd Sub\n"),
            OBJECT_NOT_SET
        );
    }

    #[test]
    fn a_with_block_over_an_unset_object_variable_raises_error_91() {
        assert_eq!(
            run_err(
                "Sub MySub()\n    Dim r As Range\n    With r\n        .Value = 1\n    End With\nEnd Sub\n"
            ),
            OBJECT_NOT_SET
        );
    }

    // ── Variant::Null — VBA's "no valid data" value ─────────────────────────
    // Every expectation below is the documented rule from Microsoft's own VBA
    // language reference (the +, -, &, comparison, And/Or/Xor/Not operator
    // pages and the If...Then...Else statement page), not elixcee's prior
    // behavior — which folded Null into Empty and so got all of them wrong.

    #[test]
    fn null_is_distinct_from_empty() {
        let vm = run(
            "Sub MySub()\n    Dim e\n    n = Null\n    a = IsNull(n)\n    b = IsEmpty(n)\n    \
             c = IsNull(e)\n    d = IsEmpty(e)\n    t = TypeName(n)\n    v = VarType(n)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Boolean(false));
        assert_eq!(vm.variables["c"], Variant::Boolean(false));
        assert_eq!(vm.variables["d"], Variant::Boolean(true));
        assert_eq!(vm.variables["t"], Variant::Str("Null".into()));
        assert_eq!(vm.variables["v"], Variant::Integer(1)); // vbNull
    }

    #[test]
    fn arithmetic_propagates_null_from_either_side() {
        let vm = run(
            "Sub MySub()\n    n = Null\n    a = n + 5\n    b = 5 + n\n    c = n - 1\n    \
             d = n * 3\n    e = 2 + 3\nEnd Sub\n",
        );
        for k in ["a", "b", "c", "d"] {
            assert_eq!(vm.variables[k], Variant::Null, "{} should be Null", k);
        }
        // Ordinary arithmetic is untouched.
        assert_eq!(vm.variables["e"], Variant::Integer(5));
    }

    #[test]
    fn concat_propagates_null_only_when_both_sides_are_null() {
        let vm = run(
            "Sub MySub()\n    n = Null\n    a = n & \"x\"\n    b = \"x\" & n\n    \
             c = n & n\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Str("x".into()));
        assert_eq!(vm.variables["b"], Variant::Str("x".into()));
        assert_eq!(vm.variables["c"], Variant::Null);
    }

    #[test]
    fn every_comparison_operator_propagates_null() {
        let vm = run(
            "Sub MySub()\n    n = Null\n    a = (5 < n)\n    b = (5 <= n)\n    c = (5 > n)\n    \
             d = (5 >= n)\n    e = (5 = n)\n    f = (5 <> n)\n    g = (n = n)\n    \
             h = (3 < 5)\nEnd Sub\n",
        );
        for k in ["a", "b", "c", "d", "e", "f", "g"] {
            assert_eq!(vm.variables[k], Variant::Null, "{} should be Null", k);
        }
        // Ordinary comparison is untouched.
        assert_eq!(vm.variables["h"], Variant::Boolean(true));
    }

    #[test]
    fn logical_operators_follow_the_documented_three_valued_tables() {
        // The two rows where Null does NOT propagate — the answer is already
        // determined without the missing operand.
        let vm = run(
            "Sub MySub()\n    n = Null\n    a = (False And n)\n    b = (True And n)\n    \
             c = (True Or n)\n    d = (False Or n)\n    e = (True Xor n)\n    \
             f = (Not n)\nEnd Sub\n",
        );
        assert_eq!(vm.variables["a"], Variant::Boolean(false));
        assert_eq!(vm.variables["b"], Variant::Null);
        assert_eq!(vm.variables["c"], Variant::Boolean(true));
        assert_eq!(vm.variables["d"], Variant::Null);
        assert_eq!(vm.variables["e"], Variant::Null);
        assert_eq!(vm.variables["f"], Variant::Null);
    }

    #[test]
    fn a_null_condition_is_treated_as_false_not_an_error() {
        // Documented on the If...Then...Else page: "If condition is Null,
        // condition is treated as False."
        let vm = run(
            "Sub MySub()\n    n = Null\n    taken = 0\n    If n Then\n        taken = 1\n    \
             Else\n        taken = 2\n    End If\nEnd Sub\n",
        );
        assert_eq!(vm.variables["taken"], Variant::Integer(2));
    }

    #[test]
    fn a_null_reaching_a_genuinely_numeric_context_raises_error_94() {
        assert_eq!(
            run_err("Sub MySub()\n    n = Null\n    x = Abs(n)\nEnd Sub\n"),
            "Invalid use of Null"
        );
    }

    #[test]
    fn a_non_numeric_string_operand_of_an_arithmetic_operator_is_a_type_mismatch() {
        // Documented on the + operator page: "One expression is a numeric
        // data type and the other is a String | A `Type mismatch` error
        // occurs." Scoped to the arithmetic operators — `to_f64`'s own
        // message, and its ~53 other call sites, are unchanged.
        assert_eq!(
            run_err(
                "Sub MySub()\n    Dim v1, v2\n    v1 = \"abc\"\n    v2 = 3\n    x = v1 + v2\nEnd Sub\n"
            ),
            "Type mismatch"
        );
    }

    #[test]
    fn a_numeric_looking_string_still_adds_rather_than_type_mismatching() {
        // The complement of the test above: `arith_to_f64` must only fire on
        // a string that genuinely can't convert.
        let vm = run(
            "Sub MySub()\n    Dim v1, v2\n    v1 = \"34\"\n    v2 = 6\n    x = v1 + v2\nEnd Sub\n",
        );
        assert_eq!(vm.variables["x"], Variant::Integer(40));
    }

    #[test]
    fn for_each_binds_the_loop_variable_as_a_live_single_cell_range() {
        let vm = run(
            "Sub MySub()\n    Range(\"A1\").Value = 4\n    Range(\"A2\").Value = 6\n    \
             Dim c As Range\n    t = 0\n    For Each c In Range(\"A1:A2\")\n        \
             t = t + c.Value\n    Next c\nEnd Sub\n",
        );
        assert_eq!(vm.variables["t"], Variant::Integer(10));
    }

    // ── run_sub's pre-flight compile-check (check::compile_check_errors) ────
    // Real VBA compiles the whole module before running any of it, and never
    // lets `On Error` trap a compile error — these confirm `run_sub` gets
    // both properties by running the check before `call_sub_def`.

    #[test]
    fn a_compile_error_prevents_even_the_entrypoints_own_earlier_statements_from_running() {
        let mut vm = Vm::new();
        let prog = parser::parse("Sub MySub()\n    x = 1\n    GoTo Nowhere\nEnd Sub\n").unwrap();
        vm.run_sub(&prog, "mysub").unwrap_err();
        assert!(!vm.variables.contains_key("x"));
    }

    #[test]
    fn a_compile_error_in_an_unrelated_sub_still_blocks_the_entrypoint_from_running() {
        // Whole-module semantics: Innocent itself has no problem, but
        // Broken's undefined label is enough to fail the whole run.
        let mut vm = Vm::new();
        let prog = parser::parse(concat!(
            "Sub Broken()\n",
            "    GoTo Nowhere\n",
            "End Sub\n",
            "Sub Innocent()\n",
            "    x = 42\n",
            "End Sub\n",
        ))
        .unwrap();
        let err = vm.run_sub(&prog, "innocent").unwrap_err();
        assert_eq!(err, "GoTo: label 'nowhere' not found");
        assert!(!vm.variables.contains_key("x"));
    }

    #[test]
    fn on_error_resume_next_does_not_catch_a_compile_error() {
        let mut vm = Vm::new();
        let prog = parser::parse(concat!(
            "Sub MySub()\n",
            "    On Error Resume Next\n",
            "    x = 1\n",
            "    GoTo Nowhere\n",
            "    y = 2\n",
            "End Sub\n",
        ))
        .unwrap();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "GoTo: label 'nowhere' not found");
        // Not just uncaught — never ran at all (the whole point of a
        // pre-flight check), unlike a genuine runtime error under Resume
        // Next, which would leave x set and skip only the failing line.
        assert!(!vm.variables.contains_key("x"));
    }

    #[test]
    fn on_error_goto_does_not_catch_an_undefined_procedure_call() {
        let mut vm = Vm::new();
        let prog = parser::parse(concat!(
            "Sub MySub()\n",
            "    On Error GoTo Handler\n",
            "    Call DoesNotExist()\n",
            "    Exit Sub\n",
            "Handler:\n",
            "    n = Err.Number\n",
            "End Sub\n",
        ))
        .unwrap();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "Sub/Function 'doesnotexist' not found");
        assert!(!vm.variables.contains_key("n"));
    }

    #[test]
    fn an_argument_count_mismatch_is_caught_before_the_call_runs() {
        let mut vm = Vm::new();
        let prog = parser::parse(concat!(
            "Sub Helper(a, b)\n",
            "    x = a + b\n",
            "End Sub\n",
            "Sub MySub()\n",
            "    Call Helper(1)\n",
            "End Sub\n",
        ))
        .unwrap();
        let err = vm.run_sub(&prog, "mysub").unwrap_err();
        assert_eq!(err, "'helper' expects 2 argument(s), got 1");
    }

    #[test]
    fn a_clean_program_runs_normally_through_the_new_pre_flight_check() {
        let vm = run("Sub MySub()\n    x = 1 + 2\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(3));
    }

    #[test]
    fn run_sub_multi_pre_flight_check_covers_every_module_not_just_the_entrypoints() {
        // Entrypoint is Innocent (module2), but module1's Broken has the
        // compile error — must still block the whole run, same as the
        // single-module case, since real VBA compiles the whole project.
        let modules = vec![
            module("module1", "Sub Broken()\n    GoTo Nowhere\nEnd Sub\n"),
            module("module2", "Sub Innocent()\n    x = 42\nEnd Sub\n"),
        ];
        let mut vm = Vm::new();
        let err = vm.run_sub_multi(&modules, "Innocent").unwrap_err();
        assert_eq!(err, "GoTo: label 'nowhere' not found");
        assert!(!vm.variables.contains_key("x"));
    }

    #[test]
    fn run_sub_multi_pre_flight_check_does_not_misflag_a_legitimate_cross_module_call() {
        // Regression guard for the exact risk this design has to avoid:
        // compile_check_errors only ever sees one module's own Program, so
        // run_sub_multi must build other_module_names per module (mirroring
        // main.rs's own multi-module `elixcee check` path) or this would
        // wrongly reject Main's call to Helper as undefined.
        let modules = vec![
            module("module1", "Sub Helper()\n    y = 1\nEnd Sub\n"),
            module(
                "module2",
                "Sub Main()\n    Call Helper()\n    x = 42\nEnd Sub\n",
            ),
        ];
        let mut vm = Vm::new();
        vm.run_sub_multi(&modules, "Main").unwrap();
        assert_eq!(vm.variables["x"], Variant::Integer(42));
        assert_eq!(vm.variables["y"], Variant::Integer(1));
    }

    // ── R1: bulk worksheet range/row API core (resolve_sheet_key,
    // sheet_used_range, next_append_row, read_rect, write_rect,
    // iter_rows_values) ─────────────────────────────────────────────────────

    #[test]
    fn resolve_sheet_key_none_returns_the_active_sheet() {
        let vm = Vm::new();
        assert_eq!(vm.resolve_sheet_key(None).unwrap(), "sheet1");
    }

    #[test]
    fn resolve_sheet_key_looks_up_case_insensitively() {
        let mut vm = Vm::new();
        vm.ensure_sheet("Sheet2");
        assert_eq!(vm.resolve_sheet_key(Some("SHEET2")).unwrap(), "sheet2");
        assert_eq!(vm.resolve_sheet_key(Some("sheet2")).unwrap(), "sheet2");
        assert_eq!(vm.resolve_sheet_key(Some("Sheet2")).unwrap(), "sheet2");
    }

    #[test]
    fn resolve_sheet_key_errors_on_an_unknown_name() {
        let vm = Vm::new();
        let err = vm.resolve_sheet_key(Some("Typo")).unwrap_err();
        assert!(err.contains("Typo"), "{err:?}");
    }

    #[test]
    fn sheet_used_range_is_none_on_an_empty_sheet() {
        let vm = Vm::new();
        assert_eq!(vm.sheet_used_range("sheet1"), None);
    }

    #[test]
    fn sheet_used_range_a_single_cell_is_both_corners() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (3, 3),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        assert_eq!(vm.sheet_used_range("sheet1"), Some(((3, 3), (3, 3))));
    }

    #[test]
    fn sheet_used_range_is_the_true_bounding_box_not_anchored_at_1_1() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (3, 3),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        vm.cells_mut().insert(
            (7, 5),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        assert_eq!(vm.sheet_used_range("sheet1"), Some(((3, 3), (7, 5))));
    }

    #[test]
    fn sheet_used_range_handles_a_sparse_sheet() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        vm.cells_mut().insert(
            (50, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        assert_eq!(vm.sheet_used_range("sheet1"), Some(((1, 1), (50, 1))));
    }

    #[test]
    fn sheet_used_range_counts_null_but_not_empty() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (10, 1),
            CellContent {
                formula: None,
                value: Variant::Null,
            },
        );
        vm.cells_mut().insert(
            (99, 1),
            CellContent {
                formula: None,
                value: Variant::Empty,
            },
        );
        assert_eq!(vm.sheet_used_range("sheet1"), Some(((10, 1), (10, 1))));
        assert_eq!(vm.read_rect("sheet1", 10, 1, 10, 1)[0][0], Variant::Null);
    }

    #[test]
    fn next_append_row_is_1_on_an_empty_sheet() {
        let vm = Vm::new();
        assert_eq!(vm.next_append_row("sheet1"), 1);
    }

    #[test]
    fn next_append_row_uses_the_real_max_on_a_sparse_sheet() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (50, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        assert_eq!(vm.next_append_row("sheet1"), 51);
    }

    #[test]
    fn read_rect_returns_empty_for_gaps() {
        let vm = Vm::new();
        let grid = vm.read_rect("sheet1", 1, 1, 2, 2);
        assert_eq!(
            grid,
            vec![
                vec![Variant::Empty, Variant::Empty],
                vec![Variant::Empty, Variant::Empty],
            ]
        );
    }

    #[test]
    fn read_rect_preserves_every_variant_kind() {
        let mut vm = Vm::new();
        let values = [
            (1, 1, Variant::Integer(42)),
            (1, 2, Variant::Float(1.5)),
            (1, 3, Variant::Str("hi".into())),
            (1, 4, Variant::Boolean(true)),
            (1, 5, Variant::Date(45366)),
            (1, 6, Variant::Error(ExcelError::DivZero)),
        ];
        for &(r, c, ref v) in &values {
            vm.cells_mut().insert(
                (r, c),
                CellContent {
                    formula: None,
                    value: v.clone(),
                },
            );
        }
        let grid = vm.read_rect("sheet1", 1, 1, 1, 6);
        for (i, (_, _, v)) in values.iter().enumerate() {
            assert_eq!(&grid[0][i], v);
        }
    }

    #[test]
    fn read_rect_on_a_formula_cell_returns_its_evaluated_value() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: Some("=1+2".to_string()),
                value: Variant::Integer(3),
            },
        );
        assert_eq!(
            vm.read_rect("sheet1", 1, 1, 1, 1)[0][0],
            Variant::Integer(3)
        );
    }

    #[test]
    fn write_rect_writes_the_exact_grid_at_the_exact_offset() {
        let mut vm = Vm::new();
        vm.write_rect(
            "sheet1",
            (2, 3),
            &[
                vec![Variant::Integer(1), Variant::Integer(2)],
                vec![Variant::Integer(3), Variant::Integer(4)],
            ],
        );
        assert_eq!(vm.get_cell(2, 3), Variant::Integer(1));
        assert_eq!(vm.get_cell(2, 4), Variant::Integer(2));
        assert_eq!(vm.get_cell(3, 3), Variant::Integer(3));
        assert_eq!(vm.get_cell(3, 4), Variant::Integer(4));
    }

    #[test]
    fn write_rect_on_a_non_active_sheet_does_not_change_active_sheet() {
        let mut vm = Vm::new();
        vm.ensure_sheet("Sheet2");
        vm.write_rect("sheet2", (1, 1), &[vec![Variant::Integer(9)]]);
        assert_eq!(vm.active_sheet, "sheet1");
        let written = vm.get_sheet_cells("sheet2").unwrap().get(&(1, 1)).unwrap();
        assert_eq!(written.value, Variant::Integer(9));
        assert_eq!(written.formula, None);
    }

    #[test]
    fn write_rect_into_a_non_anchor_merged_cell_does_not_error() {
        // B1:D1 merged; writing at C1 (a covered, non-anchor cell) must
        // succeed and store the value -- matches real Excel's own plain
        // `.Value=` behavior and PyVm::set_cell's existing lack of any merge
        // check, per docs/openpyxl-gap-audit.md's design note.
        let mut vm = vm_with_merge("sheet1", ((1, 2), (1, 4)));
        vm.write_rect("sheet1", (1, 3), &[vec![Variant::Integer(5)]]);
        assert_eq!(vm.get_cell(1, 3), Variant::Integer(5));
    }

    #[test]
    fn write_rect_into_a_protected_sheet_does_not_error() {
        // No VBA construct reachable from here that both protects a sheet
        // AND leaves it selectable for a raw write_rect call, so this
        // injects `protected_sheets` directly, the same way merge/hidden-row
        // tests inject state with no VBA syntax to create it.
        let mut vm = Vm::new();
        vm.protected_sheets.insert("sheet1".to_string());
        vm.write_rect("sheet1", (1, 1), &[vec![Variant::Integer(1)]]);
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(1));
    }

    #[test]
    fn write_rect_invalidates_the_lazy_cell_index_used_by_end_xlup() {
        // `last_nonempty_row`/`last_nonempty_col` (VBA's `End(xlUp)`/
        // `End(xlToLeft)`) are backed by a lazily-rebuilt index gated on
        // `cell_index_dirty`. `write_rect` goes through `sheet_cells_mut`,
        // which already flips that flag unconditionally -- this test pins
        // that behavior so a bulk write followed by a VBA-side `End(xlUp)`
        // in the same `Vm` can't silently see stale data.
        let mut vm = Vm::new();
        assert_eq!(vm.last_nonempty_row(1, 100), 1); // builds+caches the index at "empty"
        vm.write_rect("sheet1", (10, 1), &[vec![Variant::Integer(42)]]);
        assert_eq!(vm.last_nonempty_row(1, 100), 10);
    }

    #[test]
    fn iter_rows_values_on_an_empty_sheet_with_no_explicit_max_row_is_empty() {
        let vm = Vm::new();
        assert_eq!(
            vm.iter_rows_values("sheet1", 1, None, 1, None),
            Vec::<Vec<Variant>>::new()
        );
    }

    #[test]
    fn iter_rows_values_on_an_empty_sheet_with_an_explicit_max_row_returns_empties() {
        let vm = Vm::new();
        let grid = vm.iter_rows_values("sheet1", 1, Some(3), 1, None);
        assert_eq!(grid.len(), 3);
        assert!(grid.iter().all(|row| row == &[Variant::Empty]));
    }

    #[test]
    fn iter_rows_values_short_circuit_keys_on_max_row_not_max_col() {
        // max_col given explicitly but max_row is not -- still [] on an
        // empty sheet, proving the short-circuit is about max_row.
        let vm = Vm::new();
        assert_eq!(
            vm.iter_rows_values("sheet1", 1, None, 1, Some(3)),
            Vec::<Vec<Variant>>::new()
        );
    }

    #[test]
    fn iter_rows_values_defaults_to_the_used_range() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (2, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        let grid = vm.iter_rows_values("sheet1", 1, None, 1, None);
        assert_eq!(
            grid,
            vec![
                vec![Variant::Empty, Variant::Empty],
                vec![Variant::Empty, Variant::Integer(1)]
            ]
        );
    }

    #[test]
    fn iter_rows_values_on_a_reversed_numeric_window_is_empty_not_a_panic() {
        let vm = Vm::new();
        assert_eq!(
            vm.iter_rows_values("sheet1", 5, Some(2), 1, None),
            Vec::<Vec<Variant>>::new()
        );
    }

    #[test]
    fn iter_cols_values_on_an_empty_sheet_with_no_explicit_max_col_is_empty() {
        let vm = Vm::new();
        assert_eq!(
            vm.iter_cols_values("sheet1", 1, None, 1, None),
            Vec::<Vec<Variant>>::new()
        );
    }

    #[test]
    fn iter_cols_values_on_an_empty_sheet_with_an_explicit_max_col_returns_empties() {
        let vm = Vm::new();
        let grid = vm.iter_cols_values("sheet1", 1, None, 1, Some(3));
        assert_eq!(grid.len(), 3);
        assert!(grid.iter().all(|col| col == &[Variant::Empty]));
    }

    #[test]
    fn iter_cols_values_short_circuit_keys_on_max_col_not_max_row() {
        // max_row given explicitly but max_col is not -- still [] on an
        // empty sheet, proving the short-circuit is about max_col.
        let vm = Vm::new();
        assert_eq!(
            vm.iter_cols_values("sheet1", 1, Some(3), 1, None),
            Vec::<Vec<Variant>>::new()
        );
    }

    #[test]
    fn iter_cols_values_defaults_to_the_used_range() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (2, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        let grid = vm.iter_cols_values("sheet1", 1, None, 1, None);
        assert_eq!(
            grid,
            vec![
                vec![Variant::Empty, Variant::Empty],
                vec![Variant::Empty, Variant::Integer(1)]
            ]
        );
    }

    #[test]
    fn iter_cols_values_on_a_reversed_numeric_window_is_empty_not_a_panic() {
        let vm = Vm::new();
        assert_eq!(
            vm.iter_cols_values("sheet1", 1, None, 5, Some(2)),
            Vec::<Vec<Variant>>::new()
        );
    }

    #[test]
    fn iter_cols_values_is_the_transpose_of_iter_rows_values_on_the_same_bounds() {
        let mut vm = Vm::new();
        vm.cells_mut().insert(
            (1, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(1),
            },
        );
        vm.cells_mut().insert(
            (1, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(2),
            },
        );
        vm.cells_mut().insert(
            (2, 1),
            CellContent {
                formula: None,
                value: Variant::Integer(3),
            },
        );
        vm.cells_mut().insert(
            (2, 2),
            CellContent {
                formula: None,
                value: Variant::Integer(4),
            },
        );
        let rows = vm.iter_rows_values("sheet1", 1, Some(2), 1, Some(2));
        let cols = vm.iter_cols_values("sheet1", 1, Some(2), 1, Some(2));
        let transposed_rows: Vec<Vec<Variant>> = (0..rows[0].len())
            .map(|ci| rows.iter().map(|row| row[ci].clone()).collect())
            .collect();
        assert_eq!(cols, transposed_rows);
    }
}
