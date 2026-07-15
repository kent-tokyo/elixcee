use std::collections::{BTreeSet, HashMap, HashSet, VecDeque};

use crate::formula;
use crate::parser::ast::{CalcModeValue, CaseMatch, Expr, FuncDef, Program, SourceSpan, SpannedStmt, Stmt, SubDef, VbaBinOp, XlDir, XlEndProp};
use crate::parser::{self, EntrypointResolution};
use crate::reader::{self, SheetCell, WorkbookSheet};

/// Excel worksheet error values (#DIV/0!, #N/A, etc.)
#[derive(Debug, Clone, PartialEq)]
pub enum ExcelError {
    DivZero, // #DIV/0!
    NA,      // #N/A
    Value,   // #VALUE!
    Ref,     // #REF!
    Name,    // #NAME?
    Num,     // #NUM!
    Null,    // #NULL!
}

impl ExcelError {
    pub fn as_str(&self) -> &'static str {
        match self {
            ExcelError::DivZero => "#DIV/0!",
            ExcelError::NA      => "#N/A",
            ExcelError::Value   => "#VALUE!",
            ExcelError::Ref     => "#REF!",
            ExcelError::Name    => "#NAME?",
            ExcelError::Num     => "#NUM!",
            ExcelError::Null    => "#NULL!",
        }
    }
}

impl std::fmt::Display for ExcelError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.as_str())
    }
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
}

/// The VM's clipboard state, populated by `.Copy` and consumed by
/// `.Paste`/`.PasteSpecial` (Milestone B6b). Values are snapshotted at copy
/// time (`cells`), not re-read from the source range at paste time — this
/// matches real Excel's copy-then-mutate-then-paste semantics now that Copy
/// and Paste can be separate statements.
#[derive(Debug, Clone)]
struct ClipboardState {
    source_addr: String,
    rows: u32,
    cols: u32,
    cells: Vec<Vec<Variant>>, // [row][col], 0-based offsets from the source's top-left
    span: SourceSpan,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Variant {
    Integer(i64),
    Float(f64),
    Str(String),
    Boolean(bool),
    Date(i64),           // Excel serial date — displays as "YYYY-MM-DD"
    Error(ExcelError),   // Excel error value (#DIV/0!, #N/A, …)
    Empty,
    Array(Vec<Variant>),                  // 0-indexed 1D array
    Record(std::collections::HashMap<String, Variant>), // UDT instance (p.x, p.y, …)
}

pub fn serial_to_display(s: i64) -> String {
    // Reuse formula engine's serial_to_ymd
    let (y, m, d) = crate::formula::eval::serial_to_ymd_pub(s);
    format!("{:04}-{:02}-{:02}", y, m, d)
}

impl std::fmt::Display for Variant {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Variant::Integer(n) => write!(f, "{}", n),
            Variant::Float(v)   => write!(f, "{}", v),
            Variant::Str(s)     => write!(f, "{}", s),
            Variant::Boolean(b) => write!(f, "{}", if *b { "True" } else { "False" }),
            Variant::Date(s)    => write!(f, "{}", serial_to_display(*s)),
            Variant::Error(e)   => write!(f, "{}", e),
            Variant::Empty      => write!(f, ""),
            Variant::Array(a)   => write!(f, "[{}]", a.iter().map(|v| v.to_string()).collect::<Vec<_>>().join(", ")),
            Variant::Record(m)  => {
                let mut pairs: Vec<String> = m.iter().map(|(k, v)| format!("{}: {}", k, v)).collect();
                pairs.sort();
                write!(f, "{{{}}}", pairs.join(", "))
            }
        }
    }
}

#[derive(Debug, Clone)]
pub struct CellContent {
    pub formula: Option<String>,
    pub value: Variant,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CalculationMode {
    Automatic,
    Manual,
}

/// Signals emitted by Exit For / Exit Do / Exit Sub / Exit Function.
#[derive(Debug, Clone, PartialEq)]
pub enum ExitKind { For, Do, Sub, Function }

pub struct Vm {
    /// Per-sheet cell storage. Key is sheet name (lowercase for lookup).
    sheets: HashMap<String, HashMap<(u32, u32), CellContent>>,
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
    pub on_error_resume_next: bool,
    /// Label to jump to when an error occurs (On Error GoTo <label>).
    pub on_error_goto_label: Option<String>,
    /// Pending unconditional jump target (GoTo <label>).
    pending_goto: Option<String>,
    user_funcs: HashMap<String, FuncDef>,
    user_subs:  HashMap<String, SubDef>,
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
    /// The clipboard populated by `.Copy` and consumed by
    /// `.Paste`/`.PasteSpecial` (Milestone B6b). `None` initially, and
    /// whenever `Application.CutCopyMode` is set to `False`.
    clipboard: Option<ClipboardState>,
    /// Lowercase sheet keys currently `.Protect`ed (Milestone B6c) — same
    /// key space as `sheets`/`active_sheet`/`ensure_sheet`. Empty by
    /// default; blocks any cell-mutating statement on that sheet.
    protected_sheets: HashSet<String>,
}

impl Vm {
    pub fn new() -> Self {
        let mut sheets = HashMap::new();
        sheets.insert("sheet1".into(), HashMap::new());
        Vm {
            sheets,
            active_sheet: "sheet1".into(),
            variables: HashMap::new(),
            calc_mode: CalculationMode::Automatic,
            error_on_msgbox: false,
            print_msgbox: false,
            msgbox_log: Vec::new(),
            current_span: None,
            exit_flag: None,
            on_error_resume_next: false,
            on_error_goto_label: None,
            pending_goto: None,
            user_funcs: HashMap::new(),
            user_subs:  HashMap::new(),
            named_ranges: HashMap::new(),
            type_defs: HashMap::new(),
            col_rows: HashMap::new(),
            row_cols: HashMap::new(),
            cell_index_dirty: true,
            deadline: None,
            loop_iters: 0,
            strict_resolution: false,
            last_resolution_failure: None,
            loaded_workbook_name: None,
            clipboard: None,
            protected_sheets: HashSet::new(),
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
        if let Some(r) = parse_range_addr(addr) { return Some(r); }
        self.named_ranges.get(&addr.to_lowercase())
            .and_then(|real| parse_range_addr(real))
    }

    pub fn cells(&self) -> &HashMap<(u32, u32), CellContent> {
        self.sheets.get(&self.active_sheet).expect("active sheet must exist")
    }

    pub fn cells_mut(&mut self) -> &mut HashMap<(u32, u32), CellContent> {
        self.cell_index_dirty = true;
        self.sheets.get_mut(&self.active_sheet).expect("active sheet must exist")
    }

    fn rebuild_cell_index(&mut self) {
        let pairs: Vec<(u32, u32)> = self.cells().iter()
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
        self.sheets.entry(name.to_lowercase()).or_insert_with(HashMap::new);
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
    /// identical output. `lower` is always 0 and `upper` is `len - 1`
    /// (elixcee's arrays are always 0-based — see the module-level note on
    /// `Dim arr(1 To N)` not being tracked): this is elixcee's true bound,
    /// not a fabricated VBA-style `1 To N`.
    fn array_oob_error(&mut self, name: &str, idx: usize, len: usize) -> String {
        self.last_resolution_failure = Some(ResolutionFailureKind::ArrayIndexOutOfBounds {
            name: name.to_string(),
            index: idx as i64,
            lower: 0,
            upper: len as i64 - 1,
        });
        format!(
            "Array '{}': index {} out of bounds (len={})",
            name, idx, len
        )
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
    fn populate_from_sheets(&mut self, sheets: Vec<WorkbookSheet>) -> Vec<String> {
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
                };
                self.cells_mut().insert(
                    (row, col),
                    CellContent {
                        formula: None,
                        value,
                    },
                );
            }
            self.active_sheet = prev;
            names.push(sheet_data.name.to_lowercase());
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
        // Cache user-defined functions, subs, and type definitions.
        self.user_funcs = program.funcs.iter().map(|f| (f.name.clone(), f.clone())).collect();
        self.user_subs  = program.subs.iter().map(|s| (s.name.clone(), s.clone())).collect();
        for td in &program.type_defs {
            self.type_defs.insert(td.name.clone(), td.fields.clone());
        }
        let name = sub_name.to_lowercase();
        let sub = self.user_subs.get(&name)
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
        self.user_funcs.clear();
        self.user_subs.clear();
        for (_, program) in modules {
            for f in &program.funcs { self.user_funcs.insert(f.name.clone(), f.clone()); }
            for s in &program.subs { self.user_subs.insert(s.name.clone(), s.clone()); }
            for td in &program.type_defs { self.type_defs.insert(td.name.clone(), td.fields.clone()); }
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
        let saved: Vec<(String, Option<Variant>)> = sub.params.iter().enumerate().map(|(i, p)| {
            let old = self.variables.get(p).cloned();
            if let Some(v) = args.get(i) { self.variables.insert(p.clone(), v.clone()); }
            (p.clone(), old)
        }).collect();
        let body = sub.body.clone();
        self.exec_body(&body, |f| matches!(f, ExitKind::Sub))?;
        for (p, old) in saved {
            match old { Some(v) => { self.variables.insert(p, v); } None => { self.variables.remove(&p); } }
        }
        Ok(())
    }

    fn call_func_def(&mut self, func: &FuncDef, args: &[Variant]) -> Result<Variant, String> {
        let saved: Vec<(String, Option<Variant>)> = func.params.iter().enumerate().map(|(i, p)| {
            let old = self.variables.get(p).cloned();
            if let Some(v) = args.get(i) { self.variables.insert(p.clone(), v.clone()); }
            (p.clone(), old)
        }).collect();
        let ret_name = func.name.clone();
        let old_ret = self.variables.remove(&ret_name);
        let body = func.body.clone();
        self.exec_body(&body, |f| matches!(f, ExitKind::Function | ExitKind::Sub))?;
        let ret_val = self.variables.remove(&ret_name).unwrap_or(Variant::Empty);
        for (p, old) in saved {
            match old { Some(v) => { self.variables.insert(p, v); } None => { self.variables.remove(&p); } }
        }
        if let Some(v) = old_ret { self.variables.insert(ret_name, v); }
        Ok(ret_val)
    }

    /// Execute a body slice with label-jump support (for GoTo and On Error GoTo).
    /// The existing per-statement `exec_stmt` error catch (resume_next) is preserved.
    fn exec_body<F>(&mut self, stmts: &[SpannedStmt], is_exit: F) -> Result<(), String>
    where F: Fn(&ExitKind) -> bool
    {
        let mut i = 0;
        while i < stmts.len() {
            // Handle pending unconditional GoTo
            if let Some(label) = self.pending_goto.take() {
                match stmts.iter().position(|s| matches!(&s.stmt, Stmt::Label(l) if l == &label)) {
                    Some(pos) => { i = pos; continue; }
                    None      => return Err(format!("GoTo: label '{}' not found", label)),
                }
            }

            if let Some(ref f) = self.exit_flag {
                if is_exit(f) { self.exit_flag = None; break; }
                break; // other exit kinds bubble up
            }

            let result = self.exec_stmt(&stmts[i]); // per-stmt catch preserves resume_next behavior
            match result {
                Ok(()) => {}
                Err(e) => {
                    // On Error GoTo: jump to handler label — skipped in
                    // strict-resolution mode (`diagnose`) so the first
                    // failure always propagates instead of being redirected
                    // to a handler that would mask it.
                    if !self.strict_resolution
                        && let Some(label) = self.on_error_goto_label.take()
                    {
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

    fn exec_stmt(&mut self, spanned: &SpannedStmt) -> Result<(), String> {
        if self.exit_flag.is_some() { return Ok(()); }
        self.current_span = Some(spanned.span);
        let result = self.exec_stmt_inner(&spanned.stmt);
        match result {
            Ok(()) => Ok(()),
            // `On Error Resume Next` is not honored in strict-resolution
            // mode (`diagnose`) — see the field doc on `strict_resolution`.
            Err(_) if self.on_error_resume_next && !self.strict_resolution => Ok(()),
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
                self.cells_mut().insert((r, c), CellContent { formula: None, value: v });
            }
            Stmt::SetCalcMode(mode) => {
                let m = match mode {
                    CalcModeValue::Automatic => CalculationMode::Automatic,
                    CalcModeValue::Manual    => CalculationMode::Manual,
                };
                self.set_calc_mode(m)?;
            }
            Stmt::For { var, from, to, step, body } => {
                let mut i  = to_f64(&self.eval_expr(from)?)?;
                let to_f   = to_f64(&self.eval_expr(to)?)?;
                let step_f = match step { Some(s) => to_f64(&self.eval_expr(s)?)?, None => 1.0 };
                if step_f == 0.0 { return Err("For loop: step cannot be zero".into()); }
                'for_loop: while (step_f > 0.0 && i <= to_f) || (step_f < 0.0 && i >= to_f) {
                    self.check_deadline()?;
                    self.variables.insert(var.clone(), as_int_if_whole(i));
                    for s in body {
                        self.exec_stmt(s)?;
                        if matches!(self.exit_flag, Some(ExitKind::For)) { self.exit_flag = None; break 'for_loop; }
                        if self.exit_flag.is_some() { return Ok(()); }
                    }
                    i += step_f;
                }
            }
            Stmt::ForEach { var, range_addr, body } => {
                let ((r1, c1), (r2, c2)) = self.resolve_range_addr(range_addr)
                    .ok_or_else(|| format!("ForEach: invalid range '{}'", range_addr))?;
                'fe_outer: for r in r1..=r2 {
                    for c in c1..=c2 {
                        self.check_deadline()?;
                        let v = self.get_cell(r, c);
                        self.variables.insert(var.clone(), v);
                        for s in body {
                            self.exec_stmt(s)?;
                            if matches!(self.exit_flag, Some(ExitKind::For)) { self.exit_flag = None; break 'fe_outer; }
                            if self.exit_flag.is_some() { return Ok(()); }
                        }
                    }
                }
            }
            Stmt::If { condition, then_body, else_body } => {
                let branch = if is_truthy(&self.eval_expr(condition)?) { then_body } else { else_body };
                for s in branch { self.exec_stmt(s)?; if self.exit_flag.is_some() { return Ok(()); } }
            }
            Stmt::DoLoop { pre_cond, post_cond, body } => {
                let check = |vm: &mut Vm, cond: &Option<(bool, Expr)>| -> Result<bool, String> {
                    match cond {
                        None => Ok(true),
                        Some((is_until, expr)) => {
                            let v = vm.eval_expr(expr)?;
                            Ok(if *is_until { !is_truthy(&v) } else { is_truthy(&v) })
                        }
                    }
                };
                'do_loop: while check(self, pre_cond)? {
                    self.check_deadline()?;
                    for s in body.clone() {
                        self.exec_stmt(&s)?;
                        if matches!(self.exit_flag, Some(ExitKind::Do)) { self.exit_flag = None; break 'do_loop; }
                        if self.exit_flag.is_some() { return Ok(()); }
                    }
                    if !check(self, post_cond)? { break 'do_loop; }
                }
            }
            Stmt::SelectCase { expr, cases, else_body } => {
                let val = self.eval_expr(expr)?;
                let mut matched = false;
                'outer: for (matchers, body) in cases {
                    for m in matchers {
                        let hit = match m {
                            CaseMatch::Value(v) => vba_eq(&val, &self.eval_expr(v)?),
                            CaseMatch::Range(lo, hi) => {
                                let l = self.eval_expr(lo)?; let h = self.eval_expr(hi)?;
                                vba_cmp(&val, &l)? != std::cmp::Ordering::Less && vba_cmp(&val, &h)? != std::cmp::Ordering::Greater
                            }
                            CaseMatch::IsOp(op, rhs) => {
                                let r = self.eval_expr(rhs)?;
                                match op {
                                    VbaBinOp::Eq => vba_eq(&val, &r), VbaBinOp::Ne => !vba_eq(&val, &r),
                                    VbaBinOp::Lt => vba_cmp(&val, &r)? == std::cmp::Ordering::Less,
                                    VbaBinOp::Le => vba_cmp(&val, &r)? != std::cmp::Ordering::Greater,
                                    VbaBinOp::Gt => vba_cmp(&val, &r)? == std::cmp::Ordering::Greater,
                                    VbaBinOp::Ge => vba_cmp(&val, &r)? != std::cmp::Ordering::Less,
                                    _ => false,
                                }
                            }
                        };
                        if hit {
                            for s in body { self.exec_stmt(s)?; }
                            matched = true; break 'outer;
                        }
                    }
                }
                if !matched { for s in else_body { self.exec_stmt(s)?; } }
            }
            Stmt::ExitFor      => self.exit_flag = Some(ExitKind::For),
            Stmt::ExitDo       => self.exit_flag = Some(ExitKind::Do),
            Stmt::ExitSub      => self.exit_flag = Some(ExitKind::Sub),
            Stmt::ExitFunction => self.exit_flag = Some(ExitKind::Function),
            Stmt::OnError { resume_next } => {
                self.on_error_resume_next = *resume_next;
                self.on_error_goto_label = None;
            }
            Stmt::OnErrorGoTo(label) => {
                self.on_error_goto_label = Some(label.clone());
                self.on_error_resume_next = false;
            }
            Stmt::Label(_) => {}  // no-op during normal execution
            Stmt::GoTo(label) => {
                self.pending_goto = Some(label.clone());
            }
            Stmt::Resume { .. } => {
                // After error handler runs: clear error state, continue
                self.on_error_goto_label = None;
            }
            Stmt::CallSub { name, args } => {
                let arg_vals: Vec<Variant> = args.iter().map(|a| self.eval_expr(a)).collect::<Result<_, _>>()?;
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
            Stmt::RangeWrite { addr, is_formula, value } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let v = self.eval_expr(value)?;
                let ((r1,c1),(r2,c2)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeWrite: invalid address '{}'", addr))?;
                if *is_formula {
                    let s = vba_to_str(&v);
                    for r in r1..=r2 {
                        for c in c1..=c2 { self.set_cell_formula(r, c, &s)?; }
                    }
                } else {
                    // Batch writes: access sheet directly to avoid N dirty-flag sets
                    let sheet = self.active_sheet.clone();
                    if let Some(cells) = self.sheets.get_mut(&sheet) {
                        for r in r1..=r2 {
                            for c in c1..=c2 {
                                cells.insert((r, c), CellContent { formula: None, value: v.clone() });
                            }
                        }
                    }
                    self.cell_index_dirty = true;
                }
            }
            Stmt::RangeClear { addr, .. } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let ((r1,c1),(r2,c2)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeClear: invalid address '{}'", addr))?;
                let sheet = self.active_sheet.clone();
                if let Some(cells) = self.sheets.get_mut(&sheet) {
                    for r in r1..=r2 { for c in c1..=c2 { cells.remove(&(r, c)); } }
                }
                self.cell_index_dirty = true;
            }
            Stmt::RangeOffsetWrite { addr, row_off, col_off, value } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let v = self.eval_expr(value)?;
                let (base_r, base_c) = parse_cell_addr(addr)
                    .ok_or_else(|| format!("RangeOffsetWrite: invalid address '{}'", addr))?;
                let ro = to_f64(&self.eval_expr(row_off)?)? as i64;
                let co = to_f64(&self.eval_expr(col_off)?)? as i64;
                let row = (base_r as i64 + ro) as u32;
                let col = (base_c as i64 + co) as u32;
                self.cells_mut().insert((row, col), CellContent { formula: None, value: v });
            }
            Stmt::RangeDelete { addr } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let ((r1,_),(r2,_)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeDelete: invalid address '{}'", addr))?;
                let rows_del = r2 - r1 + 1;
                // Collect cells below the deleted range and shift them up
                let to_move: Vec<((u32,u32), CellContent)> = self.cells().iter()
                    .filter(|((r,_),_)| *r > r2)
                    .map(|((r,c),v)| ((*r,*c), v.clone())).collect();
                for ((r,c),_) in &to_move { self.cells_mut().remove(&(*r,*c)); }
                for r in r1..=r2 { self.cells_mut().retain(|&(row,_),_| row != r); }
                for ((r,c),v) in to_move {
                    self.cells_mut().insert((r - rows_del, c), v);
                }
            }
            Stmt::RangeInsert { addr } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let ((r1,_),(r2,_)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeInsert: invalid address '{}'", addr))?;
                let rows_ins = r2 - r1 + 1;
                // Shift cells at r1 and below downward
                let to_move: Vec<((u32,u32), CellContent)> = self.cells().iter()
                    .filter(|((r,_),_)| *r >= r1)
                    .map(|((r,c),v)| ((*r,*c), v.clone())).collect();
                for ((r,c),_) in &to_move { self.cells_mut().remove(&(*r,*c)); }
                for ((r,c),v) in to_move {
                    self.cells_mut().insert((r + rows_ins, c), v);
                }
            }
            Stmt::RangeSort { addr, key_col, descending } => {
                let active = self.active_sheet.clone();
                self.check_sheet_not_protected(&active, &active)?;
                let ((r1,c1),(r2,c2)) = self.resolve_range_addr(addr)
                    .ok_or_else(|| format!("RangeSort: invalid address '{}'", addr))?;
                // key_col is 1-based absolute column; convert to 0-based offset within range
                let key_off = (*key_col).saturating_sub(c1) as usize;
                let mut rows: Vec<Vec<Variant>> = (r1..=r2)
                    .map(|r| (c1..=c2).map(|c| self.get_cell(r, c)).collect()).collect();
                rows.sort_by(|a, b| {
                    let va = a.get(key_off).unwrap_or(&Variant::Empty);
                    let vb = b.get(key_off).unwrap_or(&Variant::Empty);
                    let ord = cmp_variants(va, vb);
                    if *descending { ord.reverse() } else { ord }
                });
                for (ri, row) in rows.iter().enumerate() {
                    for (ci, val) in row.iter().enumerate() {
                        self.cells_mut().insert((r1 + ri as u32, c1 + ci as u32),
                            CellContent { formula: None, value: val.clone() });
                    }
                }
            }
            Stmt::RangeCopy { src, dst } => {
                let ((r1, c1), (r2, c2)) = self.resolve_range_addr(src)
                    .ok_or_else(|| format!("RangeCopy: invalid source range '{}'", src))?;
                let cells: Vec<Vec<Variant>> = (r1..=r2)
                    .map(|r| (c1..=c2).map(|c| self.get_cell(r, c)).collect())
                    .collect();
                self.clipboard = Some(ClipboardState {
                    source_addr: src.clone(),
                    rows: r2 - r1 + 1,
                    cols: c2 - c1 + 1,
                    cells,
                    span: self.current_span.unwrap_or(SourceSpan { start: 0, end: 0 }),
                });
                if let Some(dst_addr) = dst {
                    self.do_paste(dst_addr, false)?;
                }
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
            Stmt::SheetCellWrite { sheet, row, col, value } => {
                let (key, display) = self.resolve_sheet_expr(sheet)?;
                self.check_strict_sheet_exists(&display, &key)?;
                self.check_sheet_not_protected(&key, &display)?;
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                let v = self.eval_expr(value)?;
                if !self.strict_resolution { self.ensure_sheet(&key); }
                self.sheet_cells_mut(&key).unwrap().insert((r, c), CellContent { formula: None, value: v });
            }
            Stmt::SheetRangeWrite { sheet, addr, is_formula, value } => {
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
                if !self.strict_resolution { self.ensure_sheet(sheet_name); }
                self.active_sheet = sheet_name.to_lowercase();
                self.cell_index_dirty = true;
                for s in body.clone() {
                    self.exec_stmt(&s)?;
                    if self.exit_flag.is_some() { break; }
                }
                self.active_sheet = prev;
                self.cell_index_dirty = true;
            }
            Stmt::SheetsAdd => {
                let new_name = format!("sheet{}", self.sheets.len() + 1);
                self.ensure_sheet(&new_name);
            }
            Stmt::SheetsDelete { sheet } => {
                let (key, display) = self.resolve_sheet_expr(sheet)?;
                self.check_sheet_not_protected(&key, &display)?;
                if key != self.active_sheet { self.sheets.remove(&key); }
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
            Stmt::Unsupported { .. } => {}
            Stmt::DimArray { name, sizes } => {
                let upper = to_f64(&self.eval_expr(&sizes[0])?)? as usize;
                self.variables.insert(name.clone(), Variant::Array(vec![Variant::Empty; upper + 1]));
            }
            Stmt::ReDim { name, sizes, preserve } => {
                let upper = to_f64(&self.eval_expr(&sizes[0])?)? as usize;
                let new_size = upper + 1;
                let new_arr = if *preserve {
                    if let Some(Variant::Array(old)) = self.variables.get(name).cloned() {
                        let mut a = old; a.resize(new_size, Variant::Empty); a
                    } else { vec![Variant::Empty; new_size] }
                } else { vec![Variant::Empty; new_size] };
                self.variables.insert(name.clone(), Variant::Array(new_arr));
            }
            Stmt::ArrayWrite { name, indices, value } => {
                let v = self.eval_expr(value)?;
                let idx = to_f64(&self.eval_expr(&indices[0])?)? as usize;
                let oob_len = match self.variables.get(name) {
                    Some(Variant::Array(arr)) if idx >= arr.len() => Some(arr.len()),
                    Some(Variant::Array(_)) => None,
                    _ => return Err(format!("'{}' is not an array", name)),
                };
                if let Some(len) = oob_len {
                    return Err(self.array_oob_error(name, idx, len));
                }
                if let Some(Variant::Array(arr)) = self.variables.get_mut(name) {
                    arr[idx] = v;
                }
            }
            Stmt::With { body } => {
                for s in body { self.exec_stmt(s)?; if self.exit_flag.is_some() { return Ok(()); } }
            }
            Stmt::MsgBox { message } => {
                let msg = self.eval_expr(message)?;
                // Record before checking error_on_msgbox: `messages` should
                // reflect every MsgBox the macro attempted to show, even
                // ones that are then treated as a blocking error.
                self.msgbox_log.push(msg.to_string());
                if self.error_on_msgbox { return Err(format!("MsgBox: {}", msg)); }
                if self.print_msgbox { println!("{}", msg); }
            }
            Stmt::DimRecord { var, type_name } => {
                if let Some(fields) = self.type_defs.get(type_name).cloned() {
                    let record = make_record_default(&fields, &self.type_defs);
                    self.variables.insert(var.clone(), record);
                }
                // Unknown type name → no-op (built-in type)
            }
            Stmt::DimArrayRecord { name, sizes, type_name } => {
                let upper = to_f64(&self.eval_expr(&sizes[0])?)? as usize;
                let element = if let Some(fields) = self.type_defs.get(type_name).cloned() {
                    make_record_default(&fields, &self.type_defs)
                } else {
                    Variant::Empty
                };
                self.variables.insert(name.clone(), Variant::Array(vec![element; upper + 1]));
            }
            Stmt::RecordSetNested { var, fields, value } => {
                let v = self.eval_expr(value)?;
                let target = self.variables.entry(var.clone())
                    .or_insert_with(|| Variant::Record(HashMap::new()));
                nested_set(target, fields, v);
            }
            Stmt::ArrayRecordSet { name, indices, field, value } => {
                let v = self.eval_expr(value)?;
                let idx = to_f64(&self.eval_expr(&indices[0])?)? as usize;
                let oob_len = match self.variables.get(name) {
                    Some(Variant::Array(arr)) if idx >= arr.len() => Some(arr.len()),
                    Some(Variant::Array(_)) => None,
                    _ => return Err(format!("'{}' is not an array", name)),
                };
                if let Some(len) = oob_len {
                    return Err(self.array_oob_error(name, idx, len));
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
            Stmt::WithRecord { body, .. } => {
                // Parser already substituted the variable name into each statement.
                for s in body { self.exec_stmt(s)?; if self.exit_flag.is_some() { return Ok(()); } }
            }
            Stmt::RecordSet { var, field, value } => {
                let v = self.eval_expr(value)?;
                let entry = self.variables.entry(var.clone()).or_insert(Variant::Record(std::collections::HashMap::new()));
                if let Variant::Record(m) = entry {
                    m.insert(field.clone(), v);
                } else {
                    self.variables.insert(var.clone(), Variant::Record({
                        let mut m = std::collections::HashMap::new();
                        m.insert(field.clone(), v);
                        m
                    }));
                }
            }
        }
        Ok(())
    }

    pub fn eval_expr(&mut self, expr: &Expr) -> Result<Variant, String> {
        match expr {
            Expr::Integer(n)    => Ok(Variant::Integer(*n)),
            Expr::Float(f)      => Ok(Variant::Float(*f)),
            Expr::Str(s)        => Ok(Variant::Str(s.clone())),
            Expr::Bool(b)       => Ok(Variant::Boolean(*b)),
            Expr::Var(name) => {
                if let Some(v) = self.variables.get(name) { return Ok(v.clone()); }
                // Excel built-in constants
                Ok(match name.as_str() {
                    // Calculation mode
                    "xlcalculationmanual"        => Variant::Integer(-4135),
                    "xlcalculationautomatic"     => Variant::Integer(-4105),
                    "xlcalculationsemiautomatic" => Variant::Integer(2),
                    // Direction
                    "xlup"        => Variant::Integer(-4162),
                    "xldown"      => Variant::Integer(-4121),
                    "xltoleft"    => Variant::Integer(-4159),
                    "xltoright"   => Variant::Integer(-4161),
                    // Cursor
                    "xlwait"           => Variant::Integer(2),
                    "xldefault"        => Variant::Integer(1),
                    "xlibeam"          => Variant::Integer(3),
                    "xlnorthwestarrow" => Variant::Integer(4),
                    // VB string constants
                    "vbcrlf"       => Variant::Str("\r\n".into()),
                    "vblf"         => Variant::Str("\n".into()),
                    "vbcr"         => Variant::Str("\r".into()),
                    "vbtab"        => Variant::Str("\t".into()),
                    "vbnullstring" => Variant::Str(String::new()),
                    "vbnullchar"   => Variant::Str("\0".into()),
                    // VB boolean constants (in addition to True/False literals)
                    "vbtrue"  => Variant::Boolean(true),
                    "vbfalse" => Variant::Boolean(false),
                    // VB MsgBox return values
                    "vbok"     => Variant::Integer(1),
                    "vbcancel" => Variant::Integer(2),
                    "vbyes"    => Variant::Integer(6),
                    "vbno"     => Variant::Integer(7),
                    "empty" | "null" | "nothing" => return Ok(Variant::Empty),
                    _ => return Err(format!("Undefined variable: '{}'", name)),
                })
            }
            Expr::UnaryMinus(inner) => match self.eval_expr(inner)? {
                Variant::Integer(n) => Ok(Variant::Integer(-n)),
                Variant::Float(f)   => Ok(Variant::Float(-f)),
                other => Err(format!("Unary minus on non-numeric: {}", other)),
            },
            Expr::UnaryNot(inner) => {
                Ok(Variant::Boolean(!is_truthy(&self.eval_expr(inner)?)))
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
                    let arg_vals: Vec<Variant> = args.iter().map(|a| self.eval_expr(a)).collect::<Result<_, _>>()?;
                    return self.call_func_def(&func, &arg_vals);
                }
                // Array subscript access: arr(i)
                if matches!(self.variables.get(name.as_str()), Some(Variant::Array(_))) {
                    let idx = to_f64(&self.eval_expr(args.first().ok_or_else(|| format!("Array '{}' requires index", name))?)?)? as usize;
                    let (found, len) = match self.variables.get(name.as_str()) {
                        Some(Variant::Array(arr)) => (arr.get(idx).cloned(), arr.len()),
                        _ => return Err(format!("'{}' is not an array", name)),
                    };
                    return match found {
                        Some(v) => Ok(v),
                        None => Err(self.array_oob_error(name, idx, len)),
                    };
                }
                self.eval_vba_func(name, args)
            }
            Expr::RangeRead { addr } => {
                let ((r1, c1), (r2, c2)) = self.resolve_range_addr(addr)
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
            Expr::RangeOffsetRead { addr, row_off, col_off } => {
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
            Expr::CellsFind { what, find_row } => {
                let target = self.eval_expr(what)?;
                let mut keys: Vec<(u32, u32)> = self.cells().keys().cloned().collect();
                keys.sort(); // 行優先スキャン
                for (r, c) in keys {
                    if vba_eq(&self.get_cell(r, c), &target) {
                        return Ok(Variant::Integer(if *find_row { r as i64 } else { c as i64 }));
                    }
                }
                Ok(Variant::Integer(0)) // not found
            }
            Expr::RowsCount => Ok(Variant::Integer(1_048_576)),
            Expr::ColsCount => Ok(Variant::Integer(16_384)),
            Expr::CellsEndProp { row, col, dir, prop } => {
                let r = to_cell_index(self.eval_expr(row)?, "row")?;
                let c = to_cell_index(self.eval_expr(col)?, "col")?;
                let result = match (dir, prop) {
                    (XlDir::Up,    XlEndProp::Row)    => self.last_nonempty_row(c, r),
                    (XlDir::Down,  XlEndProp::Row)    => self.first_empty_row(c, r).saturating_sub(1),
                    (XlDir::Left,  XlEndProp::Column) => self.last_nonempty_col(r, c),
                    (XlDir::Right, XlEndProp::Column) => self.first_empty_col(r, c).saturating_sub(1),
                    (XlDir::Up,    XlEndProp::Column) |
                    (XlDir::Down,  XlEndProp::Column) => c,
                    (XlDir::Left,  XlEndProp::Row)    |
                    (XlDir::Right, XlEndProp::Row)    => r,
                };
                Ok(Variant::Integer(result as i64))
            }
            Expr::RecordGet { var, field } => {
                match self.variables.get(var) {
                    Some(Variant::Record(m)) => Ok(m.get(field).cloned().unwrap_or(Variant::Empty)),
                    _ => Ok(Variant::Empty),
                }
            }
            Expr::RecordGetNested { var, fields } => {
                let mut cur = self.variables.get(var).cloned().unwrap_or(Variant::Empty);
                for f in fields {
                    cur = match cur {
                        Variant::Record(m) => m.get(f).cloned().unwrap_or(Variant::Empty),
                        _ => Variant::Empty,
                    };
                }
                Ok(cur)
            }
            Expr::ArrayRecordGet { name, indices, field } => {
                let idx = to_f64(&self.eval_expr(&indices[0])?)? as usize;
                let (found, len) = match self.variables.get(name) {
                    Some(Variant::Array(arr)) => (arr.get(idx).cloned(), arr.len()),
                    _ => return Err(format!("'{}' is not an array", name)),
                };
                match found {
                    Some(Variant::Record(m)) => Ok(m.get(field).cloned().unwrap_or(Variant::Empty)),
                    Some(other) => Ok(other),
                    None => Err(self.array_oob_error(name, idx, len)),
                }
            }
        }
    }

    fn eval_vba_func(&mut self, name: &str, args: &[Expr]) -> Result<Variant, String> {
        let vals: Vec<Variant> = args.iter().map(|a| self.eval_expr(a)).collect::<Result<_, _>>()?;
        match name {
            "int" => {
                let f = to_f64(vals.first().ok_or("INT requires 1 argument")?)?;
                Ok(as_int_if_whole(f.floor()))
            }
            "clng" | "cint" | "cbool" => {
                let f = to_f64(vals.first().ok_or("CInt/CLng requires 1 argument")?)?;
                Ok(Variant::Integer(f.round() as i64))
            }
            "cdbl" | "csng" => {
                let f = to_f64(vals.first().ok_or("CDbl requires 1 argument")?)?;
                Ok(Variant::Float(f))
            }
            "cstr" | "str" => {
                let s = vals.first().ok_or("CStr requires 1 argument")?.to_string();
                Ok(Variant::Str(s))
            }
            "val" => {
                let s = match vals.first().ok_or("Val requires 1 argument")? {
                    Variant::Str(s) => s.trim().to_string(),
                    v => v.to_string(),
                };
                let f = s.parse::<f64>().unwrap_or(0.0);
                Ok(as_int_if_whole(f))
            }
            "len" => {
                let s = match vals.first().ok_or("Len requires 1 argument")? {
                    Variant::Str(s) => s.chars().count() as i64,
                    Variant::Empty  => 0,
                    v               => v.to_string().chars().count() as i64,
                };
                Ok(Variant::Integer(s))
            }
            "left" => {
                let s = vba_to_str(vals.get(0).ok_or("Left requires 2 arguments")?);
                let n = to_f64(vals.get(1).ok_or("Left requires 2 arguments")?)? as usize;
                Ok(Variant::Str(s.chars().take(n).collect()))
            }
            "right" => {
                let s = vba_to_str(vals.get(0).ok_or("Right requires 2 arguments")?);
                let n = to_f64(vals.get(1).ok_or("Right requires 2 arguments")?)? as usize;
                let chars: Vec<char> = s.chars().collect();
                Ok(Variant::Str(chars[chars.len().saturating_sub(n)..].iter().collect()))
            }
            "mid" => {
                if vals.len() < 2 { return Err("Mid requires at least 2 arguments".into()); }
                let s = vba_to_str(&vals[0]);
                let start = (to_f64(&vals[1])? as usize).saturating_sub(1);
                let len = if vals.len() >= 3 { to_f64(&vals[2])? as usize } else { usize::MAX };
                Ok(Variant::Str(s.chars().skip(start).take(len).collect()))
            }
            "ucase" => {
                Ok(Variant::Str(vba_to_str(vals.first().ok_or("UCase requires 1 argument")?).to_uppercase()))
            }
            "lcase" => {
                Ok(Variant::Str(vba_to_str(vals.first().ok_or("LCase requires 1 argument")?).to_lowercase()))
            }
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
            "isnull" | "isempty" => {
                Ok(Variant::Boolean(matches!(vals.first(), Some(Variant::Empty) | None)))
            }
            "isnumeric" => {
                Ok(Variant::Boolean(matches!(vals.first(), Some(Variant::Integer(_)) | Some(Variant::Float(_)))))
            }
            "chr" => {
                let n = to_f64(vals.first().ok_or("Chr requires 1 argument")?)? as u32;
                char::from_u32(n)
                    .map(|c| Variant::Str(c.to_string()))
                    .ok_or_else(|| format!("Chr: invalid code {}", n))
            }
            "asc" => {
                let s = vba_to_str(vals.first().ok_or("Asc requires 1 argument")?);
                s.chars().next()
                    .map(|c| Variant::Integer(c as i64))
                    .ok_or_else(|| "Asc: empty string".into())
            }
            "instr" => {
                // InStr([start,] string1, string2 [, compare])
                let (start, s1, s2) = if vals.len() >= 3 {
                    (to_f64(&vals[0])? as usize, vba_to_str(&vals[1]), vba_to_str(&vals[2]))
                } else {
                    (1, vba_to_str(vals.get(0).ok_or("InStr requires at least 2 arguments")?),
                        vba_to_str(vals.get(1).ok_or("InStr requires at least 2 arguments")?))
                };
                let h: Vec<char> = s1.chars().collect();
                let n: Vec<char> = s2.chars().collect();
                // VBA: empty needle → return start position; start > len → return 0
                if n.is_empty() { return Ok(Variant::Integer(start as i64)); }
                let begin = start.saturating_sub(1);
                if begin >= h.len() { return Ok(Variant::Integer(0)); }
                let pos = h[begin..].windows(n.len())
                    .position(|w| w.iter().map(|c| c.to_uppercase().next().unwrap_or(*c))
                        .eq(n.iter().map(|c| c.to_uppercase().next().unwrap_or(*c))))
                    .map(|p| p + start)
                    .unwrap_or(0);
                Ok(Variant::Integer(pos as i64))
            }
            "replace" => {
                if vals.len() < 3 { return Err("Replace requires at least 3 arguments".into()); }
                let s = vba_to_str(&vals[0]);
                let old = vba_to_str(&vals[1]);
                let new = vba_to_str(&vals[2]);
                Ok(Variant::Str(if old.is_empty() { s } else { s.replace(&old as &str, &new as &str) }))
            }
            "now" | "date" | "time" => {
                Ok(Variant::Str(format!("{:?}", std::time::SystemTime::now())))
            }
            // ── Inline conditional ───────────────────────────────────────────
            "iif" => {
                if vals.len() < 3 { return Err("IIf requires 3 arguments".into()); }
                Ok(if is_truthy(&vals[0]) { vals[1].clone() } else { vals[2].clone() })
            }
            // ── Format ───────────────────────────────────────────────────────
            "format" => {
                if vals.is_empty() { return Err("Format requires at least 1 argument".into()); }
                let v = &vals[0];
                let fmt = if vals.len() >= 2 { vba_to_str(&vals[1]) } else { String::new() };
                Ok(Variant::Str(format_vba(v, &fmt)))
            }
            // ── Type inspection ──────────────────────────────────────────────
            "typename" => {
                let name = match vals.first().ok_or("TypeName requires 1 argument")? {
                    Variant::Integer(_) => "Long",
                    Variant::Float(_)   => "Double",
                    Variant::Str(_)     => "String",
                    Variant::Boolean(_) => "Boolean",
                    Variant::Date(_)    => "Date",
                    Variant::Error(_)   => "Error",
                    Variant::Array(_)   => "Variant()",
                    Variant::Empty      => "Empty",
                    Variant::Record(_)  => "Object",
                };
                Ok(Variant::Str(name.into()))
            }
            "vartype" => {
                let n: i64 = match vals.first().ok_or("VarType requires 1 argument")? {
                    Variant::Empty      => 0,
                    Variant::Integer(_) => 3,  // vbLong
                    Variant::Float(_)   => 5,  // vbDouble
                    Variant::Str(_)     => 8,  // vbString
                    Variant::Boolean(_) => 11, // vbBoolean
                    Variant::Date(_)    => 7,  // vbDate
                    Variant::Array(_)   => 8204, // vbArray + vbVariant
                    Variant::Error(_)   => 10, // vbError
                    Variant::Record(_)  => 0,  // vbEmpty as fallback
                };
                Ok(Variant::Integer(n))
            }
            // ── Array functions ──────────────────────────────────────────────
            "split" => {
                if vals.is_empty() { return Err("Split requires at least 1 argument".into()); }
                let s = vba_to_str(&vals[0]);
                let delim = if vals.len() >= 2 { vba_to_str(&vals[1]) } else { " ".to_string() };
                let parts = s.split(delim.as_str()).map(|p| Variant::Str(p.to_string())).collect();
                Ok(Variant::Array(parts))
            }
            "join" => {
                if vals.is_empty() { return Err("Join requires at least 1 argument".into()); }
                let parts = match &vals[0] {
                    Variant::Array(a) => a.iter().map(|v| vba_to_str(v)).collect::<Vec<_>>(),
                    v                 => vec![vba_to_str(v)],
                };
                let delim = if vals.len() >= 2 { vba_to_str(&vals[1]) } else { " ".to_string() };
                Ok(Variant::Str(parts.join(&delim)))
            }
            "ubound" => {
                match vals.first().ok_or("UBound requires 1 argument")? {
                    Variant::Array(a) => Ok(Variant::Integer(a.len() as i64 - 1)),
                    _ => Err("UBound: argument is not an array".into()),
                }
            }
            "lbound" => {
                match vals.first().ok_or("LBound requires 1 argument")? {
                    Variant::Array(_) => Ok(Variant::Integer(0)),
                    _ => Err("LBound: argument is not an array".into()),
                }
            }
            "isarray" => {
                Ok(Variant::Boolean(matches!(vals.first(), Some(Variant::Array(_)))))
            }
            // ── Range object (used as WSF arg) ───────────────────────────────
            "range" => {
                if let Some(Variant::Str(addr)) = vals.first() {
                    let ((r1,c1),(r2,c2)) = self.resolve_range_addr(addr)
                        .ok_or_else(|| format!("Range: invalid address '{}'", addr))?;
                    let arr = (r1..=r2).flat_map(|r| (c1..=c2).map(move |c| (r,c)))
                        .map(|(r,c)| self.get_cell(r,c)).collect();
                    Ok(Variant::Array(arr))
                } else { Err("Range: requires a string address argument".into()) }
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
        self.cells().get(&(row, col)).map(|c| c.value.clone()).unwrap_or(Variant::Empty)
    }

    pub fn set_cell_formula(&mut self, row: u32, col: u32, formula: &str) -> Result<(), String> {
        let expr  = formula::parse(formula)?;
        let value = formula::evaluate(&expr, self.cells())?;
        self.cells_mut().insert((row, col), CellContent { formula: Some(formula.to_string()), value });
        Ok(())
    }

    pub fn recalculate_all(&mut self) -> Result<(), String> {
        // Collect all formula cells and parse them
        let formula_cells: Vec<(u32, u32, formula::FormulaExpr)> = {
            self.cells().iter()
                .filter_map(|((r, c), cell)| {
                    cell.formula.as_ref().and_then(|f| formula::parse(f).ok().map(|e| (*r, *c, e)))
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
            if let Some(cell) = self.sheets.get_mut(&active).and_then(|m| m.get_mut(&(row, col))) {
                cell.value = value;
            }
        }
        // Mark index dirty once (formula values changed, End queries may be stale)
        if !formula_cells.is_empty() { self.cell_index_dirty = true; }
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
        if self.cell_index_dirty { self.rebuild_cell_index(); }
        self.col_rows.get(&col)
            .and_then(|rows| rows.range(..=max_row).next_back().copied())
            .unwrap_or(1)
    }

    /// Find the first empty row in `col` at or below `start_row` (xlDown helper).
    pub fn first_empty_row(&mut self, col: u32, start_row: u32) -> u32 {
        if self.cell_index_dirty { self.rebuild_cell_index(); }
        if let Some(rows) = self.col_rows.get(&col) {
            let mut prev = start_row.saturating_sub(1);
            for &r in rows.range(start_row..) {
                if r != prev + 1 { return prev + 1; }  // gap found
                prev = r;
            }
            prev + 1
        } else {
            start_row  // column is entirely empty
        }
    }

    /// Find the last non-empty column in `row` at or left of `max_col` (xlToLeft).
    pub fn last_nonempty_col(&mut self, row: u32, max_col: u32) -> u32 {
        if self.cell_index_dirty { self.rebuild_cell_index(); }
        self.row_cols.get(&row)
            .and_then(|cols| cols.range(..=max_col).next_back().copied())
            .unwrap_or(1)
    }

    /// Find the first empty column in `row` at or right of `start_col` (xlToRight helper).
    pub fn first_empty_col(&mut self, row: u32, start_col: u32) -> u32 {
        if self.cell_index_dirty { self.rebuild_cell_index(); }
        if let Some(cols) = self.row_cols.get(&row) {
            let mut prev = start_col.saturating_sub(1);
            for &c in cols.range(start_col..) {
                if c != prev + 1 { return prev + 1; }
                prev = c;
            }
            prev + 1
        } else {
            start_col
        }
    }
}

impl Default for Vm {
    fn default() -> Self { Self::new() }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

// ── Range address helpers ─────────────────────────────────────────────────────

fn col_letters_to_num_vm(s: &str) -> u32 {
    s.chars().fold(0u32, |acc, c| acc * 26 + (c.to_ascii_uppercase() as u32 - b'A' as u32 + 1))
}

pub fn parse_cell_addr(addr: &str) -> Option<(u32, u32)> {
    let addr = addr.trim().to_uppercase();
    let alpha_end = addr.find(|c: char| c.is_ascii_digit())?;
    if alpha_end == 0 { return None; }
    let col = col_letters_to_num_vm(&addr[..alpha_end]);
    let row: u32 = addr[alpha_end..].parse().ok()?;
    Some((row, col))
}

pub fn parse_range_addr(addr: &str) -> Option<((u32, u32), (u32, u32))> {
    let addr = addr.trim();
    if let Some(i) = addr.find(':') {
        Some((parse_cell_addr(&addr[..i])?, parse_cell_addr(&addr[i+1..])?))
    } else {
        let c = parse_cell_addr(addr)?;
        Some((c, c))
    }
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
    let map: HashMap<String, Variant> = fields.iter().map(|(name, vba_type)| {
        let default = match vba_type.as_str() {
            "integer" | "long" | "longlong" | "byte" => Variant::Integer(0),
            "single" | "double" | "currency" | "decimal" => Variant::Float(0.0),
            "boolean" => Variant::Boolean(false),
            "string"  => Variant::Str(String::new()),
            other => {
                if let Some(nested) = type_defs.get(other) {
                    make_record_default(nested, type_defs)
                } else {
                    Variant::Empty
                }
            }
        };
        (name.clone(), default)
    }).collect();
    Variant::Record(map)
}

/// Recursively set a value at the path given by `fields` inside a `Variant::Record` tree.
fn nested_set(target: &mut Variant, fields: &[String], value: Variant) {
    if fields.is_empty() { *target = value; return; }
    let field = &fields[0];
    let rest  = &fields[1..];
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
        CellRef { col, row }                 => [(*row, *col)].into(),
        Range { c1, r1, c2, r2 }             => {
            let mut s = HashSet::new();
            for r in *r1..=*r2 { for c in *c1..=*c2 { s.insert((r, c)); } }
            s
        }
        BinOp { lhs, rhs, .. }               => { let mut s = extract_cell_refs(lhs); s.extend(extract_cell_refs(rhs)); s }
        UnaryMinus(inner)                    => extract_cell_refs(inner),
        FuncCall { args, .. }                => args.iter().flat_map(extract_cell_refs).collect(),
        Number(_) | Str(_) | Bool(_)         => HashSet::new(),
    }
}

/// Topological sort of formula cells by dependency order.
/// Returns indices into `cells` in safe evaluation order.
/// Cells with no inter-formula dependencies appear first.
/// Returns `Err` if a circular reference is detected.
fn topo_sort_formulas(
    cells: &[(u32, u32, formula::FormulaExpr)],
) -> Result<Vec<usize>, String> {
    let n = cells.len();
    // map (row, col) → index in cells slice
    let pos: HashMap<(u32, u32), usize> = cells.iter().enumerate().map(|(i, (r,c,_))| ((*r,*c), i)).collect();

    // in_degree[i] = number of formula cells that i depends on
    let mut in_degree = vec![0usize; n];
    // adj[j] = list of formula cells that depend on j
    let mut adj: Vec<Vec<usize>> = vec![vec![]; n];

    for (i, (_, _, expr)) in cells.iter().enumerate() {
        for dep in extract_cell_refs(expr) {
            if let Some(&j) = pos.get(&dep) {
                if j != i { // skip self-reference
                    adj[j].push(i);
                    in_degree[i] += 1;
                }
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
            if in_degree[j] == 0 { queue.push_back(j); }
        }
    }

    if order.len() != n {
        // Circular reference detected — evaluate remaining cells in original order with a warning
        let visited: HashSet<usize> = order.iter().copied().collect();
        for i in 0..n { if !visited.contains(&i) { order.push(i); } }
        // Return Ok with best-effort order rather than hard-erroring; circular refs will show stale values
    }
    Ok(order)
}

fn vba_to_str(v: &Variant) -> String {
    match v {
        Variant::Str(s)     => s.clone(),
        Variant::Integer(n) => n.to_string(),
        Variant::Float(f)   => f.to_string(),
        Variant::Boolean(b) => if *b { "True".into() } else { "False".into() },
        Variant::Date(s)    => serial_to_display(*s),
        Variant::Error(e)   => e.as_str().to_string(),
        Variant::Empty      => String::new(),
        Variant::Array(a)   => a.iter().map(|x| vba_to_str(x)).collect::<Vec<_>>().join(", "),
        Variant::Record(_)  => "[Record]".into(),
    }
}

fn cmp_variants(a: &Variant, b: &Variant) -> std::cmp::Ordering {
    use std::cmp::Ordering::*;
    match (a, b) {
        (Variant::Integer(x), Variant::Integer(y)) => x.cmp(y),
        (Variant::Float(x),   Variant::Float(y))   => x.partial_cmp(y).unwrap_or(Equal),
        (Variant::Integer(x), Variant::Float(y))   => (*x as f64).partial_cmp(y).unwrap_or(Equal),
        (Variant::Float(x),   Variant::Integer(y)) => x.partial_cmp(&(*y as f64)).unwrap_or(Equal),
        (Variant::Str(x),     Variant::Str(y))     => x.to_lowercase().cmp(&y.to_lowercase()),
        (Variant::Empty,      Variant::Empty)       => Equal,
        (Variant::Empty,      _)                   => Less,
        (_,                   Variant::Empty)       => Greater,
        _ => vba_to_str(a).to_lowercase().cmp(&vba_to_str(b).to_lowercase()),
    }
}

fn flat_nums(vals: &[Variant]) -> Vec<f64> {
    let mut out = vec![];
    for v in vals {
        match v {
            Variant::Array(a) => out.extend(a.iter().filter_map(|x| to_f64(x).ok())),
            _ => { if let Ok(f) = to_f64(v) { out.push(f); } }
        }
    }
    out
}

fn flat_all(vals: &[Variant]) -> Vec<Variant> {
    vals.iter().flat_map(|v| match v {
        Variant::Array(a) => a.clone(),
        other             => vec![other.clone()],
    }).collect()
}

fn eval_wsf(func: &str, vals: &[Variant]) -> Result<Variant, String> {
    match func {
        "sum" => {
            let nums = flat_nums(vals);
            Ok(as_int_if_whole(nums.iter().sum::<f64>()))
        }
        "max" => {
            let nums = flat_nums(vals);
            nums.iter().cloned().reduce(f64::max)
                .map(as_int_if_whole)
                .ok_or_else(|| "WorksheetFunction.Max: no values".into())
        }
        "min" => {
            let nums = flat_nums(vals);
            nums.iter().cloned().reduce(f64::min)
                .map(as_int_if_whole)
                .ok_or_else(|| "WorksheetFunction.Min: no values".into())
        }
        "average" => {
            let nums = flat_nums(vals);
            if nums.is_empty() { return Err("WorksheetFunction.Average: no values".into()); }
            Ok(as_int_if_whole(nums.iter().sum::<f64>() / nums.len() as f64))
        }
        "count" => {
            let n = flat_all(vals).iter().filter(|v| matches!(v, Variant::Integer(_) | Variant::Float(_))).count();
            Ok(Variant::Integer(n as i64))
        }
        "counta" => {
            let n = flat_all(vals).iter().filter(|v| !matches!(v, Variant::Empty)).count();
            Ok(Variant::Integer(n as i64))
        }
        "countblank" => {
            let n = flat_all(vals).iter().filter(|v| matches!(v, Variant::Empty)).count();
            Ok(Variant::Integer(n as i64))
        }
        "countif" => {
            if vals.len() < 2 { return Err("WorksheetFunction.CountIf requires 2 arguments".into()); }
            let range = flat_all(&vals[..1]);
            let criteria = &vals[1];
            let n = range.iter().filter(|v| wsf_criteria_match(v, criteria)).count();
            Ok(Variant::Integer(n as i64))
        }
        "sumif" => {
            // SumIf(range, criteria [, sum_range])
            if vals.len() < 2 { return Err("WorksheetFunction.SumIf requires at least 2 arguments".into()); }
            let crit_range = flat_all(&vals[..1]);
            let criteria   = &vals[1];
            let sum_range  = if vals.len() >= 3 { flat_all(&vals[2..3]) } else { crit_range.clone() };
            let total: f64 = crit_range.iter().zip(sum_range.iter())
                .filter(|(cv, _)| wsf_criteria_match(cv, criteria))
                .filter_map(|(_, sv)| to_f64(sv).ok())
                .sum();
            Ok(as_int_if_whole(total))
        }
        "round" => {
            if vals.is_empty() { return Err("WorksheetFunction.Round requires arguments".into()); }
            let f = to_f64(&vals[0])?;
            let digits = if vals.len() >= 2 { to_f64(&vals[1])? as i32 } else { 0 };
            let factor = 10f64.powi(digits);
            Ok(as_int_if_whole((f * factor).round() / factor))
        }
        "abs"   => { let f = to_f64(vals.first().ok_or("Abs: no arg")?)?; Ok(as_int_if_whole(f.abs())) }
        "sqrt"  => { let f = to_f64(vals.first().ok_or("Sqrt: no arg")?)?; Ok(Variant::Float(f.sqrt())) }
        "power" => {
            if vals.len() < 2 { return Err("Power requires 2 arguments".into()); }
            Ok(as_int_if_whole(to_f64(&vals[0])?.powf(to_f64(&vals[1])?)))
        }
        "log"   => {
            let x = to_f64(vals.first().ok_or("Log: no arg")?)?;
            let base = if vals.len() >= 2 { to_f64(&vals[1])? } else { std::f64::consts::E };
            Ok(Variant::Float(x.log(base)))
        }
        "match" => {
            // Match(lookup_val, lookup_array, [match_type]) — returns 1-based position
            if vals.len() < 2 { return Err("Match: requires at least 2 arguments".into()); }
            let target  = &vals[0];
            let arr     = flat_all(&vals[1..2]);
            let pos = arr.iter().position(|v| vba_eq(v, target))
                .map(|i| Variant::Integer(i as i64 + 1))
                .unwrap_or(Variant::Error(ExcelError::NA));
            Ok(pos)
        }
        "index" => {
            // Index(array, row_num [, col_num])
            if vals.len() < 2 { return Err("Index: requires at least 2 arguments".into()); }
            let arr = flat_all(&vals[0..1]);
            let idx = (to_f64(&vals[1])? as usize).saturating_sub(1);
            Ok(arr.get(idx).cloned().unwrap_or(Variant::Error(ExcelError::Ref)))
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

fn wsf_criteria_match(v: &Variant, criteria: &Variant) -> bool {
    match criteria {
        Variant::Str(s) => {
            let s = s.trim();
            // Comparison criteria like ">5", "<>0", ">=10"
            if let Some(rest) = s.strip_prefix(">=") {
                if let Ok(n) = rest.parse::<f64>() { return to_f64(v).map_or(false, |f| f >= n); }
            } else if let Some(rest) = s.strip_prefix("<=") {
                if let Ok(n) = rest.parse::<f64>() { return to_f64(v).map_or(false, |f| f <= n); }
            } else if let Some(rest) = s.strip_prefix("<>") {
                if let Ok(n) = rest.parse::<f64>() { return to_f64(v).map_or(false, |f| f != n); }
                return vba_to_str(v).to_lowercase() != rest.to_lowercase();
            } else if let Some(rest) = s.strip_prefix('>') {
                if let Ok(n) = rest.parse::<f64>() { return to_f64(v).map_or(false, |f| f > n); }
            } else if let Some(rest) = s.strip_prefix('<') {
                if let Ok(n) = rest.parse::<f64>() { return to_f64(v).map_or(false, |f| f < n); }
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
    let dec_places = fmt.find('.').map(|i| fmt[i+1..].chars().filter(|c| *c == '0' || *c == '#').count()).unwrap_or(0);
    match v {
        Variant::Integer(n) => {
            let f = *n as f64;
            if thousands {
                // Simple thousands separator
                let int_part = format!("{}", n.abs());
                let grouped: String = int_part.chars().rev().enumerate()
                    .flat_map(|(i, c)| if i > 0 && i % 3 == 0 { vec![',', c] } else { vec![c] })
                    .collect::<String>().chars().rev().collect();
                let signed = if *n < 0 { format!("-{}", grouped) } else { grouped };
                if dec_places > 0 { format!("{}.{}", signed, "0".repeat(dec_places)) } else { signed }
            } else if dec_places > 0 {
                format!("{:.prec$}", f, prec = dec_places)
            } else {
                format!("{}", n)
            }
        }
        Variant::Float(f) => {
            if thousands {
                let int_part = format!("{}", (*f as i64).abs());
                let grouped: String = int_part.chars().rev().enumerate()
                    .flat_map(|(i, c)| if i > 0 && i % 3 == 0 { vec![',', c] } else { vec![c] })
                    .collect::<String>().chars().rev().collect();
                let signed = if *f < 0.0 { format!("-{}", grouped) } else { grouped };
                if dec_places > 0 { format!("{}.{:.prec$}", signed, f.fract().abs(), prec = dec_places)
                } else { signed }
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
        Variant::Float(f)   => Ok(*f),
        Variant::Boolean(b) => Ok(if *b { 1.0 } else { 0.0 }),
        Variant::Date(s)    => Ok(*s as f64),
        Variant::Error(e)   => Err(e.to_string()),
        Variant::Empty      => Ok(0.0),
        Variant::Str(s)     => s.parse::<f64>().map_err(|_| format!("Cannot convert '{}' to number", s)),
        Variant::Array(_)   => Err("Cannot convert array to number".into()),
        Variant::Record(_)  => Err("Cannot convert record to number".into()),
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

fn vba_eq(a: &Variant, b: &Variant) -> bool {
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

fn vba_cmp(a: &Variant, b: &Variant) -> Result<std::cmp::Ordering, String> {
    // String operands: case-insensitive lexicographic comparison (VBA default).
    if let (Variant::Str(sa), Variant::Str(sb)) = (a, b) {
        return Ok(sa.to_uppercase().cmp(&sb.to_uppercase()));
    }
    // Mixed string/number: try numeric first, fall back to string coercion.
    match (to_f64(a), to_f64(b)) {
        (Ok(fa), Ok(fb)) => fa.partial_cmp(&fb).ok_or_else(|| "Cannot compare NaN values".into()),
        _ => {
            let sa = match a { Variant::Str(s) => s.clone(), _ => format!("{:?}", a) };
            let sb = match b { Variant::Str(s) => s.clone(), _ => format!("{:?}", b) };
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

fn to_cell_index(v: Variant, label: &str) -> Result<u32, String> {
    let f = to_f64(&v)?;
    if f < 1.0 || f.fract() != 0.0 {
        return Err(format!("Cell {} must be a positive integer, got {}", label, f));
    }
    Ok(f as u32)
}

fn eval_binop(op: &VbaBinOp, l: Variant, r: Variant) -> Result<Variant, String> {
    match op {
        VbaBinOp::Add | VbaBinOp::Sub | VbaBinOp::Mul | VbaBinOp::Div => {
            let lf = to_f64(&l)?;
            let rf = to_f64(&r)?;
            let result = match op {
                VbaBinOp::Add => lf + rf,
                VbaBinOp::Sub => lf - rf,
                VbaBinOp::Mul => lf * rf,
                VbaBinOp::Div => {
                    if rf == 0.0 { return Err("Division by zero".into()); }
                    lf / rf
                }
                _ => unreachable!(),
            };
            Ok(as_int_if_whole(result))
        }
        VbaBinOp::Concat => Ok(Variant::Str(format!("{}{}", l, r))),
        VbaBinOp::Eq  => Ok(Variant::Boolean(vba_eq(&l, &r))),
        VbaBinOp::Ne  => Ok(Variant::Boolean(!vba_eq(&l, &r))),
        VbaBinOp::Lt  => Ok(Variant::Boolean(vba_cmp(&l, &r)? == std::cmp::Ordering::Less)),
        VbaBinOp::Le  => Ok(Variant::Boolean(vba_cmp(&l, &r)? != std::cmp::Ordering::Greater)),
        VbaBinOp::Gt  => Ok(Variant::Boolean(vba_cmp(&l, &r)? == std::cmp::Ordering::Greater)),
        VbaBinOp::Ge  => Ok(Variant::Boolean(vba_cmp(&l, &r)? != std::cmp::Ordering::Less)),
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
        assert_eq!(run("Sub MySub()\n    a = 42\nEnd Sub\n").variables["a"], Variant::Integer(42));
    }

    #[test]
    fn test_variable_assignment_float() {
        assert_eq!(run("Sub MySub()\n    x = 1.5\nEnd Sub\n").variables["x"], Variant::Float(1.5));
    }

    #[test]
    fn test_variable_assignment_string() {
        assert_eq!(run("Sub MySub()\n    s = \"hello\"\nEnd Sub\n").variables["s"], Variant::Str("hello".into()));
    }

    #[test]
    fn test_cell_write_literal() {
        assert_eq!(run("Sub MySub()\n    Cells(1, 1).Value = 100\nEnd Sub\n").get_cell(1, 1), Variant::Integer(100));
    }

    #[test]
    fn test_cell_write_from_variable() {
        assert_eq!(run("Sub MySub()\n    x = 99\n    Cells(2, 3).Value = x\nEnd Sub\n").get_cell(2, 3), Variant::Integer(99));
    }

    #[test]
    fn test_cell_write_string() {
        assert_eq!(run("Sub MySub()\n    Cells(1, 2).Value = \"world\"\nEnd Sub\n").get_cell(1, 2), Variant::Str("world".into()));
    }

    #[test]
    fn test_cell_empty_by_default() {
        assert_eq!(Vm::new().get_cell(1, 1), Variant::Empty);
    }

    #[test]
    fn test_multiple_cells() {
        let vm = run("Sub MySub()\n    Cells(1, 1).Value = 1\n    Cells(1, 2).Value = 2\n    Cells(2, 1).Value = 3\nEnd Sub\n");
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
    fn test_variant_display() {
        assert_eq!(Variant::Integer(42).to_string(), "42");
        assert_eq!(Variant::Float(3.14).to_string(), "3.14");
        assert_eq!(Variant::Boolean(true).to_string(), "True");
        assert_eq!(Variant::Empty.to_string(), "");
    }

    // ── Arithmetic ────────────────────────────────────────────────────────────

    #[test]
    fn test_arithmetic_assignment() {
        let vm = run("Sub MySub()\n    a = 3 + 4\n    b = 10 - 3\n    c = 2 * 5\n    d = 10 / 4\nEnd Sub\n");
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
        ).unwrap();
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
        let vm = run("Sub MySub()\n    x = 0\n    Do While x < 5\n        x = x + 1\n    Loop\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test]
    fn test_do_until_loop() {
        let vm = run("Sub MySub()\n    x = 0\n    Do Until x >= 5\n        x = x + 1\n    Loop\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test]
    fn test_do_loop_while_post() {
        // Post-check: body runs at least once even if condition is already false
        let vm = run("Sub MySub()\n    x = 99\n    Do\n        x = x + 1\n    Loop While x < 5\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(100));
    }

    // ── Select Case ───────────────────────────────────────────────────────────

    #[test]
    fn test_select_case_value() {
        let vm = run("Sub MySub()\n    x = 2\n    Select Case x\n        Case 1\n            r = \"one\"\n        Case 2\n            r = \"two\"\n        Case Else\n            r = \"other\"\n    End Select\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Str("two".into()));
    }

    #[test]
    fn test_select_case_else() {
        let vm = run("Sub MySub()\n    x = 99\n    Select Case x\n        Case 1\n            r = 1\n        Case Else\n            r = 0\n    End Select\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(0));
    }

    #[test]
    fn test_select_case_multi_value() {
        let vm = run("Sub MySub()\n    x = 3\n    Select Case x\n        Case 1, 2\n            r = 12\n        Case 3, 4\n            r = 34\n    End Select\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(34));
    }

    #[test]
    fn test_select_case_is_op() {
        let vm = run("Sub MySub()\n    x = 10\n    Select Case x\n        Case Is > 5\n            r = 1\n        Case Else\n            r = 0\n    End Select\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(1));
    }

    #[test]
    fn test_select_case_range() {
        let vm = run("Sub MySub()\n    x = 3\n    Select Case x\n        Case 1 To 5\n            r = 1\n        Case Else\n            r = 0\n    End Select\nEnd Sub\n");
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
        let vm = run("Sub MySub()\n    With Sheet1\n        .Cells(1, 1).Value = 100\n        .Cells(2, 1).Value = 200\n    End With\nEnd Sub\n");
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
    }

    #[test]
    fn test_vba_len() {
        let vm = run("Sub MySub()\n    a = Len(\"Hello\")\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(5));
    }

    #[test]
    fn test_vba_mid_left_right() {
        let vm = run("Sub MySub()\n    a = Mid(\"Hello\", 2, 3)\n    b = Left(\"Hello\", 3)\n    c = Right(\"Hello\", 2)\nEnd Sub\n");
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

    // ── ElseIf ───────────────────────────────────────────────────────────────

    #[test]
    fn test_elseif_chain() {
        let vm = run("Sub MySub()\n    x = 7\n    If x > 10 Then\n        r = 1\n    ElseIf x > 5 Then\n        r = 2\n    Else\n        r = 3\n    End If\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(2));
    }

    // ── Exit ─────────────────────────────────────────────────────────────────

    #[test]
    fn test_exit_for() {
        let vm = run("Sub MySub()\n    s = 0\n    For i = 1 To 10\n        If i > 3 Then\n            Exit For\n        End If\n        s = s + i\n    Next i\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Integer(6)); // 1+2+3
    }

    #[test]
    fn test_exit_do() {
        let vm = run("Sub MySub()\n    x = 0\n    Do\n        x = x + 1\n        If x >= 5 Then\n            Exit Do\n        End If\n    Loop While x < 100\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    // ── On Error Resume Next ──────────────────────────────────────────────────

    #[test]
    fn test_on_error_resume_next() {
        let vm = run("Sub MySub()\n    On Error Resume Next\n    a = 1\n    b = 2\n    a = 1\nEnd Sub\n");
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
            "    x = 99\n",               // should be skipped
            "    Exit Sub\n",
            "ErrH:\n",
            "    handled = 1\n",
            "End Sub\n",
        );
        let vm = run(code);
        assert_eq!(vm.variables["x"], Variant::Integer(1));   // set before error
        assert!(!vm.variables.contains_key("x") || vm.variables["x"] != Variant::Integer(99)); // not 99
        assert_eq!(vm.variables["handled"], Variant::Integer(1)); // handler ran
    }

    #[test]
    fn test_goto_unconditional() {
        let code = concat!(
            "Sub MySub()\n",
            "    a = 1\n",
            "    GoTo Skip\n",
            "    a = 99\n",  // should be skipped
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
            "    result = p.y\n",  // p.y not set → Empty → 0 in arithmetic
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
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    total = 0\n    For Each cell In Range(\"A1:A3\")\n        total = total + cell\n    Next cell\nEnd Sub\n");
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
        let vm = run("Function Square(n)\n    Square = n * n\nEnd Function\nSub MySub()\n    result = Square(7)\nEnd Sub\n");
        assert_eq!(vm.variables["result"], Variant::Integer(49));
    }

    #[test]
    fn test_function_return_value_nested() {
        let vm = run("Function Add(a, b)\n    Add = a + b\nEnd Function\nSub MySub()\n    x = Add(3, 4) + Add(1, 2)\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(10));
    }

    #[test]
    fn test_function_recursive() {
        // Factorial: 5! = 120
        let vm = run("Function Fact(n)\n    If n <= 1 Then\n        Fact = 1\n    Else\n        Fact = n * Fact(n - 1)\n    End If\nEnd Function\nSub MySub()\n    result = Fact(5)\nEnd Sub\n");
        assert_eq!(vm.variables["result"], Variant::Integer(120));
    }

    #[test]
    fn test_function_in_cell_write() {
        let vm = run("Function Double(x)\n    Double = x * 2\nEnd Function\nSub MySub()\n    Cells(1, 1).Value = Double(21)\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
    }

    #[test]
    fn test_call_sub_with_args() {
        let prog = parser::parse("Sub FillRow(rowNum, val)\n    Cells(rowNum, 1).Value = val\nEnd Sub\nSub MySub()\n    Call FillRow(3, 99)\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        vm.run_sub(&prog, "mysub").unwrap();
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(99));
    }

    // ── vb constants ─────────────────────────────────────────────────────────

    #[test]
    fn test_vb_string_constants() {
        let vm = run("Sub MySub()\n    a = \"Hello\" & vbCrLf & \"World\"\n    b = \"tab\" & vbTab & \"here\"\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Str("Hello\r\nWorld".into()));
        assert_eq!(vm.variables["b"], Variant::Str("tab\there".into()));
    }

    // ── While ... Wend ───────────────────────────────────────────────────────

    #[test] fn test_while_wend() {
        let vm = run("Sub MySub()\n    x = 0\n    While x < 5\n        x = x + 1\n    Wend\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(5));
    }

    #[test] fn test_while_wend_no_iteration() {
        let vm = run("Sub MySub()\n    x = 10\n    While x < 5\n        x = x + 1\n    Wend\n    y = 99\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(10));
        assert_eq!(vm.variables["y"], Variant::Integer(99));
    }

    // ── Const ─────────────────────────────────────────────────────────────────

    #[test] fn test_const_declaration() {
        let vm = run("Sub MySub()\n    Const MAX_ROW As Long = 100\n    x = MAX_ROW\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(100));
    }

    #[test] fn test_const_string() {
        let vm = run("Sub MySub()\n    Const PREFIX = \"ID_\"\n    s = PREFIX & \"001\"\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("ID_001".into()));
    }

    // ── Empty / Null / Nothing ────────────────────────────────────────────────

    #[test] fn test_empty_literal() {
        let vm = run("Sub MySub()\n    a = Empty\n    b = Null\n    c = Nothing\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Empty);
        assert_eq!(vm.variables["b"], Variant::Empty);
        assert_eq!(vm.variables["c"], Variant::Empty);
    }

    // ── IIf ──────────────────────────────────────────────────────────────────

    #[test] fn test_iif_true() {
        let vm = run("Sub MySub()\n    x = IIf(1 > 0, \"yes\", \"no\")\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Str("yes".into()));
    }

    #[test] fn test_iif_false() {
        let vm = run("Sub MySub()\n    x = IIf(0 > 1, 10, 20)\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(20));
    }

    // ── Format ───────────────────────────────────────────────────────────────

    #[test] fn test_format_decimal() {
        let vm = run("Sub MySub()\n    s = Format(3.14159, \"0.00\")\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("3.14".into()));
    }

    #[test] fn test_format_integer_no_dec() {
        let vm = run("Sub MySub()\n    s = Format(42, \"0\")\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("42".into()));
    }

    // ── TypeName / VarType ───────────────────────────────────────────────────

    #[test] fn test_typename() {
        let vm = run("Sub MySub()\n    a = TypeName(42)\n    b = TypeName(\"hi\")\n    c = TypeName(True)\n    d = TypeName(Empty)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Str("Long".into()));
        assert_eq!(vm.variables["b"], Variant::Str("String".into()));
        assert_eq!(vm.variables["c"], Variant::Str("Boolean".into()));
        assert_eq!(vm.variables["d"], Variant::Str("Empty".into()));
    }

    // ── Arrays ───────────────────────────────────────────────────────────────

    #[test] fn test_dim_array_write_read() {
        let vm = run("Sub MySub()\n    Dim arr(5)\n    arr(0) = 10\n    arr(3) = 99\n    a = arr(0)\n    b = arr(3)\n    c = arr(1)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(10));
        assert_eq!(vm.variables["b"], Variant::Integer(99));
        assert_eq!(vm.variables["c"], Variant::Empty);
    }

    #[test] fn test_dim_array_loop() {
        let vm = run("Sub MySub()\n    Dim arr(4)\n    For i = 0 To 4\n        arr(i) = i * 2\n    Next i\n    s = 0\n    For i = 0 To 4\n        s = s + arr(i)\n    Next i\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Integer(20)); // 0+2+4+6+8
    }

    #[test] fn test_redim_preserve() {
        let vm = run("Sub MySub()\n    Dim arr(2)\n    arr(0) = 1\n    arr(1) = 2\n    ReDim Preserve arr(4)\n    arr(3) = 99\n    a = arr(0)\n    b = arr(3)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(99));
    }

    #[test] fn test_ubound_lbound() {
        let vm = run("Sub MySub()\n    Dim arr(9)\n    u = UBound(arr)\n    l = LBound(arr)\nEnd Sub\n");
        assert_eq!(vm.variables["u"], Variant::Integer(9));
        assert_eq!(vm.variables["l"], Variant::Integer(0));
    }

    #[test] fn test_split_join() {
        let vm = run("Sub MySub()\n    arr = Split(\"a,b,c\", \",\")\n    s = Join(arr, \"-\")\n    n = UBound(arr)\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Str("a-b-c".into()));
        assert_eq!(vm.variables["n"], Variant::Integer(2));
    }

    #[test] fn test_isarray() {
        let vm = run("Sub MySub()\n    Dim arr(3)\n    a = IsArray(arr)\n    b = IsArray(42)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Boolean(true));
        assert_eq!(vm.variables["b"], Variant::Boolean(false));
    }

    // ── Application properties ────────────────────────────────────────────────

    #[test]
    fn test_app_prop_noop() {
        let vm = run("Sub MySub()\n    Application.ScreenUpdating = False\n    Application.EnableEvents = False\n    x = 1\n    Application.ScreenUpdating = True\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1));
    }

    #[test]
    fn test_xl_constants() {
        let vm = run("Sub MySub()\n    a = xlUp\n    b = xlDown\n    c = xlCalculationManual\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(-4162));
        assert_eq!(vm.variables["b"], Variant::Integer(-4121));
        assert_eq!(vm.variables["c"], Variant::Integer(-4135));
    }

    // ── Range write / copy / read ─────────────────────────────────────────────

    #[test]
    fn test_range_write_value() {
        let vm = run("Sub MySub()\n    Range(\"A1\").Value = 42\n    Range(\"B2\").Value = \"hello\"\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
        assert_eq!(vm.get_cell(2, 2), Variant::Str("hello".into()));
    }

    #[test]
    fn test_range_write_formula() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(1,2).Value = 20\n    Range(\"C1\").Formula = \"=SUM(A1:B1)\"\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 3), Variant::Integer(30));
    }

    #[test]
    fn test_range_copy() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    Range(\"A1:A3\").Copy Destination:=Range(\"B1\")\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 2), Variant::Integer(10));
        assert_eq!(vm.get_cell(2, 2), Variant::Integer(20));
        assert_eq!(vm.get_cell(3, 2), Variant::Integer(30));
    }

    #[test]
    fn test_range_read_expr() {
        let vm = run("Sub MySub()\n    Cells(5,1).Value = 99\n    x = Range(\"A5\").Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(99));
    }

    // ── WorksheetFunction ────────────────────────────────────────────────────

    #[test] fn test_wsf_sum() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    s = WorksheetFunction.Sum(Range(\"A1:A3\"))\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Integer(60));
    }

    #[test] fn test_wsf_max_min() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 5\n    Cells(2,1).Value = 3\n    Cells(3,1).Value = 8\n    mx = WorksheetFunction.Max(Range(\"A1:A3\"))\n    mn = WorksheetFunction.Min(Range(\"A1:A3\"))\nEnd Sub\n");
        assert_eq!(vm.variables["mx"], Variant::Integer(8));
        assert_eq!(vm.variables["mn"], Variant::Integer(3));
    }

    #[test] fn test_wsf_average() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    av = WorksheetFunction.Average(Range(\"A1:A2\"))\nEnd Sub\n");
        assert_eq!(vm.variables["av"], Variant::Integer(15));
    }

    #[test] fn test_wsf_countif() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 5\n    Cells(2,1).Value = 10\n    Cells(3,1).Value = 3\n    n = WorksheetFunction.CountIf(Range(\"A1:A3\"), \">4\")\nEnd Sub\n");
        assert_eq!(vm.variables["n"], Variant::Integer(2));
    }

    #[test] fn test_wsf_sumif() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 5\n    Cells(2,1).Value = 10\n    Cells(3,1).Value = 3\n    s = WorksheetFunction.SumIf(Range(\"A1:A3\"), \">4\")\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Integer(15));
    }

    #[test] fn test_wsf_application_prefix() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 7\n    Cells(2,1).Value = 3\n    s = Application.WorksheetFunction.Sum(Range(\"A1:A2\"))\nEnd Sub\n");
        assert_eq!(vm.variables["s"], Variant::Integer(10));
    }

    #[test] fn test_wsf_match() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = \"a\"\n    Cells(2,1).Value = \"b\"\n    Cells(3,1).Value = \"c\"\n    pos = WorksheetFunction.Match(\"b\", Range(\"A1:A3\"), 0)\nEnd Sub\n");
        assert_eq!(vm.variables["pos"], Variant::Integer(2));
    }

    // ── Range("A1:A10").Value 多セル読み取り ─────────────────────────────────

    #[test] fn test_range_read_multi_cell() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    x = Range(\"A1:A3\").Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Array(vec![Variant::Integer(1), Variant::Integer(2), Variant::Integer(3)]));
    }

    #[test] fn test_range_read_single_cell_backward_compat() {
        let vm = run("Sub MySub()\n    Cells(5,1).Value = 99\n    x = Range(\"A5\").Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(99));
    }

    #[test] fn test_range_read_2d_row_major() {
        // A1=1, B1=2, A2=3, B2=4 → Range("A1:B2").Value → [1,2,3,4]
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(1,2).Value = 2\n    Cells(2,1).Value = 3\n    Cells(2,2).Value = 4\n    arr = Range(\"A1:B2\").Value\n    a = arr(0)\n    b = arr(1)\n    c = arr(2)\n    d = arr(3)\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
        assert_eq!(vm.variables["c"], Variant::Integer(3));
        assert_eq!(vm.variables["d"], Variant::Integer(4));
    }

    // ── Cells.Find ───────────────────────────────────────────────────────────

    #[test] fn test_cells_find_row() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = \"apple\"\n    Cells(2,1).Value = \"banana\"\n    Cells(3,1).Value = \"cherry\"\n    r = Cells.Find(What:=\"banana\").Row\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(2));
    }

    #[test] fn test_cells_find_column() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = \"x\"\n    Cells(1,2).Value = \"y\"\n    Cells(1,3).Value = \"z\"\n    c = Cells.Find(What:=\"y\").Column\nEnd Sub\n");
        assert_eq!(vm.variables["c"], Variant::Integer(2));
    }

    #[test] fn test_cells_find_not_found() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = \"a\"\n    r = Cells.Find(What:=\"missing\").Row\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(0));
    }

    #[test] fn test_cells_find_extra_kwargs() {
        let vm = run("Sub MySub()\n    Cells(2,1).Value = 42\n    r = Cells.Find(What:=42, LookIn:=xlValues, SearchDirection:=xlPrevious).Row\nEnd Sub\n");
        assert_eq!(vm.variables["r"], Variant::Integer(2));
    }

    // ── EntireRow / EntireColumn ──────────────────────────────────────────────

    #[test] fn test_entirerow_delete() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    Range(\"A2:A2\").EntireRow.Delete\n    x = Cells(2,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(3)); // 3 が行2に移動
    }

    #[test] fn test_entirerow_clear() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 99\n    Range(\"A1\").EntireRow.Clear\n    x = Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Empty);
    }

    #[test] fn test_entirerow_clear_contents() {
        let vm = run("Sub MySub()\n    Cells(2,1).Value = 55\n    Range(\"A2\").EntireRow.ClearContents\n    x = Cells(2,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Empty);
    }

    #[test] fn test_entirecolumn_delete() {
        // パースエラーなしで実行できることを確認
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Range(\"A1:A3\").EntireColumn.Delete\n    x = 1\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1));
    }

    #[test] fn test_range_noop_hidden() {
        let vm = run("Sub MySub()\n    Range(\"A1\").Hidden = True\n    x = 1\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1));
    }

    #[test] fn test_range_noop_interior_color() {
        let vm = run("Sub MySub()\n    Range(\"A1:B2\").Interior.Color = 3\n    x = 2\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(2));
    }

    #[test] fn test_range_noop_numberformat() {
        let vm = run("Sub MySub()\n    Range(\"A1\").NumberFormat = \"0.00\"\n    x = 3\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(3));
    }

    // ── Range.Delete / Range.Insert ──────────────────────────────────────────

    #[test] fn test_range_delete_shifts_up() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Cells(3,1).Value = 3\n    Cells(4,1).Value = 4\n    Range(\"A2:A2\").Delete\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(3)); // 3 shifted up
        assert_eq!(vm.variables["c"], Variant::Integer(4)); // 4 shifted up
    }

    #[test] fn test_range_insert_shifts_down() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    Cells(2,1).Value = 2\n    Range(\"A2:A2\").Insert\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Empty); // new empty row
        assert_eq!(vm.variables["c"], Variant::Integer(2)); // shifted down
    }

    // ── Range.Sort ───────────────────────────────────────────────────────────

    #[test] fn test_range_sort_ascending() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 3\n    Cells(2,1).Value = 1\n    Cells(3,1).Value = 2\n    Range(\"A1:A3\").Sort Key1:=Range(\"A1\"), Order1:=xlAscending\n    a = Cells(1,1).Value\n    b = Cells(2,1).Value\n    c = Cells(3,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(1));
        assert_eq!(vm.variables["b"], Variant::Integer(2));
        assert_eq!(vm.variables["c"], Variant::Integer(3));
    }

    #[test] fn test_range_sort_descending() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 3\n    Cells(2,1).Value = 1\n    Cells(3,1).Value = 2\n    Range(\"A1:A3\").Sort Key1:=Range(\"A1\"), Order1:=xlDescending\n    a = Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["a"], Variant::Integer(3));
    }

    // ── Range clear / offset / multi-cell write / Sheets() ───────────────────

    #[test] fn test_range_clear_contents() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 99\n    Range(\"A1\").ClearContents\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Empty);
    }

    #[test] fn test_range_clear() {
        let vm = run("Sub MySub()\n    Range(\"A1:A3\").Value = 5\n    Range(\"A1:A3\").Clear\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Empty);
        assert_eq!(vm.get_cell(2, 1), Variant::Empty);
        assert_eq!(vm.get_cell(3, 1), Variant::Empty);
    }

    #[test] fn test_range_write_multi_cell() {
        let vm = run("Sub MySub()\n    Range(\"A1:A3\").Value = 7\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(7));
        assert_eq!(vm.get_cell(2, 1), Variant::Integer(7));
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(7));
    }

    #[test] fn test_range_offset_read() {
        let vm = run("Sub MySub()\n    Cells(2,2).Value = 42\n    x = Range(\"A1\").Offset(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    #[test] fn test_range_offset_write() {
        let vm = run("Sub MySub()\n    Range(\"A1\").Offset(2,0).Value = 99\nEnd Sub\n");
        assert_eq!(vm.get_cell(3, 1), Variant::Integer(99));
    }

    #[test] fn test_sheets_cell_write() {
        let vm = run("Sub MySub()\n    Sheets(\"Sheet1\").Cells(1,1).Value = 123\nEnd Sub\n");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(123));
    }

    #[test] fn test_sheets_cell_read() {
        let vm = run("Sub MySub()\n    Sheets(\"Sheet1\").Cells(1,1).Value = 55\n    x = Sheets(\"Sheet1\").Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(55));
    }

    #[test] fn test_worksheets_cell_write() {
        let vm = run("Sub MySub()\n    Worksheets(\"Data\").Cells(2,3).Value = 77\nEnd Sub\n");
        // Now routes to "data" sheet, not the active "sheet1"
        let cell = vm.get_sheet_cells("data").and_then(|s| s.get(&(2,3))).map(|c| c.value.clone());
        assert_eq!(cell, Some(Variant::Integer(77)));
    }

    // ── Multi-sheet (Phase 9) ────────────────────────────────────────────────

    #[test] fn test_multisheet_write_read_different_sheets() {
        let vm = run("Sub MySub()\n    Sheets(\"Sheet1\").Cells(1,1).Value = 10\n    Sheets(\"Sheet2\").Cells(1,1).Value = 20\nEnd Sub\n");
        let s1 = vm.get_sheet_cells("sheet1").and_then(|s| s.get(&(1,1))).map(|c| c.value.clone());
        let s2 = vm.get_sheet_cells("sheet2").and_then(|s| s.get(&(1,1))).map(|c| c.value.clone());
        assert_eq!(s1, Some(Variant::Integer(10)));
        assert_eq!(s2, Some(Variant::Integer(20)));
    }

    #[test] fn test_multisheet_cross_sheet_read() {
        let vm = run("Sub MySub()\n    Sheets(\"Data\").Cells(1,1).Value = 42\n    x = Sheets(\"Data\").Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(42));
    }

    #[test] fn test_with_sheets_block() {
        let vm = run("Sub MySub()\n    With Sheets(\"Sheet2\")\n        .Cells(1,1).Value = 99\n    End With\n    x = Sheets(\"Sheet2\").Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(99));
    }

    #[test] fn test_with_sheets_restores_active() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 1\n    With Sheets(\"Sheet2\")\n        .Cells(1,1).Value = 2\n    End With\n    x = Cells(1,1).Value\nEnd Sub\n");
        assert_eq!(vm.variables["x"], Variant::Integer(1)); // active sheet unchanged
    }

    #[test] fn test_sheets_add() {
        let vm = run("Sub MySub()\n    Sheets.Add\n    n = 1\nEnd Sub\n");
        assert!(vm.sheet_names().len() >= 2);
    }

    #[test] fn test_sheets_delete() {
        let vm = run("Sub MySub()\n    Sheets(\"Sheet2\").Cells(1,1).Value = 5\n    Sheets(\"Sheet2\").Delete\n    n = 1\nEnd Sub\n");
        assert!(!vm.sheet_names().contains(&"sheet2".to_string()));
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
        assert_eq!(err, "Array 'arr': index 9 out of bounds (len=4)");
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

    #[test] fn test_sheet_names() {
        let vm = run("Sub MySub()\n    Sheets(\"Alpha\").Cells(1,1).Value = 1\n    Sheets(\"Beta\").Cells(1,1).Value = 2\nEnd Sub\n");
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
        let vm = run("Sub MySub()\n    Cells(1,1).Value = 10\n    Cells(2,1).Value = 20\n    Cells(3,1).Value = 30\n    lastRow = Cells(Rows.Count,1).End(xlUp).Row\nEnd Sub\n");
        assert_eq!(vm.variables["lastrow"], Variant::Integer(3));
    }

    #[test]
    fn test_cells_end_col_left() {
        let vm = run("Sub MySub()\n    Cells(1,1).Value = \"a\"\n    Cells(1,2).Value = \"b\"\n    Cells(1,3).Value = \"c\"\n    lastCol = Cells(1,Columns.Count).End(xlToLeft).Column\nEnd Sub\n");
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
        assert_eq!(vm.variables["xi"],  Variant::Integer(0));
        assert_eq!(vm.variables["yi"],  Variant::Integer(0));
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
            "    s = o.Val\n",       // default "" for String
            "    x = o.Pt.X\n",     // default 0 for nested Integer
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
            "    d = items(0).Value\n",  // default 0
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
        assert_eq!(vm.variables["rx"],    Variant::Integer(5));
        assert_eq!(vm.variables["ry"],    Variant::Integer(10));
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
        vm.cells_mut().insert((1,1), CellContent { formula: None, value: Variant::Integer(10) });
        // Force index rebuild
        let _ = vm.last_nonempty_row(1, 1_048_576);
        assert!(!vm.cell_index_dirty);
        // Write via sheet_cells_mut (simulated by inserting through the public method)
        vm.cells_mut().insert((5,1), CellContent { formula: None, value: Variant::Integer(50) });
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
        assert!(result.is_err(), "expected the undefined-variable error to propagate");
        assert_eq!(vm.take_messages(), vec!["seen before failure".to_string()]);
    }

    #[test]
    fn test_msgbox_blocked_is_recorded_before_failing() {
        let prog = parser::parse("Sub MySub()\n    MsgBox \"blocked\"\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        vm.error_on_msgbox = true;
        let result = vm.run_sub(&prog, "MySub");
        assert!(result.is_err(), "MsgBox must still fail when error_on_msgbox is set");
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

    // ── Stmt::Unsupported executes as a true no-op ──────────────────────────

    #[test]
    fn test_unsupported_stmt_is_a_true_noop() {
        // Range("A1").NumberFormat isn't a recognized Range property, so it
        // parses to Stmt::Unsupported — confirm it doesn't error and later
        // statements still run normally, exactly like Stmt::Dim.
        let vm = run(
            "Sub MySub()\n    Range(\"A1\").NumberFormat = \"0.00\"\n    x = 3\nEnd Sub\n",
        );
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
            module("module2", "Sub Main()\n    Call Helper()\n    x = 42\nEnd Sub\n"),
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
        assert!(!vm.variables.contains_key("x"), "no execution should have happened");
    }

    #[test]
    fn run_sub_multi_rejects_a_genuine_func_collision() {
        let modules = vec![
            module("module1", "Function Foo()\n    Foo = 1\nEnd Function\nSub Main()\n    x = 1\nEnd Sub\n"),
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
        }];

        let mut vm = Vm::new();
        let names = vm.populate_from_sheets(sheets);

        assert_eq!(names, vec!["input".to_string()]);
        assert_eq!(vm.active_sheet, "input");
        assert_eq!(vm.get_cell(1, 1), Variant::Integer(42));
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
}
