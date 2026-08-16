# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

### Added

- **`@elixcee/xlsx` compatibility groundwork (Phase 0)**: investigation and scaffolding for a planned npm package that would be a drop-in replacement for `xlsx@0.18.5` (SheetJS) —
  - `docs/xlsx-compatibility-goal.md`, `docs/xlsx-architecture.md` (ADR: target crate/npm workspace shape, and concrete resolutions for the `formula`↔`vm` circular type dependency and `reader.rs`'s path-only I/O — neither executed yet), `docs/xlsx-security-model.md` (resource-limit design, prototype-pollution-safe key handling), and `docs/licensing.md` (elixcee is MIT; the `xlsx` package and its 7 transitive SheetJS dependencies are all Apache-2.0)
  - New `compat/` Node.js project (not part of the Rust build): `compat/oracle/generate-manifest.mjs` installs and introspects the real `xlsx@0.18.5` at runtime (both its CJS and ESM entrypoints) rather than guessing from documentation, producing the committed `compat/oracle/api-manifest.json`; `compat/differential/classify.mjs` defines the six-value compatibility verdict (`MATCH`/`INTENTIONAL_SECURITY_DIVERGENCE`/`UNSUPPORTED`/`BUG`/`ORACLE_AMBIGUITY`/`NONDETERMINISTIC`) future differential tests will use, with a `run-demo.mjs` proving the plumbing
  - No `elixcee` Rust source, Python binding, CLI, or test behavior changed by this milestone
- **`@elixcee/xlsx` Phase 1A-1C**: `packages/xlsx` now implements every one of the real oracle's 33 `utils.*` runtime exports — `Object.keys(XLSX.utils)` matches the oracle exactly, both content and insertion order — differential-tested against the real oracle throughout (550+ permanent public-API fixtures across `compat/differential/`, plus a separate 1831-case internal SSF-backend conformance suite) —
  - Address/workbook utilities (`encode_*`/`decode_*`, `book_new`, `book_append_sheet`, `book_set_sheet_visibility`), worksheet mutation (`aoa_to_sheet`/`sheet_add_aoa`, `json_to_sheet`/`sheet_add_json`, `sheet_add_dom`), cell lookup (`sheet_get_cell`), JSON extraction (`sheet_to_json`/`sheet_to_row_object_array`), HTML export (`sheet_to_html`), DOM table conversion (`table_to_sheet`, `table_to_book` — duck-typed against whatever DOM-like object is passed, no DOM library imported at runtime), formula extraction (`sheet_to_formulae`), cell metadata (`cell_set_hyperlink`, `cell_set_internal_link`, `cell_add_comment`, `sheet_set_array_formula`), sheet-visibility constants (`consts`), and text export (`sheet_to_csv`, `sheet_to_txt`)
  - `format_cell`/`cell_set_number_format` are backed by the real `ssf@0.11.2` engine (Phase 1B-2B) — `packages/xlsx`'s only runtime dependency, isolated behind a single adapter file (see `docs/xlsx-architecture.md`'s "SSF backend" decision and `THIRD_PARTY_NOTICES.md`); one genuine upstream indirection-table bug (numFmtIds 67-71) was found and corrected in that adapter, including a follow-up fix so the correction never shadows a caller's own `opts.table` override
  - Four DoS-shaped divergences from the real oracle, all empirically confirmed (not assumed) and registered as intentional safety divergences: `encode_col(Infinity)` (the oracle hangs; elixcee rejects non-finite indices) and a crafted full-grid `!ref` fed to `sheet_to_formulae`/`sheet_to_csv`/`sheet_to_txt`/`sheet_to_json`/`sheet_to_html` (the oracle takes 12s+ / doesn't return within 25s; elixcee caps iteration at 5,000,000 cells — sizing measured at 100K/1M/5M/10M cells, see `docs/limits.md`)
  - Six security fixes, all live-confirmed defects in the oracle itself, not hypothetical: two prototype-corruption fixes (`book_append_sheet`, Phase 1A, and `table_to_book`'s internal sheet-to-workbook construction, Phase 1C — both let a caller-controlled sheet name of `"__proto__"` reassign a `WorkBook.Sheets` object's own prototype), two `sheet_to_json` fixes (an explicit `opts.header` array containing `"__proto__"` could silently drop a primitive column value or reassign a row object's own prototype to a Date/object cell value, Phase 1B-3), and two `sheet_to_html` fixes (Phase 1C — `data-t`/`data-v`/`data-z`/`id` attributes built with zero escaping let a cell value or `opts.id` containing `"` inject a live event handler; `cell.l.Target` embedded into `href="..."` with no scheme check let a `javascript:` hyperlink execute code on click). A third `sheet_to_html` finding — `cell.h`'s raw-HTML passthrough — is reproduced, not fixed, since it is a documented, intentional field (see `docs/compatibility-known-defects.md`)
  - TypeScript types classified as EXACT/SAFE_EXTENSION/MISSING/INCOMPATIBLE against the real oracle's own `types/index.d.ts` (`docs/typescript-compatibility.md`) rather than described loosely; `table_to_sheet`/`table_to_book`/`sheet_add_dom` mirror the oracle's own `data: any` (not `HTMLTableElement`) and are compile-tested both with and without DOM lib present (`tsconfig.no-dom.json`, `test/smoke-dom.ts`)
  - `packages/xlsx`'s own `LICENSE`/`THIRD_PARTY_NOTICES.md` are now included in the npm tarball's own `files` (previously only the repo root had them, which never reached npm consumers); package version reset from a stale `0.1.0-phase1b1` to `0.0.0-development` pending a real publish candidate
  - Still explicitly out of scope: XLSX/ODS file reading or writing (`read`/`readFile`/`write*`), any Rust↔JS bridge, and npm publish
- **Hidden row/column evidence** (Milestone B7b): `diagnose`/`diagnose-workbook` now report when a `.Copy`'d range overlaps hidden rows/columns —
  - New `vm::Interval`/`vm::SheetVisibility` types, threaded from XLSX's `<row hidden="1">`/`<col min=".." max=".." hidden="1">` into `Vm.sheet_visibility` the same way `merged_ranges` already is; ODS is explicitly deferred (its reader doesn't expand `table:number-rows-repeated`, so a hidden-row flag can't map to a correct absolute row number yet)
  - New `Vm::hidden_cells_observation()` computes the evidence on demand from the existing `Vm.clipboard` + `sheet_visibility` — no new stored side channel
  - A new sibling JSON field `observations` (not folded into `root_causes`, which means "why it failed" — this isn't a failure), present only when non-empty, on both success and failure: `{"code":"RANGE_CONTAINS_HIDDEN_CELLS","certainty":"observed","range":{...},"visibility":{...},"message":"..."}`
  - `diagnose-workbook` gets the same field via `FixtureResult::Passed`/`::Failed`, though honestly no additional value over a single `diagnose` call — hidden-row/column metadata is structural (workbook layout + macro text), not input-dependent
  - Copy/Paste behavior itself is unchanged — hidden cells still copy/paste exactly as before; this is observability only, laying groundwork for `SpecialCells(xlCellTypeVisible)` (B7c)
- **Multi-area Range foundation** (Milestone B7a): `Range("A1:A3,C1:C3")`-style disjoint ranges now have an underlying model —
  - New `vm::Rect`/`vm::RangeRef` types (`{ sheet, areas: Vec<Rect> }`); the existing single-rect `parse_range_addr`/`SheetRange` and their ~11 call sites are untouched — only Copy/Paste resolve through the new `parse_multi_area_addr`
  - `.Copy` now accepts a comma-separated multi-area source; `.Paste`/`.PasteSpecial` is diagnose-only for any multi-area shape and never completes (writes cells) in v1, even a fully-matching one — 4 new classified failures instead: `MULTI_AREA_TO_SINGLE_AREA_PASTE`, `MULTI_AREA_COUNT_MISMATCH`, `MULTI_AREA_SHAPE_MISMATCH`, and the catch-all `MULTI_AREA_PASTE_UNSUPPORTED`
  - Each area's evidence is `{"address", "rows", "columns"}`, matching the completion-condition JSON's own shape
  - `Union()`, the `Areas`/`Areas.Count`/`Areas(n)` property, and `Dim rng As Range`/`Set rng = ...` object variables remain unsupported — `Variant` gained no Range variant
  - Foundation for B7b (hidden/filtered rows) and B7c (`SpecialCells(xlCellTypeVisible)`), sequenced ahead of shrinking (B5b) since most structural failures need this range model first
- **`diagnose-workbook` subcommand** (Milestone B6d): combines `test-workbook`'s (B5a) generated-case search with `diagnose`'s (B6a–B6c2) root-cause classification —
  - Reuses `test-workbook`'s exact fixture format, strategies, and deterministic `--seed`/`--case` replay; runs each case with `Vm::strict_resolution` on and enriches classifiable failures with the same `ResolutionFailureKind` → `RootCause` pipeline `diagnose` uses, via a new `pub(crate) diagnose::root_causes_json` entry point
  - New `--cases N` flag overrides the fixture's declared case count for one invocation (scoped to this subcommand; `test-workbook` itself is unchanged)
  - Output is `test-workbook`'s existing JSON shape plus one sibling `root_causes` field (`[]` when unclassified)
  - Most root causes are structural (merge/shape/protection kinds) and fire identically regardless of input — this command's actual value is for input-dependent kinds like `ARRAY_INDEX_OUT_OF_BOUNDS`, where a drawn value can flip an index in or out of bounds across cases
  - Shrinking (minimizing a failing case's inputs) is deliberately deferred to a later phase
- **Merged-cell-aware Paste diagnostics** (Milestone B6c2): `diagnose` now classifies Copy/Paste operations that conflict with a merged-cell layout —
  - `PASTE_INTO_NON_ANCHOR_MERGED_CELL`: the destination cell falls inside an existing merge but isn't that merge's own top-left cell (pasting into the top-left cell itself, the normal way to write to a merged cell, is unaffected)
  - `PASTE_PARTIAL_MERGED_RANGE`: a multi-cell destination partially overlaps one or more merges without fully containing them
  - `PASTE_MERGE_LAYOUT_MISMATCH`: the source's and destination's merged-cell layouts, compared by relative position (accounting for `Transpose:=True`), don't match
  - `WorkbookSheet.merged_ranges` parsed from XLSX `<mergeCell ref="...">` and ODS `table:number-columns-spanned`/`table:number-rows-spanned`, threaded into a new `Vm.merged_ranges` map
  - Unconditional hard errors in every mode that executes the macro (`run`/`diagnose`/`test-workbook`), matching real Excel's Error 1004 regardless of `On Error` state — same posture as B6b/B6c
  - Scope stays Paste-only; multi-area (`Areas`) ranges, hidden/filtered rows, and AutoFilter visible-cells-only copy remain deferred

## [0.1.1]

### Added

- **CLI binary** (`src/main.rs`): standalone `elixcee` executable — no Python required
  - Usage: `elixcee <vba_file> <MacroName> [--file xlsx] [--sheet name] [--output xlsx]`
  - `MsgBox` output printed to stdout; result cells printed as `A1\t<value>` per line
  - Pre-built binaries for Windows x64, Linux x64, macOS Apple Silicon on GitHub Releases
- **GitHub Actions release workflow** (`.github/workflows/release.yml`): builds CLI binaries on `bin-v*` tag push; attaches them to a GitHub Release via `softprops/action-gh-release`
- **`pub fn save_workbook`**: public Rust API for writing `.xlsx` / `.ods` from non-Python callers
- **`Vm::print_msgbox`** field: when `true`, `MsgBox` writes to stdout instead of being silently dropped
- **pyo3 optional feature**: `pyo3` is now an optional dependency behind the `python` feature; `cargo build --bin elixcee` compiles a Python-free binary; `maturin build` continues to use `features = ["python"]`
- **Math & Combinatorics**: `FACT`, `PERMUT`, `GCD`, `LCM`, `QUOTIENT`, `SIGN`
- **Statistical**: `CORREL`, `COVARIANCE.S`, `COVARIANCE.P`, `NORM.DIST`, `NORM.INV`, `T.DIST` — uses Stirling lgamma + Lentz incomplete-beta CF
- **Financial functions**: `FV`, `PV`, `NPER`, `RATE` (Newton-Raphson), `IPMT`, `PPMT`, `NPV`, `IRR`, `MIRR`, `XNPV`, `XIRR` — all share the `annuity_fv` / `compute_pmt` helpers
- **Database functions**: `DSUM`, `DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DMAX`, `DMIN` — all take `(database, field, criteria)` and reuse the existing `db_row_matches_criteria` / `resolve_db_field` infrastructure from `DGET`
- **GitHub Actions CI/CD**: `.github/workflows/publish.yml` — builds wheels for Linux x86_64/aarch64, Windows x86_64, macOS universal2, and an sdist; publishes to PyPI via OIDC Trusted Publisher on `v*` tag push
- **README_zh.md**: Simplified Chinese translation of README

### Added — JSON Agent Contract & Static Analysis (Milestones A, A.1, A.5, B1, B1.1, B2, B3, B4)

- **`--json` output** (`src/diagnostics.rs`): single machine-readable JSON object (result or error) instead of plain text — error classification (`ElixceeError`), a hand-rolled JSON writer/escaper (no serde in the release binary), and `Vm::msgbox_log` (`MsgBox` calls recorded into `messages` instead of printed directly, drained via `take_messages()` so a reused `Vm` never leaks a prior run's messages)
- **Source location tracking** (`SourceSpan`/`SpannedStmt`, char-offset based): parse and runtime errors report `{file, line, column}` in `--json` mode; non-JSON output is unchanged
- **`check` subcommand** (`src/check.rs`): static analysis without executing the macro — parse diagnostics, entrypoint existence, undefined Sub/Function call detection anywhere in the body (probes the real builtin-function dispatch table directly, so there's no allowlist to drift), and unsupported-construct/no-op detection (`I1002`), all with source locations
- **Multi-module projects**: pass more than one `.bas`/`.vbs` file to run a project spanning several modules; `Module.Sub`-qualified entrypoints (module name from `Attribute VB_Name`, else the filename); cross-module Sub/Function name collisions are rejected at load time
- **Deterministic black-box tests** (`tests/blackbox.rs`): declarative `.toml` fixtures (VBA source + CLI args + expected JSON) diffed byte-for-byte against the real binary's `--json` output; adding a new regression case needs no Rust
- **`snapshot` subcommand** (`src/snapshot.rs`): reads a `.xlsx`/`.xlsm`/`.ods` file directly (no VBA execution) and prints every sheet's non-empty cells as Markdown or JSON, with a `sheet_id`/`stable_id` pair for cross-sheet identity (not to be confused with VBA's real `CodeName`)

### Added — Property-Based Testing & Excel Operation Diagnostics (Milestones B5a, B6a, B6b, B6c)

- **`test-workbook` subcommand** (`src/testworkbook.rs`): reruns a macro against a starting workbook many times with generated boundary-value inputs (`boundary_numeric`/`boundary_string`), checking each independent case for panics, runtime errors, timeouts, and Excel error values; failures report `seed`/`case_index` for exact replay via `--seed`/`--case`
- **`diagnose` subcommand** (`src/diagnose.rs`): runs a macro once and classifies *why* Excel would reject an operation, with evidence, instead of a bare error string —
  - `WORKSHEET_NOT_FOUND` / `WORKBOOK_NOT_FOUND` / `ARRAY_INDEX_OUT_OF_BOUNDS`, with a hand-rolled Levenshtein "did you mean" suggestion (opt-in `Vm::strict_resolution` turns off the usual auto-vivify-on-write/silent-`Empty`-on-read behavior only for this command)
  - `Sheets(name).Range(addr)`, `Worksheets(idx)` numeric index, and a minimal `Workbooks(name).Worksheets(...)` all newly parseable, needed to even express the sheet-resolution scenarios this command diagnoses
  - `PASTE_SHAPE_MISMATCH` / `PASTE_WITHOUT_COPY`: a VM clipboard (`Vm.clipboard`) populated by `.Copy`/`.Copy Destination:=` and consumed by `.Paste`/`.PasteSpecial [Transpose:=]`/`Worksheets(sheet).Paste`, with both the Copy and Paste statement locations and a mechanically-derived resize suggestion
  - `SHEET_PROTECTED`: `Sheets(name).Protect`/`.Unprotect` (including `UserInterfaceOnly:=True`, which blocks manual edits but not macro writes, matching real Excel) blocks any cell-content mutation on that sheet — writes, clears, inserts, sorts, paste, delete — unconditionally in every mode, while reads are never blocked
  - Shape mismatches, empty-clipboard pastes, and writes to a protected sheet are unconditional hard errors in every mode that executes the macro (`run`/`diagnose`/`test-workbook`), matching real Excel's Error 1004/protection behavior regardless of `On Error` state

### Changed

- `pyproject.toml`: `features = ["pyo3/extension-module"]` → `features = ["python"]` to align with the new optional-feature approach
- **`diagnose`'s entrypoint is now a positional argument** (`elixcee diagnose <vba_file>... <MacroName> --file <path> [--json]`) instead of `--entrypoint <MacroName>` — matches `run` mode's own convention (entrypoint is always mandatory for both, unlike `check`, where it's optional and therefore needs an explicit flag to stay unambiguous). Breaking change; `--entrypoint` is removed, not kept as an alias.
- **PyPI package metadata**: `pyproject.toml` now declares a description, `readme`, `license`, `keywords`, `classifiers`, and `[project.urls]` (Homepage/Documentation/Repository/Issues/Changelog) — the published package previously had none of these

### Removed

- `FUNCTIONS_ja.md`: duplicate of `FUNCTIONS.md`; `README_ja.md` now links to the English reference

### Performance (Round 4)

- **`SUM` fast path**: single-range `SUM` iterates cell refs directly — no `Vec<Variant>` allocation
- **`range_nums_fast!` macro**: `AVERAGE`, `MIN`, `MAX` on a single range skip `Vec<Variant>` and collect `f64` directly
- **`RangeWrite` / `RangeClear` dirty-flag batching**: writes go directly to the sheet map; `cell_index_dirty` set once after the loop instead of once per cell

### Tests

503 unit tests (↑ from 329) + `tests/cli_json.rs` (14) + `tests/cli_check.rs` (15) + `tests/blackbox.rs` (1 test scanning 12 `.toml` fixtures) + `tests/cli_snapshot.rs` (7) + `tests/cli_test_workbook.rs` (7) + `tests/cli_diagnose.rs` (12) + `tests/prop_tests.rs` (17)

---

## [0.1.0] — Initial Release

### Added — VBA Parser & Interpreter

- **Sub / End Sub** with parameter passing
- **Variable assignment** and arithmetic expressions
- **Cell read/write** via `Cells(row, col).Value` and `Range("A1").Value`
- **For / Next** loops with optional `Step`
- **For Each** iteration over cell ranges
- **If / ElseIf / Else / End If** conditional branches
- **Do While / Loop** and **While / Wend** loops
- **Select Case** with value, range (`To`), and comparison (`Is`) patterns
- **Exit For**, **Exit Do**, **Exit Sub**, **Exit Function**
- **Function / End Function** with return values; **Call** statement
- **On Error Resume Next**, **On Error GoTo label**, **Resume**, **GoTo**
- **With / End With** blocks (plain and `With Sheets("name")`)
- **Const** declarations; `Option Explicit` / `Option Base` ignored
- **Dim** variable declarations; `Dim arr(n)` and `ReDim [Preserve]` arrays
- **Type ... End Type** user-defined types with typed field initialization
- **Public / Private / Friend / Static** modifiers on Sub/Function (modifier ignored)
- **Debug.Print** / **Debug.Assert** as no-ops
- **MsgBox** — configurable skip or RuntimeError
- Range operations: `ClearContents`, `Clear`, `Copy`, `Delete`, `Insert`, `Sort`, `Offset.Value`
- Sheet operations: `Sheets.Add`, `Sheets.Delete`, `Sheets("name").Cells`
- `Application.Calculation` (Manual / Automatic); `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `StatusBar`, `Cursor`, `CutCopyMode` as no-ops
- `WorksheetFunction.*` prefix forwarding to formula engine
- `Cells(Rows.Count, col).End(xlUp).Row` and related `.End(dir).Row/Column` — indexed with `BTreeSet` for O(log n) performance

### Added — Named Ranges

- `Range("A1:B5").Name = "MyName"` registers a workbook-level named range
- All Range operations (Read/Write/Clear/Delete/Insert/Sort/Copy/ForEach) transparently resolve named range strings

### Added — Formula Engine (200+ functions)

#### Arithmetic & Statistical
`SUM`, `AVERAGE`, `AVERAGEIF`, `AVERAGEIFS`, `MIN`, `MAX`, `MINIFS`, `MAXIFS`,
`COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS`, `COUNTBLANK`,
`SUMIF`, `SUMIFS`, `SUMPRODUCT`, `PRODUCT`, `MEDIAN`, `MODE.MULT`,
`LARGE`, `SMALL`, `RANK`, `PERCENTILE` / `PERCENTILE.INC`, `PERCENTRANK` / `PERCENTRANK.INC`,
`ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `TRUNC`, `MOD`,
`RAND`, `RANDBETWEEN`, `SUBTOTAL`, `AGGREGATE`

#### Statistical
`STDEV` / `STDEV.S`, `STDEVP` / `STDEV.P`, `VAR` / `VAR.S`, `VARP` / `VAR.P`

#### Math & Trigonometry
`ABS`, `SQRT`, `POWER`, `EXP`, `LN`, `LOG`, `LOG10`, `PI`,
`SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `DEGREES`, `RADIANS`,
`FLOOR` / `FLOOR.MATH`, `CEILING` / `CEILING.MATH`, `MROUND`

#### Logical
`IF`, `IFS`, `SWITCH`, `AND`, `OR`, `NOT`, `XOR`, `IFERROR`

#### Text
`LEFT`, `RIGHT`, `MID`, `LEFTB`, `RIGHTB`, `MIDB`,
`LEN`, `LENB`, `UPPER`, `LOWER`, `PROPER`, `TRIM`,
`FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`,
`CONCATENATE`, `CONCAT`, `TEXTJOIN`, `TEXT`, `VALUE`, `EXACT`,
`CHAR`, `UNICHAR`, `CODE`, `UNICODE`, `ASC`, `JIS`

#### Date & Time
`DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `DAYS`,
`EDATE`, `EOMONTH`, `DATEDIF`, `DATEVALUE`,
`TIME`, `TIMEVALUE`, `HOUR`, `MINUTE`, `SECOND`,
`NETWORKDAYS`, `NETWORKDAYS.INTL`, `WORKDAY.INTL`

#### Lookup & Reference
`VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `LOOKUP`,
`INDEX`, `MATCH`, `XMATCH`, `CHOOSE`,
`ROW`, `COLUMN`, `INDIRECT`, `OFFSET`, `ADDRESS`

#### Information
`ISBLANK`, `ISERROR`, `ISERR`, `ISNA`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, `ISNONTEXT`

#### Array / Spill
`FILTER`, `UNIQUE`, `SORT`, `SORTBY`, `SEQUENCE`, `TRANSPOSE`,
`TOCOL`, `TOROW`, `WRAPCOLS`, `WRAPROWS`, `RANDARRAY`

#### Lambda & Higher-Order
`LET`, `LAMBDA`, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`

### Added — Formula Engine Features

- **Topological sort** for formula recalculation (`topo_sort_formulas`): formulas are evaluated in dependency order; circular references fall back to best-effort ordering
- **Application.Calculation** mode — Manual suppresses recalc; switching to Automatic triggers full recalc
- A1-notation and R1C1-notation cell references; range references (`A1:B10`)
- DBCS byte semantics (`LENB`, `LEFTB`, `RIGHTB`, `MIDB`) matching Excel's 2-byte-per-CJK rule
- Excel 1900 leap-year bug compatibility in date serial arithmetic

### Added — Python API

| Method | Description |
|---|---|
| `Vm(on_msgbox=)` | Create a VM; `"skip"` or `"error"` on MsgBox |
| `vm.run(vba, name)` | Execute a Sub |
| `vm.set_cell(r, c, v)` / `get_cell(r, c)` | 1-based cell read/write |
| `vm.cells()` | All non-empty cells as `{(r, c): value}` |
| `vm.cells_df()` | Active sheet as pandas DataFrame (requires pandas) |
| `vm.variables()` | VBA variables as `{name: value}` |
| `vm.set_cell_formula(r, c, f)` | Set and evaluate a formula string |
| `vm.set_cell_formula_batch(d)` | Batch formula set: `{(r,c): formula}` |
| `vm.recalculate()` | Re-evaluate all formula cells |
| `vm.set_sheet(name)` / `active_sheet()` / `sheet_names()` | Sheet management |
| `vm.get_sheet(name)` | Cells of a named sheet |
| `vm.save_workbook(path)` | Save to `.xlsx` or `.ods` |
| `vm.named_ranges` | Dict of registered named ranges |
| `elixcee.run_macro(vba, name)` | One-shot macro runner |
| `elixcee.load_workbook(path)` | Load `.xlsx` / `.ods` into a `Vm` |

- `Variant::Date` → Python `datetime.date` conversion
- `Variant::Error` → Python `ExcelError` class with `.code` attribute (bidirectional)
- Type stubs `elixcee.pyi` for IDE completion

### Added — File I/O

- **Read**: `.xlsx`, `.xlsm`, `.ods` — hand-written XML parser (no calamine at runtime)
- **Write**: `.xlsx` — hand-written XML + zip; `.ods` — hand-written XML + zip
- Multi-sheet support: all sheets loaded on `load_workbook`; saved on `save_workbook`

### Performance

- `Cells.End` searches (`xlUp`, `xlDown`, `xlToLeft`, `xlToRight`) use a lazy `BTreeSet` index — O(log n) per query after O(n) rebuild on cell mutation
- Zero-copy formula parse caching via `recalculate_all` with topological ordering

### Dependencies (runtime)

| Library | Purpose |
|---|---|
| `pyo3` | Python bindings |
| `zip` | XLSX / ODS archive read-write |

`calamine` is kept as a `[dev-dependencies]` oracle for diff-testing the hand-written reader.

### Tests

299 unit tests covering parser, formula engine, VM interpreter, file round-trips, and diff tests against calamine.
