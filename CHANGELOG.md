# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [0.15.0] - 2026-08-30

Safe Style Engine for controlled workbook formatting edits. Added `set_number_format`,
`set_style` for font/fill/border/alignment/protection, `copy_style`, named-style
application, and row/column default styles. Style records are deduplicated and existing
shared styles are never mutated in place; pending edits are chained through one safe
resolution pipeline so combined formatting changes are preserved.

Theme-color minting and effective theme-color resolution remain out of scope.

## [0.14.0] - 2026-08-30

Dependency-aware structural editing for formulas and cell metadata. Row/column
insert-delete now rewrites same-sheet and sheet-qualified formula references, and
sheet rename updates qualified formulas and defined names. Added safe same-sheet
`move_range` with formula translation and atomic validation. Structural edits and
range moves now preserve or transform merged ranges, hidden row/column intervals,
cell styles, number formats, row heights, and column widths. XLSX save/reload also
preserves formula cells without cached values and worksheet AutoFilter metadata.

The release deliberately excludes structured table, data-validation, chart, and
comment editing; those remain later milestones.

## [0.12.0] - 2026-08-27

Root `elixcee` (Rust crate + Python package) only -- `elixcee-types`/`elixcee-wasm`/
`@elixcee/xlsx` all unaffected (see `[Unreleased]` below for `@elixcee/xlsx`'s own
independent, still-unpublished work). Eight independent Python API additions against
`docs/openpyxl-gap-audit.md`'s priority list -- R1 (bulk worksheet range/row API), P1
core 3 (sheet rename/move, row/col insert-delete glue, read-only merged-cell access), P1
remainder (`iter_cols`, `sort_range`, merge create/remove), P2's first slice (hidden
row/col read/write), P2's second slice (`copy_sheet`), P2's third slice
(`defined_names`, read-only), P2's fourth slice (`sheet_state`, read-only), and P2's
fifth slice (`row_height`/`column_width`, read-only) -- all below -- unrelated to each
other. A `FIND()` crash fix, found by the P2 fifth slice round's own fuzz CI job
(unrelated to what that round actually changed), is also below. Released from a
`release-0.10.0`-branch base (`v0.11.0`) with these eight rounds cherry-picked on top,
deliberately excluding this repo's own still-unreleased `0.10.0-D`/`t="e"` work below
(same reasoning `0.11.0` used before it -- see the "Packaging note" in ROADMAP.md).

### Root crate (Python binding): R1 -- bulk worksheet range/row API

Seven new Python methods close the highest-value gap identified against openpyxl (see the
new `docs/openpyxl-gap-audit.md`, which scores openpyxl's full API surface against what
`elixcee` exposes today and records P1/P2/P3 follow-up candidates): `get_range(addr,
sheet=None)`/`set_range(addr, values, sheet=None)` (rectangular read/write, 1-based A1
notation), `append_row(values, sheet=None)` (writes past the sheet's true max used row --
correct on a sparse sheet, not a populated-row count), `iter_rows(min_row=1, max_row=None,
min_col=1, max_col=None, sheet=None)` (values-only, defaults to the used range, `[]` on a
totally empty sheet unless `max_row` is given explicitly), and `max_row`/`max_column`/
`calculate_dimension` (all `None`, never `0`/`"A1:A1"`, on a sheet with zero non-empty
cells). Every method takes `sheet` as a keyword; `None` means the active sheet, and an
explicit sheet name never changes which sheet is active.

New `Vm`-core primitives (`src/vm/mod.rs`, PyO3-agnostic, unit-tested via plain `cargo
test`): `resolve_sheet_key`, `sheet_used_range` (Empty-exclusion bounding box, matching
`cells()`/`get_sheet()`'s convention, not `cells_df`'s divergent one), `next_append_row`,
`read_rect`/`write_rect`, `iter_rows_values`. `set_range`/`append_row` convert and
shape-validate their entire input into a scratch buffer before writing anything, so a
validation failure can't partially apply. A literal `"="`-prefixed string is stored as-is,
never promoted to a formula (`set_cell_formula`/`set_cell_formula_batch` remain the only
way to set one). Address parsing reuses `elixcee-types::parse_range_addr` as-is; a new,
ungated `validate_range_addr` wrapper in `src/lib.rs` adds `$`-stripping, multi-area
rejection, and reversed/zero-row-col rejection as explicit `ValueError`s, closing three
disclosed shared-parser gaps only for calls made through this API -- the shared parser
itself is untouched, recorded in the gap-audit doc for whoever picks it up next.

Writing into a non-anchor cell of a merged range, or into a protected sheet, is
deliberately **not** blocked by `set_range`/`append_row` -- matches `PyVm::set_cell`'s
existing (equally unchecked) behavior; introducing a stricter Python-only rule with no
VBA-side precedent was rejected as inconsistent with the rest of the binding. See the
gap-audit doc's "Implementation notes for R1" for this and two other disclosed,
out-of-scope gaps (a `cells_df` used-range inconsistency, no upper-bound guard against a
pathological full-column/full-row address).

New `compat/differential-python/` harness (stdlib `unittest`, `openpyxl` as a test-only
oracle -- `pyproject.toml` still declares no runtime/test Python dependencies) compares
`get_range`/`iter_rows`/`append_row` against openpyxl's own read of a real fixture, and
pins one real, expected divergence rather than silently matching it: a merged range's
non-anchor cells are excluded from `calculate_dimension`'s bounding box (no value of their
own), while openpyxl's `dimensions` mirrors the real XLSX `<dimension>` element, which
Excel widens to the merge's full span regardless.

### Root crate (Python binding): P1 core 3 -- sheet rename/move, row/col insert-delete glue, merged-cell read

The next slice of `docs/openpyxl-gap-audit.md`'s priority list after R1. Seven new Python
methods: `rename_sheet(old_name, new_name)`, `move_sheet(name, new_index)` (absolute
0-based position, matching `set_sheet`'s own convention -- not openpyxl's relative-offset
`Worksheet.move_sheet(offset)`), `insert_rows(idx, amount=1, sheet=None)`/`delete_rows`/
`insert_cols`/`delete_cols` (Python glue over the existing `0.11.0` VBA-only handlers,
Excel-grid bounds checked: 1,048,576 rows / 16,384 columns), and `merged_cells(sheet=None)
-> list[str]` (read-only).

`rename_sheet` atomically re-keys all 8 lowercase-keyed per-sheet `Vm` maps (`sheets`,
`sheet_order`, `active_sheet`, `merged_ranges`, `sheet_visibility`, `cell_style_indices`,
`cell_number_formats`, `worksheet_origins`) -- more bookkeeping than initially scoped (see
the gap-audit doc's "Implementation notes for P1 core 3"), since nothing renamed a sheet
before this round. Renaming the active sheet is supported (it stays active under the new
name), not a silent no-op; renaming a protected sheet is rejected outright. `move_sheet`
is the missing complement to `ensure_sheet_at` (which only positions a newly-created
sheet) -- reordering an *existing* sheet had no primitive at all before this round.

Because `<definedName localSheetId="N">` is positional, `move_sheet` reordering
`sheet_order` could otherwise leave a saved workbook's `<definedNames>` pointing at the
wrong sheet; separately, a `<definedName>`'s own TEXT can reference a sheet by name (e.g.
`Sheet1!$F$5`), which `rename_sheet` doesn't rewrite, so a rename could leave that text
dangling too. Both are fixed by extending the existing deletion-only passthrough guard
(`src/lib.rs`) with a new `Vm::defined_names_may_be_stale` flag, set by both `move_sheet`
and `rename_sheet`, checked alongside the existing check. The `rename_sheet` half was
missed in the first pass -- caught in a follow-up review against `fixture4` (the one real
fixture with genuine `<definedNames>` content; the original tests only used fixtures
without any) -- and fixed with its own integration test verifying `<definedNames>` is
actually absent from a real saved output after a rename, not just that an internal flag
got set.

`insert_rows`/`delete_rows`/`insert_cols`/`delete_cols` are built on new
`insert_rows_on_sheet`/`delete_rows_on_sheet`/`insert_cols_on_sheet`/`delete_cols_on_sheet`
`Vm`-core siblings of the existing active-sheet-only private methods, which become
one-line delegators -- none of their 8 existing VBA call sites needed touching. The delete
siblings use a single-pass `retain` rather than the original's two-phase
remove-then-band-delete-then-reinsert; an earlier draft of that collapse used the wrong
predicate (kept stale data at the pre-shift position while also inserting a shifted copy)
and was caught and fixed in review, with a regression test pinning the correct one. Does
**not** shift merged ranges, hidden-row/col markers, cell styles/number formats, or
formula references -- a pre-existing VBA-engine limitation, now Python-reachable,
disclosed rather than silently inherited.

`merged_cells` reuses a newly-factored-out `merge_rect_to_a1` helper, now shared with
`save_xlsx_impl`'s own `<mergeCell ref="...">` writer (previously an inline `format!`,
duplicated once this method needed the same conversion).

New `compat/differential-python/sheet_ops_check.py`, same stdlib-`unittest`-plus-openpyxl
structure as `bulk_range_check.py`, compares `rename_sheet`/`move_sheet`/`merged_cells`
against openpyxl's own read of the same real fixture after a save/reload round trip.
Row/col insert-delete gets no differential coverage -- the disclosed fidelity gap means an
openpyxl comparison would correctly fail on exactly the cases worth testing.

See `docs/openpyxl-gap-audit.md`'s "Implementation notes for P1 core 3" for the full
account of this round's disclosed gaps, including two new ones surfaced while
implementing rename (`remove_sheet`'s own pre-existing 6-map leak on delete, left
unfixed; the residual `Sheets.Add(before:=...)` `<definedNames>` gap the `move_sheet` fix
doesn't close).

### Root crate (Python binding): P1 remainder -- iter_cols, sort_range, merge create/remove

The last three items `docs/openpyxl-gap-audit.md` still tagged `P1`. Four new Python
methods: `iter_cols(min_row=1, max_row=None, min_col=1, max_col=None, sheet=None)`
(column-major values-only iteration, the transposed sibling of `iter_rows`),
`sort_range(addr, key_col, descending=False, header=False, sheet=None)` (not from
openpyxl, which has no sort primitive of its own -- exposes the existing VBA
`Range(addr).Sort` statement's exact behavior to Python), and `merge_cells(addr,
sheet=None)`/`unmerge_cells(addr, sheet=None)` (create/remove a merge).

`iter_cols` is `Vm::iter_cols_values`, built on the same `read_rect`/`sheet_used_range`
primitives as `iter_rows_values`, short-circuiting on `max_col`'s explicitness instead of
`max_row`'s.

`sort_range` required extracting `Stmt::RangeSort`'s previously fully-inlined,
active-sheet-only sort algorithm into a new sheet-parameterized `Vm::sort_range_on_sheet`
(built on `read_rect`/`write_rect`) -- the VBA dispatch arm shrinks to resolve-address +
protection-check + delegate, with all 4 pre-existing `test_range_sort_*` tests passing
unmodified. `PyVm::sort_range` validates `key_col` against the range's own column span
explicitly (`ValueError`) rather than inheriting the VBA path's silent `saturating_sub`
clamp on an out-of-range key column, and enforces the same 1,048,576-row/16,384-column
ceiling `insert_rows`/`delete_rows` already do -- unlike `get_range`/`iter_rows` (a
large-but-harmless allocation), an oversized address here writes into the saved file.
Deliberately does **not** check sheet protection, matching `set_range`'s existing bulk
cell-value-write precedent.

`merge_cells`/`unmerge_cells` needed zero writer changes -- `save_xlsx_impl` already emits
`<mergeCell>` mechanically from whatever's in `merged_ranges` with no validation of its
own, so the new API only manages that map. `merge_cells` rejects a single-cell address and
any merge that would overlap an existing one on the same sheet, reusing `rects_overlap`
(Milestone B6c2's Copy/Paste conflict-detection primitive) rather than the Copy/Paste-
specific `check_merge_conflicts`; the overlap check runs before the map is touched, so a
rejected merge never leaves a stray empty entry behind. `unmerge_cells` requires an exact
rect match, erroring rather than silently no-opping (matching `rename_sheet`/`move_sheet`/
`delete_sheet`'s existing convention). Same address-bounds ceiling as `sort_range`. Does
**not** touch cell values in the covered range either way.

`compat/differential-python/bulk_range_check.py` gained an `iter_cols` comparison against
openpyxl's own `ws.iter_cols()`; `sheet_ops_check.py` gained a `merge_cells`/
`unmerge_cells` round-trip comparison plus direct pins of the PyO3-layer bound checks
(which have no Rust unit test of their own, living in `#[cfg(feature = "python")]` glue
rather than `Vm`-core logic). `sort_range` gets no differential coverage -- openpyxl has
no sort primitive to compare against.

### Root crate (Python binding): P2 first slice -- hidden row/col read/write

The first item off `docs/openpyxl-gap-audit.md`'s P2 list. Four new Python methods:
`hidden_rows(sheet=None)`/`hidden_columns(sheet=None)` (sorted, flattened 1-based row/
column numbers -- expanded, not interval-form) and `set_row_hidden(row, hidden=True,
sheet=None)`/`set_column_hidden(col, hidden=True, sheet=None)` (hide or unhide a single
row/column).

Reading and hiding needed no new algorithmic work: `Vm.sheet_visibility`'s existing
`Interval`-run storage (Milestone B7b) and the writer's already-mechanical `<col
hidden="1">`/`<row hidden="1">` emission meant `hidden_rows_on_sheet`/
`hidden_columns_on_sheet` are a flatten-and-dedup over existing state, and the *hide* half
of `set_row_hidden_on_sheet`/`set_column_hidden_on_sheet` just pushes a new single-unit
interval (a no-op if the unit's already covered, so hiding twice never produces a stray
duplicate). *Unhiding* a single row/column needed genuinely new work: splitting whatever
interval currently covers it (dropped entirely, shrunk from one end, or split into two
flanking intervals if the unit sits strictly inside a wider range) via a new
`remove_unit_from_intervals` free function -- the existing `visible_runs` helper computes
visible gaps across a whole range and discards which specific hidden interval produced
each gap, so it wasn't reusable here. Unhiding an already-visible unit, or a unit on a
sheet with no `sheet_visibility` entry at all, is a no-op that does not create a stray
empty entry -- matches `merge_cells`' own validate-before-mutating convention. Both
setters enforce the same 1,048,576-row/16,384-column ceiling `insert_rows`/`sort_range`/
`merge_cells` already do.

`compat/differential-python/sheet_ops_check.py` gained a `HiddenRowsAndColumnsAgreeWithOpenpyxl`
class: agreement with openpyxl on the real fixture's pre-existing hidden row 5/column D,
a newly-hidden row/column round-tripped through a save/reload, and unhiding the fixture's
pre-existing hidden row.

### Root crate (Python binding): P2 second slice -- copy_sheet

The second item off `docs/openpyxl-gap-audit.md`'s P2 list. One new Python method:
`copy_sheet(source_name, new_name)`, duplicating a sheet's cells, merges, hidden-row/col
state, cell styles, and cell number formats into a brand-new sheet.

Reuses `rename_sheet`'s own per-sheet-map list directly (`sheets`/`merged_ranges`/
`sheet_visibility`/`cell_style_indices`/`cell_number_formats`/`worksheet_origins`) but
`get()`-then-`clone()`-then-`insert()` instead of `remove()`-then-`insert()` -- no new
algorithmic work, since `rename_sheet` (P1 core 3) had already discovered that exact list
the hard way. The copy gets a brand-new `WorksheetOrigin` with only
`original_display_name` set, mirroring `ensure_sheet`'s own shape for a sheet with no
loaded-file origin, so the existing from-scratch-sheet writer path handles it with zero
new logic.

Deliberately appends the copy at the end of the sheet order rather than positioning it
immediately after the source (unlike openpyxl's own `copy_worksheet`) -- inserting
anywhere before the end would shift every later sheet's positional index, the same risk
`move_sheet`'s `defined_names_may_be_stale` flag exists to guard against for a reorder; an
append avoids the risk entirely, so `copy_sheet` correctly leaves that flag untouched. Use
`move_sheet` afterward if exact placement next to the source matters. Does not copy sheet
protection status (the copy is always unprotected) and does not change the active sheet.

`compat/differential-python/sheet_ops_check.py` gained a `CopySheetAgreesWithOpenpyxl`
class. Discovered while writing it, unrelated to `copy_sheet` itself and pre-existing:
`Vm::sheet_names()` returns sheets alphabetically sorted, not in `sheet_order`/
tab-position order -- undocumented but real, not a regression, not fixed here (see
ROADMAP.md's known gaps).

### Root crate (Python binding): P2 third slice -- defined_names (read-only)

The third item off `docs/openpyxl-gap-audit.md`'s P2 list. One new Python method:
`defined_names() -> dict[str, str]`, reading every `<definedName
name="...">TEXT</definedName>` in the loaded workbook's `xl/workbook.xml` into `{name:
raw_text}` (e.g. `{"MyRange": "Sheet1!$A$1:$A$3"}`).

Confirmed before writing any code that `Vm.named_ranges` -- VBA's own runtime table,
populated only by `Range(addr).Name = "x"` -- is never populated from a loaded file, so
reading defined names needed a genuinely new parser: `reader::xlsx_defined_names`, modeled
directly on the existing `xlsx_shared_strings` streaming pattern, no new parsing
infrastructure required. Deliberately read-only (no create/delete this round) and
deliberately returns each name's raw formula text rather than a resolved sheet+address --
elixcee's formula engine has no cross-sheet reference syntax (`=Sheet2!A1`) to resolve
that text against, and real XLSX additionally allows a sheet-scoped (`localSheetId`) name
to shadow a workbook-scoped one of the same name, which a flat map can't represent either
(both flatten into one map, last-encountered-in-document-order silently wins on a
collision -- disclosed, not solved, no fixture exercises it).

Re-reads the source file's ZIP on every call (mirroring `save_xlsx_impl`'s own passthrough
re-read at save time) rather than caching at load time -- a pure reporting view of the
file's current `<definedNames>`. Returns `{}` if no workbook is loaded; raises
`ValueError` if a workbook WAS loaded but its source file is no longer readable (a
genuinely different failure mode from "nothing to report").

`compat/differential-python/sheet_ops_check.py` gained a `DefinedNamesAgreesWithOpenpyxl`
class, using `fixture4_hyperlink_comment_name.xlsm` (the one real fixture with genuine
`<definedNames>` content) to confirm exact agreement with openpyxl's own
`wb.defined_names` dict.

### Root crate (Python binding): P2 fourth slice -- sheet_state (read-only)

The fourth item off `docs/openpyxl-gap-audit.md`'s P2 list. One new Python method:
`sheet_state(name) -> str`, reading a sheet's whole-tab visibility as `"visible"`,
`"hidden"`, or `"veryHidden"` -- matching openpyxl's own `ws.sheet_state` string
vocabulary exactly, no translation needed. Name-addressed (case-insensitive) like
`rename_sheet`/`copy_sheet`, not "current sheet"-defaulted; raises `ValueError` on an
unknown sheet name rather than silently returning `"visible"`.

Confirmed a real, independent, pre-existing bug while researching this round: neither the
reader nor the writer has ever handled XLSX's `<sheet state="...">` attribute at all --
loading a real file with a hidden or veryHidden sheet and saving it, even a completely
no-op save, silently reverted every sheet to visible. Not introduced by this round; just
discovered by it, and now pinned by a differential-python test that asserts the current
broken behavior explicitly rather than leaving it as an unverified claim.

Deliberately read-only this round -- no `set_sheet_state` yet. Every real fixture under
`compat/oracle-excel-com/fixtures/` and `compat/corpus/` was checked; none has a hidden or
veryHidden sheet, and this project's hard gate is no writer code for a structural OOXML
element without real fixture evidence. `copy_sheet` was extended to also copy the
source's visibility state onto the new sheet (its ninth per-sheet map to re-key on
`rename_sheet`, eighth to copy on `copy_sheet`), matching every other field it already
copies.

New `SheetState` enum (`Visible`/`Hidden`/`VeryHidden`) on the Rust side; the reader's
`xlsx_workbook_sheets` now also captures the `state` attribute. `tests/xlsx_roundtrip.rs`
gained a `synthetic_three_sheet_workbook_with_states` helper (a real fixture can't
exercise this) and two tests using it. `compat/differential-python/sheet_ops_check.py`
gained a `SheetStateAgreesWithOpenpyxl` class -- its fixture is built with openpyxl
itself (which can freely write `ws.sheet_state = "hidden"`), compared against elixcee's
read of the same file, plus the round-trip-loses-state regression test described above.

### Root crate (Python binding): P2 fifth slice -- row_height/column_width (read-only)

The fifth item off `docs/openpyxl-gap-audit.md`'s P2 list. Two new Python methods:
`row_height(row, sheet=None) -> float | None` / `column_width(col, sheet=None) -> float |
None`, sheet-parameterized like `hidden_rows`/`hidden_columns` (not name-addressed like
`sheet_state`) since these are row/column-level queries within a sheet.

Confirmed zero prior representation anywhere (not read, stored, or written), and
confirmed the writer's gap is worse than `sheet_state`'s: `xlsx_worksheet_xml`'s
`<row>`/`<cols>` emission is fully regenerated from `Vm.sheet_visibility` alone on EVERY
save -- not passthrough, not even an opaque fragment -- so a loaded file's row heights
and column widths are unconditionally dropped, not just sometimes. Pinned by a
differential-python regression test that checks the saved file's raw XML directly:
openpyxl's own `column_dimensions[letter].width` auto-vivifies a default-13.0 entry on
first `[]` access even for a column the file never set, which would have made a naive
comparison pass for the wrong reason.

Deliberately read-only, same reason as `sheet_state`: zero real fixtures have a genuine
custom row height or column width (a first grep pass falsely suggested `ht=` existed on
real `<row>` elements -- a substring false positive from `<sheetFormatPr
defaultRowHeight="15">`; fixture1's only `<col>` is the already-known hidden column D
with `width="0"`, not real data).

Two independent value types, not one enum like `sheet_state`: `row_heights:
HashMap<u32, f64>` (per-row, sparse) and `column_widths: Vec<(u32, u32, f64)>`
(range-shaped like `hidden_columns`, with a value attached) -- confirmed live that real
producers don't always coalesce same-width columns into one range either (openpyxl wrote
three separate single-column `<col>` elements for three identically-widthed columns, not
one `<col min="2" max="4">`). This pushed `rename_sheet`'s per-sheet-map re-key count from
9 to 11, and `copy_sheet`'s copied-field count from 7 to 9. `customHeight="1"`/
`customWidth="1"` are both required for `ht`/`width` to actually apply in real Excel; a
bare `ht`/`width` without the flag is not recorded, pinned by dedicated unit tests rather
than assumed from the spec.

`tests/xlsx_roundtrip.rs` gained a `synthetic_sheet_with_row_heights_and_column_widths`
helper and two tests. `compat/differential-python/sheet_ops_check.py` gained a
`RowHeightAndColumnWidthAgreeWithOpenpyxl` class, fixture built with openpyxl itself.

### Root crate: `FIND()` panicked on an empty search string

Found by the P2 fifth slice round's own `fuzz` CI job -- unrelated to what that round
actually changed; the fuzz corpus discovers whatever bugs already exist in the code it
runs against, regardless of the current PR's diff. `func_find`'s `.windows(n_chars.len())`
panics ("window size must be non-zero", a real `slice::windows` requirement) when the
search string is empty, e.g. `FIND("","abc")`. Fixed by matching trivially at the start
position for an empty search string -- the same convention VBA's `InStr` already uses in
this codebase -- plus a start-beyond-haystack guard for the identical
out-of-bounds-slice risk. Verified by temporarily reverting the fix and confirming all
three new tests fail with the exact original panic, including a reproduction using the
fuzz corpus's own crash bytes.

## [0.11.0] - 2026-08-26

Root `elixcee` (Rust crate + Python package) only: seven real-world bugs/gaps reported
against `0.10.1` (GitHub #2–#8), all from the same reporter building small Excel "RPA
action" wrappers on top of elixcee instead of pulling in openpyxl as a second
Excel-handling dependency. Minor, not patch: two genuinely new PyO3 methods
(`delete_sheet`, `get_cell_number_format`) and one new keyword argument (`set_sheet`'s
`index`) are real API additions, not just bug fixes. `elixcee-types`/`elixcee-wasm`/
`@elixcee/xlsx` all unaffected.

**#8 — `EntireColumn.Delete` deleted the row instead of the column (the most severe of
the seven: silent wrong-data deletion, not a no-op).** `Stmt::RangeDelete`/`RangeInsert`
carried no axis at all — the parser mapped `EntireRow.Delete` and `EntireColumn.Delete`
to the exact same AST node, so both always shifted by row. Fixed by adding an `Axis`
(`Row`/`Column`) to both statements; `Vm` gained `delete_cols`/`insert_cols` as the
column-axis mirror of the pre-existing `delete_rows`/`insert_rows`.

**#7 — `EntireRow.Insert`/`EntireColumn.Insert`/`Rows(n).Insert`/`Columns(n).Insert` were
silent no-ops.** `EntireRow`/`EntireColumn` only ever recognized `.Delete`/`.Clear`/
`.ClearContents` — `.Insert` fell through to `Stmt::Unsupported`. `Rows(n)`/`Columns(n)`
weren't recognized as statement-starting keywords at all. Fixed using the same `Axis`
infrastructure as #8: `EntireRow`/`EntireColumn` now handle `.Insert`, and new
`Rows(index)`/`Columns(index)` statement parsing (an `Expr`-typed index, like
`Cells(row, col)`, so a variable index works — unlike `Range(...)`'s string-literal-only
addressing) backs `Stmt::RowColDelete`/`RowColInsert`.

**#6 — `Range.Sort` ignored `Header:=xlYes`, sweeping the header row into the sort.** The
parser never captured `Header:=` at all. Fixed: `Stmt::RangeSort` gained a `header: bool`
field: `true` excludes the range's first row from both the sort and the write-back, which
stays exactly where it was. `Header:=xlNo`/omitted is unchanged (VBA's real default,
`xlGuess`, isn't modeled — no report/fixture evidence either way).

**#2 — a sheet created via `set_sheet()` (or VBA's `Sheets.Add`) round-tripped its name in
lowercase, ASCII names only.** The same bug class as the already-fixed GitHub #1 (display
name lowercased on save), but for a sheet with no `WorksheetOrigin` from a loaded file —
`ensure_sheet` never recorded one, so the writer's display-name fallback used the
lowercased internal key. Non-ASCII names were never affected (`to_lowercase()` is a no-op
on e.g. Japanese), which is exactly why this went unnoticed until a plain ASCII name was
tried. Fixed: `ensure_sheet` now records the caller's as-written name into
`WorksheetOrigin.original_display_name` at creation time.

**#3 — no direct sheet-deletion API, and no position control on sheet creation.**
Deleting a sheet required building and running a VBA snippet through `Vm::run()` — a much
heavier tool than the structural operation warranted. Added `Vm::delete_sheet(name)` /
Python's `delete_sheet(name)`, sharing its actual removal code with VBA's
`Sheets(name).Delete` (so 0.10.0-D4's save-time reachability pruning behaves identically
regardless of which caller deleted the sheet) but — unlike the VBA path, which never
validates existence — raising a clear error on an unknown name instead of silently
no-opping. `set_sheet()`/Python's `set_sheet()` gained an optional `index` (0-based),
placing a newly-created sheet at that position instead of always appending; ignored if the
sheet already exists (this VM still has no sheet-reorder primitive at all).

**#4 — no way to read a date-formatted cell as a date, or to see its number format at
all.** `get_cell()` returns a date-formatted cell as the raw Excel serial number (e.g.
`45366`), matching openpyxl's own underlying storage but not its converted-on-read
behavior — and elixcee had no way to expose the format string a caller could use to
convert it themselves either. Fixed via the reporter's own preferred option (exposing the
format, not auto-converting `get_cell` — a caller-visible type change would be a breaking
change a patch-adjacent fix shouldn't make): new `Vm::get_cell_number_format`/Python's
`get_cell_number_format(row, col)`, resolving a cell's `s="N"` style index through
`xl/styles.xml`'s `<cellXfs>` to either a custom `<numFmt formatCode="...">` definition or
the ECMA-376 built-in numFmtId table (a fixed, published constant — not something this
project's usual "no writer code until a fixture shows the shape" rule applies to). The
path-based reader (`WorkbookSheet`, backing `Vm`/`load_workbook_file`) gained a
`cell_number_formats` field for this; the pre-existing buffer/WASM path already resolved
`numFmtId` per cell but never turned it into a format string.

**#5 — `Range.AutoFilter` was a silent no-op, even with `Field`/`Criteria1` given.**
Partially fixed, with the scope boundary disclosed rather than silently bypassed: the
VM-side effect (hiding rows whose `Field`-th column, 1-based relative to the given range's
own left edge per real VBA's convention, doesn't match `Criteria1`) is now implemented,
reusing the same `Vm.sheet_visibility`/`<row hidden="1">` machinery a loaded file's own
hidden rows already round-trip through — verified end to end, including the save. The
`<autoFilter ref="...">` element itself (the dropdown-arrow UI state) is deliberately
**not** persisted: no real fixture in this repo has one, and this project's own hard gate
is no writer code for an OOXML structural element without fixture evidence (this was
already an open, disclosed item — ROADMAP.md's former "B5" note). A bare `.AutoFilter`
(no `Field`/`Criteria1`) remains a real no-op, matching real Excel (nothing to visibly
turn on without the element).

## [Unreleased]

Root `elixcee` (Rust crate + Python package): `0.10.0-D`, the last slice of the Lossless
Worksheet Preservation milestone `[0.10.0]` (below) — the first three slices, plus an
unrelated dependency-security fix, shipped as `elixcee` `0.10.0`; the unbound-`r:`-prefix
regression that `0.10.0` introduced was fixed and released separately as `[0.10.1]`
(below). The error-typed-cell fix in this section is a genuinely new, independent fix —
not yet released in any version (see `[0.12.0]`'s own "Packaging note" reference in
ROADMAP.md for what that blocks). `@elixcee/xlsx` (still unpublished,
`0.0.0-development`/`private: true`, no `publishConfig`): see its own two entries below
for exactly what's implemented, plus a CI observability addition for the shared WASM
bridge — both independent of the root crate's own `[0.12.0]` release above.

`0.10.0-D` (relationship-backed features, including the actual fix for
`SOURCE_REFERENCE_LOSS`): design decided (origin-based worksheet part naming — an existing
sheet's output part name stays `WorksheetOrigin.original_part_name` rather than being
renumbered by position; see `docs/xlsx-worksheet-preservation-0.10.0-design.md` §10 for the
full `WorksheetOutputPlan` design and D1–D4 breakdown).

**`D1` (output plan + tests, done).** New `WorksheetOutputPlan` +
`plan_worksheet_output(sheet_names, origins, reserved_part_numbers)`: an existing sheet
keeps its own `original_part_name` verbatim regardless of save-time position; a new sheet
gets `max(reserved) + 1`, where `reserved` is every `sheetN.xml` number that ever existed
in the source (including a deleted sheet's number, scanned from the raw passthrough ZIP
entries) — never reusing a freed number. `build_xlsx_content_types` /
`build_xlsx_workbook` / `build_xlsx_workbook_rels` and the per-sheet write loop all now
consume `&[WorksheetOutputPlan]` instead of separately re-deriving `sheet{i+1}.xml` /
`sheetId` / `r:id` at each call site. This is a real bug fix, not just a rename: before
this change, a surviving sheet's content was written to a position-derived part name
while its own passthrough `.rels` file stayed at its original name, so deleting an
earlier sheet could orphan a later sheet's relationship file — confirmed via a git-stash
A/B comparison against the pre-fix code, not just inferred. No relationship-backed
restoration yet (`r:id` references / `SOURCE_REFERENCE_LOSS` itself are still `D2`/`D3`),
per the approved D1 scope boundary.

One independent bug found and fixed while building end-to-end coverage for D1's
fresh-part-name branch: `Stmt::SheetsAdd` named a new sheet from `self.sheets.len() + 1`
alone, with no collision check — any workbook with a gap in sheet numbering (most
commonly: delete a middle sheet, then `Sheets.Add`) computed a name that collided with a
later surviving sheet, and `ensure_sheet()` no-ops on an existing key, so the `Add`
silently produced nothing. Fixed by probing upward until a free name is found. Confirmed
pre-existing via `git blame` (`72b5cc38`, 2026-06-21), unrelated to `0.10.0`.

**`<tableParts>` restored (first `SOURCE_REFERENCE_LOSS` fix, done).** Turned out D1 had
already satisfied `D2`'s stated goal ("carry a surviving sheet's `.rels` through at the
original part name with `r:id`s unchanged") as a side effect: the generic passthrough
loop has copied worksheet `.rels` files byte-identical since `0.9.0`, and D1 made them
land at the correct co-located part name — confirmed by running
`check_source_references()` against `fixture3` and finding every violation was "the
sheet doesn't reference the rId", never "the `.rels` differs from the original". So there
was no separate D2 slice to implement; the whole remaining gap was splicing the reference
itself back into the regenerated worksheet — that's what this does, for `<tableParts>`
specifically (schema position confirmed: right after `<pageMargins>`, since nothing
between the two is emitted yet). Gated on `rels_survived` (the sheet `is_existing` AND
its own `.rels` genuinely present in this save's passthrough set) — splicing a reference
whose `.rels` didn't survive would emit a dangling `r:id`, a real Excel repair warning
strictly worse than the prior silent inertness. `fixture3`'s `check_source_references()`
verdict is now `CLEAN`; `fixture4`/`fixture5` still report `SOURCE_REFERENCE_LOSS` for
hyperlink/vmlDrawing/drawing `r:id`s — unaffected, next slices.

`<drawing>`/`<legacyDrawing>`, `<hyperlinks>` (a rewrite of `0.10.0-B4`'s existing
relationship-free-only filtering, not a new addition), `<pageSetup r:id>` (only if a
fixture actually has one), and `D4` (rename/reorder/delete/add + reachability-based
deleted-part cleanup) remain not started, in that fixture-evidence order.

**`<drawing>`/`<legacyDrawing>` restored, same mechanism and `rels_survived` gate as
`<tableParts>`.** `fixture5_chart_image_freeze_print.xlsm`'s only worksheet-level
relationship is its `<drawing r:id>` — `check_source_references()` now reports `CLEAN`
for this fixture too. `fixture4_hyperlink_comment_name.xlsm`'s `<legacyDrawing r:id>`
(VML comment shapes) is restored as well, though that fixture's `.rels` also carries an
r:id-backed hyperlink, left unrestored at this point — `SOURCE_REFERENCE_LOSS` remained
on `fixture4` until the next change.

**`<hyperlinks>` r:id children restored — all 7 real fixtures now `CLEAN`, the last
`SOURCE_REFERENCE_LOSS` gap closed.** This one is a rewrite of `0.10.0-B4`'s shipped
behavior, not a new addition: B4 unconditionally excluded every r:id-bearing
`<hyperlink>` child (no relationship-graph reconnection existed at the time). Now that
`rels_survived` exists, `reader::extract_hyperlinks(xml, include_relationship_backed:
bool)` keeps location-only children unconditionally (unchanged from B4) and r:id-backed
ones only when `rels_survived` is true. `fixture4`'s hyperlink is r:id-backed (external
URL, `TargetMode="External"`) — exactly the shape B4's own negative test asserted must
NOT survive; that test is rewritten to assert the opposite (`fixture4` is now fully
`CLEAN`, its last violation cleared). `SOURCE_REFERENCE_LOSS` is eliminated from the
entire current fixture set: all 7 fixtures report `CLEAN` across every
`mechanical_check.py` category.

**`D4` (deleted-sheet reachability cleanup, done) — closes Known gaps item 15.** A
deleted sheet's own worksheet-level `.rels` (and whatever it exclusively pointed at —
tables, drawings, comments) used to survive a save as an orphan: byte-identical, but
unreferenced by anything in the output, invisible to `check_roundtrip()`'s structural
checks alone. Fixture→checker→writer, as always: `check_deleted_sheet_cleanup()` (a
package-reachability BFS over the source's own relationship graph, computed independently
of whatever the writer actually did) was written and self-test-verified first; the Rust
writer's `deleted_sheet_prunable_parts` is a direct port of the same algorithm, wired into
the existing passthrough-building loop so a prunable part never enters `passthrough` in
the first place — no separate cleanup pass needed, since `carried_overrides` and both
`carry_over_rels` calls already only keep what's still present. A part shared between the
deleted sheet and anything else (a surviving sheet, or a workbook-level relationship) is
correctly kept; a part exclusively reachable from the deleted sheet is correctly pruned,
transitively (its own `.rels`, and one level further for whatever that points at). Two
real bugs in the reachability computation were found and fixed on the Python checker side
before the Rust port even started: naively using `xl/workbook.xml` as a "reachable
elsewhere" root walks its own unfiltered `.rels`, which in the source still lists the
deleted sheet, silently reintroducing everything reachable from it — fixed by threading an
`exclude` set through every hop of the BFS, not just the initial roots. Verified against
the exact real scenario that exposed item 15 earlier this session (a fixture with a
relationship-bearing sheet deleted): the orphaned `.rels` is now genuinely absent, both
`check_roundtrip()` (no false positive) and the new checker report `CLEAN`. A dedicated
shared-vs-exclusive-target scenario was also run end to end through the CLI.

Sheet rename/reorder — two rows in the design doc's required test-case table — are
deliberately marked N/A: this `Vm` has no rename/reorder primitive (only `Sheets.Add`/
`Delete`), and adding VBA statement support purely to make a test-table row reachable
would invert this project's own hard gate (`src/vm/mod.rs`'s own stated position: "building
it now would be validated by nothing").

**Plain `<pageSetup>` restored (done).** Checked all 7 real fixtures for the `<pageSetup
r:id>` shape before starting — none has it; `fixture5`'s own `<pageSetup paperSize="9"
orientation="portrait" horizontalDpi="0" verticalDpi="0"/>` has no `r:id` at all, and no
fixture's `.rels` declares a `printerSettings` relationship. `r:id`-backed `<pageSetup>`
stays genuinely blocked on this project's own hard gate (fixture evidence required before
writer code). But `fixture5`'s plain `pageSetup` was itself a real, previously-uncaught
bug: never added to `_INLINE_WORKSHEET_ELEMENTS`, so it was silently lost on every save,
invisible to every existing checker category. New `check_page_setup()` — deliberately not
folded into `check_inline_worksheet_elements()`, same reasoning as
`check_internal_hyperlinks()` (already excluded for the identical hazard on
`<hyperlinks>`): unlike `sheetViews`/`pageMargins`/etc, `CT_PageSetup` genuinely CAN carry
an `r:id` per the real XSD, so a blanket present/absent check would false-positive the
day a real `r:id`-backed fixture shows up. This check only ever looks at an `r:id`-free
original `<pageSetup>`. New `reader::root_tag_has_rid()` gates the writer side the same
way: a plain `pageSetup` restores unconditionally (no relationship dependency, same as
`pageMargins`); one with `r:id` is dropped entirely, staying unrestored until a real
fixture justifies the same `rels_survived` gate every other relationship-backed element
uses. All 7 real fixtures now report `CLEAN` across every `mechanical_check.py` category.

This closes `0.10.0-D`'s only remaining open item for the current fixture set. Real-Excel
reopen verification remains not done for `tableParts`/`drawing`/`legacyDrawing`/
`hyperlinks`/`D4`/plain `pageSetup`.

### Root crate + `@elixcee/xlsx`: error-typed cells (`t="e"`) round-trip as real errors, not strings

ROADMAP.md Known gaps item 14, found live during `0.10.0-C`'s real-Excel verification
(fixture5's `D8`, a real `#VALUE!` cell) and pre-existing since well before `0.10.0`
(`git blame`: `72b5cc38`, 2026-06-21). `reader.rs`'s `SheetCell` enum had no `Error`
variant, so `xlsx_parse_cell` treated `t="e"` identically to `t="str"` — both became
`SheetCell::Str`. On save the error text was written into `xl/sharedStrings.xml` as an
ordinary string (`t="s"`), so the cell displayed the same text in Excel but was no longer
an error-typed cell underneath.

Fixed by threading a new `SheetCell::Error(ExcelError)`/`elixcee_types::ExcelError::FromStr`
through the reader, `Vm::populate_from_sheets`, and the writer, the same way
`Variant::Error` already is at the VBA-runtime level: `xlsx_cell_xml` now emits `t="e"`
with the literal error string in `<v>`, never shared-string indexed — confirmed against
real Excel's own output, which never puts e.g. `"#VALUE!"` in `sharedStrings.xml` either.
An unrecognized error string (a newer dynamic-array error like `#SPILL!`) falls back to a
plain string rather than guessing at a wrong code.

`@elixcee/xlsx`'s `read()` (via `crates/elixcee-wasm`) gets the same fix: error cells now
come back as `{t:"e", v:<BIFF numeric code>, w:<display string>}`, matching the real
`xlsx` oracle's own shape exactly (confirmed live: even reading a real Excel-authored
`t="e"` cell through `XLSX.read()`, the display string only ever appears in `.w`, never
`.v`). New differential case (`compat/differential/xlsx-read.test.mjs`, all 7 classic
error codes) — read differential count is now 34/34 MATCH (up from 33/33).

### CI: WASM artifact size observability

`packages/xlsx`'s WASM bridge size was already measured (`scripts/wasm-smoke.mjs` step 5)
but only printed to the console, with no baseline to compare against. Now diffed against a
committed baseline (`crates/elixcee-wasm/wasm-size-baseline.json`, updated by hand when a
size change is intentional) and written to the CI step summary; the vendored WASM bridge
build is also uploaded as a CI artifact. Observation only — no pass/fail threshold yet;
that needs a few normal builds' worth of data first.

### `@elixcee/xlsx` — `write()`/`writeFile()`/`writeFileSync()`

Independent of the root `elixcee` crate: no Rust changes, no new npm dependency.
`bookType: "xlsx"` only, output `type: "buffer" | "array" | "base64"`, producing a real
OOXML ZIP via a hand-rolled ZIP/XML writer (no zip/xml-builder dependency added) —
strings/numbers/booleans/dates/formulas, multiple worksheets, merges, sheet visibility,
hidden rows/columns, basic number formats, safe XML escaping. Unsupported input (a
non-`"xlsx"` `bookType`, an unrecognized `type`, an unsupported cell shape/type, a
non-finite numeric/formula-cached value, an oversized declared `!ref`) throws an explicit
`ELIXCEE_*` error, never silently ignored or truncated.

- **`packages/xlsx/src/internal/xlsx-writer.cjs`** (new) — the OOXML XML generator:
  `[Content_Types].xml`, both `.rels` parts, `docProps/{core,app}.xml`, `xl/workbook.xml`,
  `xl/worksheets/sheetN.xml`, `xl/styles.xml`. Output is deliberately constrained to
  shapes `src/reader.rs` (elixcee's own reader) already parses, verified by reading
  `reader.rs` directly — inline strings (not shared strings), a small built-in
  numFmtId table plus custom `<numFmts>` entries (164+) for anything else.
- **`packages/xlsx/src/internal/zip-writer.cjs`** (new) — a hand-rolled ZIP archive
  writer (local file headers, central directory, end-of-central-directory record,
  table-based CRC-32) with a deterministic fixed epoch, so two `write()` calls on the
  same `WorkBook` produce byte-identical output. Platform-agnostic by design: no
  `Buffer`, every byte buffer is a plain `Uint8Array` built with `DataView`/
  `TextEncoder` — real browsers never had `Buffer` regardless of bundler, so the shared
  writer core is built to work on both platforms from the start. DEFLATE compression is
  supplied by the caller as an optional callback rather than required internally (falls
  back to STORED, a legal ZIP/OOXML method, when omitted — this is what lets the browser
  entry reuse the same writer with no `zlib` access at all).
- **`compat/differential/xlsx-write.test.mjs`** (new) — 36 MATCH + 1 disclosed
  UNSUPPORTED case (`bookType: "ods"`, registered in `classify.mjs`'s
  `UNSUPPORTED_ALLOWLIST`), covering all three round-trip directions (own write -> own
  read, own write -> oracle read, oracle write -> own read) against a fourth,
  independently-computed baseline (oracle write -> oracle read); plus standalone checks
  for OOXML ZIP/XML structural validation (CRC-32, balanced XML, `[Content_Types].xml`/
  `.rels` cross-references), 12 malformed-workbook rejection cases, output-type
  agreement (buffer/array/base64 carry identical bytes), write-determinism, a real
  filesystem round trip for `writeFile`/`writeFileSync`, and the browser entry's
  behavior (both throwing `ELIXCEE_UNSUPPORTED_IN_BROWSER` for `writeFile`/
  `writeFileSync`, and `write()` itself working with no filesystem).
- `compat/differential/metadata.test.mjs` extended: `write`/`writeFile`/`writeFileSync`
  now among the 39/39 exports checked (name/length/property-descriptor/CJS-ESM-identity
  against the oracle), plus a `writeFile === writeFileSync` aliasing check.

### `@elixcee/xlsx` — make writer bundles work in Node ESM and browsers

**Two real bundler bugs found by actually bundling and running the code, not assumed, and
both fixed at the source**:

1. An esbuild `--format=esm --platform=node` bundle can never synchronously `require()`
   anything reached through CJS-origin code — confirmed neither a lazy require,
   `require('node:zlib')`, nor `--external:zlib` changes this; the documented, correct
   pattern is marking the whole package `external` (`--packages=external`), verified
   end-to-end and pinned as a permanent regression check in `scripts/wasm-smoke.mjs`
   (step 6).
2. An esbuild `--platform=browser` bundle refused to even build at all with a
   `require('zlib')` reachable anywhere in its module graph (dead code included, since
   esbuild can't tree-shake CommonJS `module.exports` properties). Fixed by isolating the
   Node-only `zlib.deflateRawSync` wrapper into its own new file,
   **`packages/xlsx/src/internal/deflate-node.cjs`**, and stubbing that exact path (plus
   bare `zlib`) to `false` in `package.json`'s `browser` field — the same mechanism
   already used for `elixcee_wasm.node.cjs`. This works because `browser`-field
   path-remapping happens at module-resolution time, before the stubbed file's contents
   are ever parsed; moving the `require('zlib')` around *within* `index.cjs` (tried
   first) did not work, since `index.cjs` itself is wholesale-included in the browser
   bundle graph via `index.browser.mjs`'s re-export of its other, browser-safe exports.

- `scripts/wasm-smoke.mjs` extended (step 6): `bundleAndRunWrite`/`runWriteBundle` verify
  all four combinations — inlined-ESM-must-throw, inlined-CJS-must-run,
  externalized-ESM-must-run, externalized-CJS-must-run — pinned as a permanent regression
  check for bug 1 above.
- `scripts/browser-smoke.mjs` extended: the bundled entry now calls `write()` then
  `read()` and asserts the round trip, plus a build-time assertion that the bundle
  contains zero `zlib` references at all — verified against a real headless Chrome
  process, not just a passing build.
- `scripts/pack-consumer-smoke.mjs` extended: a shared `WRITE_ROUNDTRIP` snippet exercises
  `write()`/`writeFile()`/`writeFileSync()` from inside a real `npm pack` + `npm install`,
  both from CJS and ESM consumers, plus a new step for `writeFile()`/`writeFileSync()`
  against a real filesystem.
- `docs/xlsx-architecture.md` — new "Phase D: `write()`'s Node-builtin bundling posture"
  section documents both bugs, why each fix works, and why bug 2's fix (isolating the
  Node-only `zlib` access) is a different problem from bug 1's (ESM+Node package
  externalization) and needs a different solution.

### Root crate: formula reference rewriting on row/column insert-delete (0.14.0-A / 0.14.0-A2)

`insert_rows_on_sheet`/`delete_rows_on_sheet`/`insert_cols_on_sheet`/`delete_cols_on_sheet`
(and their Python-bound `insert_rows`/`delete_rows`/`insert_cols`/`delete_cols`) now rewrite
formula cell-references workbook-wide instead of leaving every formula's text stale. Precise
scope, not "cross-sheet formulas are supported":

- **Supported**: an unqualified reference (`=A10`) shifts when its own formula's cell lives
  on the sheet being edited; a sheet-qualified reference (`=Sheet2!A10`, `='Sales 2026'!A10`,
  `='Bob''s Data'!A10` — quoting/escaping/case-insensitive identity all handled) shifts
  whenever it *names* the edited sheet, regardless of which sheet hosts the formula. A
  reference landing inside a deleted band becomes `#REF!` (a range with only one corner
  inside the band shrinks instead of collapsing, matching real Excel); the sheet qualifier,
  if any, is preserved through both cases.
- **Not supported**: cross-sheet formula *evaluation* — `evaluate()` explicitly refuses any
  formula containing a sheet-qualified reference rather than ever silently reading the
  active sheet's cell as if it were the qualified one; `recalculate_all` skips such formulas
  the same way it already skips unparseable ones, so one cross-sheet formula can't abort a
  whole-workbook recalculation. `set_cell_formula` still can't be used to *author* a new
  cross-sheet formula (it evaluates immediately) — only a formula already present (e.g.
  loaded from a file) benefits from the rewrite. External workbook references
  (`[Book2.xlsx]Sheet1!A1`), 3D references (`Sheet1:Sheet3!A1`), and range move are all
  still unimplemented (sheet rename reference-following is a separate entry below).
- Implemented as targeted text splicing over parser-tracked reference spans
  (`formula::parse_with_refs`/`shift_references`), not a general AST-to-formula-text
  serializer — everything outside a changed reference (operators, function names, literals,
  whitespace, unaffected references, unrelated sheets' formulas) is left byte-for-byte
  untouched. `$`-absolute cell/range reference parsing (`FormulaExpr::CellRef`/`Range` gained
  `abs_col`/`abs_row`) is a prerequisite this round also depends on.
- 60+ new tests across the parser, rewriter, evaluator, and VM layers, plus a real
  save→reload→re-save integration test confirming a plain re-save doesn't shift a reference
  a second time.
- **Discovered, unrelated, unfixed**: the XLSX writer silently drops a cell entirely —
  formula text included — whenever its cached value is `Variant::Empty`, regardless of
  whether a formula is present (reproduces for an ordinary same-sheet formula with no
  cross-sheet reference involved at all, e.g. `=IF(FALSE,1)`; pre-existing, not introduced by
  this work). Tracked as discovered work, not fixed here.

### Root crate: `rename_sheet` rewrites formula qualifiers referencing the renamed sheet

Follow-up to 0.14.0-A2 above, reusing its qualifier parser exactly as planned there.
`rename_sheet` now rewrites every formula reference qualified with the OLD sheet name,
workbook-wide, to the new one (`=Sheet1!A1` on any sheet becomes `=NewName!A1` after
`Sheet1` is renamed to `NewName`) — requoted/escaped per the new name's own requirements
(e.g. `'Sales 2026'!A1` when the new name needs quoting), regardless of how the old
reference was written. Fixes a real dangling-reference gap 0.14.0-A2 introduced: once
qualified references started parsing, a rename that left them unrewritten made them
silently resolve to nothing rather than simply failing to parse as before.

- Unqualified references are never touched, even on the renamed sheet itself — `=A1`
  still means "this same sheet", whatever it's now called.
- A case-only rename (`"Sheet1"` → `"SHEET1"`) still updates existing qualifiers to the
  new display casing, matching real Excel.
- A formula this parser can't parse at all (external workbook references, 3D references)
  is left completely untouched, same as the structural-edit rewrite above.
- `<definedName>` text referring to a sheet by name remains unrewritten — a separate
  mechanism, still out of scope; see `internal_docs/openpyxl-gap-audit.md`.
- New shared `Vm::rewrite_formulas_workbook_wide` helper backs both this and the
  structural-edit rewrite (only what gets rewritten differs) — range move is its planned
  third caller.
- 16+ new tests (rewriter targeting/quoting tests, VM wiring tests) plus a real
  save→reload integration test confirming the renamed qualifier persists through a file.

### Root crate: fix silent formula loss for cells with no cached value

Correctness fix, independent of the 0.14.0-A reference-rewrite work above (though
discovered while writing its integration tests). A formula cell with no cached value —
freshly typed and not yet recalculated, or a cross-sheet reference this engine
deliberately doesn't evaluate (0.14.0-A2) — was silently dropped ENTIRELY on save,
formula text included, whenever its value was `Variant::Empty`. "No cached result" and
"no formula" are different things; this made a real formula vanish from the saved file
with no error or warning. The bug was two-sided, on both the writer and the reader:

- **Writer** (`xlsx_cell_xml`, `src/lib.rs`): a cell with `Variant::Empty` and a formula
  now still emits `<c r="..."><f>...</f></c>` — no `<v>` element at all, rather than
  fabricating a placeholder value (`<v>` is optional per the OOXML schema; this is also
  what openpyxl itself does, modulo a self-closing `<v/>` it adds and elixcee doesn't
  need to match). A cell with `Variant::Empty` and NO formula is still omitted entirely,
  unchanged.
- **Reader** (`Vm::populate_from_sheets`, `src/vm/mod.rs`): even before this fix, the XML
  parser already correctly extracted a formula-only cell's `<f>` text into its own
  `formulas` map, independent of whether `<v>` was present — but the cell-population loop
  only ever walked the separate `cells` map (populated from `<v>`/inline-string content),
  so a `(row, col)` present only in `formulas` was silently skipped, even for a real file
  that already had this shape. A new second pass now inserts a `CellContent { formula,
  value: Variant::Empty }` for every formula-only cell `cells` doesn't already cover.
- Not a formula-evaluation or recalculation feature — the cell's value is `Variant::Empty`
  before and after this fix; only its *survival* through save/reload changed.
- 4 new integration tests (`tests/xlsx_roundtrip.rs`): the exact failure matrix
  (`=IF(FALSE,1)`, a string-literal formula, a same-sheet reference, a cross-sheet
  reference), an emission-matrix regression guard (formula+Integer/String/Boolean/Error
  cached values still round-trip unchanged; a plain empty formula-less cell is still
  omitted), and two consecutive saves not losing the formula the second time.
- **Not verified against a real Excel-authored fixture** — no real fixture in this repo
  contains a formula-only cell with no cached value, and authoring one requires actual
  Microsoft Excel, which (as documented throughout this project's history) has no
  scriptable path on this machine. The fix targets a schema-valid, openpyxl-observed
  shape rather than a speculative one, but this remains open verification, not silently
  claimed as done — see ROADMAP.md.

### Root crate: range-move formula-reference translation (0.14.0-A4, Stage 2 of 4)

`formula::translate_references_for_move` — the formula rewriter for a same-sheet range
move (cut-paste tracking). Not yet wired into `Vm`; no `move_range` API exists yet, so
this function has no caller in this round (Stage 3, "cell move API接続", wires it up).
Semantics researched and confirmed against real Microsoft documentation before writing
any code (real Excel has no scriptable path on this machine, same standing constraint as
above) — see `internal_docs/range-move-0.14.0-a4-design.md` for the full research and the
design decision this implements.

- Every reference (unqualified, or qualified naming the moved sheet itself) whose target
  cell falls inside the move's source rectangle translates by the move offset — this is
  the SAME mechanism whether the referencing formula's own cell is inside or outside the
  moved rectangle, matching real Microsoft's documented behavior ("the cell references
  within the formula stay the same" for what's left behind; a reference elsewhere
  "follows" a cell that moved) rather than two separate internal/external rules. `$`
  absolute-reference flags are preserved through the translation, matching real Excel.
- A range reference (e.g. `SUM(A2:D2)`) with **both** corners inside the source rectangle
  translates as a whole. **Neither** corner inside is a no-op. **Exactly one** corner
  inside — confirmed by Microsoft Community Hub as a real, narrower-than-expected shrink
  behavior in the one specific sub-case where the destination is still inside the same
  range, unconfirmed for the general case — is reported as `MoveRewrite::Ambiguous`
  rather than guessed at; the eventual `move_range` caller (Stage 3) must reject the
  *entire* move when this occurs, not just skip the one formula, since real Excel's
  correct output for this shape is unverified and a silent wrong guess would change what
  a formula computes (same severity class as the empty-cached-value bug above).
- Scoped to same-sheet moves only this round, deliberately not extended to workbook-wide
  qualified-reference following the way 0.14.0-A2 extended insert/delete — whether the
  "follows" mechanism applies identically across sheets is itself one of the design doc's
  disclosed open questions (§4-B), not assumed. Cross-sheet range move is an explicit,
  disclosed follow-up, not attempted here.
- A formula this parser can't parse at all (external workbook references, 3D references)
  is reported as `Err`, same non-fatal "leave this formula untouched" contract as
  `shift_references`/`rename_sheet_references` — distinct from `Ambiguous`, which is
  fatal to the whole move.
- 16 new unit tests in `src/formula/rewrite.rs` covering: inside/outside/both-corners/
  one-corner-ambiguous cases, absolute-flag preservation, a moved formula's own internal
  reference (still uses the same follow mechanism, not a separate relative-offset rule),
  negative offsets (move up/left), self-qualified vs. other-sheet-qualified references,
  multiple references in one formula, a reversed range reference (`B10:A1`), and parse
  errors still propagating as `Err`.

### Root crate + Python: `move_range` (0.14.0-A4, Stage 3 of 4 — cell-move API)

`Vm::move_range_on_sheet` and its Python-facing `move_range(addr, rows=0, cols=0,
sheet=None)` wire up Stage 2's rewriter above to a real, callable move operation.
Same-sheet only (see the design doc's disclosed cross-sheet open question). Verified
against the actual built Python extension, not just `cargo test` — a reference following a
move, an ambiguous move being rejected, and a real save→reload→recalculate round trip
were all exercised through `maturin develop --release` and a real `.xlsx` file, not only
unit-tested.

- **Validate-before-mutate, matching `merge_cells`'s existing precedent**: every formula
  cell on the sheet is scanned first via `formula::translate_references_for_move`, and the
  *whole* move is rejected with `Err`/`ValueError` — before a single cell is touched — the
  moment any formula reports `MoveRewrite::Ambiguous`. Only once the scan clears does
  formula-reference rewriting apply, followed by the physical cell relocation.
- **Source/destination overlap is handled atomically**: every source cell is read into a
  scratch `Vec` and removed from the sheet before any destination write, so a move whose
  destination overlaps its own source (e.g. shifting a column down by one row) can't
  clobber a not-yet-relocated source cell mid-move. This is new plumbing, not a reuse of
  `copy_areas_to_clipboard`/`ClipboardState` — that mechanism is copy-paste/values-only,
  not formula-aware, and moving is not copying. A pre-existing cell at the destination
  that isn't itself part of the move is silently overwritten, matching real Excel's own
  paste behavior.
- **Bounds validation lives at the Python-facing layer** (`src/lib.rs`), matching
  `merge_cells`/`sort_range`'s existing division of responsibility with the `Vm` core: the
  source address, and — this is the one genuinely new check this round needed — the
  *destination* rectangle's far corner (`dest_r1 + (r2-r1)`, computed from `rows`/`cols`),
  are both checked against Excel's real grid limits (1,048,576 rows / 16,384 columns)
  before calling into `Vm`. Missing this would let a large source range moved near the
  sheet edge translate some of its rows/columns past the real limit while silently
  succeeding for the others — the destination's *far* corner, not just its near one, is
  what actually needs checking.
- Does **not** move `merged_ranges`/`sheet_visibility`/`cell_style_indices`/
  `cell_number_formats` — the same disclosed 0.14.0-B gap `insert_rows_on_sheet` already
  has, not new to this round. Cached `.value`s are left stale, same as every other
  structural edit.
- 11 new `Vm`-level unit tests (`src/vm/mod.rs`): plain-value relocation, a moved
  formula's outside reference staying put, an outside formula following a reference into
  the moved block, a moved formula's own internal reference using the same follow
  mechanism, the ambiguous case rejecting the whole move with *nothing* mutated
  (confirmed by re-reading every cell involved after the rejected call), a self-overlapping
  move, overwriting a pre-existing unrelated destination cell, a zero-offset no-op, an
  unknown-sheet error, a qualified reference to a different sheet staying untouched, and
  merged ranges staying untouched.
- New `elixcee.pyi` stub entry for `move_range`.

### Root crate: merged-cell transform on structural edit and range move (0.14.0-B Phase 2)

`insert_rows`/`delete_rows`/`insert_cols`/`delete_cols`/`move_range` now transform
`merged_ranges` too, closing the gap the entries above disclosed ("merged ranges staying
untouched"). Real Excel's own behavior for the most common shape — a merge with only
*some* of its rows/columns falling inside a deleted band — has no dedicated Microsoft
documentation and came back unconfirmed after a targeted research pass (same
Microsoft-documentation-based method as 0.14.0-A4's range-move research, since this
machine can't run Excel); see `internal_docs/cell-metadata-transform-0.14.0-b-design.md`
§5 for the full findings and confidence levels per case, and §7 for the decision below.

- **Insert/delete**: reuses `formula::shift_bound_low`/`shift_bound_high` — the *exact*
  arithmetic a formula range already uses for insert/delete, now exposed as `pub(crate)`
  for this purpose (Phase 1, PR #23) — applied to a merge's row or column bounds on
  whichever axis is edited. For the cases real Excel's behavior is actually confirmed or
  reasonably well-supported (a merge entirely inside a deleted band is destroyed; an
  insert landing strictly inside a merge grows it; an insert at a merge's edges shifts or
  leaves it as expected), this matches. For the one case that came back genuinely
  unconfirmed — a delete removing only *some* of a merge's rows/columns — this applies
  the clamp arithmetic anyway as a **disclosed, unverified-against-real-Excel best-effort
  shape**, matching the precedent already set by the formula-empty-cached-value fix
  (`[Unreleased]`/PR #20): decided explicitly by the user rather than assumed, after two
  other options (reject the edit outright; leave only this shape untouched) were laid out
  with their tradeoffs.
- A merge that would collapse entirely, or survive but shrink to a single cell on both
  axes, is dropped rather than kept — `merge_cells` itself already refuses to create a
  single-cell "merge", so keeping one here would be inconsistent with this engine's own
  rule regardless of what real Excel does for this exact shape.
- **Range move**: a merge fully inside the moved rectangle translates as a whole; fully
  outside is untouched; a merge with only *partial* overlap makes the *whole move* fail
  (`Err`/`ValueError`, nothing mutated) rather than guessing — the same "reject rather
  than guess" precedent already established for a partially-overlapping formula range
  reference (`MoveRewrite::Ambiguous`). A moved merge landing on an existing, unrelated
  merge is rejected the same way, independent of the research question above — it follows
  directly from `merge_cells`'s own already-enforced overlap rule.
- Validate-before-mutate: `move_range_on_sheet`'s merge check runs in the same
  scan-before-any-mutation phase as its existing formula-reference check, so a move
  rejected for either reason leaves cells, formulas, and merges all completely unchanged.
- Verified through the actual built Python extension (`maturin develop --release`), not
  just `cargo test`: `insert_rows` shifting a merge, `move_range` translating one,
  `move_range` rejecting a partial-overlap case with the merge confirmed unchanged
  afterward, and a real `save_workbook` → `load_workbook` round trip.
- 9 new `Vm`-level unit tests. Updated one existing unit test and one real-fixture
  integration test that had pinned the OLD "merges are not shifted" behavior as the
  disclosed gap this round closes (both now assert the new, correct behavior instead).

### Root crate: hidden row/column interval transform on structural edit (0.14.0-B Phase 3)

`insert_rows`/`delete_rows`/`insert_cols`/`delete_cols` now transform `sheet_visibility`'s
hidden-row/column intervals too. Only the axis actually being edited is touched — a row
insert/delete can't affect which *columns* are hidden and vice versa, unlike a merge,
which is 2D and can be affected on either axis.

- Reuses the exact same `shift_bound_low`/`shift_bound_high` arithmetic as merges and
  formula ranges. Unlike a merge, there's no degenerate-size drop case — a hidden interval
  spanning a single row or column is a perfectly ordinary state (`set_row_hidden`'s own
  intervals already look like this), not something this engine's own API refuses to
  create — so an interval only disappears here if the clamp collapses it entirely (the
  whole hidden band fell inside a deleted band).
- **No range-move counterpart, deliberately** — hidden state belongs to the row/column
  itself, not to the cell content that moves through it, so `move_range_on_sheet` (which
  only relocates cell contents) has nothing to do with this map. Confirmed with a test
  that moving a range over a hidden row leaves that row's hidden state untouched.
- 7 new `Vm`-level unit tests, plus one updated real-fixture integration test (the same one
  Phase 2 updated) to assert the hidden row now shifts alongside the merge.
- Verified through the actual built Python extension: `insert_rows` shifting a hidden row,
  `delete_rows` leaving an unrelated hidden column alone, `move_range` never touching
  hidden state, and a real `save_workbook` → `load_workbook` round trip.
- `cell_style_indices`/`cell_number_formats` are 0.14.0-B's remaining Tier 1 fields, not
  yet transformed (next phase).

### Root crate: per-cell style/number-format transform (0.14.0-B Phase 4, Tier 1 complete)

`insert_rows`/`delete_rows`/`insert_cols`/`delete_cols`/`move_range` now transform
`cell_style_indices` and `cell_number_formats` too — the last two of 0.14.0-B's four Tier 1
fields (`merged_ranges`/`sheet_visibility` shipped in Phases 2/3). This closes Tier 1 of
the design doc entirely.

- Unlike merges/hidden intervals (both range-shaped), these two maps are keyed by exact
  `(row, col)` — the same shape as a formula's single-cell reference, not a range — so they
  reuse `formula::shift_cell_coord`/`CellShift` (now `pub(crate)`) instead of
  `shift_bound_low`/`shift_bound_high`. A key whose target cell falls inside a deleted band
  is dropped entirely — there's no surviving cell for its style/format to belong to.
- One generic helper (`shift_keyed_cell_map<V>`) backs both maps for structural edit, since
  they share the exact same `HashMap<(row, col), V>` shape.
- **Range move**: a style/number-format belongs to the cell it's on, so it moves with it,
  exactly like `CellContent` itself already does. Unlike merges, there's no
  ambiguous-partial-overlap case possible here — a single cell is either inside the moved
  rectangle or it isn't, no "in-between" shape exists for a point. A pre-existing entry at
  the destination that isn't itself part of the move is overwritten (moved entries are
  applied after stationary ones, so this is deterministic regardless of `HashMap` iteration
  order), matching `CellContent`'s own overwrite behavior on a move.
- 8 new `Vm`-level unit tests.
- Verified through the actual built Python extension: since these two fields are read-only
  from Python (populated only from a loaded file), verification used a real
  `openpyxl`-authored fixture with a date-formatted cell — `insert_rows` moving the format
  to the correct new cell, the old cell's format correctly gone, and a real
  `save_workbook` → `load_workbook` round trip preserving it afterward.

### Root crate: fix row height / column width being dropped on every save

A loaded file's row heights and column widths were dropped on **every** save,
unconditionally — a separate, pre-existing correctness bug from 0.14.0-B's own transform
work above, requested and fixed directly rather than as part of that phased plan.
`build_xlsx_sheet`'s `<row>`/`<cols>` emission was fully regenerated from
`Vm.sheet_visibility` alone; it now also reads `Vm.row_heights`/`column_widths`.

- `<row>` now carries `customHeight="1" ht="..."` and `<col>` carries
  `customWidth="1" width="..."` merged onto the *same* element as `hidden="1"` when a
  hidden interval and a size entry share the exact same `(min, max)` range — restoring the
  combined shape a single source `<col>`/`<row>` element can carry (the reader already
  parses both attributes off one element independently, unaffected by this change). A range
  mismatch, or an unrelated entry appearing alone, lands as an independent, non-merged
  element instead, matching how real producers (confirmed via `openpyxl`) don't always
  coalesce ranges either.
- This is a **preservation** fix — an already-loaded value now survives a save — not new
  **write** support: `set_row_height`/`set_column_width` (authoring a brand-new value from
  scratch) remain deferred. Zero real fixtures in this repo have genuine custom row
  height/column width data to validate that from-scratch writer shape against, a different,
  still-open question from correctly re-emitting a value the reader already extracts
  correctly (same standing "no real fixture" constraint documented since P2's fifth slice).
- New `Vm`-level unit test (`row_height_and_column_width_survive_a_save_and_reload`,
  `src/lib.rs`) covering a row/column with both a custom size *and* hidden state together
  (exercising the merged-attribute path, not just the size-only path), plus one with
  neither (confirming no stray default value appears). The differential-python test that
  pinned the old bug is inverted and renamed
  (`test_row_height_and_column_width_survive_an_elixcee_save`) and now also exercises a
  second save to confirm no re-drop or duplicate emission.
- 6 pre-existing real-fixture integration tests updated: fixture1's hidden column D
  (`width="0" hidden="1" customWidth="1"` in the source) now correctly round-trips its
  `customWidth`/`width="0"` alongside `hidden="1"`, where the old, buggy writer only ever
  emitted the `hidden="1"` half — the updated assertions reflect this as a genuine
  correctness improvement, not a relaxed check.
- Not verified against a real Excel-authored fixture with genuine data, for the reason
  above — targets a schema-valid, `openpyxl`-observed shape, disclosed as such rather than
  silently claimed as fully verified, matching the precedent set by the
  formula-empty-cached-value fix.

### Root crate: row height / column width transform (0.14.0-B Tier 2, unblocked by the fix above)

`insert_rows`/`delete_rows` now shift `row_heights`; `insert_cols`/`delete_cols` now shift
`column_widths` — the last two 0.14.0-B fields, closing Tier 2. Axis-only by construction
(a row edit can't affect column widths and vice versa) and, like `sheet_visibility`, never
touched by `move_range` — both belong to the row/column itself, not to moving cell content.
Row heights reuse `shift_cell_coord` (single-index shape); column widths reuse
`shift_bound_low`/`shift_bound_high` (range shape, no degenerate-size drop needed — a
single-column width is ordinary, unlike a merge). 8 new `Vm`-level unit tests; verified
against the built Python extension including a save/reload round trip.

### Root crate: fix `<autoFilter>` being silently dropped on every save

Found while scoping 0.14.0-C (`internal_docs/structured-object-transform-0.14.0-c-scoping.md`):
a loaded file's `<autoFilter>` was completely destroyed on every save, not merely stale —
confirmed empirically against a real `openpyxl`-authored fixture. Fixed as byte-preservation
only, no new `Vm` state or Python API (create/remove/filter-type API is `0.16.0`).

- `<autoFilter>` has no `r:id` (confirmed against openpyxl's own writer schema-order
  docstring, `worksheet/_writer.py`), unlike `tableParts`/`drawing`, so it needs no
  `rels_survived` gate — same unconditional treatment as `sheetFormatPr`/`dataValidations`.
  New `OpaqueWorksheetFragments::auto_filter` field, captured via the existing
  `reader::extract_raw_element` and spliced back verbatim, whole-element (children included).
- Schema position matters: `CT_Worksheet` orders `autoFilter` **before** `mergeCells`, not
  after — verified against openpyxl's `write_tail` order, not guessed, and confirmed live
  against a real round trip with both present.
- 3 new differential-python tests (`AutoFilterSurvivesAnElixceeSave`): a bare `ref`, a
  `filterColumn` with real filter values, and a same-sheet merge + unrelated cell edit
  (schema-position + no-collateral-damage check).

### Root crate: `rename_sheet` rewrites `<definedNames>` text instead of dropping it wholesale

Scoped and implemented after 0.14.0-C's retirement
(`internal_docs/defined-names-rename-preservation-scoping.md`): `rename_sheet` used to set
`defined_names_may_be_stale`, dropping the entire `<definedNames>` passthrough on ANY
rename, even for a name that never referenced the renamed sheet. It now rewrites each
`<definedName>`'s stale sheet-qualifier text in place, per name — `move_sheet`'s own
`localSheetId`-staleness case (a different, position-based failure mode) is unchanged and
still drops wholesale, since no state tracks the original load-time sheet-position order to
recompute it against.

- New `Vm.sheet_renames_since_load: HashMap<String, String>` (original lowercased name →
  current display name, collapsing a sheet renamed more than once to one entry) — needed
  because `<definedNames>` is never mirrored into `Vm` state, only re-read from the
  *original* source file's raw bytes at save time.
- Two-path rewrite per `<definedName>` value: the existing formula-reference rewriter
  (`rename_sheet_references`, unchanged) handles a plain reference and a genuine
  formula-valued name; a new, narrower reference-list grammar
  (`rewrite_reference_list_for_renames`) covers what that can't parse at all —
  comma-separated multi-area lists and full-row/full-column references
  (`Sheet1!$1:$3,Sheet1!$A:$A`), the near-universal real shape of `_xlnm.Print_Titles` and a
  common shape for `_xlnm.Print_Area`. A value neither path can confirm safe is dropped
  individually (not the whole block) only if it plausibly mentions a renamed sheet at all.
- Real-fixture verified: `fixture4`'s genuine `Sheet1!$F$5` named range and `fixture5`'s
  real, real-Excel-verified `_xlnm.Print_Area` (`Sheet1!$E$3`) both now survive a rename
  with the qualifier rewritten, previously both were dropped outright.
- Print_Titles' full-row/full-column shape verified against a real `openpyxl`-authored
  fixture (`ws.print_title_rows`/`print_title_cols`), including a second save-reload cycle
  with no further rename.

### Root crate: fix stale docstrings on `insert_rows`/`delete_rows`/`insert_cols`/`delete_cols`/`move_range`

Doc-only, no behavior change. These five methods' docstrings still claimed the pre-0.14.0-A/B
behavior (no merge/hidden-marker/style/number-format/dimension shifting on structural edit;
no merge/style/number-format translation on range move) — stale since those transforms
shipped. Corrected to describe what each method actually does today.

### Root crate: fix a from-scratch `Vm()`'s minimal `styles.xml` rejected by `openpyxl` on reopen

`XLSX_STYLES` (the minimal stylesheet a from-scratch `elixcee.Vm()` — no loaded source file —
emits, since it has no real `styles.xml` to pass through) had two bare `<fill/>` elements with
no `<patternFill>`/`<gradientFill>` child. `openpyxl.load_workbook()` rejected this on reopen
with `TypeError: expected <class 'openpyxl.styles.fills.Fill'>`. Fixed to match openpyxl's own
from-scratch default shape exactly (verified against a real `openpyxl.Workbook()` save):
`<fill><patternFill/></fill>` (index 0), `<fill><patternFill patternType="gray125"/></fill>`
(index 1). The other bare elements in the same minimal stylesheet (`<font/>`, `<border/>`,
`<xf/>`) do not trigger the same rejection — confirmed by a full from-scratch-save → openpyxl-
reopen → openpyxl-resave → reopen round trip, left as-is. Covered by a new differential-python
test (`FromScratchVmProducesAnOpenpyxlReadableStylesheet`); a stale comment in the same test
file (and this method's own module docstring) describing the bug as still-open is corrected.

### Root crate: `delete_sheet`/VBA `Sheets(...).Delete` no longer leak stale entries into 7 per-sheet maps

`remove_sheet` only ever cleaned `sheets`/`sheet_order` on delete, leaving a dead entry
under the deleted sheet's old (lowercased) key in `merged_ranges`, `sheet_visibility`,
`cell_style_indices`, `cell_number_formats`, `sheet_states`, `row_heights`, and
`column_widths` — harmless today (the stale key is never looked up again), but real, and
now fixed to mirror the same map list `rename_sheet` already re-keys. `worksheet_origins`
is deliberately **not** cleaned: `deleted_sheet_prunable_parts` and `no_sheet_was_deleted`
(`src/lib.rs`) both detect a deletion by diffing this map against the current sheet list,
so clearing it would blind both checks — confirmed the hard way, an initial version of
this fix that cleaned all 8 broke both mechanisms until narrowed to 7.

### Root crate: `set_number_format` — 0.15.0-A, Safe Style Engine's first slice

`xl/styles.xml` was 100% opaque byte-for-byte passthrough until now (or the hardcoded
`XLSX_STYLES` minimal default for a from-scratch `Vm()`) — no VBA statement in this VM
ever mutated a cell's style, so nothing needed to parse it. `set_number_format(addr,
format_code, sheet=None)` is the first thing that does, and it's real, from-scratch
write-path work, not a thin wrapper:

- `<cellXfs>` is parsed into opaque per-`<xf>` raw spans (attributes + any
  `<alignment>`/`<protection>` children captured verbatim, never interpreted) — the same
  "capture the raw span, don't parse it" shape already proven for `OpaqueWorksheetFragments`
  (hyperlinks, data-validations, `autoFilter`), applied to per-cellXf records instead of
  per-worksheet elements.
- A target format string resolves to an existing built-in id (0-49), an existing custom
  `<numFmt>` this file already defines, or a newly-minted custom id (164+, XML-escaped) —
  inverting the read-only `resolve_number_format`/`BUILTIN_NUMBER_FORMATS` machinery that
  already existed for reading.
- **No shared-style mutation, ever**: setting a cell's format clones its current `<xf>`
  with only `numFmtId` changed, then finds an existing byte-identical record to reuse or
  appends a new one — an existing `<xf>` (possibly shared by many untouched cells) is
  never mutated in place. Verified concretely: two cells sharing style index 0 (General)
  on a from-scratch `Vm()`, only one formatted — the other survives a save+reload as
  General, not silently corrupted to the new format.
- The loaded-file and from-scratch-`Vm()` cases share one code path:
  `resolve_pending_number_formats` treats `XLSX_STYLES` as just another starting document
  with the exact same `<numFmts>`+`<cellXfs>` shape a loaded file has, resolved only when
  a pending edit actually exists — the common (untouched) case still writes the original
  bytes back byte-for-byte, unchanged.
- Real-fixture verified for the common path: `fixture1_values_styles_merge_hidden.xlsm`'s
  genuine `numFmtId="4"` (`#,##0.00`) is correctly reused, not duplicated, when the same
  format is requested on a different cell. No real-Excel fixture has a *custom* `<numFmt>`
  though — minting one is a deliberate, disclosed exception to this project's "no writer
  code without a real fixture" hard gate (the same gate still blocking `set_row_height`/
  `set_column_width`/hidden-sheet-write): `<numFmt numFmtId="N" formatCode="...">`'s shape
  is unambiguous ECMA-376 §18.8.30 spec text with zero real-world producer variance to
  discover, so an `openpyxl`-authored synthetic fixture, verified via a real
  `openpyxl.load_workbook()` reopen, stands in.
- Immediate read-after-write consistency: `get_cell_number_format` reflects a pending
  edit right away, no save/reload needed — but the actual `<cellXfs>`/`<numFmts>`
  resolution is deferred to `save_workbook()`, since the starting styles document (a
  loaded file's raw bytes, or `XLSX_STYLES`) is only available there; `Vm` itself never
  holds those bytes. Calling `save_workbook()` twice back to back re-resolves from the
  exact same starting point both times — nothing is consumed or mutated persistently by
  a save, so repeated saves are naturally idempotent.
- **New, disclosed limitation** (pre-existing, not introduced by this change — confirmed
  by checking the reader, which already captures a value-less cell's `s="N"` into
  `raw_style_indices` regardless of whether it has content): a cell with no value by save
  time has no persisted format on disk, even though a read right after `set_number_format`
  still reports it — the writer only ever emits a `<c>` element for a cell present in
  `Vm`'s value map. A loaded file's own genuinely empty, pre-formatted cell is dropped the
  same way on any save today, completely independent of this feature. See ROADMAP.md's
  known gaps.
- Explicitly out of scope, per this milestone's own phase boundary: font/fill/border/
  alignment/protection (0.15.0-B) and copy-style/named-style/theme (0.15.0-C) — every
  other attribute on a cell's `<xf>` is copied verbatim, never interpreted. No
  style-explosion diagnostic threshold either — deferred to whenever real usage data
  exists to calibrate one against, ships without one for this first cut.

### Root crate: `set_style` — 0.15.0-B, Safe Style Engine's second slice (font/fill/border/alignment/protection)

`set_style(addr, font=None, fill=None, border=None, alignment=None, protection=None,
sheet=None)`. Extends 0.15.0-A's `<cellXfs>`-only find-or-append into `<fonts>`/
`<fills>`/`<borders>` too — three genuinely different record-merge shapes, not one
generic trick reused three times:

- **Font**: clones the cell's current `<font>` record, only touching the requested
  properties (`bold`/`italic`/`underline`/`strike` as an explicit `val="1"`/`val="0"`
  child — matches a real `openpyxl`-authored file's own convention, verified directly
  rather than assumed; `size`/`color`/`name` as value children) — every other property
  survives untouched. Verified against `fixture4`'s real, in-use hyperlink font
  (underlined, theme-colored, sized): setting only `bold` leaves the theme color and
  everything else exactly as they were.
- **Fill**: `{"type": "solid", "color": "..."}` REPLACES the whole fill record rather
  than merging (unlike font/border) — matches how Excel's own single-color fill picker
  works, not a patch onto a prior gradient/pattern. `fgColor` gets the visible color,
  `bgColor` gets the real `indexed="64"` "no second color" sentinel — the classic
  fgColor/bgColor-for-solid-fills convention this project had never touched before,
  confirmed against a real `openpyxl`-authored fixture's actual bytes before trusting it.
  Only literal RGB/ARGB hex colors (6-digit RGB, alpha assumed opaque, or 8-digit ARGB) —
  theme-relative color *minting* is 0.15.0-C's job; copying an existing theme color
  forward when cloning a record that isn't being recolored stays free.
- **Border**: clones the current `<border>`, touching only the named side(s)
  (`left`/`right`/`top`/`bottom`/`diagonal`, each with `style`/`color`) — inserted at
  their real, order-significant `CT_Border` schema position when a side is added fresh
  (not just appended at the end), so an out-of-order side never risks a real Excel
  repair warning.
- **Alignment/protection**: unlike font/fill/border, these are inline `<xf>` children,
  not a separate table — merged onto the EXISTING `<alignment>`/`<protection>` child's
  attributes (not replaced wholesale), since every real fixture in this project already
  carries `vertical="center"` on every `<xf>`; a naive replace would silently drop it.

**Mandatory correctness fix, not optional cleanup**: 0.15.0-A's `resolve_pending_number_formats`
ran exactly once per save, producing the one `(new_xml, effective_indices)` pair the
final write used. Bolting `set_style`'s own resolve pass on independently, starting fresh
from the same original bytes, would have silently discarded whichever of the two features
ran first whenever one cell got both a `set_number_format` and a `set_style` edit before
one save — a real silent-data-loss shape, not hypothetical (formatting a cell's number and
its visual style together is an ordinary edit sequence). Fixed by chaining:
`resolve_pending_style_attrs` now takes the number-format pass's own `(new_xml,
effective_indices)` as ITS starting point instead of the raw bytes, so both passes'
changes land in the final output. `pending_number_formats` itself (0.15.0-A, already
shipped) is untouched — a new sibling `pending_style_attrs` field holds `set_style`'s own
log, resolved second.

Same "no shared-style mutation" safety as 0.15.0-A, now spanning four tables instead of
one — verified concretely again with two cells sharing a style index, only one styled.
`set_style` calls on the same cell before one save accumulate (a `fill=...` call after an
earlier `font=...` call keeps both) rather than the later call wiping the earlier one.

Real-fixture status, disclosed per-property rather than bundled into one blanket
exception (font-authoring and vertical-alignment have real grounding; fill/border/
protection/most-alignment properties have none in this project's fixtures — same profile
as `set_row_height`/`set_column_width`'s still-open half): user granted the same
ECMA-376-spec-text exception 0.15.0-A's custom-numFmt path got, for all four, verified via
`openpyxl`-authored synthetic fixtures and real `openpyxl.load_workbook()` reopens.

Explicitly out of scope: theme-relative color minting, copy-style/named-style (0.15.0-C),
and the still-deferred style-explosion diagnostic threshold.

**New, disclosed limitation, pre-existing and unrelated to this change** (found while
verifying, on a plain from-scratch `Vm()` with zero style edits): `XLSX_STYLES` (the
minimal default stylesheet for a from-scratch `Vm()`, since PR #32's `<fill>` fix) has no
`<cellStyles>` element at all, unlike every real Excel/`openpyxl`-authored file — `CT_StyleSheet`
allows this (`cellStyles` is optional), but `openpyxl.load_workbook()` emits a
`UserWarning: Workbook contains no default style, apply openpyxl's default` on reopen. Not
a data-loss bug (every real value/style survives; confirmed across every test in this
round) and not a schema violation, just a cosmetic gap in the from-scratch default's
completeness — not fixed here, out of this phase's scope. See ROADMAP.md's known gaps.

### Root crate: `copy_style` and `set_style(..., named_style=...)` — 0.15.0-C1

The roadmap's single `0.15.0-C` line bundled five differently-sized features (copy style,
row/column default style, named style, theme-color resolution); scoping split it before
implementation, same discipline as every prior phase. This ships the two small,
well-grounded pieces — `copy_style` and named-style APPLY — as their own sub-phase.
Row/column default style and theme-color mint/read remain unscheduled (real work, neither
free nor required for this milestone; see `internal_docs/style-engine-0.15.0-c-design.md`).

- **`copy_style(source, dest, sheet=None)`**: copies a single cell's complete style (font,
  fill, border, number format, alignment, protection — everything, matching Excel's own
  Format Painter) onto every cell in `dest`. Pure index aliasing — `dest` cells simply
  point at whatever style index `source` resolves to, no new `<xf>`/font/fill/border
  record is ever minted. Resolved LAST at save time, after `set_number_format`'s and
  `set_style`'s own passes, so it automatically picks up a `set_style`/`set_number_format`
  edit made on `source` earlier in the same session, even before any save — verified
  concretely (style a cell, copy it immediately, confirm the copy shows the NEW style, not
  a stale one). A later `copy_style`/`set_style`/`set_number_format` call on the same
  destination cell always wins, but between two different features on the same cell,
  `copy_style` always applies last regardless of call order — a documented fixed-pass-
  order rule, not true chronological tracking.
- **`set_style(..., named_style="...")`**: points a cell at an EXISTING named style
  already defined in the loaded file's own `<cellStyles>` (e.g. `"Hyperlink"`, or a real
  file's own locale-specific spelling — Japanese-authored files spell it "ハイパーリンク").
  Bakes the named style's font/fill/border/number-format/alignment/protection directly
  onto the cell's `<cellXfs>` entry AND sets `xfId` — matching real Excel's own behavior,
  confirmed against `fixture4`'s real `xfId="1"`/`fontId="2"` cell (a naive xfId-only
  pointer-set would rely on inheritance real Excel itself doesn't use). Resolved FIRST on
  a `set_style` call, so `named_style="Hyperlink", font={"bold": True}` on the same call
  bakes the named style in, then bolds on top of it — verified against the real fixture.
  Defining a brand-new/undocumented named style is explicitly out of scope (Excel's own
  ~20-30 builtin style definitions are product design choices, not spec text — a
  categorically weaker basis for a real-fixture-gate exception than every prior one this
  project has granted; would need real fixture evidence per builtin style, which this
  project doesn't have beyond `builtinId` 0/8). An unknown name raises at
  `save_workbook()` time (deferred resolution, like every other style edit) — surfaces as
  `OSError`, since `save_workbook` maps every save failure to that regardless of cause,
  not `ValueError` as a first draft of this docstring incorrectly claimed before real
  verification caught it.
- Same "no shared-style mutation" safety as 0.15.0-A/B, reverified for `copy_style`
  specifically: retargeting one cell's style never touches another cell still sharing the
  cell's OLD style index.
- Real-fixture verified end to end: `copy_style` against `fixture1`'s real number-format
  dedup case; `named_style` apply against `fixture4`'s real, in-use "ハイパーリンク"
  style, including its theme-colored font surviving intact; a second save-reload cycle for
  stability.

### Root crate: `set_row_style`/`set_column_style` — 0.15.0-C2, row/column default style

- `set_row_style(row, font=, fill=, border=, alignment=, protection=, named_style=,
  sheet=)` / `set_column_style(col, ...)` — same kwarg shape as `set_style`, but sets an
  entire row/column's DEFAULT style (`<row s=".." customFormat="1">` / `<col
  style="..">`) rather than one cell range. A cell's own `set_style` always wins over its
  row/column's default — this project's job is only to persist both facts independently
  (`row_styles`/`column_styles` are separate `Vm` fields from `cell_style_indices`),
  never to resolve precedence itself.
- First row/column-level WRITE API in this project (`set_row_height`/`set_column_width`
  still don't exist — read-only, still blocked on the same missing-real-fixture gate).
  Reuses the exact `<xf>` find-or-append machinery `set_style` already built — a
  row/column style resolves through the identical clone-merge-dedup pipeline, just
  stored in `row_styles`/`column_styles` afterward instead of `cell_style_indices`.
- Chained into the SAME styles.xml resolve pipeline as `set_number_format`/`set_style`
  (not a fourth independent pass) — a cell touched by `set_style` and its row touched by
  `set_row_style` before one save both resolve correctly against the same font/fill/
  border/cellXf tables, matching the mandatory-chaining discipline 0.15.0-B's own scoping
  established for the number-format/style-attrs pair.
- `column_styles` is range-shaped (`(min, max, style_index)`, like `column_widths`) —
  setting ONE column's style splits any existing range containing it into up to two
  remaining sub-ranges (before/after) plus a fresh singleton for the touched column,
  mirroring how real Excel itself fragments a `<cols>` run when one column's formatting
  changes. `row_styles` is a plain per-row index map, like `row_heights`.
- Shifts on structural edit via two new near-duplicates of
  `shift_row_heights_for_structural_edit`/`shift_column_widths_for_structural_edit` (same
  logic, `u32` style index instead of `f64` dimension). `rename_sheet`/`remove_sheet`/
  `copy_sheet` extended to re-key/clear/copy the two new maps, following PR #33's own
  per-sheet-map-list discipline from day one.
- **Real-fixture verification exception, explicitly granted** (this project's usual
  hard gate — no complex writer without a real Excel-authored fixture — is crossed here
  with zero real-fixture grounding of any kind, the weakest evidentiary footing of any
  style-engine phase so far; granted on the same basis as every prior exception this
  session: `<col style>`/`<row customFormat>`'s attribute shape is unambiguous ECMA-376
  spec text with no producer variance to discover). The real `openpyxl`-authored
  attribute convention (`<row r=".." customFormat="1" s="..">`, `<col min=".." max=".."
  style="..">`) was verified against real bytes before implementing, not assumed.
- Verified: a real openpyxl-authored fixture's existing row/column default style
  survives an otherwise-unrelated save; `set_row_style`/`set_column_style` on both a
  from-scratch and a loaded `Vm()`, reopened cleanly with `openpyxl`; the mandatory
  chaining case (`set_number_format` + `set_style` + `set_row_style` all before one
  save, all three present after reload); a cell's own explicit style surviving distinct
  from its row's default; shared-style-mutation safety; a second save-reload cycle.

### Root crate: fix `<conditionalFormatting>` being silently dropped on every save

Found while scoping `0.16.0` (Tables, Filters and Rules) — the exact same bug shape as
the `autoFilter` fix above, just never previously disclosed. Confirmed empirically:
loaded `fixture3_table_validation_conditional.xlsm` (a real `cellIs` rule), saved, and
the output had zero `conditionalFormatting` occurrences, while `dataValidations`/
`tableParts` on the same fixture survived the same save fine.

- Root cause: an existing comment in `build_xlsx_sheet` explained `conditionalFormatting`
  needed "separate consideration" before joining the unconditional opaque-fragment list
  (it can reference `xl/styles.xml`'s `<dxfs>` via `dxfId`), but that consideration was
  never actually done — it simply never joined `phoneticPr`/`dataValidations`'s
  passthrough loop, nor was it captured anywhere else.
- Verified the "separate consideration" is a non-issue in practice: `<dxfs>` is never
  referenced by any `<cellXfs>`/`<fonts>`/`<fills>`/`<borders>` resolve pass (0.15.0-A
  through 0.15.0-C2) — each only ever `replacen`s the specific container it targets, so
  every sibling element, including `<dxfs>`, is carried through byte-identical regardless
  of any style mutation. A preserved rule's `dxfId` stays valid indefinitely.
- Fixed the same way as `autoFilter`: added to `OpaqueWorksheetFragments` as an
  unconditional, non-relationship-backed passthrough fragment. Unlike `autoFilter`/
  `dataValidations` (each occurs at most once per sheet), `CT_Worksheet` allows
  `conditionalFormatting` to repeat (one block per distinct range/rule-set) — a new
  `reader::extract_all_raw_elements` captures every occurrence in document order, not
  just the first, so a sheet with more than one conditional-formatting range doesn't
  silently lose all but one.
- Real schema position (immediately after `phoneticPr`, before `dataValidations`)
  confirmed against `fixture3`'s and `fixture4`'s actual bytes, not just the pre-existing
  comment's own claim — both agree on this order.
- Preservation only, matching `autoFilter`'s own scope: no structured `Vm` state, no
  create/edit API for conditional-formatting rules (that's real `0.16.0` feature work, a
  separate, much larger effort covered by its own scoping doc).
- Verified: `fixture3`'s real rule survives an otherwise-unrelated save and a second
  save-reload cycle; a sheet with two distinct `conditionalFormatting` blocks (an
  `openpyxl`-authored synthetic fixture) round-trips both; an unrelated `set_style` call
  (mutates `<cellXfs>`, never `<dxfs>`) leaves a preserved rule's `dxfId` untouched.

### Root crate: read-only `tables()` — 0.16.0-A1, Tables' first slice

First sub-phase of `0.16.0-A` (Tables), itself the first sub-phase of the
`0.16.0 — Tables, Filters and Rules` milestone (its own single roadmap line turned out to
bundle six differently-sized features, per a dedicated scoping pass — this slice is the
smallest, lowest-risk one: parse an existing table into real `Vm` state and expose it
read-only, with zero change to the write path).

- `vm.tables(sheet=None)` — every table on a sheet as a list of dicts (name,
  `display_name`, `ref`, header/totals-row info, `style_name`, the nested `autoFilter`'s
  `ref` if present, and a `columns` list). Parses `xl/tables/tableN.xml` via each sheet's
  own `.rels` file (`<tablePart r:id="...">` → `../tables/tableN.xml`, resolved the same
  way any other worksheet-scoped relationship is) — `xlsx_rels` generalized from a
  worksheet-only filter to take a `Type` suffix, reused for `"/table"`.
- **The headline finding, not just an implementation note**: preserving an UNMODIFIED
  table through an otherwise-unrelated save already worked before this PR, for free — the
  generic unknown-part passthrough loop, automatic content-type carry-over, the existing
  `rels_survived`-gated `<tableParts>` splice, and the existing generic
  deleted-sheet-reachability pruning all already handle it with zero table-specific code.
  This PR is a pure `Vm`-side READ projection on top of that — it does not touch, and must
  not change, that existing byte-identical-when-unchanged guarantee.
- Structured references (`Table1[@Qty]`, `[#Totals]`) and calculated-column *authoring*
  are entirely out of scope, mirroring 0.14.0-A's own cross-sheet-formula-evaluation
  exclusion — nothing in this codebase's formula parser has any structured-reference
  grammar today, and building it would be comparably large to everything else in Tables
  combined. `TableColumn::calculated_column_formula` is captured as raw, unparsed,
  unevaluated text; an *existing* calculated column survives verbatim.
- New `Vm.tables: HashMap<sheet, Vec<TableDef>>`, shifted on any structural edit (BOTH
  axes, like `merged_ranges` — a table's `ref` is a 2D rect, not a row/column-only
  dimension) via `shift_tables_for_structural_edit`/`shift_table_rect` (the same clamp
  arithmetic as `shift_merge_rect`, minus its merge-specific single-cell-collapse rule —
  a table has no "must span 2+ cells" invariant). `rename_sheet` now re-keys 19 per-sheet
  maps (up from 18); `remove_sheet`/`copy_sheet` extended to match.
- **Bug caught by the real-fixture verification itself, fixed before merge**: the first
  cut only shifted a table's own `ref`, leaving a nested `<autoFilter>`'s `ref` stale
  after a structural edit — found by the verification script, not assumed correct. Fixed:
  `auto_filter_ref` now shifts identically to `ref_range` (or drops if it collapses on its
  own), with a dedicated regression test.
- Verified against `fixture3_table_validation_conditional.xlsm`'s real, complete table
  (`テーブル1`, 3 columns, `TableStyleMedium2`, a nested bare `<autoFilter>`): `tables()`
  matches every real field; an unmodified table's `table1.xml` bytes survive an unrelated
  save byte-identical (and a second save-reload cycle); `openpyxl` still opens the result
  and still sees the table; a row insert above the table shifts both `ref` and
  `auto_filter_ref` correctly; a column insert shifts the column axis only.
- 0.16.0-A2 (edit an existing table) and 0.16.0-A3 (create a new table from scratch, the
  three-part linkage built from nothing) remain, per the scoping doc's own recommended
  split.

### Root crate: fix a value-less, pre-formatted cell being dropped from the worksheet entirely on save

`build_xlsx_sheet`'s cell-emission loop was built purely from `Vm`'s value map, but a
value-less cell (e.g. a merged-cell anchor styled but never given a value — `fixture1`'s
own `B1:C1`) never gets a value-map entry at all; only `cell_style_indices` (populated
unconditionally from the raw `s="N"` attribute on load) knows it exists. The cell's own
`<c>` element, and its style, silently vanished on every save regardless of
`set_number_format`/`set_style`.

- Fixed by also consulting the RESOLVED effective style-index map (the same map that
  already accounts for any pending `set_number_format`/`set_style` edit) when building
  the per-row cell list, synthesizing a value-less `<c r="..." s="N"/>` for any cell
  present there but absent from the value map — matching real Excel's own shape for a
  formatted-but-empty cell.
- Verified against `fixture1_values_styles_merge_hidden.xlsm`'s real `B1`/`C1` (both
  value-less, `s="2"`, the merged-cell anchor pair): reproduced the drop on the
  pre-fix code first, confirmed both cells now survive an otherwise-unrelated save and a
  second save-reload cycle.

### Root crate: fix a from-scratch `Vm()` triggering openpyxl's "no default style" warning on reopen

A from-scratch `Vm()`'s minimal `XLSX_STYLES` had no `<cellStyles>` element at all —
schema-legal, but `openpyxl.load_workbook()` raises `UserWarning: Workbook contains no
default style` on reopen. Non-fatal, no data loss, but spurious for every from-scratch
`Vm()` save with zero style edits.

- Reproduced first: dumped a real from-scratch `Vm()`'s save and confirmed the warning
  fires against the actual output bytes, not a hand-built approximation.
- Compared against a real `openpyxl.Workbook()`'s own from-scratch default `xl/styles.xml`
  directly — it always includes `<cellStyles count="1"><cellStyle name="Normal" xfId="0"
  builtinId="0" hidden="0"/></cellStyles>`, positioned right after `<cellXfs>` (matching
  `CT_Stylesheet`'s real child sequence). Added the equivalent to `XLSX_STYLES`, using the
  English `"Normal"` name (real Excel/openpyxl's own from-scratch default), not a
  locale-varied loaded-file name.
- Re-confirmed against the same real output bytes: the warning is gone.

## [0.10.1] - 2026-08-24

Root `elixcee` (Rust crate + Python package) only: `0.10.0` → `0.10.1`, a single targeted
bug fix, no new functionality. `elixcee-types`/`elixcee-wasm`/`@elixcee/xlsx` all
unaffected.

**A workbook whose source binds the OOXML relationships namespace to a prefix other than
the conventional `r:` (e.g. `xmlns:rel="..."` + `rel:id="..."` on `<sheet>` — fully valid
OOXML; namespace binding is about the URI, not the prefix spelling) round-tripped through
`0.10.0` into a file with an unbound `r:` prefix, rejected outright by any strict XML
consumer (openpyxl/lxml, Excel itself) — not a lossy passthrough, a hard parse failure.**
`build_xlsx_workbook`'s `<sheet r:id="...">` always hardcodes the literal `r:` prefix, but
the root `<workbook>` tag's own namespace declarations were carried through verbatim from
the source, with no guarantee the source declared `r:` at all. Reported against the
published `0.10.0` wheel; reproduced by rewriting a real openpyxl-authored file's
`xl/workbook.xml` (`xmlns:r`→`xmlns:rel`, `r:id`→`rel:id`, confirmed the rewritten input
is itself valid OOXML) and round-tripping it through both a local build and the actual
PyPI `0.10.0` wheel — output declared `xmlns:rel` but used `r:id="rId1"`, openpyxl raised
`ParseError: unbound prefix`. Six other scenarios (the original GitHub #1 repro, a
from-scratch workbook, multi-sheet, chained double-save, in-place save) were checked first
and found **not** reproducible — this is specifically an alternate-relationships-prefix
issue, not a broader regression.

Fixed by a new `reader::ensure_r_prefix_bound()`, applied before a source's root attribute
string is reused: if `xmlns:r` is already correctly bound, nothing changes; if it's
simply absent (the realistic case), the correct binding is appended; if `r:` is already
bound to some unrelated URI (essentially never seen in real files), the source's root
attrs are not reused at all, falling back to the writer's own safe hardcoded default
rather than risk a wrong rebind. In short: arbitrary relationship-namespace prefixes are
supported; a malformed workbook with no sheet relationship identifier at all (under any
prefix) is rejected at load time, unrelated to this fix (`<sheet>`'s relationship id is a
required attribute per the OOXML schema — no real producer omits it).

Full regression sweep clean (see the fix commit for the complete list); all 7 real
fixtures rerun with no new regressions (`fixture3`/`4`/`5` still show the already-known,
already-disclosed `SOURCE_REFERENCE_LOSS` gap `0.10.0-D` — unreleased, unrelated — not a
regression from this fix). Before release, six additional cases were verified end-to-end
against the real published PyPI `0.10.1` wheel and a real `cargo check` against the live
crates.io `0.10.1`: the standard-prefix case; the alternate-prefix case; `xmlns:r` already
correctly bound (no spurious growth); `xmlns:r` bound to an unrelated URI (falls back
without ever producing a duplicate `xmlns:r` declaration — confirmed via direct XML
inspection, not just successful parsing); and save-as/in-place/two-consecutive-saves, all
reopening cleanly in openpyxl. Released as `elixcee` `0.10.1` on PyPI, crates.io, and as a
GitHub Release (`bin-v0.10.1`, 3 platform binaries), all from the same commit. The
original reporter independently re-verified both of GitHub issue #1's fixes plus this
namespace fix against the published `0.10.1` wheel and closed the issue themselves.

## [0.10.0] - 2026-08-24

Root `elixcee` (Rust crate + Python package) only: `0.9.0` → `0.10.0`. `elixcee-types` stays
`0.3.0` (no changes this round), `elixcee-wasm` stays `0.1.0` (never published,
`publish = false`), `@elixcee/xlsx` stays unpublished/`private:true`/`0.0.0-development` —
none of them have any public-surface change this round; `@elixcee/xlsx`'s own in-progress
work stays under `[Unreleased]` above, untouched by this release.

This ships the first three slices (`0.10.0-A`/`0.10.0-B`/`0.10.0-C`) of the `0.10.0`
Lossless Worksheet Preservation milestone, all real-Excel reopen-verified with 0 repair
warnings. `0.10.0-D` (relationship-backed elements — the actual fix for
`SOURCE_REFERENCE_LOSS`) remains unreleased; see `[Unreleased]` above.

### Root crate: dev-dependency security fix, plus a real writer bug it uncovered

`cargo audit` (run for the first time in this project) found three RustSec advisories in
`Cargo.lock`: `RUSTSEC-2026-0204` (`crossbeam-epoch`, bench-only via `criterion`/`rayon`,
fixed with a plain `cargo update`) and two HIGH-severity (7.5) `quick-xml` advisories
(`RUSTSEC-2026-0195` memory-exhaustion DoS, `RUSTSEC-2026-0194` quadratic-time DoS) reached
via `calamine` — a `[dev-dependencies]`-only differential-testing oracle, never shipped to
users (`src/reader.rs`'s own header comment: elixcee's real XLSX/ODS reader "replaces
calamine as a runtime dependency"). Fixed by bumping `calamine` `0.24` → `0.36`.

That bump changed the oracle's own parsing behavior enough to fail an existing differential
test — calamine 0.36 trims leading/trailing whitespace on a shared-string `<t>` element that
lacks `xml:space="preserve"`, where 0.24 didn't. Tracing the raw XML elixcee itself writes
confirmed this isn't an oracle regression to route around: `build_xlsx_shared_strings` never
emitted `xml:space="preserve"`, so any string with leading/trailing whitespace round-tripped
ambiguously through *any* strict XML consumer — real Excel included. This is the writer-side
counterpart of the `xml:space="preserve"`-on-read fix already applied to `<v>` cells. Fixed by
emitting `xml:space="preserve"` on `<t>` whenever a shared string's content differs from its
own `trim()`, with a direct unit test on the raw XML output.

`cargo audit` now reports zero advisories. It's now wired permanently into CI (`rust-quality`
job, `cargo audit --version 0.22.1 --locked` after a from-scratch tool install, no
`continue-on-error`) rather than the one-off local run above — alongside `cargo fmt --all
--check`, `cargo clippy --workspace --all-targets -- -D warnings` (with and without the
`python` feature), and `RUSTDOCFLAGS="-D warnings" cargo doc --workspace --no-deps --features
python --document-private-items`, none of which had a CI job of their own before.

### `0.10.0-A`/`0.10.0-B`/`0.10.0-C` — Lossless Worksheet Preservation

Worksheet XML was always fully regenerated on save (`build_xlsx_sheet`), so anything elixcee
doesn't parse that lives *inside* a `<worksheet>` element — as opposed to a separate ZIP
part, already handled by `0.8.0`'s passthrough — was silently lost. This closes that for
worksheet-level and workbook-level elements, under a hard gate carried through every step
below: no writer code for an element until a real Excel-authored fixture demonstrates it,
its XSD sequence is confirmed against the actual ECMA-376 schema, and `mechanical_check.py`
has a negative test for its loss. Full design in
`docs/xlsx-worksheet-preservation-0.10.0-design.md`. Relationship-backed elements
(`<tableParts>`/`<drawing>`/`<legacyDrawing>`/r:id-backed `<hyperlinks>`) are `0.10.0-D`,
not part of this release — see `[Unreleased]` above.

**`0.10.0-A` (foundation, done).** `WorksheetOrigin` (original `sheetId`/`workbook.xml`
`r:id`/part name) now threads from `reader::WorkbookSheet` through `Vm` to the writer, so a
sheet's original `sheetId` survives a save instead of being renumbered by position every
time — closing a gap `snapshot.rs`'s own `stable_id` doc comment had already disclosed.
`mechanical_check.py` gained `check_source_references()` and a new `SOURCE_REFERENCE_LOSS`
violation category, distinct from the pre-existing `ORPHANED_PART`/`DANGLING_RELATIONSHIP`: a
worksheet-level relationship's `.rels` entry and target part can both survive a save
byte-identical while the regenerated worksheet XML no longer references the `r:id` that
activates it — confirmed systemic across every fixture with a worksheet-level relationship at
all (not yet fixed; that's `0.10.0-D`). Two new real Excel-authored fixtures added (internal
`location=` hyperlink, real freeze pane — the repo had neither before). One independent bug
found and fixed along the way: `Vm::new()`'s default empty `"sheet1"` wasn't cleared before
loading a real workbook, so any workbook whose sheets are never literally named `"Sheet1"`
gained a spurious extra empty sheet on every save.

**`0.10.0-B` (inline worksheet elements, functionally done).** An opaque-fragment mechanism —
capture an element's raw source XML, splice it back at the correct schema position, never
parse or reconstruct it — applied one element at a time, each slice independently checker-
verified (a new `INLINE_ELEMENT_LOSS`/`INTERNAL_HYPERLINK_LOSS` category per shape), fixture-
verified, and reopened in real Excel with 0 repair warnings before being called done:
`<sheetViews>` (freeze panes, active-cell selection), `<sheetPr>`/`<sheetFormatPr>`/
`<phoneticPr>`/`<dataValidations>`, `<pageMargins>`, and internal hyperlinks. The last one
needed a different mechanism than the rest: a `<hyperlinks>` container can mix
relationship-free `location=` children with `r:id`-backed ones that stay out of scope until
`0.10.0-D`, so it's reconstructed from filtered children (confirmed via the real
`CT_Hyperlinks` XSD: its `<hyperlink>` child is `minOccurs="1"`, so an all-`r:id` source must
omit the container entirely rather than emit an empty `<hyperlinks/>`) instead of copied
whole. Deliberately left out (not blocking): `<autoFilter>` (no fixture has it as a
standalone worksheet element yet) and row/column style properties beyond hidden state.

**Sheet order/writer bug found and fixed while scoping `0.10.0-C`.** `save_xlsx_impl` derived
its entire sheet-iteration order — worksheet part naming (`sheetN.xml`), `<sheets>` element
order, `sheetId` assignment — from `Vm::sheet_names()`, which sorts alphabetically. Every
existing fixture happened to already be alphabetical (`Sheet1`/`2`/`3`), so a save silently
reordering a workbook's tabs (e.g. "Zebra" then "Alpha" round-tripping as "Alpha" then
"Zebra", with no macro touching sheets at all) had never been caught. This is exactly the
kind of loss this milestone is chartered to close, and it directly blocked `0.10.0-C`:
`<bookViews>`'s `activeTab`/`firstSheet` are position indices, so carrying them through
opaque-fragment passthrough would have been wrong from the first non-alphabetical save.
Fixed with a new, separate `Vm::sheet_order` (insertion-ordered, kept in sync with `sheets`
by `ensure_sheet` — the single choke point behind every sheet-introducing call site — and by
`Sheets(...).Delete`) that `save_xlsx_impl` now reads instead. `Vm::sheet_names()` itself is
left unchanged (still alphabetical): it also drives `Sheets(i)`/`Worksheets(i)` numeric-index
resolution at VBA runtime, a separate, already-documented fidelity gap
(`docs/agent-contract.md`) this fix does not touch.

**A second, related bug found in the same pass: sheet display-name case wasn't preserved
either.** `build_xlsx_workbook` wrote sheet names straight from their lowercased `Vm` lookup
key (the only key space every per-sheet map uses), so `Sheet1` silently round-tripped as
`sheet1` — a visible tab-label change on every save, again with no macro involvement. Fixed
by adding `WorksheetOrigin.original_display_name` (the name exactly as written in the
source, alongside the existing `original_sheet_id`/`original_workbook_rel_id`/
`original_part_name`), which the writer now prefers over the lowercased key.

**`0.10.0-C` (workbook-level preservation, done).** Split into slices by
position-dependence, same discipline as `0.10.0-B`. **C1 (done):** `<workbookPr>`,
`<calcPr>`, `<extLst>`, and the root `<workbook>` tag's own namespace declarations —
all position-independent (no dependency on sheet order/count), opaque-fragment
passthrough, same mechanism as `0.10.0-B`. New `check_workbook_elements()` and
`WORKBOOK_ELEMENT_LOSS` category in `mechanical_check.py` (a workbook.xml-level check,
distinct from `INLINE_ELEMENT_LOSS`'s per-sheet-name matching — workbook.xml is a
single, fixed-path part). All 7 fixtures confirmed `WORKBOOK_ELEMENT_LOSS` before the
writer change, `CLEAN` after. **C2 (done):** `<bookViews>`. `<workbookView>`'s
`activeTab`/`firstSheet` are sheet-position indices, which in principle need their own
carry-over design if they ever hold a non-default value — but all 7 fixtures were
checked first and none of them sets either attribute (both default to 0 per the real
XSD), so a plain verbatim copy is correct against all current evidence. No gating
machinery was built for the unevidenced case: doing so ahead of a real fixture that
actually exercises it would be exactly the speculative abstraction this milestone's hard
gate exists to prevent. Shares C1's `WORKBOOK_ELEMENT_LOSS` category (a whole-element
copy, same as `workbookPr`/`calcPr`/`extLst` — no dedicated extraction logic needed,
unlike internal hyperlinks' filtered-children approach). **C3 (done, simplified):**
`<definedNames>` (print area/print titles included — `_xlnm.Print_Area` is a
`<definedName>` with a special reserved name, fixture5 has a real example). Unlike C2,
`localSheetId` (a 0-based index into `<sheets>`) DOES have real fixture evidence
(fixture5's `_xlnm.Print_Area localSheetId="0"`), so C2's "no evidence, ship verbatim"
reasoning doesn't apply here — but fixture5 is single-sheet, so the actual failure mode
(a delete shifting a *surviving* sheet's effective position) isn't fixture-evidenced
either, only reproducible with a synthetic multi-sheet fixture. Shipped a simplified,
conservative rule instead of per-name `localSheetId` remapping: `<definedNames>` is
carried verbatim only when every sheet present at load time is still present at save
time (`Sheets(...).Delete` never ran); if any sheet was deleted, the whole element is
dropped rather than risk a stale reference. Sheet *additions* don't affect this — new
sheets only ever append, so existing positions stay valid. New `check_defined_names()`
in `mechanical_check.py` (a dedicated function, not folded into
`check_workbook_elements()`'s plain presence check — this one needs to know the
correct answer differs depending on whether a sheet was deleted, checked by comparing
`<sheets>` between original and output). Verified both directions in self-test and
against a real CLI round-trip of a synthetic two-sheet fixture with a
`Sheets(...).Delete` macro.

**`0.10.0-C` real-Excel verified.** `fixture4`/`fixture5`, save-as and in-place, reopened in
Mac Excel: 0 repair warnings across all 3 output files. `fixture4`'s defined name (`test`,
workbook scope, `=Sheet1!$F$5`, comment `test desu!!!`) and `fixture5`'s `_xlnm.Print_Area`
(`Sheet1!$E$3`) both confirmed byte-for-byte correct in Excel's own Name Manager;
`Print_Area`'s print preview showed exactly the (empty) `E3` cell rather than the sheet's
real data table, confirming the print area is actually functioning, not just present as
inert XML. `0.10.0-C` is complete — mechanical-check-verified and real-Excel-verified.

One independent, pre-existing bug found during this verification (unrelated to `0.10.0-C`,
not fixed): a cell holding a real Excel error value (`t="e"`, e.g. `#VALUE!`) round-trips as
plain text, not an error — `SheetCell` (the file reader's cell type) has no `Error` variant,
so `t="e"` is read the same as `t="str"`. Confirmed pre-existing via `git blame`
(`72b5cc38`, 2026-06-21, well before `0.10.0` started). See `ROADMAP.md`'s Known gaps item
14.

## [0.9.0] - 2026-08-22

Root `elixcee` (Rust crate + Python package) only: `0.8.0` → `0.9.0`. `elixcee-types` stays
`0.3.0`, `elixcee-wasm` stays `0.1.0`, `@elixcee/xlsx` stays unpublished/`private:true`/
`0.0.0-development` — none of them have any public-surface change this round, and
`@elixcee/xlsx` is untouched by this round entirely (see `[Unreleased]` above for its own,
independent work).

**First real Microsoft-Excel-validated round trip (`0.9.0-A`, see `ROADMAP.md`).** Five
sanitized, Microsoft-Excel-for-Mac-authored `.xlsm` fixtures, each edited via elixcee, saved
both ways (save-as and in-place), and reopened in real Excel: 0 repair warnings, 0
`vbaProject.bin` loss, 0 relationship breakage, 0 in-place-save failures. Found and fixed
three real bugs none of the prior synthetic-fixture tests exercised — formula flattening,
orphaned relationships, and a wrong `.xlsm` content type that made Excel refuse to open the
file outright (see the "Root crate: safe round-trip, milestone 4" section below). New
`compat/oracle-excel-com/mechanical_check.py` (structural OOXML validator, self-tests against
7 deliberately corrupted cases before trusting any real result) and 5 real Excel-authored
fixtures under `compat/oracle-excel-com/fixtures/pristine/`.

**Explicitly not validated this round, stated precisely in `README.md`/`README_ja.md`/
`README_zh.md` rather than implied by a blanket claim**: post-save VBA macro execution (blocked
by a Mac Excel VBA license/environment error that reproduces on an untouched file — neither
confirmed working nor confirmed broken, not elixcee's own round-trip result either way); and
worksheet-embedded features (tables, data validation, conditional formatting, hyperlinks,
comments, defined names, charts, images, print settings) — `build_xlsx_sheet`/
`build_xlsx_workbook` still fully regenerate their XML on every save, so anything embedded
there that elixcee doesn't itself model is silently dropped (though the underlying ZIP part
usually survives as inert, orphaned bytes, and this never causes a repair warning). Already
disclosed as a `0.8.0` Non-goal, confirmed live here with real fixtures rather than newly
discovered — closing it is `0.10.0`'s job. The originally-scoped 10-consecutive-cycle
exit criterion is superseded, not literally met: a 5-cycle chained in-place stress test on
the same file (harder than 5 independent cycles, since any accumulating corruption would
compound) stayed clean through a real Excel reopen, judged sufficient in place of running
every fixture to the full 10. Full results:
`compat/oracle-excel-com/results/0.9.0-A_{results.json,summary.md}`.

`cargo test --workspace` 961/961 (up from 955 at `[0.8.0]`), `cargo build --release
--workspace` clean, `cargo check --features python --lib` clean, every GitHub Actions job
green on `master` before this bump.

### CLI: `elixcee --version`/`-V`

The CLI had no way to print its own version at all — found while verifying the `bin-v0.8.0`
GitHub Release binaries by hand (`gh release download` + run), where `--version` turned out
to be an unrecognized flag with no substitute (`--help` doesn't print it either). Fixed:
`elixcee --version`/`-V` now prints `elixcee <CARGO_PKG_VERSION>` and exits 0 — reads
`env!("CARGO_PKG_VERSION")` at compile time, so it can never drift from `Cargo.toml`. New
`tests/cli_version.rs` (2 tests, spawning the real built binary, matching the existing
`tests/cli_*.rs` convention).

### VBA: `Call <Sub>` without parentheses

Real VBA's `Call` grammar is `Call name [(argumentlist)]` — the parentheses are optional,
required only when passing arguments. `Call Foo` (a valid, zero-argument call) failed to
parse (`"expected LParen, got Newline"`), while `Call Foo()` and bare `Foo` (no `Call`
keyword at all) both already worked. Found and disclosed, not fixed, during `0.7.0` release
verification while writing a fresh-venv smoke test; confirmed unrelated to that round's own
changes (the parsing function hadn't been touched since the 2026-06-21 hand-written-parser
rewrite). Fixed: `parse_call_stmt` now checks for `(` before consuming it, the same pattern
this parser already uses for other optional constructs, defaulting to an empty argument list
when absent. Two new tests (zero-argument parenless call; a parenless call followed by
another statement, guarding against over/under-consuming the line).

### Root crate: safe round-trip, milestone 4 — three bugs found opening real-Excel output in real Excel

`0.8.0`'s safe-round-trip work (milestones 1–3) was only ever verified against hand-built
synthetic fixtures. `0.9.0`'s real-Excel validation (see `ROADMAP.md`) authored the first
genuinely Excel-produced `.xlsm` round-trip fixture and found three bugs none of the
synthetic fixtures happened to exercise — the third made Microsoft Excel refuse to open
the saved file outright, not even a repair prompt.

1. **Formula flattening.** `WorkbookSheet` (the struct backing the CLI `--file` path and
   PyO3 `load_workbook()`) had no field for per-cell formula text at all — `read_xlsx`
   already extracted it via the shared `xlsx_sheet_cells` parser, but discarded that half
   converting down from the buffer-API's `BufferSheet`. Every load flattened every
   formula cell to its last cached value; every subsequent save wrote that stale literal
   back with no `<f>` element, silently and permanently — editing any one cell dropped
   every *other* cell's formula on save. Fixed: `WorkbookSheet.formulas` (new field)
   threads the formula text through to `Vm::CellContent.formula` on load (keeping the
   file's own cached value, not recomputing — elixcee's formula engine doesn't cover
   Excel's full function surface) and `xlsx_cell_xml` now emits `<f>` before `<v>` when
   present. Shared-formula follower cells (`<f t="shared" si="N"/>`, no inline text) are
   a pre-existing, documented `reader.rs` limitation, unchanged by this fix.
2. **Orphaned relationships.** `xl/_rels/workbook.xml.rels` and `_rels/.rels` are both
   writer-owned (fully regenerated from a fixed template) — the template only ever
   emitted the relationships it already knew about (worksheet/sharedStrings/styles/
   vbaProject, officeDocument), with no mechanism to carry over any other kind.
   `xl/theme/theme1.xml` and `docProps/{app,core}.xml` passed through byte-identical but
   lost their relationships, becoming orphaned parts. Fixed: new `carry_over_rels()`
   parses the source's own rels files (`reader::workbook_rels_decls`, new) and re-emits
   any relationship whose target survived as a passthrough part, skipping types the
   writer already owns.
3. **Wrong content type for a non-macro `.xlsm`.** `build_xlsx_content_types` chose
   `workbook.xml`'s content type from whether the *source* had a VBA project, not from
   the *output* extension. A real Excel-authored `.xlsm` with zero VBA content still
   declares `application/vnd.ms-excel.sheet.macroEnabled.main+xml` — the macro-enabled
   type is a property of the file format, not current content. Any `.xlsm` output with no
   VBA project (the common case) declared the plain type instead, which Excel treats as a
   fatal extension/format mismatch. Now driven by the output path's own extension.

All three fixed and covered by new `tests/xlsx_roundtrip.rs` regression tests. New
`compat/oracle-excel-com/mechanical_check.py` — a pure-stdlib structural OOXML validator
(content-types, relationship completeness in both directions, formula preservation,
vbaProject byte-identity) that's the fast, Excel-independent primary signal for this and
future real-Excel validation rounds; its own self-test deliberately corrupts 7 different
ways before trusting any real result. Five real Microsoft-Excel-authored `.xlsm` fixtures
added under `compat/oracle-excel-com/fixtures/pristine/`. Full validation results in
`compat/oracle-excel-com/results/0.9.0-A_{results.json,summary.md}`.

## [0.8.0]

Root `elixcee` (Rust crate + Python package) only: `0.7.0` → `0.8.0`. `elixcee-types` stays
`0.3.0` (no public-surface change this round — every change lives in the root crate's own
`src/reader.rs`/`src/lib.rs`/`src/vm/mod.rs`), `elixcee-wasm` stays `0.1.0` (no real source
change; one test fixture needed a one-line mechanical fix to keep compiling against a new
`WorkbookSheet` field — see the milestone-2 section below). `@elixcee/xlsx` stays
unpublished/`private:true`/`0.0.0-development`, untouched by this round.

**"Safe round-trip", first three slices — a deliberately partial release, not the full
scope originally proposed for this direction.** The three sections below (unknown-part +
`xl/vbaProject.bin` passthrough; per-cell style-index + `xl/styles.xml` passthrough; merged
ranges + hidden rows/columns write-back) are shipped as `0.8.0` on their own merits — each
independently closes a real "your changes destroyed the workbook" failure mode. Still
genuinely unimplemented, deferred to a later release, not a scope cut discovered late: named
ranges (`<definedNames>`), tables/hyperlinks/comments/data-validation/freeze-panes/
print-and-page-setup embedded in worksheet XML, and richer workbook metadata. `cargo test
--workspace` 955/955 (up from 952 at `[0.7.0]`), `cargo build --release --workspace` clean,
`cargo check --features python --lib` clean, every GitHub Actions job green on `master`
before this bump. Verification for all three sections below is structural/synthetic-fixture
only (`tests/xlsx_roundtrip.rs`) — no real Microsoft-Excel-authored `.xlsm` exists in this
repo yet; see `tests/fixtures/xlsm_roundtrip/README.md` for where one slots in later.

### Root crate: safe round-trip, milestone 1 — unknown-part passthrough + `xl/vbaProject.bin` preservation

First slice of a longer "read an existing workbook, run/modify it, write it back without
destroying anything Excel put there" direction (see `ROADMAP.md`'s new gap 13 and
`docs/xlsx-architecture.md`'s new "Root-crate writer: regenerate vs. preserve-and-merge"
section for the full design). Until now, `save_xlsx_impl` (CLI `--output`, PyO3
`save_workbook()`) discarded the entire original file and regenerated a brand-new minimal
workbook from scratch on every save — most damagingly, `--output foo.xlsm` silently
produced a non-macro-enabled `.xlsx`-shaped file, losing `xl/vbaProject.bin` outright.

- **New `reader::read_raw_zip_entries`/`reader::content_type_decls`** (`src/reader.rs`) —
  read a source file's raw ZIP entries and its `[Content_Types].xml` declarations at save
  time only (not cached at load time, so `check`/`snapshot`/`diagnose`/`test-workbook`,
  which never save, pay zero extra cost).
- **`save_xlsx_impl`** (`src/lib.rs`) now merges writer-owned parts (still regenerated from
  `Vm` state exactly as before — `[Content_Types].xml`, workbook/rels/sharedStrings/styles,
  every `xl/worksheets/*.xml` matched by pattern, not a name list, so non-sequential
  surviving sheet parts don't leak through as stale orphans) with everything else copied
  through byte-for-byte from the source, including `xl/vbaProject.bin` when the output is
  `.xlsm` (dropped when it isn't, mirroring Excel's own Save-As-`.xlsx` behavior).
  `[Content_Types].xml`'s declarations for passed-through parts are carried over from the
  source's own declarations (exact `Override` match, then extension `Default`), not
  guessed — a hardcoded `Default Extension="bin"` would have mis-declared sibling parts
  like `xl/printerSettings/printerSettings1.bin` as a VBA project.
- **New `Vm::loaded_workbook_path`** field, set by `load_workbook_file` and PyO3's
  `load_workbook()`, is what makes the source re-readable at save time.
- **Deliberately not done in this slice** (see the docs section above for the full list):
  sheets are always fully regenerated, never diffed against the original; named ranges,
  tables/hyperlinks/comments/data-validation/freeze-panes embedded in worksheet XML,
  styles beyond existing numFmt handling, charts/images, streaming, and `.ods` passthrough
  all remain out of scope. `@elixcee/xlsx` and `crates/elixcee-wasm` untouched.
- **Tests**: new `tests/xlsx_roundtrip.rs`, 3 tests against hand-built synthetic `.xlsm`/
  `.xlsx` fixtures (no real Excel-authored `.xlsm` exists in this repo yet — see
  `tests/fixtures/xlsm_roundtrip/README.md` for the documented slot to add one later) plus
  a manual CLI smoke test of the realistic in-place `--file foo.xlsm --output foo.xlsm`
  overwrite case. `cargo test --workspace` 955/955 (up from 952).

### Root crate: safe round-trip, milestone 2 — per-cell style-index preservation + `xl/styles.xml` passthrough

Second slice of the same direction (see `docs/xlsx-architecture.md`'s "Root-crate writer:
regenerate vs. preserve-and-merge" section, "Slice 2" subsection, and `ROADMAP.md`'s item
13). Milestone 1's passthrough mechanism could technically have carried `xl/styles.xml`
through unchanged, but on its own that would have been pointless: the writer never emitted
a cell's `s="N"` style-index attribute at all, so every cell's font/fill/border formatting
was lost on every save regardless of whether the style *definitions* survived.

- **New `reader::WorkbookSheet::raw_style_indices`** — a cell's raw `s="N"` index, captured
  unconditionally whenever present, independent of the existing `style_ids` numFmtId
  resolution (a style index can carry font/fill/border info under a General number format).
  Threaded into new **`Vm::cell_style_indices`** by `populate_from_sheets`, same per-sheet
  pattern as `merged_ranges`.
- **`xlsx_cell_xml`** (`src/lib.rs`) now re-emits a surviving cell's original `s="N"`
  unchanged on every `<c>` arm.
- **Always safe, not just usually safe**: no VBA statement in this VM ever mutates a cell's
  style — `Range.Interior.Color =`/`.NumberFormat =` are explicit no-ops
  (`test_range_noop_interior_color`/`test_range_noop_numberformat`, pre-existing tests in
  `src/vm/mod.rs`). A cell's original style index is therefore still correct after any VBA
  edit to that cell's value; a brand-new cell simply has no entry to inherit.
- **`xl/styles.xml` conditional passthrough** — a distinct mechanism from milestone 1's
  general passthrough loop, not a generalization of it: stays in `is_writer_owned_part`'s
  fixed set, but its content is now the source's own bytes when available, falling back to
  the hardcoded minimal stylesheet only when there's no passthrough source. Deliberately
  paired with the style-index change above in the same slice: a cell's `s="N"` is only
  meaningful against the exact stylesheet it was captured from.
- **Tests**: no new test file — the 3 tests from milestone 1 (`tests/xlsx_roundtrip.rs`)
  extended in place (styled edited/untouched/brand-new cells, `xl/styles.xml` byte-identity,
  including across the in-place-overwrite case). `cargo test --workspace` still 955/955 (test
  *count* unchanged; existing tests got stronger assertions, not new tests).
- Deliberately not done: everything milestone 1 already deferred, still deferred — see that
  section above and `docs/xlsx-architecture.md`'s non-goals list. In particular, this VM
  still cannot *author or change* a style from VBA at all; only *preserving* an existing
  cell's style survived this milestone.

### Root crate: safe round-trip, milestone 3 — merged ranges and hidden rows/columns written back

No new reader work — `Vm::merged_ranges`/`Vm::sheet_visibility` were already populated by
`populate_from_sheets` and used elsewhere in the VM, but `build_xlsx_sheet` never emitted
`<mergeCells>` or a `<row>`/`<col>` `hidden="1"` attribute at all (confirmed live: zero
matches grepping `src/lib.rs` for `mergeCells`/`hidden` before this slice) — every save of a
workbook with merges or hidden rows/columns silently dropped them, independent of any
unknown-part-passthrough concern.

- `<cols>` (hidden columns, before `<sheetData>`), `<row hidden="1">` (including synthesizing
  an empty `<row r="N" hidden="1"/>` for a hidden row with no cell data — hidden-ness lives on
  the element itself, so an absent `<row>` reads as visible), and `<mergeCells>` (after
  `</sheetData>`) — all correctly OOXML-schema-ordered.
- `Vm::merged_ranges`/`Vm::sheet_visibility` promoted from private to `pub(crate)` so
  `save_xlsx_impl` can read them directly.
- **Tests**: no new test file — the same 3 tests extended again (merge/hidden-column/hidden-row
  assertions added to the flagship test, a merge-survival assertion added to the
  in-place-overwrite test). `cargo test --workspace` still 955/955.
- Narrows (doesn't close) the "embedded inside worksheet XML" gap milestone 2 noted: merges
  and hidden rows/columns now survive; hyperlinks, data validation, conditional formatting,
  freeze panes, and print/page setup — also worksheet-XML-embedded — still don't.

## [0.7.0]

Root `elixcee` (Rust crate + Python package) and `elixcee-types` only: `elixcee-types`
`0.2.0` → `0.3.0` (new, additive `Variant::VbaArray` enum variant — `Variant::Array` itself
is unchanged, but any code exhaustively matching on `Variant` needs updating, hence the
minor bump), root `elixcee` `0.6.0` → `0.7.0`, root `Cargo.toml`'s own
`elixcee-types = { ..., version = "0.3.0" }` dependency pin updated to match (not checked by
`scripts/check-versions.sh`, which only compares root `Cargo.toml` against `pyproject.toml`
— this pin needs a manual edit at every bump touching `elixcee-types`). `elixcee-wasm` stays
`0.1.0` (no source changes — grep-confirmed it references none of `vm::`/`Variant::`/
`VbaArray` directly and already compiled clean against the new types — but its build
output, vendored into `@elixcee/xlsx`, was refreshed to reflect these VM/type changes; see
`build(wasm): refresh packaged artifacts for array and writer changes` in git history).
Three independent Rust-side items: real multi-dimensional VBA arrays; call-frame-scoped
`On Error` with `Err.Source`/`Err.HelpFile`/`Err.HelpContext`/full 5-argument `Err.Raise`;
and moving undefined-procedure calls, argument-count mismatches, and undefined `GoTo`
labels to a compile-time, `On Error`-uncatchable check — plus, found as necessary
groundwork for the third item, a pre-existing parameter-parsing bug affecting any macro
using `ByVal`/`ByRef`/`Optional`/`ParamArray`. `cargo test --workspace` 952/952 (up from
872 before this round — 80 new tests: 826 lib + 82 integration + 25 `elixcee-types` + 19
`elixcee-wasm`), `cargo build --release --workspace` clean, `compat/vba-semantics` 386
cases (0 `BUG`/0 `UNCLASSIFIED`, `KNOWN_LIMITATION` 14 — down from 16), `compat/corpus` 581
scenarios (0 `UNEXPLAINED`/0 `MISMATCH`), every GitHub Actions job green on `master` before
this bump. `RuntimeErrorKind`/a typed `RuntimeError` struct (this round's own
lower-priority item) is still not started — deliberately deferred, not an oversight.

### Real multi-dimensional VBA arrays

Previously, `Dim arr(3, 2)` allocated storage as if it were 1-D (dimension 1's element count
only), so `arr(1, 1)` and `arr(1, 2)` silently aliased the same storage slot and
`UBound(arr, 2)` returned dimension 1's bound regardless of which dimension was actually asked
for — disclosed as `KNOWN_LIMITATION` in `compat/vba-semantics` (`two_dimensional_array_second_index_is_silently_dropped`,
`ubound_second_dimension_argument_ignored`), now fixed and reclassified.

- **New `elixcee_types::VbaArray`/`ArrayBound`** (`crates/elixcee-types/src/lib.rs`) — flat,
  row-major storage (`idx = idx * bound.len() + (sub - bound.lower)`, first dimension varies
  slowest) with real per-dimension bounds, replacing 1-D-only storage for every `Dim`-declared
  array. New `Variant::VbaArray(VbaArray)` enum variant — additive, `Variant::Array(Vec<Variant>)`
  itself unchanged (still used for Range-value multi-cell reads, `formula::eval`'s array-formula
  results, and `DimArrayRecord`/`ArrayRecordSet` storage, none of which have per-dimension bounds
  to track). Element count is overflow-checked (`checked_mul`) and capped at 10,000,000 elements,
  surfacing real VBA's own "Out of memory" wording rather than a Rust-side allocation panic.
- **`LBound`/`UBound`** now honor the dimension argument per-dimension; `Option Base` applies
  independently to every dimension; `Erase` resets elements while preserving all dimensions'
  bounds; **`ReDim Preserve`** is correctly restricted to real VBA's own rule — only the *last*
  dimension's *upper* bound may change, every other dimension (and the last one's own lower
  bound) must stay identical or it's Error 9 (`redim_preserve` in `src/vm/mod.rs`, found and
  fixed its own bug during this round: an earlier version of the check missed that the last
  dimension's *lower* bound is equally protected, not just the non-last dimensions).
- Shape preserved through variable assignment, function-argument passing, and function-return
  values; `Array()`/`Split()` migrated to `VbaArray` while keeping their externally-observable
  0-based rank-1 shape unchanged.
- **PyO3 bindings** (`src/lib.rs`): new `vba_array_to_py` recursively reshapes flat `VbaArray`
  storage into nested Python lists matching the array's real dimensional shape — verified against
  a real `maturin`-built wheel, not just `cargo test`.
- `crates/elixcee-wasm` needed no changes — grep-confirmed it references none of `vm::`/
  `Variant::`/`VbaArray` directly, and it already compiled clean.

### Call-frame-scoped `On Error`

Previously, `On Error Resume Next`/`On Error GoTo <label>` state was a single `Vm`-wide flag —
a callee's own body could see and mistakenly try to resolve a caller's still-active `GoTo`
label, and (found and fixed as part of the same rework) a callee's remaining statements kept
running under a caller's `On Error Resume Next` even after the callee itself failed, since the
catch fired inside the callee's own `exec_stmt`, not at the call site.

- **New `Vm::call_stack: Vec<CallFrame>`**, each frame holding its own `ErrorMode` (`Disabled`/
  `ResumeNext`/`GoTo(String)`), replacing the old `on_error_resume_next: bool`/
  `on_error_goto_label: Option<String>` fields. Pushed/popped around every `call_sub_def`/
  `call_func_def` invocation, so a callee always starts with `Disabled` regardless of the
  caller's own mode — matching real VBA (error handling doesn't inherit into a callee). A
  `GoTo` handler is consumed (reset to `Disabled`) the moment it fires, matching real VBA: a
  second failure while already inside a handler propagates to the caller rather than
  re-entering the same handler.
- **Deliberate behavior change**: under the old flag, a caller's `On Error Resume Next` catching
  an error from inside a called Sub let that Sub's *remaining* statements keep running (the
  catch happened inside the callee's own body). Now the error propagates out of the callee
  entirely and is caught at the `Call` statement in the caller's own frame — the callee's
  remaining statements do not run. This is the correct real-VBA behavior, but a macro that
  depended on the old leniency will observe the difference.
- Incidental fix: `run_sub`/`run_sub_multi` never reset the old `on_error_resume_next`/
  `on_error_goto_label` fields between runs on a reused `Vm` (the Python bindings' own usage
  pattern) — `call_stack.clear()` at the start of each run closes that.

### `Err.Source`/`Err.HelpFile`/`Err.HelpContext`, full `Err.Raise`

- **`Err.Source`/`Err.HelpFile`/`Err.HelpContext`** added as readable properties (`Expr::ErrSource`/
  `ErrHelpFile`/`ErrHelpContext`), joining the existing `Err.Number`/`Err.Description`.
  **`Err.Raise`** now accepts and threads through all five real positional arguments (`Number,
  Source, Description, HelpFile, HelpContext`), correctly handling a bare comma skipping any of
  the last four at any position (`Err.Raise 513, , "text"` still means Number=513,
  Description="text", not Source="text"). **`Err.Clear`** now resets all five properties, not
  just Number/Description.
- Not done: the richer `RuntimeError`/`RuntimeErrorKind` struct this phase's own spec also
  asked for (a `span`/`kind`-classified error type, replacing string-matching-based
  `classify_vba_error_number`) — the five `Err.*` properties above are still backed by flat
  `Vm` fields (`err_number`/`err_description`/`err_source`/`err_help_file`/`err_help_context`),
  not a unified struct. Deliberately deferred, not an oversight.

### Undefined-procedure calls, argument-count mismatches, and undefined `GoTo` labels are now compile errors

Real VBA fails a macro that calls an undefined Sub/Function, passes the wrong number of
arguments to one it can see, or `GoTo`s a label that doesn't exist — *before* running a single
statement, and never lets `On Error` trap it (a whole-project compile phase, not a runtime
check). Previously all three were ordinary runtime errors here, reachable partway through
execution and (incorrectly, relative to real VBA) catchable by an active `On Error Resume
Next`/`GoTo`.

- **New `check::compile_check_errors`** (`src/check.rs`) walks the whole program (every Sub and
  Function, not just the ones the entrypoint's call chain actually reaches — matching real
  VBA's whole-project compile-then-run semantics) for exactly these three conditions, reusing
  the same `is_resolvable` logic `check`'s existing undefined-name diagnostic (E1002) already
  used. **`Vm::run_sub`/`run_sub_multi`** now run this once, before `call_sub_def`, and return
  its finding as an ordinary `Err(String)` — the "uncatchable by `On Error`" property comes for
  free from running it before any statement (including any `On Error`) has executed.
  Multi-module runs build each module's own cross-module name set the same way `elixcee check`'s
  own multi-module path already did, so a legitimate unqualified cross-module call isn't
  misreported as undefined.
- A deliberately-unimplemented `WorksheetFunction.*` call (e.g. `.TextJoin`) is still reported,
  but with the exact message its real dispatch path (`eval_wsf`) would give at runtime — a new
  `vm::builtin_call_error` helper asks the VM itself instead of guessing generic wording, so
  `check`'s pre-flight rejection always matches word-for-word what actually running it would
  have said.
- **`elixcee check` learned the same two checks** (new diagnostic codes `E1008`/
  `argument_count_mismatch`, `E1009`/`undefined_label` — see `docs/agent-contract.md`), closing
  a gap this round's own testing surfaced: without this, `elixcee check` could report a program
  clean that `run_sub`'s new pre-flight pass would then refuse to run at all. Every violation is
  reported (not just the first, unlike `run_sub`'s own short-circuit-on-first-violation pass).
- **Measured, not assumed, performance impact**: `is_known_builtin_function`/
  `builtin_call_error` each construct a throwaway `Vm` and run a real dispatch probe per
  unresolved name, and the whole pre-flight check now re-runs on every `run_sub` call —
  relevant since `test-workbook` reruns the same macro across many generated cases. Measured
  with a `test-workbook` fixture calling 10 distinct built-in/`WorksheetFunction` names
  (deliberately adversarial — a typical macro has far fewer), 3000 cases: roughly 5% slower
  than the pre-this-round build (~0.68s → ~0.72s wall-clock for the whole run, ~13µs/case). Not
  optimized (no memoization of the builtin probe) — the absolute cost is small and the fixture
  used to measure it overstates a typical macro's actual builtin-call density.
- **Deliberately not checked**: "invalid assignment target" (e.g. calling a Function's result on
  the left of `=` as if it were an array element) — `name(args) = value` parses unconditionally
  as `Stmt::ArrayWrite` regardless of whether `name` is a real array or (invalidly) a Function
  name, and telling those apart isn't syntactically decidable without type inference this
  project stays out of by design. Also not checked: a cross-module call's argument count (this
  pass only ever sees one module's own declared Sub/Function arities), and a recursive call's
  own argument count (a procedure's own name is already in its local scope, so it's never
  treated as a checkable external call).
- **Found while building this**: `Sub Foo(ByVal x As Integer)` used to silently misparse —
  `parse_params` had no special handling for the `ByVal`/`ByRef`/`Optional`/`ParamArray`
  keywords, so `consume_ident()` swallowed `byval` itself as a bogus extra parameter name,
  making `Foo` a real *2*-parameter Sub (`["byval", "x"]`). Calling `Foo(5)` bound `5` to the
  phantom `byval` parameter and left `x` unbound — `x` inside `Foo`'s body raised "Undefined
  variable: 'x'", not a type/argument error, so this was easy to miss. Pre-existing, unrelated
  to any array/call-frame work above; found because it would have made the new argument-count
  check wrong for any macro using these (very common) parameter modifiers. Fixed:
  `ByVal`/`ByRef` are now recognized and discarded (this VM already treats every parameter as
  effectively by-value, with no `ByRef` write-back modeled for *any* parameter, so discarding
  the keyword is correct, not a simplification); `Optional`/`ParamArray` are not implemented and
  now fail with a clear parse error instead of the same silent misparse — a deliberate behavior
  change, not a regression, for any macro that happened to declare one of these keywords before.

## [0.6.0]

Root `elixcee` (Rust crate + Python package) only — `elixcee-types` stays `0.2.0` (no
public-surface change this round: the new array lower-bound tracking lives in a `Vm`-side
`HashMap`, not on `Variant::Array` itself, so nothing semver-relevant moved), `elixcee-wasm`
stays `0.1.0` (no source changes to `crates/elixcee-wasm/src`; its vendored build output was
regenerated to pick up the `src/reader.rs` fix below), and `@elixcee/xlsx` stays
`0.0.0-development`/unpublished/`"private": true` (none touched, though its vendored WASM
artifact was refreshed too). Four independent items, in the order they were authorized: real
VBA `Err` object semantics; `compat/vba-semantics`'s 386-case suite and `compat/corpus`'s
581-scenario suite wired into a new `compat-vba` CI job (previously local-only); four of five
disclosed array-declaration/bounds gaps fixed, the fifth (`UBound`'s dimension argument,
needing real multi-dimensional array storage) deliberately deferred with its own disclosure
corrected after its registry entry turned out to make a false claim; and `src/reader.rs`'s
`xml:space="preserve"` whitespace-trimming defect on `t="str"` cells, fixed
(`compat/differential:read` 30 MATCH + 3 disclosed → 33/33 MATCH). Full detail in the four
sections below. Plus one bug found during this round's own pre-release verification, not part
of the four authorized items: the Python `elixcee.load_workbook()` binding panicked with
`"active sheet must exist"` on any sheet named the way Excel itself defaults to naming one
(`"Sheet1"`, capital S) — a hand-rolled duplicate of the CLI's sheet-population loop had never
picked up a mixed-case-sheet-name fix the CLI path got back in July. Fixed by routing through
the same already-tested helper the CLI uses instead of maintaining a second copy. Pre-existing
since that July commit, unrelated to any 0.6.0-phase work. `cargo test --workspace` 872/872,
`compat/vba-semantics` 386 cases (0 `BUG`/0 `UNCLASSIFIED`, 16 `KNOWN_LIMITATION` — down from
19), `compat/corpus` 581 scenarios (0 `UNEXPLAINED`/0 `MISMATCH`) all green as of this bump;
every real GitHub Actions job (including the new `compat-vba` job's first real run) green on
`master` before this bump.

### `Err` object: `Err.Number` / `Err.Description` / `Err.Clear` / `Err.Raise`

First item of the 0.6.0 phase. `On Error Resume Next`/`On Error GoTo <label>` already
existed but had no way for the running macro to inspect *what* error was caught — the
single most common real-world idiom this blocked was
`On Error Resume Next : <risky op> : On Error GoTo 0 : If Err.Number <> 0 Then ...`.

- **New `Err.Number`/`Err.Description` expressions**, `Err.Clear`/`Err.Raise` statements
  (`src/parser/ast.rs`, `src/parser/mod.rs`, `src/vm/mod.rs`). Parser recognition is
  guarded on the exact member name (`Err.Number`/`Err.Description`/`Err.Clear`/`Err.Raise`
  specifically), matching the existing `ThisWorkbook`/`ActiveWorkbook` precedent — a
  genuine user variable named `err` with an unrelated field (`err.code = 1`) still parses
  as ordinary assignment/field access, unaffected (test:
  `a_bare_err_variable_is_unaffected_by_err_object_parsing`).
- **`Vm::err_number`/`err_description`** are set at both existing error-catch sites
  (`On Error Resume Next`'s per-statement catch, `On Error GoTo <label>`'s jump) via a new
  `classify_vba_error_number(msg: &str) -> (i64, String)`. Only maps a handful of
  elixcee-internal message strings that are **confirmed exact matches against Microsoft's
  own long-stable, publicly documented VBA runtime error constants** (unchanged since
  VB6 — a fact independent of this project's lack of a live Excel/VBA oracle, see
  `ROADMAP.md`'s Known gap #1): Division by zero → 11, Subscript out of range → 9, Type
  mismatch → 13, Invalid procedure call or argument → 5, Invalid use of Null → 94, Object
  variable or With block variable not set → 91. Everything else elixcee itself raises
  (undefined variable, sheet/sub/workbook not found, etc.) defaults to 1004
  ("Application-defined or object-defined error", real VBA's own generic catch-all for
  Excel-object-related failures) — a disclosed default, not independently confirmed per
  condition. Several of those (calling an undefined Sub/Function, in particular) would
  actually be a *compile*-time failure in real VBA, never reaching `On Error` at runtime
  at all — a known, disclosed divergence, not fixed here.
- **`Err.Raise Number[, Source][, Description]`** parses real VBA's full positional-slot
  grammar, including the idiomatic `Err.Raise 513, , "custom text"` form that skips
  `Source` — a naive two-positional-argument implementation would misread that
  `"custom text"` as `Source` instead of `Description`, since real VBA's slot order is
  (Number, Source, Description, HelpFile, HelpContext), not (Number, Description).
  `Source` is parsed (so this can't happen) but not modeled as a readable property —
  `Err.Source` doesn't exist here, matching this project's existing choice not to model a
  VBA project/module naming concept elsewhere. `HelpFile`/`HelpContext` aren't parsed at
  all. Raising without an explicit `Description` fills in the real VBA description text
  for the numbers above, or the 1004 catch-all text otherwise.
- `Err.Number`/`Err.Description` reset to `0`/`""` at the start of each `run_sub`/
  `run_sub_multi` call and on `Err.Clear`. Deliberately does **not** auto-clear on `On
  Error GoTo 0`/a fresh `On Error` statement — the common idiom above relies on
  `Err.Number` surviving past `On Error GoTo 0` to be inspectable at all, and the exact
  real-VBA clearing rule around `Resume`/`On Error` re-statements wasn't independently
  confirmed, so this stays conservative rather than guessing.
- 16 new tests (10 `src/vm/mod.rs`, 6 `src/parser/mod.rs` covering AST shape, the
  `Err.Raise`-skips-`Source` case at both layers, and a regression test confirming
  `pending_raised_error` — the side channel `Err.Raise` uses to preserve its own
  number/description across the generic error-classification path — can't leak into
  an unrelated later error: it's consumed synchronously by the first `On Error Resume
  Next`/`GoTo` catch on the same unwind, before any other statement can run) —
  `cargo test --workspace` 857/857
  (857 = 741 lib + 1 + 15 + 16 + 5 + 14 + 7 + 7 + 17 + 15 + 19, every test binary summed),
  no regressions in the 581-scenario corpus classifier or the 386-case `vba-semantics`
  suite (still 0 `BUG`/0 `UNCLASSIFIED`, same 19 `KNOWN_LIMITATION`).
- **Known limitation found while verifying the above, not fixed here (pre-existing,
  unrelated to `Err.Raise` specifically — reproduces with any error, e.g. `1 / 0`):**
  `On Error GoTo <label>` set in a caller does not run its handler if the error instead
  occurs inside a *called* Sub/Function that has no `On Error` of its own. Root cause:
  `on_error_goto_label` is a single `Vm`-wide field with no per-call-frame scoping, so
  the callee's own `exec_body` sees the caller's still-armed label, tries (and fails) to
  find it among the callee's own statements, and returns a synthetic "label not found"
  error instead of ever reaching the caller's handler. Needs a real design decision
  (save/restore `on_error_goto_label` — and likely `on_error_resume_next` — per call
  frame) rather than a local patch, so it's deliberately left for a dedicated fix rather
  than folded into this feature commit.

### Array declaration/bounds gaps: `Dim arr(lo To hi)`, `Option Base 1`, `Dim arr()`, `Erase`

Fixes four of the five `array_bounds` `KNOWN_LIMITATION` cases in `compat/vba-semantics/`
(19 → 16 — see that suite's own CHANGELOG-adjacent note below for the fifth, newly
disclosed rather than fixed):

- **`Dim arr(2 To 8)`** — an explicit non-zero lower bound — now parses (`ArrayDim { lower:
  Option<Expr>, upper: Expr }` replaces a bare `Expr` per dimension in `DimArray`/`ReDim`'s
  AST) and is honored: `LBound(arr)` is `2`, and `arr(2)`/`arr(8)` address the real first/
  last elements. `Option Base 1` — previously parsed and silently discarded at module level
  — now sets the default lower bound for declarators that don't give an explicit `lo To hi`
  (`Program.option_base`, read by `Dim`/`ReDim` at execution time). Storage stays a flat
  `Vec<Variant>` (`elixcee-types`' public `Variant::Array` is untouched — no semver bump):
  the lower bound is tracked separately, per array *variable name*, in a new `Vm`-side
  `array_lower_bounds` map (`LBound`/`UBound`/`ArrayWrite`/array-subscript reads all resolve
  arrays by name already, so this needed no public-surface change). An array value with no
  name to key on — `Split()`/`Array()`'s return, or any array-valued expression not bound to
  a `Dim`'d variable — defaults to lower bound 0, unchanged from before.
- **`Dim arr()`** (empty parens, a dynamic array sized later by `ReDim`) now parses — the
  declarator's dimension list is simply empty — and creates an unsized placeholder array
  ReDim can then legally resize, matching the one documented use this suite tests (`Dim
  arr()` immediately followed by `ReDim arr(5)`). elixcee doesn't model the stricter real-
  VBA behavior of raising "Subscript out of range" if `UBound`/an element is accessed
  *before* the first `ReDim` — not exercised by any case, not attempted.
- **`Erase arr`** — verified (checked the pre-change `parse_ident_stmt`, not inferred from
  the old registry entry's "IsEmpty is still False" description) to have had no `Erase`
  statement dispatch at all: `erase` wasn't a recognized keyword, so `erase arr` fell all
  the way through to the generic "bare identifier statement" fallback and became a
  `Stmt::Unsupported` no-op. Is now a real `Stmt::Erase { name }`: resets every element of a
  fixed-size array back to `Empty` in place, leaving its bounds untouched (matching real
  VBA's documented behavior for a statically-declared array). Real VBA's comma-separated
  `Erase a, b` form isn't parsed — no case needs it.
- `array_oob_error`'s `ArrayIndexOutOfBounds` diagnostic evidence used to hardcode
  `lower: 0` unconditionally; now reports the array's actual lower bound and the VBA-facing
  index that was attempted (not an internal, bound-shifted one), for the two call sites that
  now track a real bound. The two UDT-array call sites (`DimArrayRecord`/`ArrayRecordSet`/
  `ArrayRecordGet` — `Dim arr(10) As MyType`) are unaffected and still report `lower: 0`,
  matching their existing (unchanged) always-0-based behavior.
- **Found while verifying the above, and separately disclosed (not fixed — see below):**
  the fifth `array_bounds` `KNOWN_LIMITATION` case's own description ("UBound(arr, 2)
  ignores its dimension argument ... even though the array's own storage genuinely is
  two-dimensional") turned out to be **factually wrong**. elixcee's array storage is
  genuinely 1-D: `Dim arr(3, 2)` only ever allocates dimension 1's 4 elements (dimension
  2's size is parsed and discarded), and every array write/read (`Stmt::ArrayWrite`, the
  `Expr::FuncCall` array-subscript read path) indexes using only the first index
  expression — a second or later index is silently ignored on *both* sides, so `arr(2,
  0) = 111` followed by `arr(2, 1) = 222` overwrites the same element (confirmed live:
  both `arr(2,0)` and `arr(2,1)` then read back `222`). The suite's own
  `two_dimensional_array_write_and_read_round_trips` case had been cited as evidence 2-D
  addressing worked — it passed only because its single write and single read happened to
  use the *same* second index on both sides, a coincidental round-trip that never exercised
  the collision. Renamed to `two_dimensional_array_second_index_is_silently_dropped`,
  reshaped to actually discriminate (write two elements differing only in the second
  index, confirm the first wasn't clobbered — it currently is), and registered as a new,
  previously-undisclosed `KNOWN_LIMITATION`. `ubound_second_dimension_argument_ignored`'s
  own `knownLimitation` text is corrected to stop citing the now-corrected sibling case as
  proof of working 2-D storage. Fixing this for real needs shape metadata and stride
  arithmetic in both the write and read paths — comparable in scope to this project's other
  deferred Variant-surface work (see the Date/Time note elsewhere in this file) — and was
  deliberately not attempted alongside the smaller, independent lower-bound-tracking fixes
  above.
- 13 net new tests (8 `src/vm/mod.rs`; `src/parser/mod.rs` net +5 — 6 added, replacing the
  now-stale `test_option_base_ignored` with two narrower tests that assert the captured
  value instead of just "didn't break parsing") — `cargo test --workspace`
  870/870 (870 = 754 lib + 1 + 15 + 16 + 5 + 14 + 7 + 7 + 17 + 15 + 19, every test binary
  summed), `cargo clippy -p elixcee-types --all-targets -- -D warnings` clean, `cargo build
  --release --workspace` clean, `cargo check --features python --lib` clean (only the
  pre-existing, disclosed pyclass deprecation warning). `compat/corpus/` unaffected (581
  scenarios, 0 `UNEXPLAINED`/0 `MISMATCH`). `compat/vba-semantics/`: 386 cases, 0 `BUG`/
  0 `UNCLASSIFIED`, 16 `KNOWN_LIMITATION` (down from 19: four fixed, one newly disclosed
  as described above) — see `compat/vba-semantics/README.md`'s "Current state" for the
  full breakdown.

### `src/reader.rs`'s `xml:space="preserve"` whitespace defect on `t="str"` cells

`xlsx_sheet_cells`'s `<v>`-text handler called `xlsx_parse_cell(text.trim(), ...)`
unconditionally — for a `t="str"` cell (whose literal text lives directly in `<v>`, unlike
`t="s"` shared-string cells or inline `<is><t>` strings, neither of which this call site
ever trims), that silently dropped significant leading/trailing whitespace whenever the
source XML marked it with `xml:space="preserve"`. Confirmed live against
`compat/corpus/workbooks/with_text.xlsx`'s own raw `sheet1.xml`: cell A3 is `<c t="str"><v
xml:space="preserve">  padded  </v></c>`, read back as `"padded"` instead of `"  padded
  "`. Disclosed since the round that found it via `compat/differential/classify.mjs`'s
`UNSUPPORTED_ALLOWLIST` (`XML_SPACE_PRESERVE_DEFECT`, registered under both `read` and
`readFile`); reachable through `read()`/`readFile()`/`readFileSync()` alike, since the
latter two are thin wrappers over `read()`.

Fix: `xlsx_sheet_cells` now reads `<v>`'s own `xml:space` attribute (a new `v_preserve_space`
local, re-read fresh on every `<v>` open — no stale carry-over from a previous cell in the
same row) and skips the trim when it's `"preserve"`, matching plain XML `xml:space`
semantics rather than special-casing `t="str"` specifically. Real Excel/SheetJS writers
never emit this attribute on a numeric/boolean `<v>` (whitespace is never meaningful there),
so this doesn't change default behavior for any realistic file — a regression test confirms
a numeric `<v xml:space="preserve">42</v>` still parses even though `f64::parse` itself
rejects surrounding whitespace, which the fix's unconditional (not `t`-gated) skip could in
principle have broken for a pathological input.

Both `UNSUPPORTED_ALLOWLIST` entries (`with_text.xlsx:xml_space_preserve_trimmed` under
`read` and `readFile`) are removed — the allowlist is empty again — and the now-dead
`unsupportedCaseId` plumbing threading them through
`compat/differential/xlsx-read.test.mjs`'s `with_text.xlsx` fixture cases is dropped too,
matching this project's established precedent for closing a disclosed reader defect (see
this same file's `classify.mjs` comment history). `differential:read`: 33/33 MATCH, 0
disclosed (was 30 MATCH + 3 disclosed). Vendored WASM artifact
(`packages/xlsx/src/internal/wasm/`) rebuilt via `crates/elixcee-wasm/build.sh` so
`@elixcee/xlsx`'s `read()` actually carries the fix — `wasm:smoke` and
`differential:utils`/`:ssf-format`/`:metadata` all still clean, confirming no regression
from the rebuild itself.

2 new tests in `src/reader.rs` — `cargo test --workspace` 872/872 (872 = 756 lib + 1 + 15 +
16 + 5 + 14 + 7 + 7 + 17 + 15 + 19, every test binary summed), `cargo clippy -p
elixcee-types --all-targets -- -D warnings` clean, `cargo build --release --workspace`
clean. `compat/corpus/` and `compat/vba-semantics/` unaffected (verified by re-running both
after the reader.rs change — neither exercises this XML shape).

## [0.5.0]

Root `elixcee` (Rust crate + Python package) **and** `elixcee-types` (`0.1.0` → `0.2.0`,
a minor bump, not a patch — see "`elixcee-types` 0.2.0" below for why). `elixcee-wasm`
stays `0.1.0` (no source changes; its vendored build output was regenerated to pick up the
VM changes below, and its own `elixcee-types`/`elixcee` path dependencies carry no version
requirement to update), and `@elixcee/xlsx` stays
`0.0.0-development`/unpublished/`"private": true` (none touched). Built via two parallel,
disjoint-scope worktree branches — VBA structural
semantics (parser/VM: colon-statement separator, `Variant::Null` with documented
propagation rules, real object-variable unset/`Nothing` state, a runtime `With` stack) and
`@elixcee/xlsx` consumer/browser validation (a real packed-tarball install smoke, a real
headless-Chrome smoke, bundle-safe WASM loading, `readFile()`) — merged after each was
independently reverified, then integration-regression-tested together as a whole,
surfacing and fixing one real interaction bug neither branch's own tests caught (a bare
`.member` inside a single-line `If` nested in a `With` body). Does not claim Microsoft
Excel validation anywhere. Full detail in `[Unreleased]`'s two sections below (VBA
structural semantics; `@elixcee/xlsx` real-consumer and real-browser validation) and the
single-line-`If`/`With` interaction fix that follows them.

All gates green before this bump: `cargo test --workspace` (724 passing),
`cargo build --release --workspace`, `cargo check --features python --lib`,
`cargo clippy -p elixcee-types -- -D warnings`, a real `maturin build --release` wheel
installed into a fresh venv with `Null`/object-`Nothing`-alias-safety/the `With` stack
re-verified through the actual Python API post-install (not just `cargo check`); the
`compat/vba-semantics/` suite (386 cases, 0 `BUG`/0 `UNCLASSIFIED`, 19 `KNOWN_LIMITATION`,
down from 28, deterministic across 2 runs); the existing 581-scenario `compat/corpus/` suite
(0 `UNEXPLAINED`/0 `MISMATCH`, unchanged); `compat/differential/`'s utils/SSF/read+readFile/
metadata suites (all passing, read+readFile now 30 MATCH + 3 disclosed); a fresh
`wasm-pack` rebuild of both targets verified via the real packed-tarball consumer smoke and
a real headless-Chrome smoke (not just Node simulating the `browser` export condition).

### `elixcee-types` 0.2.0

- **Added `Variant::Null`** to the public `Variant` enum (see "VBA structural semantics"
  below for what it's for). **This is a public-enum-variant addition, not a purely additive
  change** — any downstream consumer doing an exhaustive `match` on `Variant` (rather than
  ending in a `_ =>` catch-all) fails to compile against this version until it adds a
  `Variant::Null` arm. Bumped `0.1.0` → `0.2.0` (a real minor bump, not left at `0.1.0`)
  specifically because of this — `elixcee` `0.5.0` depends on `elixcee-types = "0.2.0"`
  (previously `"0.1.0"`), so a `cargo build`/`cargo publish` of `elixcee` against the old
  `elixcee-types` `0.1.0` on crates.io would fail to resolve `Variant::Null` at all. No
  other public API surface changed.

### VBA structural semantics

Four language-level gaps closed, each verified against Microsoft's own VBA language
reference (fetched live, not recalled) before being encoded as an expectation.
`compat/vba-semantics/` grew **301 → 386 cases**, with `KNOWN_LIMITATION` **28 → 19**
(nine genuinely fixed, annotations removed rather than weakened — never by changing what a
case expects). `compat/corpus/`'s 581-scenario regression baseline stays at 0 `UNEXPLAINED`
/ 0 `MISMATCH`; `cargo test --workspace` passes; `report.json` is byte-identical across two
consecutive runs.

#### Added

- **The `:` multi-statement-per-line separator** — `a = 1: b = 2: c = 3`, `label1: a = 1`,
  `MsgBox "x": Exit Sub`, `For i = 1 To 3: … : Next i`, and a single-line `If`'s own
  `:`-separated Then/Else statement lists (per the If…Then…Else reference: "One or more
  statements separated by colons; executed if condition is True", example
  `If A > 10 Then A = A + 1 : B = B + A : C = C + B`). Handled in the parser via the
  tokenizer's existing `Tok::Colon`, **never** as a pre-tokenize `:`→newline rewrite —
  which would corrupt a colon inside a string literal, break the `label:` form, and mangle
  the single-line `If`'s clause boundary. All three are pinned by tests. Each
  colon-separated statement keeps its own `SourceSpan`, so `--json`'s `location` still
  points at the individual statement, not the line.
- **`Variant::Null`** — VBA's "no valid data" value, now genuinely distinct from `Empty`
  (an uninitialized Variant). Implements the documented rules: arithmetic propagates Null
  from either side (and *before* operand coercion, so `5 / Null` is Null, not a
  Division-by-zero error); `&` propagates only when *both* sides are Null (a single Null is
  a zero-length string); all six comparison operators produce Null, including
  `Null = Null`; `And`/`Or`/`Xor`/`Not` follow their three-valued truth tables, in which
  Null does *not* uniformly propagate (`False And Null` is False, `True Or Null` is True);
  `If Null Then` treats the condition as False (documented, not an error); `IsNull` and
  `IsEmpty` are now separate questions; `TypeName(Null)`/`VarType(Null)` are `"Null"`/`1`;
  and a Null reaching a genuinely numeric context raises error 94, `Invalid use of Null`.
  Adds **no new external surface** — Null serializes exactly as `Empty` already does (JSON
  `null` / Python `None` / blank cell), so `--json`, the Python bindings and the xlsx/ods
  writers are unchanged.
- **`ObjectRef::Nothing`** — a real unset/Nothing state for object variables.
  `Dim r As Range|Worksheet|Workbook` registers the name as declared-but-unset;
  `Set r = Nothing` assigns the null reference (it used to silently no-op); every
  member-access path raises real VBA's error 91 text, `Object variable or With block
  variable not set`, from one shared constant. `Set r = Nothing` clears only `r` — a
  `Set r2 = r` alias made earlier stays live and still reads and writes through to the same
  Range. `<var> Is Nothing` now parses and reflects each variable's own state (only the
  `Is Nothing` shape; a general `a Is b` is still unparsed rather than guessed at).
- **New stable error code `E1007`/`object_variable_not_set`**, documented in
  `docs/agent-contract.md` — a genuinely new error condition, not free-text reuse of an
  existing code.
- **`Array(...)`** builtin — builds a zero-based Variant array from its arguments.

#### Changed

- **`With`-target resolution is now a runtime mechanism, not a parse-time textual
  rewrite.** The target expression is captured unevaluated (`ast::WithTarget`), evaluated
  **once** on block entry, and pushed onto `Vm::with_stack`; a bare `.member` is a
  first-class statement and expression form (`Stmt::WithDot`/`Expr::WithDot`) resolved
  against the innermost entry wherever it appears in the AST. Consequences:
  `With Cells(r, c)` (any computed target) works; a bare `.member` works at any nesting
  depth inside `If`/`For`/`Do`/`Select Case` in the body; reassigning a target variable
  inside the body no longer could (and still cannot) retarget the block; nesting restores
  each outer target in turn; and the stack is popped on *every* exit path, including
  `Exit Sub`/`Exit For` and a runtime error, so a target can't leak into whatever runs
  next. The parser's `with_target`/`with_range_target` fields and the `Stmt::WithRecord`
  variant are gone.
- **`With ws` (a Worksheet-typed object variable) now qualifies `.Cells(r, c)` to that
  worksheet.** It previously wrote to whatever sheet happened to be active — a real,
  previously-undisclosed bug, surfaced by the runtime-stack work.
- **`For Each c In Range(...)` binds the loop variable as a live single-cell Range object**
  as well as a plain value, so `c.Value` reads that cell. It previously fell through to the
  UDT path and silently yielded `Empty`. Found by `compat/corpus/` reacting to the
  `Dim c As Range` change above, not by source audit.
- **A non-numeric string operand of `+`/`-`/`*`/`/`/`^` raises `Type mismatch`**, real
  VBA's documented wording ("One expression is a numeric data type and the other is a
  String | A `Type mismatch` error occurs"). Applied narrowly, via a new `arith_to_f64`
  wrapper used only by `eval_binop`'s `Add|Sub|Mul|Div|Pow` arm — the shared `to_f64`
  helper and its ~53 other call sites, each with its own correct wording for its own
  failure, are untouched. That blast radius was the exact reason this stayed disclosed
  rather than fixed when it was first found. **Not** extended to `\`/`Mod`, which go
  through `to_i64_rounded` and keep the previous wording; the rule cited above is from the
  `+` operator page, and widening it further would re-enter the blast radius this fix was
  scoped to avoid.
- `Dim x: x = 5` now parses — the declarator's trailing-syntax tolerance loop was swallowing
  the `:` separator. Found by a new suite case, not by source audit.
- **A bare `.member` branch inside a single-line `If` nested in a `With` body now runs.**
  `parse_stmt` gained a `Tok::Dot` arm for the runtime With-stack work above, but
  `parse_single_line_if_branch`'s own dispatch checked only `Tok::Ident` and was never
  updated to match — so `If .Value > 0 Then .Value = .Value + 1` inside `With Range("A1")`
  silently degraded to `Stmt::Unsupported` (no parse error, but the assignment never ran).
  Same bug *class* as the pre-existing `Range()`/`Cells()`-in-single-line-`If` fix (a
  single-line-`If` branch dispatch lagging behind block-form `parse_stmt`'s own statement
  coverage) — found during integration by manually exercising a README code sample, not by
  either subagent's own test suite, which is exactly the kind of interaction gap that can
  slip between two disjoint-scope changes that never ran against each other until merged.

### `@elixcee/xlsx` — real-consumer and real-browser validation

Closes the gap between "the differential suites pass" and "a real npm consumer, and a real
browser, actually works" — every prior check reached the package via a relative import into
`packages/xlsx/src`, or (for the `"browser"` export condition) via Node simulating that
condition, never an actual browser process.

#### Added

- **`XLSX.readFile()`/`readFileSync()`** — one function under both names (matching the real
  `xlsx` package's own identity: same `.name`, `.length`, key order), wrapping the existing
  byte-buffer `read()`. Differential-tested file-by-file against the real `xlsx@0.18.5`
  oracle, with and without `cellStyles`/`cellDates`. Throws `ELIXCEE_UNSUPPORTED_IN_BROWSER`
  from the browser entry point rather than faking a filesystem. `write*` remains
  unimplemented.
- **A packed-tarball consumer smoke test** (`packages/xlsx/scripts/pack-consumer-smoke.mjs`,
  `npm run pack:consumer`) — runs a real `npm pack`, `npm install`s the exact `.tgz` into a
  throwaway package under `os.tmpdir()`, and exercises `require()`, `import`, a TypeScript
  compile, `XLSX.read()`, CJS/ESM export-set identity, and the `"browser"` export condition
  entirely from inside that install — asserting the resolved paths land under the throwaway
  `node_modules/@elixcee/xlsx`, not a relative path back into this repo. Every earlier check
  in this project could have passed while the actual published tarball was broken; this one
  can't.
- **A real headless-browser smoke test** (`packages/xlsx/scripts/browser-smoke.mjs`,
  `npm run browser:smoke`) — launches an actual local Chrome/Chromium process (via its own
  `--dump-dom`, no browser-driver dependency added — evaluated and rejected
  playwright-core/puppeteer-core/chrome-remote-interface as unnecessary weight for "load one
  page, read one result"), serves an esbuild browser bundle over real `node:http`, and reads
  `XLSX.read()`'s result back out of the page's own DOM: sheet names, a real cell value, an
  exported-function count, zero page-observable console/uncaught errors, zero non-200
  responses for any page-referenced resource. **Distinct from, and strictly more than, the
  pre-existing `node --conditions=browser` check** (still present, in `wasm:smoke`) — that
  one is Node simulating an export condition and proves nothing about a browser; this one is
  a real browser. Neither is described as the other anywhere in code, CI step names, or this
  entry. Safari is not covered and not claimed.
- **CI**: the packed-tarball smoke joins `node-js` (both Node versions); the real-Chrome
  smoke and a CJS *and* ESM esbuild-bundle smoke (as distinct steps) join `wasm`, along with
  a diagnostic step that prints whatever browser the runner image actually provides, so a
  missing-Chrome failure is self-explanatory from the job's own log rather than a guess.
  `compat/differential/`'s own `classify.mjs`/`normalize.mjs` self-checks — existing package
  scripts that pin the exact contents of the disclosed-divergence registries — are now
  wired into CI too; they never ran there before.

#### Fixed

- **The Node/CJS WASM loader's `.wasm` lookup is no longer `__dirname`-relative.**
  `elixcee_wasm.node.cjs` (wasm-pack's own generated code, not hand-written) located its
  compiled WASM via a path relative to its own file location — bundle-*output*-relative once
  a consumer bundled it, not source-relative. ESM bundle output has no `__dirname` at all (a
  hard `ReferenceError`, not a silent failure); CJS bundle output technically had `__dirname`
  but pointed at the wrong directory, so it only worked if the consumer manually copied
  `elixcee_wasm_bg.wasm` next to their bundle. Fixed by inlining the compiled WASM as base64
  directly into the Node loader too (`crates/elixcee-wasm/build-node-inline.mjs`, mirroring
  the technique `build-browser-inline.mjs` already used for the browser build) — generated by
  `build.sh`, never hand-patched, so a fresh rebuild reproduces the committed artifact
  byte-for-byte. No `.wasm`-copy step is required for CJS *or* ESM bundling anymore, and
  browser bundling — previously broken outright (`esbuild --platform=browser` failed
  resolving `fs`) — now works too. The raw `elixcee_wasm_bg.wasm` is no longer vendored
  separately (both loaders already carry the bytes; shipping it too would double-ship
  263 KB), and `package.json` gained a `browser` field stubbing the Node loader out of
  browser bundles, so a browser consumer pays for the WASM payload once, not twice.
  Synchronous `read()` is unaffected — no `await init()` introduced anywhere.
  **Package-size impact**, measured against 0.4.0: packed tarball 339,098 → 380,005 bytes
  (+12.1%), unpacked 741,304 → 835,712 (+12.7%); the WASM payload itself is unchanged at
  263 KB (only its containers, base64-inlined, grew). No hard size gate — recorded so a
  future round can judge whether it grows *further* without basis.

#### Discovered, disclosed (not fixed — outside this round's scope)

- **`src/reader.rs` trims every cell's text unconditionally**
  (`xlsx_parse_cell(text.trim(), …)`), ignoring the `xml:space="preserve"` attribute real
  XLSX XML uses to mark significant leading/trailing whitespace — a cell whose real value is
  `"  padded  "` reads back as `"padded"`. Confirmed live against
  `compat/corpus/workbooks/with_text.xlsx` cell A3 (oracle: `"  padded  "`; elixcee:
  `"padded"`). Reachable through both `read()` and `readFile()`. Registered in
  `compat/differential/classify.mjs`'s `UNSUPPORTED_ALLOWLIST` (3 cases, one root cause) with
  a full writeup rather than silently excluded from the fixture set — the classifier's own
  self-check pins the exact entry count, so it can't go stale unnoticed. Fixing it means
  honoring `xml:space` on the `<t>` element rather than trimming at the call site, and
  re-checking the trim isn't load-bearing for the numeric/boolean paths sharing
  `xlsx_parse_cell` — `src/reader.rs` is shared surface, not `@elixcee/xlsx`-specific, so
  this is recorded for whoever next touches the reader, not fixed here.

## [0.4.0]

Root `elixcee` (Rust crate + Python package) only — `elixcee-types` and `elixcee-wasm`
both stay `0.1.0` (no source changes this release; `elixcee-wasm`'s vendored build output
was regenerated to pick up the fixes below, but its own `src/` wasn't touched) and
`@elixcee/xlsx` stays `0.0.0-development`/unpublished. Covers this round's `compat/vba-
semantics/` expansion (208 → 301 cases, 0 `BUG`/0 `UNCLASSIFIED`, 28 disclosed
`KNOWN_LIMITATION`s — see that entry below and `compat/vba-semantics/README.md` for the
full breakdown), the new `wasm` CI job, and several real behavior-changing bug fixes found
while building the suite (Boolean arithmetic, `WorksheetFunction` numeric coercion, `Empty`
equality, single-line-`If` statement recognition) — all gates green before this bump:
`cargo test --workspace` (683 passing), `cargo build --release --workspace`,
`cargo check --features python --lib`, `cargo clippy -p elixcee-types -- -D warnings`, the
existing 581-scenario `compat/corpus/` suite (0 `UNEXPLAINED`/0 `MISMATCH`, unchanged), and
the new `wasm`/existing `node-js` CI jobs both confirmed green on real GitHub Actions, not
just locally.

### Added

- **`compat/vba-semantics/`, a new VBA value-correctness suite** — a genuinely different
  question from `compat/corpus/`'s own "does elixcee run without erroring": is the VALUE
  elixcee produces the one real, documented VBA semantics says it should be. Needs no
  oracle — `reference/*.mjs` are small, independently-checkable pure-JS reference
  implementations of documented real VBA semantics (banker's rounding, `Str()`'s
  leading-space quirk, `Val()`'s prefix parsing, `And`/`Or`/`Xor`/`Not`'s logical-vs-bitwise
  split, ...), used to compute cases' expected outcomes programmatically. Six-
  verdict classification (`MATCH_DOCUMENTED_SEMANTICS`/`EXPECTED_ERROR`/`NONDETERMINISTIC`/
  `KNOWN_LIMITATION`/`BUG`/`UNCLASSIFIED`); `BUG`/`UNCLASSIFIED` both gate at 0. Started at
  208 cases across 12 categories; grew to **301 cases across 18** in the same round that
  added the `+`-vs-`&` operator-coercion, comparison-operator-coercion, `Select Case`
  matching, `With`-block-resolution, and array-bounds categories — each expected value
  sourced from Microsoft's own VBA language reference, fetched live while writing the
  cases, not recalled from memory. Current state: 253 `MATCH_DOCUMENTED_SEMANTICS` + 18
  `EXPECTED_ERROR` + 2 `NONDETERMINISTIC` + 28 `KNOWN_LIMITATION` = 301, 0 `BUG`,
  0 `UNCLASSIFIED`. All 28 `KNOWN_LIMITATION` cases are divergences found while building
  this suite and not fixed this round (several *other* divergences found the same way
  *were* fixed — see "Fixed" below); grouped by root cause with the full breakdown in
  `compat/vba-semantics/README.md`.
- **CI now runs `@elixcee/xlsx`'s own tests.** `.github/workflows/ci.yml` gained a `node-js`
  job (Node 20/22 matrix): `packages/xlsx`'s TypeScript typecheck (with and without the DOM
  lib present) and all four `compat/differential/` suites (`utils`/`ssf-format`/`read`/
  `metadata`). Previously none of this ran anywhere except a developer's own machine, despite
  every command already working — verified live before wiring each one in, not assumed from
  this file's own previously-claimed numbers.
- **CI also now builds and smoke-tests the WASM bridge from scratch.** A new `wasm` job runs
  both `wasm-pack build --target nodejs` and `--target web` fresh (the `node-js` job above
  only ever consumed the already-vendored/committed copy — a build-breaking change to
  `crates/elixcee-wasm`/`src/reader.rs` had no CI signal until now), then runs the new
  `packages/xlsx/scripts/wasm-smoke.mjs`: a Node synchronous `read()` call, the `"browser"`
  export condition resolving *and actually running* (via `node --conditions=browser`,
  self-referencing the package by name — more than a resolution check, but still Node
  simulating the condition; no real browser executes anywhere in this project's CI, and no
  Safari support is claimed anywhere), a minimal `esbuild` bundle with an in-bundle
  `XLSX.read()` call, and the current WASM binary size logged (263,204 bytes as of this
  round) — recorded, not gated against any threshold (no prior baseline exists to compare
  against). `esbuild` is `packages/xlsx`'s one new devDependency for this (pinned to `^0.28`,
  past the version with the known dev-server CORS advisory — irrelevant to this project's
  usage, which only ever calls its one-shot `build()`, never `serve()`, but avoided anyway).
  One real, previously-undisclosed consumer caveat found while writing the bundle check: the
  Node/CJS WASM loader (`elixcee_wasm.node.cjs`, wasm-pack's own generated code, not
  hand-written) locates its `.wasm` file via a `__dirname`-relative path, which becomes
  bundle-output-relative once bundled — a consumer bundling this package's Node entry needs
  to bundle to CJS (ESM output has no `__dirname` at all, a hard `ReferenceError`) and copy
  `elixcee_wasm_bg.wasm` next to their bundle output, or externalize the loader. Not fixed
  this round (would mean patching wasm-pack's own generated boilerplate); documented in
  `wasm-smoke.mjs`'s header comment and `ROADMAP.md`.
- **`packages/xlsx/scripts/audit-pack-contents.mjs`**, also wired into the new `node-js` CI
  job — asserts what `npm pack` would actually publish (every required file present —
  `LICENSE`, `README.md`, `THIRD_PARTY_NOTICES.md`, the four public entry points; nothing
  forbidden — `node_modules/`, `test/`, `.git`, `tsconfig*`; nothing unexpected under
  `src/internal/`), checked against `npm pack --dry-run --json`'s own real file list, not a
  reimplementation of npm's inclusion rules. Didn't exist at all before — a manual dry-run
  was clean (17 files, 338.8 kB), but nothing asserted this in CI.
- **`packages/xlsx/README.md`** — a package-level README, previously absent (npm's registry
  page would have shown only the `description` field, which opened with an unqualified
  "Drop-in replacement for xlsx" and never disclosed `write*`/`readFile` are unimplemented).
  States current scope honestly: what's implemented (all 33 `utils.*` exports, `SSF`,
  `XLSX.read()`, each with its own differential-testing numbers), what isn't
  (`write*`/`readFile`), and points to `THIRD_PARTY_NOTICES.md`/`docs/compatibility-known-
  defects.md` for licensing and disclosed divergences. `description` in `package.json`
  updated to match (no longer opens with an unqualified drop-in-replacement claim).
  Confirmed via `npm pack --dry-run` that the new README is actually included in the
  tarball (npm does this automatically regardless of the `files` array). This closes one of
  three concrete `packages/xlsx` alpha-publish blockers found this round — the other two
  (`"private": true`, missing `publishConfig.access`) are a real publishability policy
  decision, deliberately left alone here, not a mechanical fix.

### Fixed

- Array out-of-bounds errors used elixcee's own diagnostic wording (`"Array 'arr': index N
  out of bounds (len=N)"`) instead of real VBA's actual runtime error 9 message,
  `"Subscript out of range"` — found and disclosed as a `compat/vba-semantics/`
  `KNOWN_LIMITATION` when that suite first ran, fixed in the same round rather than left
  registered. Safe to change: `docs/agent-contract.md` already documents `message` as free
  text, not a stable/matchable field (`code`/`kind` are); `diagnose`/`diagnose-workbook`
  already read the rich per-failure detail (array name, index, bounds) from a structured
  `ResolutionFailureKind` side channel set alongside this string, not by parsing it — so
  nothing that actually depends on the old wording broke.
- All 3 READMEs' "XLSX.read()" section still claimed the browser-target WASM artifact
  "isn't wired into the package's public API yet" — true as of Phase 2B, but Phase 2C
  (already shipped in 0.3.0) added exactly that wiring (a `"browser"` export condition).
  Found while writing `packages/xlsx/README.md` (see "Added" above) and cross-checked
  against this file's own Phase 2C entry, which already correctly described the fix — only
  the top-level READMEs had gone stale. Corrected to state the real remaining caveat: the
  browser entry point assumes bundled consumption (its shared code has a CJS
  `require('ssf')`), not that it's unwired.
- `Dim x` (and `Dim x As <builtin type>`) was a complete no-op — the variable name was
  never recorded at all, so `IsEmpty(x)`/`x + 5`/any read before assignment hit "Undefined
  variable" instead of real VBA's `Empty`. An extremely common real-world VBA idiom
  (`Dim x`, then `If IsEmpty(x) Then ...` before ever assigning it), found on the very
  first run of the new `compat/vba-semantics/` suite (see "Added" above), not by
  source-code audit. `x` now registers as a real `Empty`-valued variable when `Dim`'d.
- `Val()` required its argument to parse as a number in its *entirety* — `Val("123abc")`
  was `0`, not real VBA's `123`. Real VBA's `Val()` parses a leading numeric prefix and
  stops at the first character that doesn't fit, only returning `0` when there's no valid
  numeric prefix at all. Found while designing the new `compat/vba-semantics/` value-
  correctness test suite — the same "never independently verified against documented
  semantics" bug class
  as `IsNumeric`. Scoped to the core grammar (optional sign, digits, one decimal point);
  real VBA's documented embedded-whitespace-stripping inside the numeric prefix
  (`Val("1 2 3")` == `123`) isn't attempted — no evidence it's needed.
- `Str()` was grouped with `CStr()` and shared its implementation — but real VBA's `Str()`
  reserves a leading space for the sign position on a non-negative number (`Str(459)` is
  `" 459"`, not `"459"`), a real, documented behavior difference from `CStr(459)` == `"459"`,
  not an alias of it. Found in the same systematic pass as `IsNumeric` below. Now its own
  arm, scoped to numeric inputs (the only case `Str()` is documented for); anything else
  falls back to the same plain formatting `CStr` uses. Previously untested; now covered.
- `IsNumeric` only checked whether its argument was already an `Integer`/`Float` Variant —
  `IsNumeric("123")` was `False`, missing real VBA's numeric-string recognition entirely.
  Found by a systematic pass over `eval_vba_func` for the same "grouped/never independently
  tested" bug class as `CBool`/`CInt`/`CLng` above. Now also accepts a string that parses as
  a plain decimal/scientific-notation number (after trimming whitespace) and `Empty`
  (coerces to 0 in a numeric context, matching real VBA). Deliberately not chasing VBA's
  fuller numeric-string grammar (currency symbols, locale-specific decimal separators,
  parenthesized negatives) — no evidence any of that is needed, and this project doesn't
  guess at locale-specific parsing rules. Previously entirely untested; now covered.
- `CInt`/`CLng` used Rust's default round-half-away-from-zero (`f64::round()`) instead of
  real VBA's banker's rounding (round-half-to-even) — `CInt(0.5)` was `1`, not `0`. Found by
  auditing for the same bug class the `Round()` fix (below) had already been fixed for:
  `to_i64_rounded` (used by `\`/`Mod` operand coercion) already documented "the same
  round-half-to-even ... that CLng/Round use," but `CInt`/`CLng`'s own arm never actually
  used it. Now reuses that exact existing helper — the `test_vba_clng` test had silently
  computed a tie-case value (`CLng(-2.5)`) without ever asserting on it, which is likely how
  this went unnoticed; that assertion is filled in now, plus dedicated tie-case coverage.
- `Round(number, negativeDigits)` (e.g. `Round(1234.5, -2)`) silently returned a plausible
  answer instead of erroring — real VBA's own `Round()` raises "Invalid procedure call or
  argument" for a negative digit count (unlike `WorksheetFunction.Round`/Excel's `ROUND()`
  formula, which both accept negative digits to round left of the decimal point). Found and
  disclosed, not fixed, in the 0.3.0 round; fixed now.
- `Now`/`Date`/`Time` returned a Rust debug-formatted `SystemTime { tv_sec: ..., tv_nsec:
  ... }` string regardless of which of the three was called — visibly wrong if ever
  displayed or compared, not just imprecise. `Date()` now returns a real `Variant::Date`
  matching the actual system clock (Excel-serial epoch math, same `25569` constant the
  formula engine's own `NOW()` already uses); `Time()`/`Now()` return a numerically correct
  `Variant::Float` (0.0-1.0 for `Time()`, serial-plus-fraction for `Now()`) rather than a
  `Variant::Date`, since `Variant::Date` is whole-day-only (`i64`) in this codebase and
  can't carry a sub-day component without a shared-type change — so `TypeName(Time())`/
  `TypeName(Now())` report `"Double"`, not real VBA's `"Date"`. A disclosed, narrower gap
  than the debug-string bug, not a silent one.
- The bare no-parens form (`Date` without `()`, real VBA allows omitting `()` on zero-arg
  functions) didn't parse as a function call at all — `Expr::Var("date")` always hit
  "Undefined variable". Found alongside the fix above, fixed in the same round: a bare
  identifier now falls back to calling `Date`/`Now`/`Time` as zero-arg functions only after
  every other variable/constant lookup fails — scoped to exactly these three names (the only
  `eval_vba_func` entries that accept zero arguments) rather than a general "any unrecognized
  identifier might be a function call" rule, so a genuine variable-name typo still errors the
  same way it always did (verified with a regression test).
- `Range(...)`/`Cells(...)`/`MsgBox`/etc. weren't recognized inside a single-line `If`'s
  Then/Else branch — only identifier-led statements were, so `If x > 0 Then Exit Sub Else
  Range("A1").Value = 1` mis-parsed the Else branch as an array write to a variable
  literally named "range", failing with "Cannot convert 'A1' to number". Found by
  `compat/vba-semantics/` on exactly this shape, not by source audit. Fixed by extracting
  the full statement dispatch (previously duplicated as a narrower subset for single-line
  `If`) into one shared function used by both the block-form and single-line-`If` parsers.
  That extraction briefly regressed assignments to a variable literally named after a block
  keyword (`do = 0`, `select = 1`, ...) — caught by the existing property test before
  shipping, fixed by re-ordering the "bare `name = ...` is always assignment" check ahead
  of the block-construct keyword dispatch.
- VBA's `+`/comparison operators coerced Boolean `True` to `1.0` instead of VBA's own
  documented internal value of `-1` (`CInt(True)` is `-1` in real VBA) — `True + 5` was `6`,
  not the documented `4`. Found via `compat/vba-semantics/`'s operator-coercion matrix,
  fetched from Microsoft's own VBA language reference rather than recalled from memory.
  Fixing this then silently changed `WorksheetFunction.Sum`/`Max`/`Min`/`Average`/`SumIf`/
  `Round`/`Abs`/`Sqrt`/`Power`/`Log`/`Index` too (`WorksheetFunction.Sum(True, True)` went
  from `2` to `-2`) — wrong, since `Application.WorksheetFunction` bridges into Excel's own
  calculation engine and must keep using Excel's `TRUE = 1` coercion (matching a worksheet
  formula), not VBA's own arithmetic rule. Caught in the same round by checking every other
  caller of the shared coercion helper before considering the fix complete; `WorksheetFunction.*`
  now has its own copy of just the Boolean arm.
- The `=`/`<>` operators had no rule for comparing `Empty` against a number or string —
  `0 = Empty`/`"" = Empty` both fell through to an unconditional `False`, inconsistent with
  `<`/`>` on the exact same operand pairs (which already correctly treated `Empty` as `0`).
  Real VBA documents `Empty` as numeric-comparing as `0` and string-comparing as `""` for
  every comparison operator, not just some of them. Found via the same operator-coercion
  matrix.

## [0.3.0]

Root `elixcee` (Rust crate + Python package) only — `elixcee-types` stays `0.1.0`
(unchanged, no source changes this release) and `@elixcee/xlsx` stays
`0.0.0-development`/unpublished. Verified via a fresh-venv install of a locally-built
wheel, not just `cargo test` (see `Fix`/`Sgn`/`Round`/`CBool` below, all re-checked
through the real Python API after install).

### Added

- **`Fix`, `Sgn`, and `Round` VBA functions.** Root-caused via an automated pass over the
  581-scenario corpus's remaining non-parse-error failures (see below), not guessed at:
  28 of 41 turned out to be `Unknown VBA function` for these three, previously-missing,
  ordinary built-in functions (13 `Sgn`, 12 `Fix`, 3 `Round`) — not deliberate negative
  tests (2 more of the 41 were the related `CBool` bug below, and 1 was a low-value
  `Timer()` left unimplemented; see "Known gaps" in `ROADMAP.md` for the full 41-way
  breakdown). `Fix` truncates toward zero (unlike `Int`, which floors — `Fix(-3.9)` is `-3`, not
  `-4`). `Sgn` returns -1/0/1. `Round` uses real VBA's own banker's rounding
  (round-half-to-even), which is a genuinely different function from
  `WorksheetFunction.Round`/Excel's `ROUND()` formula (round-half-away-from-zero) — `Round`
  does *not* alias or share an implementation with the pre-existing `WorksheetFunction.Round`
  arm; verified both give different, each individually correct, answers on the same tied
  input (`Round(2.5)` is `2`; `WorksheetFunction.Round(2.5)` is `3`).

### Fixed

- `CBool` was grouped with `CLng`/`CInt` and returned a numeric `Variant::Integer` via the
  same numeric-coercion path they use — so `CBool(5)` returned `5` typed `Long`, not `True`
  typed `Boolean` (`TypeName` confirms this live), and `CBool("True")`/`CBool("False")`
  errored outright trying to parse the literal string as a number. Found while implementing
  `Fix`/`Sgn`/`Round` above (same corpus failures also involved `CBool`), not something the
  corpus itself flagged directly — no scenario happened to check `CBool`'s return *type*, only
  whether the string-literal call errored. Now its own arm: a genuine `"true"`/`"false"`
  string (case-insensitive) converts directly, anything else numeric-coerces to boolean, and
  the result is always a real `Variant::Boolean`.
- Single-line `If cond Then stmt [Else stmt]` (no `End If`) now parses — previously
  unsupported at all (`parse_if` unconditionally required a newline right after `Then`).
  Identifier-led statements (assignment, sub call, array/field write — whatever
  `parse_ident_stmt` already covers) are recognized inline, and `Exit For|Do|Sub|Function`
  / `GoTo <label>` are handled explicitly rather than routed through `parse_ident_stmt` —
  the first implementation didn't do this and silently turned `If done Then Exit Sub` into
  a no-op that let execution fall through instead of exiting, caught in review before this
  ever shipped (verified live: `y = 99` after `If x > 0 Then Exit Sub` no longer runs).
  Anything still unrecognized degrades to `Stmt::Unsupported`, same precedent as
  `parse_set`'s unmodeled-target fallback and the identical fallback an ordinary
  unparenthesized bare sub call already hits in block-form VBA — not a new risk this adds.
  This was discovered, not hunted for: it's what the comma-`Dim` fix below unmasked on the
  4 corpus scenarios that fix's own parse-error count didn't reach 0 on. With this fix,
  the corpus's parse-error count is genuinely 0/581 (verified by rerunning
  `compat/corpus/run-elixcee.mjs` — the *committed* `compat/corpus/results/` snapshot
  still shows the pre-fix numbers, since that file is regenerated on demand, not on every
  commit; don't read it as current without rerunning it).
- Comma-separated `Dim`'s built-in/bare-declarator branch (below) lost its old tolerance
  for trailing per-declarator syntax it doesn't model (e.g. `Dim s As String * 10`'s
  fixed-length-string suffix) when the comma loop was added — the first implementation
  returned immediately instead of consuming up to the next comma, so that syntax now
  hard-failed `eat_eol()` instead of being silently skipped like it always was. Caught in
  review before shipping; both the fixed-length-string case and its combination with a
  second comma-declarator (`Dim s As String * 10, i As Integer`) are covered by tests.
- `Not` now does a real bitwise complement on numeric operands (`Not 5` is `-6`, matching real
  VBA), instead of coercing the operand to truthy/falsy first. Only a genuine `Boolean`
  operand still gets logical negation — the same logical-vs-bitwise split `And`/`Or`/`Xor`
  (Phase 2C) already used, `Not` just hadn't been reconciled with it yet (see CHANGELOG's own
  0.2.0 Known limitations: `Not 5 And 3` used to diverge from real VBA's `2`; it now matches).
- Comma-separated multi-declarator `Dim` (`Dim a As Integer, b As Range`) — 0.2.0's last
  documented gap on the 581-scenario VBA corpus's own parse-error surface. Previously, a
  declarator with a non-built-in type (e.g. `b As Range`) returned from `parse_dim` as soon
  as it finished, leaving `, nextDecl` unconsumed; the statement dispatcher's `eat_eol()`
  then hit the stray comma and hard-failed the whole macro. `parse_dim` now loops over every
  comma-separated declarator (`parse_dim_declarator`), wrapping 2+ into a new `Stmt::DimMulti`
  that the VM/`check`/name-resolution passes execute or inspect by replaying each inner
  declarator through the exact same code path a single-declarator `Dim` already used — no
  new semantics, just no longer losing the rest of the line. Verified against the real
  corpus, not just new unit tests: elixcee's own parse-error count on the 581 scenarios goes
  from 8 to 4.
Everything below was previously listed under `[Unreleased]`; this release closes that
section rather than adding new scope. Developed in two internal phases (2B, then 2C after
an integration review found real gaps in 2B's first pass): 2B added the VBA object model
and a working `XLSX.read()` MVP; 2C closed the parser-level and `read()`-completeness gaps
that review surfaced. See "Compatibility" and "Known limitations" below before reading
this as a finished 90-point milestone — the VBA-macro-vs-Microsoft-Excel axis was never
attempted this release (no Windows/Excel environment available), and that gap alone means
a full compatibility claim can't be made on this release's evidence, however solid the
other axes are.

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
- **Range object variables, Union, Areas, SpecialCells, and multi-area Paste** (Milestone B7c): the VBA object-reference layer on top of the B7a/B7b foundation —
  - `Dim rng As Range` / `Set rng = Range(...)` now work with real reference semantics: `Set`-assigned variables live in a new `Vm.object_variables` namespace (`ObjectRef`), kept separate from `Vm.variables` (`Variant`s) rather than adding a `Variant::Object` variant — `Variant` is defined in the shared `elixcee-types` crate the WASM bridge also consumes. Because `ObjectRef::Range` is just B7a's `RangeRef` (coordinates, no cell values), two variables holding the same `RangeRef` already alias the same cells in `Vm.sheets` — real `Set` reference semantics with no `Rc<RefCell<_>>` needed
  - `Union(range1, range2, ...)` combines ranges into one multi-area `RangeRef`; `.Areas.Count` and `.Areas(n)` (1-based) enumerate them — `.Value`/`.Formula`/`.Areas.Count` reuse the existing generic `<var>.<field>` grammar (disambiguated at runtime by which namespace holds `var`), `.Copy`/`.Areas(n)`/`.SpecialCells(...)` are new grammar
  - `SpecialCells(xlCellTypeVisible)` consumes B7b's `sheet_visibility` directly, splitting each area into the Cartesian product of its maximal visible row-bands and column-bands
  - Multi-area Paste: the one shape both sides multi-area with matching `Areas.Count` and per-area shapes now actually executes (pairwise, in order) instead of only diagnosing — every other multi-area shape (count/shape mismatch, or either side single-area) is unchanged and stays diagnose-only via the existing `MULTI_AREA_*` root causes. `Transpose:=True` on this shape still falls through to the diagnose-only path rather than silently writing un-transposed data; merged-cell conflict checking isn't applied to it
  - `ActiveSheet` works as a dynamic sheet qualifier (`ActiveSheet.Range(...)`, `.Cells(...)`, ...); `ThisWorkbook`/`ActiveWorkbook` need no new resolution logic (elixcee only ever loads one workbook) and parse as a plain `Worksheets(...)`/`Sheets(...)` reference
  - Not supported: `Set ws = ActiveSheet` / `Set wb = ThisWorkbook` as Worksheet/Workbook object variables (only `Range` object expressions are `Set`-able); a multi-area source pasted into a single-area destination (or the reverse) actually executing; any new `--json` field (this milestone changes `MULTI_AREA_PASTE_UNSUPPORTED`'s firing condition but adds no new field or code)
- **Hidden row/column evidence** (Milestone B7b): `diagnose`/`diagnose-workbook` now report when a `.Copy`'d range overlaps hidden rows/columns —
  - New `vm::Interval`/`vm::SheetVisibility` types, threaded from XLSX's `<row hidden="1">`/`<col min=".." max=".." hidden="1">` into `Vm.sheet_visibility` the same way `merged_ranges` already is; ODS is explicitly deferred (its reader doesn't expand `table:number-rows-repeated`, so a hidden-row flag can't map to a correct absolute row number yet)
  - New `Vm::hidden_cells_observation()` computes the evidence on demand from the existing `Vm.clipboard` + `sheet_visibility` — no new stored side channel
  - A new sibling JSON field `observations` (not folded into `root_causes`, which means "why it failed" — this isn't a failure), present only when non-empty, on both success and failure: `{"code":"RANGE_CONTAINS_HIDDEN_CELLS","certainty":"observed","range":{...},"visibility":{...},"message":"..."}`
  - `diagnose-workbook` gets the same field via `FixtureResult::Passed`/`::Failed`, though honestly no additional value over a single `diagnose` call — hidden-row/column metadata is structural (workbook layout + macro text), not input-dependent
  - Copy/Paste behavior itself is unchanged — hidden cells still copy/paste exactly as before; this is observability only, laying groundwork for `SpecialCells(xlCellTypeVisible)` (B7c)
- **Multi-area Range foundation** (Milestone B7a): `Range("A1:A3,C1:C3")`-style disjoint ranges now have an underlying model —
  - New `vm::Rect`/`vm::RangeRef` types (`{ sheet, areas: Vec<Rect> }`); the existing single-rect `parse_range_addr`/`SheetRange` and their ~11 call sites are untouched — only Copy/Paste resolve through the new `parse_multi_area_addr`
  - `.Copy` now accepts a comma-separated multi-area source; `.Paste`/`.PasteSpecial` was diagnose-only for every multi-area shape in v1, even a fully-matching one — 4 new classified failures instead: `MULTI_AREA_TO_SINGLE_AREA_PASTE`, `MULTI_AREA_COUNT_MISMATCH`, `MULTI_AREA_SHAPE_MISMATCH`, and the catch-all `MULTI_AREA_PASTE_UNSUPPORTED` (the fully-matching case has since started executing — see Milestone B7c below)
  - Each area's evidence is `{"address", "rows", "columns"}`, matching the completion-condition JSON's own shape
  - `Union()`, the `Areas`/`Areas.Count`/`Areas(n)` property, and `Dim rng As Range`/`Set rng = ...` object variables were unsupported at this milestone — `Variant` gained no Range variant (still true: see Milestone B7c below for how these landed without one)
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
- **Sync `XLSX.read(bytes)` MVP and the `elixcee-wasm` bridge** (Phase 2B): a real, working file-reading entry point for `@elixcee/xlsx`, backed by WebAssembly, callable with no `await init()` —
  - `src/reader.rs`'s `read_workbook` was generalized (pure extraction, no behavior change to the path-based entry point) into `read_workbook_from_archive<R: Read + Seek>` plus a new `pub fn read_workbook_from_bytes(bytes: &[u8])`; `zip::ZipArchive` was already generic, so this needed no new dependency
  - New `crates/elixcee-wasm` crate (first real use of `wasm-bindgen` in this project — deferred until this exact phase per `docs/xlsx-architecture.md`'s ADR). Node gets `wasm-pack --target nodejs` glue (`readFileSync` + synchronous `WebAssembly.Module`/`Instance` construction — genuinely synchronous, verified live by `require()`-ing it and calling the export with no `await`); the browser target inlines the compiled `.wasm` as base64 into its own glue and calls `initSync` itself rather than depending on a bundler's `.wasm` loader (verified: esbuild does not resolve `.wasm` imports by default) — both design choices come from a feasibility spike recorded in `docs/xlsx-architecture.md`'s "Phase 2B-0" section, including its one open item (an oft-cited "Safari enforces a ~4KB sync-compile ceiling" claim could not be substantiated from current MDN docs and is reported unverified, not as fact)
  - `XLSX.read(data, opts)` added to `packages/xlsx` — accepts a `Buffer`/`Uint8Array` or `opts.type === 'base64'`; differential-tested against the real oracle via real file-format round-trips (`compat/differential/xlsx-read.test.mjs`): 9 MATCH + 2 registered `UNSUPPORTED` out of 11 cases, on the scope it actually claims (`SheetNames`, per-sheet `!ref`/`!merges`, per-cell `{t,v}` — no formulas, no styles/dates, no `!rows`/`!cols` yet)
  - `zip`'s default features (`zstd-sys`, `getrandom`'s aes-crypto path) don't compile for `wasm32-unknown-unknown` in this toolchain; trimmed to `default-features = false, features = ["deflate"]` after confirming (by grep) the codebase never uses another compression method — a real fix, not a workaround
  - Node-only for now: the browser-target artifact is built and verified at the bridge level, but `packages/xlsx`'s public `read()` doesn't yet dispatch to it (no `browser` export condition wired up)
- **Oracle-neutral VBA differential corpus infrastructure** (Phase 2B): a reusable, backend-swappable differential-testing pipeline for VBA macro execution, under `compat/corpus/` —
  - Backend-agnostic scenario schema (`compat/corpus/SCHEMA.md`), 581 generated scenarios across 25 categories, an elixcee runner, a LibreOffice UNO runner, a normalizer, and a classifier (`compat/corpus/classify.mjs`, its own file — reuses `compat/differential/classify.mjs`'s verdict vocabulary and anti-laundering discipline rather than importing it, since that file's registries are keyed to the npm API surface, not VBA scenarios) with one new verdict this domain needed: `ORACLE_UNAVAILABLE`
  - Every result record carries an explicit `oracle` field (`"libreoffice"` today; `"microsoft_excel"` is defined in the schema but has never produced a record — see "Compatibility" below). LibreOffice and Excel results are never merged into one number by this pipeline
  - An Excel COM adapter **contract** (`compat/oracle-excel-com/CONTRACT.md`, I/O schema, PowerShell scaffold, Windows execution instructions) is defined but explicitly marked `UNVERIFIED` — no Windows/Excel environment exists in this project's current toolchain to run it
- **VBA foundational syntax — `Mod`/`\`/`^`/logical operators/typed `Function`/`With Range(...)`/`Set` object references** (Phase 2C): closes gaps a Phase 2B integration review found by direct execution — ordinary VBA syntax, not advanced object-model features, that previously stopped a macro at the parse stage —
  - `Mod`, `\` (integer division, real VBA round-half-to-even rounding of each operand before dividing — e.g. `5 \ 0.5` divides by `0`, a real division-by-zero, not a bug), `^` (exponentiation, left-associative, binds tighter than unary minus), and infix `And`/`Or`/`Xor`/`Not` in expressions, all at real VBA operator precedence (`^` > unary `-` > `*`/`/` > `\` > `Mod` > `+`/`-` > `&` > relational > `Not` > `And` > `Or` > `Xor` — pinned by a test asserting `2 + 3 * 2 ^ 2 == 14`)
  - `With Range("A1:B2") ... End With` (previously only `With Sheets(...)` worked); typed `Function` parameters and return types (`Function f(x As Integer) As Integer`)
  - `Set ws = ActiveSheet` / `Set wb = ThisWorkbook` are now real object variables (a new `ObjectRef::Worksheet`/`ObjectRef::Workbook`, alongside B7c's `ObjectRef::Range`) rather than a silent no-op — `ws.Range(...)`/`wb.Worksheets(...)` work through them afterward
  - Measured against the frozen 581-scenario corpus from Phase 2B: parse-error count **132 → 8**. The 8 remaining are all comma-separated multi-declarator `Dim` (`Dim a As Integer, b As Range`, only one declarator per `Dim` parses today) — not one of this phase's items, flagged as the next highest-value parser fix since it's now the entire remaining corpus parse-error surface
  - `&` was moved to its own, correctly-lower precedence tier below `+`/`-` (previously equal precedence) to match real VBA — a small behavior change beyond the 8 named items, required by the precedence table above
- **`XLSX.read()` completeness — formulas, dates, dimension, hidden rows/cols, browser wiring** (Phase 2C): closes every gap Phase 2B's `read()` MVP left open —
  - Formula text (`.f`) captured from `<f>...</f>`; `!rows`/`!cols` (hidden row/column metadata, already used by VBA's `SpecialCells` since B7b) now surfaced in `read()`'s output, gated behind `opts.cellStyles` to match the oracle's own gating
  - `.w` (formatted display string) and date-typed cells (`t:'d'`) — the largest item, requiring `styles.xml` (`<numFmts>`/`<cellXfs>`) and `<workbookPr date1904>` parsing that `reader.rs` did not do at all before this phase — landed in full, not partially: `.w` always computed, `.z` gated behind `opts.cellNF`, `t:'d'` gated behind `opts.cellDates` and a date-like resolved format
  - A real oracle inconsistency was found and deliberately reproduced, not "fixed": in a `date1904` workbook, the oracle's `.w` display string reflects the 1462-day epoch shift but its `cellDates` `.v` Date object does not, because the oracle's own read-direction date conversion doesn't accept `date1904` while its write-direction one does
  - `packages/xlsx`'s `exports["."]` gained a `"browser"` condition routing to the already-built inlined-bytes/`initSync` artifact from Phase 2B — confirmed live via a real subprocess `import.meta.resolve()` + `read()` call under `--conditions=browser`, not just "should work." The browser entry point still assumes bundled consumption (its shared code has a CJS `require('ssf')`) — not literal no-build `<script type=module>` usage
  - Both Phase 2B `UNSUPPORTED_ALLOWLIST` entries under `'read'` (empty-string cells, `<dimension>`) removed now that their underlying `reader.rs` defects are fixed — the registry is empty again

### Changed

- `read_workbook`'s XLSX-archive-consuming body is now generic over `R: Read + Seek` internally (`read_workbook_from_archive`); the public `read_workbook(path)` signature and behavior are unchanged
- Root `Cargo.toml`'s `zip` dependency narrowed from its default feature set to `deflate`-only (see above)
- Binary string-concatenation (`&`) is now its own precedence tier, below `+`/`-` and above relational operators, matching real VBA (see Phase 2C above)

### Fixed

- A silent-wrong-result bug in the new (this release) matching-shape multi-area Paste: `Transpose:=True` was being ignored, writing un-transposed data instead of either transposing correctly or erroring. Caught during the same milestone's self-review, before ever reaching a released version — fixed with a regression test (`matching_shape_multi_area_paste_with_transpose_still_errors_instead_of_silently_mis_pasting`)
- `ssf@0.11.2`'s numFmtId 67-71 indirection-table bug (see Phase 1B-2B above) — carried forward from before this release, listed here for completeness since it ships in 0.2.0's first tagged release
- Empty-string cells (`<c t="str"><v></v></c>`) were silently dropped instead of read as `{t:'s', v:''}`; `<dimension>` was never parsed, so `!ref` always came from the populated-cell bounding box even when a file legitimately declared a wider `<dimension>` (Phase 2C, both shared by `read_workbook` and `read_workbook_from_bytes`)
- `<col hidden="true">` wasn't recognized (only `hidden="1"`) — the oracle's own writer emits `"true"` for columns but `"1"` for rows, an xsd:boolean inconsistency that silently dropped hidden-column detection until caught via a live round-trip (Phase 2C)
- `worksheet_json` always emitted a colon-form `!ref` (`"B2:B2"`) even for single-cell sheets; the oracle collapses `start === end` to a colon-less ref (`"A1"`) — surfaced once Phase 2C's `.w`/date fixtures started using single-cell sheets (Phase 2C)
- An object-qualifier parsing bug (`<var>.Worksheets(`/`.Sheets(`) that could misfire without guarding on an immediate `(` — caught in the same phase's self-review before release (Phase 2C)

### Compatibility

Two independent oracle-differential efforts ran across 2B and 2C. Read both before treating any "compatibility" claim below as broader than what it says.

- **`@elixcee/xlsx` vs. the real `xlsx@0.18.5` npm package** (`compat/differential/`, Node-side, oracle always available): 512 MATCH + 14 registered intentional divergences on the `utils.*` surface, 1831/1831 MATCH on the SSF number-format engine, 34/34 export metadata matches, and `read()` now at **19/19 MATCH, 0 UNSUPPORTED, 0 BUG, 0 UNCLASSIFIED** (up from 2B's 9 MATCH + 2 UNSUPPORTED, and on a widened comparison scope that now includes `!rows`/`!cols`, `.w`, and `.z`). This axis has a real, complete oracle and these numbers are direct measurements.
- **VBA macro execution vs. LibreOffice** (`compat/corpus/`, oracle: `"libreoffice"`): of 581 generated scenarios, only **1 produced an actual `MATCH` comparison** and 2 were `NONDETERMINISTIC`; **578 are `ORACLE_UNAVAILABLE`** — LibreOffice, driven headless via `getScriptProvider().getScript().invoke()` in this project's sandboxed environment, hangs indefinitely (confirmed >90s, no CPU activity) on any `Range`/`Cells` access, which is most of what the corpus exercises. Non-object-model code runs and compares fine (proven by a dedicated smoke scenario). **This is a real, reproducibly-measured negative result, not a partial success**, and it is **unchanged by Phase 2C** — fixing the LibreOffice hang was explicitly out of scope this phase (it doesn't raise elixcee's own product quality, only this one oracle's usability). What Phase 2C *did* measure against the same 581 scenarios is elixcee's own parse-error rate in isolation (132 → 8, see "Added" above) — a real, useful signal, but not a LibreOffice-comparison signal.
- **VBA macro execution vs. Microsoft Excel**: **not attempted, at all, across either phase.** No Windows or licensed Excel environment exists in this project's toolchain. `compat/oracle-excel-com/`'s adapter is a contract only — treat every "LibreOffice" result above as informative but **not** a proxy for Excel compatibility; LibreOffice's own VBA support is its own compatibility layer, not Microsoft Excel.

### Known limitations

- `Not` still evaluates via boolean-truthy coercion, while `And`/`Or`/`Xor` (Phase 2C) do real bitwise math — so `Not 5 And 3` doesn't match real VBA's bitwise result (`2`). Phase 2C's own scope was adding the operators, not reconciling `Not`'s pre-existing evaluation semantics with them.
- Comma-separated multi-declarator `Dim` (`Dim a As Integer, b As Range`) isn't parsed — only one declarator per `Dim` statement. This is now the entire remaining parse-error surface on the 581-scenario corpus (8/581).
- Matching-shape multi-area Paste is the only multi-area Paste shape that executes; every other combination (count/shape mismatch, single↔multi either direction) remains diagnose-only.
- The LibreOffice headless `Range`/`Cells` hang described under "Compatibility" is unresolved — root-caused (headless UNO script invocation, not a scenario-specific issue) but deliberately not fixed this release (out of scope both phases), so the corpus cannot yet produce broad VBA-vs-LibreOffice compatibility signal in this environment.
- The Excel COM adapter is a contract and PowerShell scaffold only; it has never been run, against anything, by anyone, in this project's history to date. This is the single largest remaining gap toward any formal Excel-compatibility claim.
- `@elixcee/xlsx`'s browser `read()` entry point assumes bundled consumption (its shared code `require()`s `ssf`, a CJS-only dependency) — verified via a real subprocess resolving the `"browser"` export condition, not via an actual bundler build (none is installed in this project's toolchain).

## [0.1.2]

Root `elixcee` (Rust crate + Python package) only — `elixcee-types` stays
`0.1.0`/unpublished and `@elixcee/xlsx` stays `0.0.0-development`/unpublished;
neither is part of this release. No public Rust path, CLI behavior, Python
API, or `--json` output shape changed — this release is accuracy,
structure, and test coverage, not new capability.

### Fixed

- **`xml_unescape` (`src/reader.rs`)**: only decoded the 5 named XML entities
  (`&amp;`/`&lt;`/`&gt;`/`&quot;`/`&apos;`), not numeric character
  references (`&#65;`/`&#x41;`); its chained `.replace()` calls also
  double-unescaped input like the literal text `&amp;lt;` (first pass turns
  it into `&lt;`, the very next pass then corrupts that into `<`).
  Rewritten as a single forward pass that decodes numeric references too,
  with the entity-body search bounded to a small window so a run of
  unterminated `&` stays O(n), not O(n²)
- **ODS `table:number-columns-repeated`/`table:number-rows-repeated`**
  (`src/reader.rs`): never read at all — ODS's sparse-representation
  mechanism, used by real producers (LibreOffice) for any run of matching
  cells/rows, not just trailing empty ones, so a real value following a
  repeated block landed at the wrong row/column. Tracked as an arithmetic
  skip rather than a literal expansion loop, so a pathological repeat count
  costs O(1)

### Internal

- **`elixcee-types` crate** (`crates/elixcee-types/`): `ExcelError`,
  `Variant`, `CellContent`, `serial_to_display`, `parse_cell_addr`/
  `parse_range_addr`, and the date-serial helpers `serial_to_ymd`/
  `is_leap`/`days_in_month` extracted from `src/vm/mod.rs`/
  `src/formula/eval.rs` into a new std-only, zero-dependency workspace
  member — the precondition (per `docs/xlsx-architecture.md`'s ADR) for a
  future crate that depends on elixcee's value types without pulling in
  the full VBA parser/VM. Every existing `crate::vm::X` path still resolves
  via re-export; no public module path changed. Root `Cargo.toml` becomes a
  non-virtual workspace root (`fuzz/` explicitly excluded — different
  edition, its own `cargo fuzz` toolchain)
- **Mechanical `clippy` fixes** across the crate (let-chain collapses,
  `.get(0)` → `.first()`, `% 2 == 0` → `.is_multiple_of()`, `Option::
  map_or(true, _)` → `is_none_or()`, redundant closures/casts, 5
  `needless_range_loop` rewrites, 2 duplicated `#[test]` attributes that
  were silently double-registering the same test) — logic-preserving only;
  `docs`/`tasks/todo.md`'s remaining clippy backlog (8 `approx_constant`
  false positives on literal test values, 1 `too_many_arguments`) is
  unchanged, tracked separately, not a release blocker

### Tests

- **Real-producer E2E fixtures** (`tests/fixtures/e2e/`): `.xlsx`/`.ods`
  generated by real LibreOffice (not hand-crafted), read via both
  `calamine` (independent oracle) and elixcee's own reader and asserted
  equal — closes the "zero binary fixtures from a real office suite" gap
  from the `@elixcee/xlsx` Phase 0 investigation. Verified to actually
  catch the two `Fixed` bugs above (both new tests fail against the
  pre-fix reader on this real-producer input, pass after)
- **15 new unit tests** directly in `elixcee-types`, covering the extracted
  surface at its new crate boundary (previously only indirect coverage via
  the much larger `vm`/`formula` test suites)

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
