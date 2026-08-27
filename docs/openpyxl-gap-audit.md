# openpyxl gap audit

`elixcee` is not trying to become openpyxl. Its product core is VBA macro execution,
formula evaluation, and safe OOXML round-tripping — a workbook is always fully loaded into
a mutable `Vm` because arbitrary VBA code needs random access to any cell at any time.
openpyxl's core is a general-purpose spreadsheet file library with no execution model at
all. The two overlap heavily on one thing: ordinary, no-VBA "read some cells, write some
cells, save" RPA work, which today either burns one PyO3 round trip per cell (`get_cell`/
`set_cell`) or forces a caller to add `openpyxl` as a second dependency just for bulk
access.

This document audits openpyxl's public API surface against what `elixcee`'s Python
binding (`PyVm` in `src/lib.rs`) actually exposes today, category by category, and scores
each gap so that "openpyxl has it" is never by itself a reason to build it. Only the
highest-value, lowest-risk category — bulk worksheet range/row access — is built this
round (see "R1" below); everything else is recorded for a later round, or explicitly
never.

**Scoring key** (1–5 each, except Impl/Corruption risk where higher = more dangerous):
User value, Frequency in Excel RPA, Fit with elixcee's product identity, Implementation
risk, OOXML corruption risk. Existing fixture evidence and existing internal primitive
reuse are yes/no. Recommended milestone is `R1` (this round), `P1`/`P2`/`P3` (later,
matching the user's own priority bands), or `Not planned`.

Confirmed elixcee `PyVm` surface as of `0.11.0` (grep-verified, `src/lib.rs`):
`get_cell`/`set_cell`, `cells()`/`get_sheet(name)`/`cells_df()`, `set_cell_formula`/
`set_cell_formula_batch`/`recalculate()`, `set_sheet(name, index=None)`/`delete_sheet(name)`/
`active_sheet()`/`sheet_names()`, `get_cell_number_format(row, col)`, `save_workbook(path)`,
module-level `run_macro`/`load_workbook`. Nothing else — confirmed by grepping every
`fn` in `src/lib.rs` against every style/structural keyword (hyperlink, comment,
defined_name, data_validation, conditional, table, chart, image, drawing, page_setup,
protect, freeze, merge, font, fill, border, rename, move_sheet, copy_sheet): only
`get_cell_number_format` matches anything in that list, and it's read-only.

## 1. Workbook / worksheet lifecycle

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Create/select sheet, position control | — | — | — | — | — | — | — | **Done** (`set_sheet`) |
| Delete sheet | — | — | — | — | — | — | — | **Done** (`delete_sheet`) |
| Rename an existing sheet | 4 | 3 | 5 | 2 | 2 | yes (all 7) | yes (`WorksheetOrigin`) | **Shipped** (`rename_sheet`, P1 core 3) |
| Move/reorder an existing sheet | 3 | 2 | 4 | 2 | 2 | yes | yes (`sheet_order`) | **Shipped** (`move_sheet`, P1 core 3) |
| Copy a sheet (same workbook only) | 3 | 2 | 3 | 3 | 3 | partial | partial | **Shipped** (`copy_sheet`, P2) |
| Sheet visibility (hidden/veryHidden) | 2 | 2 | 4 | 2 | 2 | no | no | **Shipped, read-only** (`sheet_state`, P2) |

`set_sheet`/`delete_sheet` already cover create/select/delete. Rename and move shipped in
the P1 core 3 round as `rename_sheet`/`move_sheet`. **Rename's actual risk was higher than
this table's Impl-risk score of 2 suggested**: it isn't "update the key in `sheets`/
`sheet_order`/`worksheet_origins`" (3 maps) as first assumed — it's 8 lowercase-keyed
per-sheet `Vm` maps that all needed atomic re-keying (`sheets`, `sheet_order`,
`active_sheet`, `merged_ranges`, `sheet_visibility`, `cell_style_indices`,
`cell_number_formats`, `worksheet_origins`), or a rename would silently drop a renamed
sheet's merge/hidden/style state. `move_sheet` also turned out to need a genuinely new
primitive — nothing reordered an *existing* sheet before this round; `ensure_sheet_at`
only positions a newly-created one. Both are still pure in-memory `Vm` state changes with
no new OOXML element, matching R1's low-corruption-risk profile — the *implementation*
risk score undersold the bookkeeping, not the *corruption* risk.

Copy shipped as `copy_sheet`, reusing `rename_sheet`'s own per-sheet-map list directly
(clone-and-insert instead of remove-and-insert) — genuinely close to this table's Impl-risk
score of 3 once `rename_sheet` had already done the harder work of *discovering* that list.
`WorksheetOrigin`'s all-`Option` shape (already exercised by `ensure_sheet`-created sheets)
meant the copy's own origin needed zero new writer logic. See "Implementation notes for P2:
copy_sheet" below for the one design decision this round made deliberately — appending the
copy rather than positioning it next to the source, to avoid the same positional
`<definedName localSheetId="N">`-staleness risk `move_sheet` already guards against.
Sheet visibility (whole-tab hidden/veryHidden, distinct from `sheet_visibility`'s row/col
intervals despite the name collision) shipped read-only as `sheet_state` (P2, fourth
slice). It had zero existing representation anywhere in the reader or writer, confirmed
during this round's research — including a real, independent, pre-existing bug: the
writer never emitted `state="..."` at all, so loading a file with a hidden sheet and
saving it (even a no-op save) silently un-hid it. Write support (`set_sheet_state`) is
deliberately deferred: no real fixture in this repo has a hidden/veryHidden sheet to
validate the writer shape against, and generating one needs either a one-time manual
grant of Excel's sandboxed file-access dialog (macOS) or a hand-built fixture — see
"Implementation notes for P2: sheet_state" below.

## 2. Cell / range access

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Rectangular range read | 5 | 5 | 5 | 1 | 1 | n/a (in-memory) | yes | **R1** (`get_range`) |
| Rectangular range write | 5 | 5 | 5 | 2 | 1 | n/a | yes | **R1** (`set_range`) |
| Append a row | 4 | 5 | 5 | 1 | 1 | n/a | yes | **R1** (`append_row`) |
| Row iteration (values-only) | 4 | 4 | 5 | 1 | 1 | n/a | yes | **R1** (`iter_rows`) |
| Used-range bounds / dimension string | 4 | 4 | 5 | 1 | 1 | n/a | yes | **R1** (`max_row`/`max_column`/`calculate_dimension`) |
| Column iteration (`iter_cols`) | 2 | 2 | 4 | 1 | 1 | n/a | yes | **Shipped** (`iter_cols`, P1 remainder) |
| `Cell` object model (style/comment/hyperlink attached to a returned cell) | 2 | 2 | 2 | 4 | 2 | no | no | **Not planned** |

This is the category R1 closes almost entirely — it needs zero writer/OOXML changes (a
pure in-memory `Vm` cell-map read/write, serialized unchanged by the existing writer),
which is why it's the lowest-risk category in this whole audit by construction, and why
it was picked first regardless of what the rest of this table says. `iter_cols` turned out
to be exactly the trivial follow-up this row predicted (same primitives, transposed) — the
P1 remainder round shipped it as `Vm::iter_cols_values` + `PyVm::iter_cols`. A full
openpyxl-style `Cell` object (an object
that carries `.value`/`.number_format`/`.font`/`.comment`/`.hyperlink` all at once) is
explicitly **not planned** — it would require the full style/comment/hyperlink write
architecture from categories 5/10 to exist first, and would invert `elixcee`'s existing
value contract (plain Python values, not wrapper objects) for every consumer of
`get_cell`/`cells()` at once.

## 3. Row / column operations

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Insert/delete rows/cols (Python-native) | 3 | 3 | 4 | 2 | 1 | n/a | yes (`Vm::delete_rows`/`insert_rows`/`delete_cols`/`insert_cols`) | **Shipped** (`*_on_sheet` + PyVm glue, P1 core 3) |
| Row height / column width (read/write) | 2 | 2 | 3 | 3 | 3 | no | no | **Shipped, read-only** (`row_height`/`column_width`, P2) |
| Hidden row/col (read/write, Python-native) | 2 | 2 | 4 | 2 | 2 | yes | yes (`sheet_visibility`) | **Shipped** (`hidden_rows`/`hidden_columns`/`set_row_hidden`/`set_column_hidden`, P2) |
| Outline/grouping level | 1 | 1 | 2 | 3 | 3 | no | no | **Not planned** |

Row/column insert-delete already had a full, tested implementation at the VBA-execution
layer (`Vm::delete_rows`/`insert_rows`/`delete_cols`/`insert_cols`, added for GitHub #7/#8
in `0.11.0`); the P1 core 3 round added sheet-parameterized `*_on_sheet` siblings and thin
`PyVm` glue (`insert_rows`/`delete_rows`/`insert_cols`/`delete_cols`, matching openpyxl's
own `idx`/`amount` naming) — genuinely close to "nearly free glue" as this row predicted.
It does **not** shift `merged_ranges`/`sheet_visibility`/`cell_style_indices`/
`cell_number_formats`/formula references — a pre-existing VBA-engine limitation, now
Python-reachable (see "Implementation notes for P1 core 3" below). Row height/column
width had zero internal representation at all before this round (not read, not stored,
not written) — real new surface area, not glue, confirmed by research; shipped
read-only as `row_height`/`column_width` (P2, fifth slice), following the same
read-first precedent `sheet_state`/`defined_names`/`merged_cells` already established.
The writer's gap turned out *worse* than `sheet_state`'s: `<row>`/`<cols>` are fully
regenerated from `sheet_visibility` alone on every save, so a loaded file's row
heights/column widths are dropped UNCONDITIONALLY, not just sometimes — see
"Implementation notes for P2: row height / column width" below.

Hidden row/col (read/write) shipped as the first P2 slice, and this row's "primitive
reuse: yes" framing held up on the read side but undersold the write side, a smaller
version of the same pattern `rename_sheet`/`sort_range` hit before it: `sheet_visibility`'s
existing `Interval`-run storage and the writer's already-mechanical `<row hidden="1">`/
`<col hidden="1">` emission made *reading* and *hiding* genuinely close to free, but
*unhiding* a single row/column needs to split whatever interval currently covers it — code
that didn't exist anywhere in the codebase (the existing `visible_runs` helper computes
visible gaps for a whole range and discards which specific hidden interval produced them,
not reusable for identity-preserving single-unit removal). See "Implementation notes for
P2: hidden row/col" below.
Grouping/outline levels have no evidence of demand and touch a genuinely unexplored part
of the schema — not planned absent a concrete request.

## 4. Formula / error / date handling

Already broadly comparable, no large gap. `set_cell_formula`/`set_cell_formula_batch`/
`recalculate()` cover formula writing; a `t="e"` fix on `master` makes `ExcelError` (a
typed object, `code` attribute) round-trip real error cells, but is not yet part of any
released version (ROADMAP.md Known gaps item 14; see the "Packaging note" for why) — once
released, arguably a stronger contract than openpyxl's own (which surfaces an error
cell's cached value as a plain string, with no typed wrapper). Date handling is
openpyxl's weaker spot too, not
elixcee's: both libraries return a date-formatted cell as a raw serial number unless the
caller consults the cell's number format (`get_cell_number_format`, `0.11.0`) or an
`is_date`-style heuristic. No action recommended here; noted for completeness only.

## 5. Cell styles (font / fill / border / alignment / number format / protection)

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Read number format (existing) | — | — | — | — | — | — | — | **Done** (`get_cell_number_format`) |
| Write number format only | 3 | 3 | 3 | 3 | 4 | yes | partial | **P2** |
| Font / fill / border / alignment (read) | 2 | 2 | 2 | 3 | 2 | yes | no | **P3** |
| Font / fill / border / alignment (write) | 2 | 2 | 1 | 5 | 5 | yes | no | **Not planned (without a redesign)** |
| Named styles | 1 | 1 | 1 | 4 | 4 | no | no | **Not planned** |

This is openpyxl's single largest surface area, and the highest-corruption-risk category
in this entire audit. `elixcee`'s current writer strategy is to preserve an existing
`styles.xml` (and the rest of the style graph — `cellXfs`, `numFmts`, fonts/fills/borders)
byte-for-byte via passthrough, only ever *reading* resolved style info
(`get_cell_number_format`). Writing even one new font/fill requires either mutating a
shared, indexed style table in place (risking every other cell that references the same
`cellXfs` index) or growing it correctly (new `numFmt`/`font`/`fill`/`xf` entries,
correctly cross-referenced) — a fundamentally different, riskier writer architecture than
anything else in this document, exactly as flagged in the user's own framing. A narrow
"write only the number-format string for one cell, in a style-table-append-safe way"
slice is plausibly a contained P2; writing font/fill/border/alignment for real needs a
dedicated design round of its own, not a slice of this one.

## 6. Merged cells

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Read a sheet's merged ranges | 3 | 3 | 4 | 1 | 1 | yes (`fixture1`) | yes (`merged_ranges`) | **Shipped** (`merged_cells`, P1 core 3) |
| Create/remove a merge | 3 | 2 | 3 | 2 | 3 | yes | partial | **Shipped** (`merge_cells`/`unmerge_cells`, P1 remainder) |

`merged_ranges` (`src/vm/mod.rs:693`) was already fully populated from any loaded file and
already round-tripped on save — there was simply no Python getter for it (not even
read-only). The P1 core 3 round added a read-only `merged_cells(sheet=None) -> list[str]`,
genuinely nearly-free glue as this row predicted, identical in spirit to R1's own
`get_range`. Creating/removing a merge was considered for that round (an earlier draft of
this doc floated bundling it in) but the user scoped P1 core 3 to read-only, re-scoping
create/remove to P2 at the time — creating a *new* merge needs the writer to correctly
emit a `<mergeCell>` for a range that didn't have one in the source, which sounded like
real new writer surface. The P1 remainder round shipped it anyway, re-scoped back to P1 at
the user's own request, and it turned out meaningfully de-risked from how this row
originally framed it: the writer already emits `<mergeCell>` mechanically from whatever's
in `merged_ranges` with zero validation of its own, so no writer changes were needed at
all — only `Vm`-side map management (`merge_cells`/`unmerge_cells`), reusing
`rects_overlap` (Milestone B6c2's Copy/Paste conflict-detection primitive) for overlap
rejection instead of writing new geometry-math from scratch.

## 7. Defined names

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Read defined names | 3 | 2 | 4 | 1 | 1 | yes (`fixture4`) | partial | **Shipped** (`defined_names`, P2) |
| Create/delete a defined name | 2 | 1 | 3 | 3 | 3 | yes | no | **P2** |

Defined names already survive round-trip (`0.10.0-C` slice 3, delete-gated passthrough) —
confirmed preserved, but there was no Python-visible representation of them at all (not
even internal to `Vm` as a queryable structure; they lived purely as passthrough XML).
Reading required actually parsing `<definedNames>` into a real internal structure for the
first time — genuinely as modest as this row predicted: a new `reader::xlsx_defined_names`
streaming parser, modeled directly on the existing `xlsx_shared_strings` pattern (same
crate, same technique, no new parsing infrastructure). Deliberately read-only and
deliberately unresolved — returns each name's raw formula text (e.g. `"Sheet1!$A$1:$A$3"`)
rather than a resolved `(sheet, address)` tuple, since elixcee's formula engine has no
cross-sheet reference syntax to resolve that text against, and real XLSX additionally
allows a sheet-scoped (`localSheetId`) name to shadow a workbook-scoped one of the same
name, which a flat map can't represent anyway — see "Implementation notes for P2:
defined_names" below for the full account. Create/delete remains P2: a real write path
into `<definedNames>` (rather than reading the existing passthrough blob) is a different,
larger piece of work not attempted this round.

## 8. Tables / AutoFilter / sorting

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Python-native `.sort_range(...)` | 2 | 2 | 3 | 1 | 1 | n/a | yes (`Vm`'s VBA `RangeSort` handler) | **Shipped** (`sort_range`, P1 remainder) |
| Python-native AutoFilter (VM-effect only, matching `0.11.0`'s VBA scope) | 2 | 2 | 3 | 2 | 2 | n/a | yes | **P2** |
| Table (`<table>` part) creation | 2 | 2 | 2 | 4 | 4 | no (no fixture has a from-scratch table) | no | **P3 / user's own explicit deferral** |

Sort and AutoFilter already exist as VBA statements (`0.11.0`). This row's "thin wrapper,
cheap" framing undersold `sort_range`'s actual implementation cost the same way the P1
core 3 round's `rename_sheet` row did: the entire sort algorithm was inlined directly in
`Stmt::RangeSort`'s dispatch arm, active-sheet-only, not a standalone method to wrap —
shipping `sort_range` required first extracting it into a sheet-parameterized
`Vm::sort_range_on_sheet` (see "Implementation notes for P1 remainder" below). AutoFilter
remains P2, unaffected by this round; a thin Python-native wrapper calling the same
internal handlers is still expected to be cheap there, matching this audit's general theme
of "glue over an existing VBA primitive is cheap, a brand-new OOXML element is not." Table
*creation* is explicitly on the user's own P3 list and stays there — this project's
existing hard gate (no writer code for a structural OOXML element without real fixture
evidence of its shape) applies directly, and no fixture with a from-scratch table exists.

## 9. Data validation / conditional formatting

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Data validation creation | 2 | 2 | 2 | 4 | 4 | partial (`fixture3` has one, pre-existing only) | no | **P3 / user's own explicit deferral** |
| Conditional formatting creation | 1 | 1 | 2 | 4 | 4 | partial (`fixture3`) | no | **Not planned** |

Both exist only as passthrough-preserved content in fixtures that already had them at
load time (`fixture3_table_validation_conditional.xlsm`) — creating either from scratch
hits the same hard-gate/high-corruption-risk profile as table creation above. Data
validation is on the user's own explicit P3 list; conditional formatting isn't on any of
the user's priority bands and has no signal of demand — recorded, not planned.

## 10. Comments / hyperlinks

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Read hyperlinks | 3 | 3 | 4 | 2 | 2 | yes (`fixture4`/`fixture6`) | partial | **P2** |
| Create/modify a hyperlink | 2 | 2 | 3 | 3 | 3 | yes | partial | **P2** |
| Read/write comments (notes) | 2 | 2 | 3 | 3 | 4 | yes (`fixture4`) | no | **P3** |

Internal (location-only) hyperlinks already round-trip via `0.10.0-B4`'s relationship-free
restoration, released. External (r:id-backed) hyperlinks round-trip too on `master`, via
`0.10.0-D`'s relationship-backed restoration -- but `0.10.0-D` is not yet part of any
released version (see ROADMAP.md's "Packaging note"). There is still no Python-visible
read of either kind. Comments
are VML-backed (`<legacyDrawing>`) — a genuinely more awkward OOXML shape than a plain
`r:id` hyperlink, and higher corruption risk to write, hence P3 over P2.

## 11. Charts / images / drawings

Read-only passthrough preservation exists on `master` (`0.10.0-D`, relationship-graph
reachability), but is not yet part of any released version (see ROADMAP.md's "Packaging
note"). Zero creation/modification API of any kind, regardless. Explicitly out of scope
for this round's
non-goals, and arguably **not planned indefinitely** rather than merely deferred — chart/
image *authoring* is one of openpyxl's largest subsystems in its own right, has a poor fit
with a VBA-execution-centric product identity (nobody drives Excel chart creation from
Python instead of from the VBA they're trying to emulate), and every prior "why not X"
decision in this project's own history has favored VBA-execution-parity work over
GUI-authoring-parity work. Not scored — this is a "why this project probably never builds
it" note, not a milestone candidate.

## 12. Page setup / print area

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Read page setup / print area | 2 | 2 | 3 | 2 | 2 | yes (`fixture5`) | partial | **P3** |
| Write page setup / print area | 1 | 1 | 2 | 3 | 3 | yes | no | **Not planned** |

Already preserved on round-trip (`0.10.0-C`/`D`); low frequency in RPA use cases relative
to everything above it, hence P3 for read-only, and no signal of demand for write.

## 13. Workbook / worksheet protection

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Read/set sheet protection from Python | 2 | 2 | 3 | 2 | 2 | no | partial (`protected_sheets`, fully private today) | **P3** |

VBA's `.Protect`/`.Unprotect` already exist and set `protected_sheets`
(`src/vm/mod.rs:684`), but that field has no `pub`/`pub(crate)` visibility at all, and (per
this session's own investigation for R1) protection enforcement itself is inconsistent —
14 VBA-statement call sites honor it, the Python binding honors it nowhere. Exposing it to
Python without first deciding whether the Python binding should honor it too is a
half-measure; recorded as P3 pending that decision, not scored higher.

## 14. Read-only / write-only streaming

**Not planned — structural mismatch with elixcee's product identity, not a
missing-feature gap.** openpyxl's `read_only=True`/`write_only=True` modes exist so a
caller can stream a huge workbook in/out without holding it entirely in memory, at the
cost of losing random access. `elixcee`'s entire execution model is the opposite: a VBA
macro can reference any cell on any sheet at any point during execution, so `Vm` holding
the complete workbook in memory isn't an implementation shortcut to later remove — it's a
load-bearing assumption the whole VBA interpreter depends on. Streaming would require an
entirely separate, execution-model-free code path with none of the VM's guarantees,
closer to a second product than a feature of this one. Recorded so it isn't silently
reproposed later without this context.

## Summary: recommended order after R1

1. **R1 (shipped)**: bulk worksheet range/row API — category 2, closes it almost
   entirely, zero writer changes, lowest risk in this document.
2. **P1 core 3 (shipped)**: sheet rename/move (category 1), Python-native row/col
   insert-delete wrapping the existing `0.11.0` VBA handlers (category 3), read-only
   merged-cell access (category 6). Rename/move turned out to be more than "cheap glue"
   (see "Implementation notes for P1 core 3" below) — real, if contained, new bookkeeping.
   `iter_cols` (category 2) and Python-native `.sort_range(...)` (category 8), both also
   tagged P1 in this document's tables, were deliberately deferred out of this round's
   scope, not forgotten.
3. **P1 remainder (shipped)**: `iter_cols` (category 2, genuinely as cheap as predicted),
   Python-native `.sort_range(...)` (category 8, more implementation cost than predicted —
   see "Implementation notes for P1 remainder" below), and merge create/remove (category 6,
   re-scoped back to P1 from the P2 landing it got when P1 core 3 stayed read-only —
   turned out meaningfully *less* new writer surface than that P2 re-scoping assumed).
4. **P2, first slice (shipped)**: hidden row/col read/write (category 3). Reading and
   hiding were genuinely close to free as this row's table predicted (`sheet_visibility`'s
   existing storage, the writer's already-mechanical `hidden="1"` emission); unhiding a
   single row/column needed genuinely new interval-splitting logic the table's score didn't
   surface — see "Implementation notes for P2: hidden row/col" below.
5. **P2, second slice (shipped)**: `copy_sheet` (category 1). Genuinely close to this
   table's own cost estimate once `rename_sheet` had already discovered the per-sheet-map
   list it reuses — see "Implementation notes for P2: copy_sheet" below for the one design
   decision worth disclosing (appending rather than positioning next to the source, to
   avoid a positional `<definedName>`-staleness risk).
6. **P2, third slice (shipped)**: `defined_names` read-only (category 7). Genuinely as
   modest as this table's own row predicted — a new streaming parser modeled directly on
   the existing `xlsx_shared_strings` pattern, no new parsing infrastructure. Deliberately
   read-only and deliberately unresolved (raw formula text, not a resolved sheet+address) —
   see "Implementation notes for P2: defined_names" below.
7. **P2, fourth slice (shipped)**: `sheet_state` read-only (category 1's other row).
   Confirmed a real, independent, pre-existing bug in the same research pass: the writer
   never emitted `state="..."` at all, so a loaded file's hidden sheet silently reverted to
   visible on ANY save (pinned by a differential-python test, not just disclosed in prose).
   Write support (`set_sheet_state`) deliberately deferred — zero real fixtures in this repo
   have a hidden/veryHidden sheet, and this project's hard gate is no writer code for a
   structural OOXML element without real fixture evidence. See "Implementation notes for P2:
   sheet_state" below for the fixture-generation path found (and not yet taken) this round.
8. **P2, fifth slice (shipped)**: `row_height`/`column_width` read-only (category 3's
   other row). Confirmed zero prior representation, matching this row's own prediction —
   but also confirmed the writer's gap is *worse* than `sheet_state`'s: `<row>`/`<cols>`
   are fully regenerated from `sheet_visibility` alone on every save, so a loaded file's
   row heights/column widths are dropped UNCONDITIONALLY (not just on some saves), pinned
   by a differential-python test. Two independent value types, not one enum like
   `sheet_state` — `row_heights: HashMap<u32, f64>` (per-row) and `column_widths:
   Vec<(u32, u32, f64)>` (range-shaped like `hidden_columns`, with a value attached) —
   pushed `rename_sheet`'s per-sheet-map re-key count from 9 to 11. Write support
   deferred for the same reason as `sheet_state`'s: zero real fixtures have a genuine
   custom row height or column width (fixture1's only `<col>` is a hidden column with
   `width="0"`, not real data). See "Implementation notes for P2: row height / column
   width" below.
9. **P2, remaining**: number-format-only writing (category 5, the narrowest possible
   slice of the style engine), defined-name create/delete (category 7's other row — a
   real write path into `<definedNames>`, different and larger than the read-only slice
   already shipped), sheet-visibility *write* support (`set_sheet_state`, blocked on real
   fixture evidence), row-height/column-width *write* support (`set_row_height`/
   `set_column_width`, same blocker), Python-native AutoFilter (category 8 — confirmed to
   need the same `Stmt::RangeAutoFilter`-extraction treatment `sort_range` did, plus its
   own signature-design decision since `field`/`criteria1` are VBA `Expr` AST nodes, not
   plain values — don't take its "thin wrapper" framing at face value, the same optimism
   this doc's table showed for `rename_sheet`/`sort_range`/hidden-row-write before each
   one turned out to need more than glue), hyperlink read/write (category 10 — confirmed
   to need two separate pieces of new work, not one: parsing `<hyperlink>` elements into a
   queryable structure AND joining `r:id`-backed ones against the sheet's own `.rels` file,
   a lookup that today only happens ad hoc for survival-checking).
10. **P3**: font/fill/border/alignment *read* (not write), comments, page-setup read, sheet
    protection exposure (each needs its own follow-up design decision, noted above).
11. **Not planned**: full style-engine writing, named styles, table/data-validation/
    conditional-formatting *creation*, chart/image authoring, outline/grouping, streaming
    modes, and a wrapper `Cell` object model — each either fights this project's own hard
    gates (no writer code without fixture evidence), fights its product identity (VBA
    execution needs full random access; chart/image authoring has no VBA-emulation angle),
    or both.

---

## Implementation notes for R1

Gaps and deliberate scope boundaries discovered while designing the bulk worksheet API
(`get_range`/`set_range`/`append_row`/`iter_rows`/`max_row`/`max_column`/
`calculate_dimension`), disclosed here rather than silently absorbed or silently fixed as
a side effect of an unrelated feature:

**Three pre-existing gaps in the shared address parser
(`crates/elixcee-types::parse_cell_addr`/`parse_range_addr`), closed only for calls made
through the new bulk API, left untouched everywhere else.** This is a shared function
with many call sites across VBA parsing/formula evaluation; fixing it project-wide is a
separate concern from this round, recorded here for whoever picks it up next:
- `col_letters_to_num_vm` does an unchecked `u32` subtraction that **underflows** on a
  leading `$` (Excel's absolute-reference syntax, e.g. `"$A$1"`) — no existing caller
  strips `$` before calling into this function today. The new API's `validate_range_addr`
  wrapper (`src/lib.rs`) strips `$` before delegating, closing the risk for its own call
  path only.
- `parse_cell_addr("A0")` succeeds, returning `(0, 1)` — row `0` is accepted rather than
  rejected as invalid. The new API's wrapper explicitly rejects row/col `0`.
- `parse_range_addr("C3:A1")` (a reversed range) is accepted as-is rather than rejected —
  inconsistent with `reader.rs`'s own separate `parse_dimension_ref`, which already
  rejects a reversed ref. The new API's wrapper rejects a reversed range too, matching
  that existing precedent, but the shared parser itself still accepts one.

**`cells_df`'s used-range convention diverges from `cells()`/`get_sheet()`'s, and this
round does not reconcile them.** `cells_df` (`src/lib.rs`) computes its own max row/column
by including `HashMap` entries whose `.value` is `Variant::Empty`; `cells()`/`get_sheet()`
(and the new `sheet_used_range` this round adds, feeding `get_range`/`iter_rows`/
`max_row`/`max_column`/`calculate_dimension`) exclude them. A `Vm` can therefore report a
different used-range extent from `cells_df()` than from the new methods, for the exact
same underlying state. Pre-existing, not introduced by this round; noted so it isn't
mistaken for a new bug later.

**`Variant::Null` counts toward a sheet's used-range bounding box; `Variant::Empty` does
not — even though both render as Python `None`.** This matches `cells()`/`get_sheet()`'s
existing behavior (neither of those two variants is new), but it's a real, disclosed
surprise: a sheet whose only "content" is a single VBA-assigned `Null` at, say, `Z100`
will report `max_row() == 100`, `get_range` on that cell returns `None`, and there is no
way from the returned `None` alone to tell why that cell counted.

**No upper-bound/size guard on `get_range`/`iter_rows`/`set_range`.** A pathological
full-column or full-row address (e.g. `"A1:XFD1048576"`, ~2.3 billion cells) will attempt
to allocate and iterate that many cells rather than erroring quickly. Not implemented in
this round absent concrete evidence anyone actually does this — added here as a disclosed
limitation, matching this project's own stated pattern of not preemptively guarding
without evidence of real-world impact.

---

## Implementation notes for P1 core 3

Gaps and deliberate scope boundaries discovered while implementing sheet rename/move and
row/col insert-delete glue, disclosed here rather than silently absorbed or fixed as a
side effect of an unrelated feature:

**Row/col insert-delete does not shift merged ranges, hidden-row/col markers, cell
styles/number formats, or formula cell-reference text.** This is a pre-existing
limitation of the underlying VBA engine (`Vm::insert_rows`/`delete_rows`/`insert_cols`/
`delete_cols`, and now their `*_on_sheet` siblings) — real Excel shifts all of these when
a row/column is inserted or deleted; `elixcee` doesn't, and didn't before this round
either. Making these functions Python-reachable surfaces the gap to a new audience, so
it's stated here explicitly rather than silently inherited. Pinned as an executable fact
by `insert_rows_on_a_merged_and_hidden_row_sheet_does_not_shift_the_merge_or_hidden_markers`
(`tests/xlsx_roundtrip.rs`).

**`rename_sheet` does not rewrite formula or `<definedName>` text that refers to the
sheet by its old name — only the `<sheet name="...">` tab label changes.** `elixcee`'s
formula engine has no cross-sheet cell-reference syntax (`=Sheet2!A1`) today, so this
can't corrupt a formula. A `<definedName>` whose *text* names the old sheet (as opposed to
its *`localSheetId`*, which is positional) is a real risk for the file's next reader
(Excel, or anything else that resolves that text against a live sheet name) — mitigated,
not by rewriting the text, but by dropping the whole `<definedNames>` element on any
rename, the same way a deletion already does (see "dropped wholesale" below). Also does
not validate Excel's real sheet-name rules (31-character limit, illegal characters
`: \ / ? * [ ]`, reserved/duplicate-after-truncation names) beyond rejecting an
empty/whitespace-only name — matches `set_sheet`'s pre-existing total lack of name
validation, not a new regression.

**`remove_sheet` leaves stale entries in 6 of the 8 per-sheet maps `rename_sheet` had to
learn to re-key atomically.** `remove_sheet` (`src/vm/mod.rs`) only cleans `sheets` and
`sheet_order` on delete; `merged_ranges`, `sheet_visibility`, `cell_style_indices`,
`cell_number_formats`, `worksheet_origins`, and `protected_sheets` all keep a dead entry
under the deleted sheet's old key. Harmless today (the key is never looked up again), but
a real, pre-existing gap surfaced while designing `rename_sheet`'s own re-key list.
Deliberately **not** fixed in this round — the user was offered "fix this first and share
the re-key list with `rename_sheet`" as an option and chose to proceed with rename alone
instead, so this stays a known gap for a future round rather than an incidental fix.

**`<definedNames>` passthrough is now also guarded against `move_sheet`-caused reordering
AND `rename_sheet`-caused staleness, closing two gaps that P1 core 3 would otherwise have
introduced.** A `<definedName localSheetId="N">` is a positional index into `<sheets>`;
the existing save-time guard (`src/lib.rs`) already dropped `<definedNames>` passthrough
once any sheet was deleted (its `localSheetId`s could no longer be trusted), but said
nothing about *reordering* — before this round, nothing could reorder an existing sheet,
so the gap was latent. Separately, a `<definedName>`'s own TEXT can reference a sheet by
name (e.g. `Sheet1!$F$5`), which `rename_sheet` doesn't rewrite, so a renamed sheet could
leave that text dangling. **Both were initially missed**: the first implementation of this
fix set a flag only from `move_sheet`, not `rename_sheet` — caught in a second review pass
against a fixture (`fixture4`) that actually has `<definedNames>` content, since neither
this round's original tests nor `mechanical_check.py`'s pipeline exercised that
combination. Fixed by a single `Vm::defined_names_may_be_stale` flag, set by both
`move_sheet` and `rename_sheet`, checked alongside the existing deletion check — dropping
the whole element wholesale rather than attempting a surgical `localSheetId`
renumbering or defined-name-text rewrite, consistent with the deletion case's own
established choice. **One related gap remains, not closed by this fix**: VBA's
`Sheets.Add(before:=...)` can also shift existing sheets' positions without deleting
anything, and nothing tracks that today either — pre-existing, not introduced by this
round, and a real fix would need snapshotting the workbook's load-time sheet order for
comparison, which doesn't exist anywhere today. Pinned by
`move_sheet_drops_defined_names_that_would_have_stale_positional_indices` and
`rename_sheet_drops_defined_names_that_would_reference_the_old_name`
(`tests/xlsx_roundtrip.rs`).

---

## Implementation notes for P1 remainder

Gaps and deliberate scope boundaries discovered while implementing `iter_cols`,
`sort_range`, and merge create/remove, disclosed here rather than silently absorbed or
fixed as a side effect of an unrelated feature:

**`sort_range`'s "thin wrapper" framing in this doc's table undersold the real
implementation cost, the same pattern as `rename_sheet` in the P1 core 3 round.** The
entire sort algorithm (address resolution, row gathering, comparison, write-back) was
inlined directly in `Stmt::RangeSort`'s VBA dispatch arm — active-sheet-only, with no
standalone method underneath it to call from Python. Shipping `sort_range` required first
extracting that body into a new sheet-parameterized `Vm::sort_range_on_sheet`, built on the
existing `read_rect`/`write_rect` primitives from R1 rather than the original's manual
per-cell loops. The VBA dispatch arm now just resolves the address, checks protection, and
delegates — all 4 pre-existing `test_range_sort_*` tests pass with zero modification,
confirming the extraction preserved VBA's existing behavior exactly.

**`sort_range`'s Python API deliberately diverges from the VBA path on an out-of-range
`key_col`, rather than inheriting its silent clamp.** The original inline code computed
`key_col.saturating_sub(c1)` with no bounds check — a `key_col` below the range's own `c1`
silently saturates to offset `0` and sorts by the range's first column instead of erroring.
This is preserved as-is for `Stmt::RangeSort` (an existing, tested VBA behavior this round
had no mandate to change) and pinned by a dedicated unit test
(`sort_range_on_sheet_with_an_out_of_range_key_col_clamps_via_saturating_sub`) specifically
so it can't be "fixed" by accident later without a conscious decision. `PyVm::sort_range`,
having no prior behavior to preserve, instead raises `ValueError` naming both the bad
`key_col` and the range's actual column span.

**Neither `sort_range` nor `merge_cells` had an upper-bound guard on the address before
this round; both got the same 1,048,576-row/16,384-column ceiling `insert_rows`/
`delete_rows` already enforce (P1 core 3), added at the `PyVm` layer.** R1's own
`get_range`/`iter_rows` were deliberately left unguarded (see "Implementation notes for
R1" above) on the reasoning that a pathological address there just costs a large,
self-inflicted allocation. `sort_range` and `merge_cells` are different: an unbounded
address doesn't just cost time, it writes real geometry (a bulk value rewrite, or a
`<mergeCell>` spanning the address) into the file that gets saved — a persisted-corruption
path, not a transient cost, so the same ceiling other write-shaped methods already use was
applied here too rather than left as a disclosed gap.

**Merge create/remove turned out to need zero writer changes, contrary to this doc's own
"real new writer surface" framing when it was P2-scoped after P1 core 3.** `save_xlsx_impl`
already emits `<mergeCell ref="...">` mechanically from whatever `merged_ranges` holds, with
no validation of its own — a new create/remove API only needed to correctly manage that
map. `merge_cells` rejects a single-cell address and any overlap with an existing merge on
the same sheet, reusing `rects_overlap` (Milestone B6c2's Copy/Paste conflict-detection
primitive, already sheet-agnostic) rather than the Copy/Paste-specific
`check_merge_conflicts` (which is `&mut self` and writes a diagnostic side channel, the
same reasoning R1/P1 core 3 used to justify their own independent resolvers over reusing
`require_sheet_exists`/`check_sheet_not_protected`). The overlap check runs before
`merged_ranges` is mutated, so a rejected merge cannot leave a stray empty entry behind for
a sheet that previously had none. `unmerge_cells` requires an exact rect match and errors
on a partial/no match rather than silently no-opping, matching `rename_sheet`/`move_sheet`/
`delete_sheet`'s existing "must not silently no-op on an unknown target" convention.
Neither method touches cell values in the covered range — this VM's merge geometry and
cell values were already orthogonal by design (`write_rect`/`set_range` already permit
writing into a non-anchor merged cell without error; this round applies the same
precedent in the other direction).

---

## Implementation notes for P2: hidden row/col

Gaps and deliberate scope boundaries discovered while implementing `hidden_rows`/
`hidden_columns`/`set_row_hidden`/`set_column_hidden`, the first P2 slice, disclosed here
rather than silently absorbed or fixed as a side effect of an unrelated feature:

**Reading and hiding were genuinely close to free, as this doc's table predicted; unhiding
was not, the same undersold-cost pattern `rename_sheet` and `sort_range` hit before it.**
`Vm.sheet_visibility`'s `hidden_rows`/`hidden_columns: Vec<Interval>` already existed
(Milestone B7b, built for `SpecialCells(xlCellTypeVisible)`), and the writer already emits
`<col min=".." max=".." hidden="1">`/`<row r=".." hidden="1">` purely mechanically from
whatever's in those lists, with no validation — so `hidden_rows_on_sheet`/
`hidden_columns_on_sheet` (flatten intervals into a sorted `Vec<u32>`) and the *hide* half
of `set_row_hidden_on_sheet`/`set_column_hidden_on_sheet` (push a new single-unit interval,
a no-op if the unit's already covered) needed no new algorithmic work at all. *Unhiding* a
single row/column, though, needs to split whatever interval currently covers it — dropped
entirely if it's a single-unit interval, shrunk from one end, or split into two flanking
intervals if the unit sits strictly inside a wider hidden range (e.g. a loaded fixture's
`Interval{1,10}` becomes `[{1,4},{6,10}]` after unhiding row 5). The existing
`visible_runs` helper (`src/vm/mod.rs`) computes visible *gaps* across a whole range and
discards which specific hidden interval produced each gap, so it wasn't reusable for this
identity-preserving, single-unit removal — a new free function, `remove_unit_from_intervals`,
was needed.

**Hiding an already-hidden unit is a no-op, not a duplicate interval push, and unhiding an
already-visible unit (or a unit on a sheet with no `sheet_visibility` entry at all) is a
no-op that does not create a stray empty entry.** The first was caught before writing any
code, not after: a naive "always push a new single-unit interval" hide implementation would
have made `hidden_rows_on_sheet`'s own flattened output correct regardless (a `BTreeSet`
collapses the duplicate), but would silently leave two overlapping intervals describing the
same row in `sheet_visibility` itself — harmless today, but exactly the kind of
easy-to-miss state divergence this project's own house rule (validate/check before
mutating a map, not after) exists to prevent, the same rule `merge_cells` follows for its
own overlap check. The second matches `merge_cells`'s identical convention for the same
reason: `.entry(key).or_default()` on an unhide call would insert an empty
`SheetVisibility` for a sheet that had no hidden-row state at all, purely from a call that
changed nothing.

**`hidden_columns()`'s return shape is a real API choice, not an implementation detail.**
Columns are stored as intervals (mirroring `<col min="..." max="...">`), and openpyxl's own
`ws.column_dimensions` is keyed by column *letter* (`"D"`), not number. This API returns
plain, expanded 1-based column numbers (matching `hidden_rows()`'s own shape and the
setters' own number-based parameters) rather than letters or interval tuples — chosen for
symmetry with `hidden_rows()` and because every other numeric row/col API in this binding
(`insert_rows`, `sort_range`'s `key_col`, etc.) already uses plain numbers, not letters.

**No guard against a pathological full-sheet hide** (e.g. Excel's own `<col min="1"
max="16384" hidden="1">` shape for "hide all columns," or the row equivalent) — flattening
such an interval into individual numbers would eagerly materialize up to 1,048,576/16,384
entries. Not implemented in this round absent concrete fixture evidence anyone actually
does this, matching R1's own precedent for `get_range`/`iter_rows`'s unbounded-address gap.

---

## Implementation notes for P2: copy_sheet

Gaps and deliberate scope boundaries discovered while implementing `copy_sheet`, disclosed
here rather than silently absorbed or fixed as a side effect of an unrelated feature:

**This table's own cost estimate held up, for once — but only because `rename_sheet`
(P1 core 3) had already paid the discovery cost.** `copy_sheet` reuses the exact same
per-sheet-map list `rename_sheet` needed to learn the hard way (`sheets`, `merged_ranges`,
`sheet_visibility`, `cell_style_indices`, `cell_number_formats`, `worksheet_origins`) —
`get()`-then-`clone()`-then-`insert()` instead of `remove()`-then-`insert()`, with no new
algorithmic work anywhere. `WorksheetOrigin`'s all-`Option` shape, already exercised by
every `ensure_sheet`-created sheet (GitHub #2's own fix), meant the copy's brand-new origin
(`original_display_name` set, everything else `None`) hits an already-correct writer code
path with zero new logic. Had `rename_sheet` not already existed and disclosed its own
8-map re-key list, this round would likely have repeated the same discovery cost.

**Deliberately appends the copy at the end of `sheet_order` rather than positioning it
immediately after the source, unlike openpyxl's own `copy_worksheet`.** Inserting a new
sheet anywhere before the end of `sheet_order` shifts every later sheet's positional index
by one — the exact same risk `move_sheet`'s own `defined_names_may_be_stale` flag exists to
guard against for a *reorder*. An append never changes any existing sheet's index, so it
was chosen specifically to avoid needing that flag at all for `copy_sheet` — verified by a
dedicated test (`copy_sheet_does_not_flag_defined_names_as_stale`) pinning that the flag
stays `false`. A caller who wants the copy positioned next to its source can compose
`copy_sheet` with the existing `move_sheet` rather than this method growing its own
placement parameter — consistent with `set_sheet`'s `index` and `move_sheet` already
covering "position a sheet" as a separate, general concern.

**Does not copy sheet protection status; the copy is always unprotected.** No fixture
evidence or concrete signal exists for what "copying a protected sheet" should mean (does
the copy inherit the protection, or should protecting be a separate, deliberate act on the
new sheet?) — deferred rather than guessed, matching this project's own restraint pattern
for genuinely ambiguous, evidence-free decisions.

**Discovered while writing this round's differential-python coverage, unrelated to
`copy_sheet` itself: `Vm::sheet_names()` returns sheets alphabetically sorted, not in
`sheet_order`/tab-position order.** Undocumented in both the Rust doc comment and the
Python docstring; not a regression (it predates this round entirely), just never
previously exercised by a differential test against a real multi-sheet fixture. Not fixed
here — changing an existing, unversioned method's ordering contract is out of scope for a
sheet-copy feature and could be a breaking change for any existing caller relying on
alphabetical order. Recorded so it isn't rediscovered as a surprise later.

---

## Implementation notes for P2: defined_names

Gaps and deliberate scope boundaries discovered while implementing `defined_names`,
disclosed here rather than silently absorbed or fixed as a side effect of an unrelated
feature:

**Confirmed before writing any code: `Vm.named_ranges` is NOT a loaded file's defined
names.** It's a completely separate table, populated only by the VBA runtime statement
`Range(addr).Name = "x"` (`Stmt::RangeName`) and consulted only by `resolve_range_addr`/
`resolve_multi_area_addr` to expand an active-sheet-only address string during macro
execution. `populate_from_sheets` never touches it. Worth stating plainly since it would
have been an easy, wrong assumption to reuse it — "read defined names" turned out to be
genuinely zero reused work, not partially done, matching this table's own "partial"
primitive-reuse rating for the right reason (the passthrough machinery is reusable, the
resolution table is not).

**`<definedNames>` had no queryable in-memory representation anywhere — only a raw,
opaque XML blob carried from source to output by `save_xlsx_impl`'s own passthrough
logic (`OpaqueWorkbookFragments`).** That passthrough logic doesn't even live on `Vm`
between load and save: the raw `xl/workbook.xml` text is re-read from the source file
path at *save* time (`reader::read_raw_zip_entries`), not cached anywhere after
`load_workbook_file` returns. `Vm::defined_names()` follows the same re-read-on-demand
pattern rather than adding a new eagerly-populated `Vm` field — it's a pure reporting
view of what the file currently says, re-derived from `Vm.loaded_workbook_path` (already
existed, used for the in-place-save feature) each call. This means it surfaces a `ValueError`
if the source file is no longer readable (moved/deleted after loading) rather than
silently reporting `{}` — a genuinely different failure mode from "no workbook loaded at
all," which does return `{}`.

**Deliberately does not resolve a defined name's text into a `(sheet, address)` tuple.**
A `<definedName>`'s TEXT is a sheet-qualified reference like `Sheet1!$F$5` — elixcee's
formula engine has no cross-sheet reference syntax (`=Sheet2!A1`) anywhere today, and there
is no existing "split sheet-name!address" parser to reuse. Writing one from scratch for
this feature alone was judged real, separable new work, not folded in here. Returning the
raw text (matching openpyxl's own `wb.defined_names[name].value`, confirmed via a
differential test against the real fixture) is a smaller, honestly-scoped read.

**Sheet-scoped and workbook-scoped names both flatten into the same map, undistinguished.**
Real XLSX allows a `localSheetId`-scoped (sheet-local) name to shadow a workbook-scoped
name of the same string — a flat `HashMap<String, String>` cannot represent that collision
at all, let alone resolve it correctly. `xlsx_defined_names` (the reader-level parser)
collects both kinds under their own `name` attribute in document order; `Vm::defined_names`
collapses them with whichever the reader encounters LAST silently winning. Disclosed, not
solved — no fixture in this repo exercises the collision case, so there's no concrete
shape to design around yet.

## Implementation notes for P2: sheet_state

Gaps and deliberate scope boundaries discovered while implementing `sheet_state`,
disclosed here rather than silently absorbed or fixed as a side effect of an unrelated
feature:

**Confirmed a real, independent, pre-existing bug: whole-tab visibility was not read OR
written anywhere before this round.** `xlsx_workbook_sheets` (the `<sheet>`-element
parser) never captured the `state="..."` attribute, and the writer's `<sheets>` emission
never wrote it. The practical consequence: loading any real file with a hidden or
veryHidden sheet and saving it — even a completely no-op save — silently un-hid every
sheet. This has nothing to do with `sheet_state` shipping this round; it was already
broken and would still be broken if this round had shipped nothing at all. Pinned by a
differential-python test (`test_sheet_state_does_not_yet_survive_an_elixcee_save`) that
asserts the CURRENT broken behavior explicitly, so a future writer fix is a deliberate,
visible change to that test rather than a silent behavior shift nobody notices.

**Zero real fixtures have a hidden/veryHidden sheet, which is why write support
(`set_sheet_state`) is deferred, not shipped alongside the read side.** Checked all 7
pristine fixtures under `compat/oracle-excel-com/fixtures/` and everything under
`compat/corpus/` — none has `state=` on any `<sheet>` element. This project's hard gate
(no writer code for a structural OOXML element without real fixture evidence) isn't met.
A path to get real evidence was found but not taken this round: Mac Excel's AppleScript
dictionary (`Excel.sdef`) exposes a worksheet-level `visible` property (`XlSheetVisibility`:
`sheet visible`/`sheet hidden`/`sheet very hidden`), live-verified to read/write correctly
with no VBA-project access needed — but saving through it hits macOS's sandboxed
"Grant File Access" dialog, which needs one human click and can't be scripted around.
Unblocking it needs either a one-time manual grant (open Excel, do one normal Save into
the target fixture folder) or a hand-built fixture (2-3 sheets, one hidden, one
veryHidden, via Excel's UI or `ActiveWorkbook.Sheets("X").Visible = xlSheetVeryHidden`).

**The read side's test fixture is synthetic (hand-built XML via `ZipWriter`), matching
this test file's own established pattern for shapes no real fixture demonstrates** — the
same technique `synthetic_three_sheet_workbook` already uses for rename/move-order tests.
A separate `synthetic_three_sheet_workbook_with_states` helper was added rather than
extending the existing one with a new parameter, which would have touched 4+ unrelated
call sites for a shape they don't care about. The differential-python side instead builds
its fixture with openpyxl itself (which can freely write `ws.sheet_state = "hidden"`) —
simpler than replicating the same synthetic-ZIP technique in Python, and openpyxl is
already the established test-only oracle for this whole file.

**Design choices, matching established conventions rather than copying openpyxl
blindly:** `sheet_state(name)` is name-addressed like `rename_sheet`/`copy_sheet`/
`delete_sheet` (not "current sheet"-defaulted like `hidden_rows`/`hidden_columns`), since
visibility is inherently a question about a specific, often non-active, sheet. It raises
`ValueError` on an unknown name rather than silently returning `"visible"` — openpyxl's
own `ws.sheet_state` can't make that distinction at all (it's a plain attribute read off
an already-resolved `Worksheet` object), but this project's own "explicit error over
silent wrong behavior" convention (`sort_range`'s `key_col`, `merge_cells`'s address
bounds) applies here too. The string vocabulary (`"visible"`/`"hidden"`/`"veryHidden"`)
matches openpyxl's own exactly, confirmed live during research — no translation needed.
`copy_sheet` was extended to also copy the source's visibility state (its now-ninth
per-sheet map to re-key on `rename_sheet`, eighth to copy on `copy_sheet`), matching the
"copy everything else" precedent every other field on that method already sets, absent
any concrete signal pointing the other way.

## Implementation notes for P2: row height / column width

Gaps and deliberate scope boundaries discovered while implementing `row_height`/
`column_width`, disclosed here rather than silently absorbed or fixed as a side effect of
an unrelated feature:

**Confirmed the writer's gap is worse than `sheet_state`'s.** `sheet_state`'s bug was
conditional in spirit (the writer just never read or wrote the attribute at all, so a
save silently lost it). Row height/column width hit the same root cause but with a
starker consequence: `xlsx_worksheet_xml`'s `<row>`/`<cols>` emission is **fully
regenerated from `Vm.sheet_visibility` alone**, not passthrough, not even an opaque
fragment (unlike `<sheetFormatPr>`/`<sheetViews>`/`<sheetPr>`, which already are carried
as opaque blobs). A loaded file's row heights and column widths are dropped on every
single save, full stop — pinned by a differential-python regression test
(`test_row_height_and_column_width_do_not_yet_survive_an_elixcee_save`) that inspects the
saved file's raw XML directly (not `openpyxl`'s `column_dimensions[letter].width`, which
turned out to auto-vivify a default-13.0 entry on first `[]` access even for a column the
file never set — an openpyxl implementation artifact that would have made the regression
test pass for the wrong reason if trusted blindly).

**Zero real fixtures have a genuine custom row height or column width, confirmed by
direct inspection, not assumed.** A first grep pass across the 7 pristine fixtures
falsely suggested `ht=` existed on `<row>` elements — a substring false positive:
`<sheetFormatPr defaultRowHeight="15">` (a workbook-wide default, already opaquely
preserved, unrelated to any individual row) contains the literal text `ht="`. Redone with
a `<row ...>`-anchored check: zero real `ht=`/`customHeight` attributes anywhere. Only
`fixture1` has any `<col>` element at all, and it's the already-known hidden column D
(`width="0" hidden="1" customWidth="1"`) — `width="0"` is how Excel represents a hidden
column, not a real custom width. This project's hard gate (no writer code for a
structural OOXML element without real fixture evidence) isn't met for either value, so
write support (`set_row_height`/`set_column_width`) is deferred the same way
`set_sheet_state` was.

**Two independent value types, not one enum.** Unlike `sheet_state` (three fixed
variants), row height is inherently per-row (`row_heights: HashMap<u32, f64>`, sparse —
only rows with an explicit `customHeight="1"` height get an entry) and column width is
range-shaped like `hidden_columns` but with a value attached
(`column_widths: Vec<(u32, u32, f64)>`). This pushed `rename_sheet`'s per-sheet-map
re-key count from 9 to 11, and `copy_sheet`'s copied-field count from 7 to 9 — confirmed
live (via openpyxl) that real producers don't always coalesce a run of identically-widthed
columns into one `<col min="B" max="D">` range either: setting the same width on columns
B/C/D one at a time produced three separate single-column `<col>` elements in openpyxl's
own output, not one coalesced range. `column_width_on_sheet`'s linear scan over
`(min, max, width)` triples handles both shapes correctly regardless.

**`customHeight="1"`/`customWidth="1"` are both required for `ht`/`width` to actually
apply in real Excel** — a bare `ht` without the flag (some producers emit this for an
auto-fit row) is not recorded as an explicit height, confirmed via a dedicated unit test
(`xlsx_sheet_cells_ignores_ht_without_custom_height`) rather than assumed from the spec
alone.

**API shape**: `row_height(row, sheet=None) -> Optional[float]` /
`column_width(col, sheet=None) -> Optional[float]`, matching `hidden_rows`/
`hidden_columns`'s sheet-parameterized (not name-required) convention rather than
`sheet_state`'s name-addressed one — row/column-level queries within a sheet are this
project's own established shape for that family, distinct from whole-sheet-level queries.
No bound-checking on `row`/`col` (unlike `set_row_hidden`'s write-side Excel-grid-limit
check): an out-of-range lookup on a sparse map just costs a `None`, the same
no-cost-to-check-large-values precedent R1's `get_range`/`iter_rows` already established
for reads.
