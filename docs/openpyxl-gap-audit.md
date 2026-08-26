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
| Rename an existing sheet | 4 | 3 | 5 | 2 | 2 | yes (all 7) | yes (`WorksheetOrigin`) | **P1** |
| Move/reorder an existing sheet | 3 | 2 | 4 | 2 | 2 | yes | yes (`sheet_order`) | **P1** |
| Copy a sheet (same workbook only) | 3 | 2 | 3 | 3 | 3 | partial | partial | **P2** |
| Sheet visibility (hidden/veryHidden) | 2 | 2 | 4 | 2 | 2 | no | no | **P2** |

`set_sheet`/`delete_sheet` already cover create/select/delete. Rename and move are the
two genuinely missing lifecycle primitives with a clean, low-risk implementation path:
`sheet_order` (`src/vm/mod.rs:598`) already models tab order, and `WorksheetOrigin`
already threads a sheet's original display name/part name through to the writer — a
rename is "update the key in `sheets`/`sheet_order`/`worksheet_origins` consistently," a
reorder is "permute `sheet_order`." Both are pure in-memory `Vm` state changes with no new
OOXML element involved, matching R1's own risk profile — good P1 candidates for the very
next round. Copy-in-same-workbook needs to duplicate a sheet's `WorksheetOrigin`/cell map/
merges/hidden-row state consistently, a bit more surface area, hence P2.

## 2. Cell / range access

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Rectangular range read | 5 | 5 | 5 | 1 | 1 | n/a (in-memory) | yes | **R1** (`get_range`) |
| Rectangular range write | 5 | 5 | 5 | 2 | 1 | n/a | yes | **R1** (`set_range`) |
| Append a row | 4 | 5 | 5 | 1 | 1 | n/a | yes | **R1** (`append_row`) |
| Row iteration (values-only) | 4 | 4 | 5 | 1 | 1 | n/a | yes | **R1** (`iter_rows`) |
| Used-range bounds / dimension string | 4 | 4 | 5 | 1 | 1 | n/a | yes | **R1** (`max_row`/`max_column`/`calculate_dimension`) |
| Column iteration (`iter_cols`) | 2 | 2 | 4 | 1 | 1 | n/a | yes | **P1** |
| `Cell` object model (style/comment/hyperlink attached to a returned cell) | 2 | 2 | 2 | 4 | 2 | no | no | **Not planned** |

This is the category R1 closes almost entirely — it needs zero writer/OOXML changes (a
pure in-memory `Vm` cell-map read/write, serialized unchanged by the existing writer),
which is why it's the lowest-risk category in this whole audit by construction, and why
it was picked first regardless of what the rest of this table says. `iter_cols` is a
trivial follow-up (same primitives, transposed) once R1 lands, deferred only because it
wasn't in the user's original R1 request. A full openpyxl-style `Cell` object (an object
that carries `.value`/`.number_format`/`.font`/`.comment`/`.hyperlink` all at once) is
explicitly **not planned** — it would require the full style/comment/hyperlink write
architecture from categories 5/10 to exist first, and would invert `elixcee`'s existing
value contract (plain Python values, not wrapper objects) for every consumer of
`get_cell`/`cells()` at once.

## 3. Row / column operations

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Insert/delete rows/cols (Python-native) | 3 | 3 | 4 | 2 | 1 | n/a | yes (`Vm::delete_rows`/`insert_rows`/`delete_cols`/`insert_cols`) | **P1** |
| Row height / column width (read/write) | 2 | 2 | 3 | 3 | 3 | no | no | **P2** |
| Hidden row/col (read/write, Python-native) | 2 | 2 | 4 | 2 | 2 | yes | yes (`sheet_visibility`) | **P2** |
| Outline/grouping level | 1 | 1 | 2 | 3 | 3 | no | no | **Not planned** |

Row/column insert-delete already has a full, tested implementation at the VBA-execution
layer (`Vm::delete_rows`/`insert_rows`/`delete_cols`/`insert_cols`, added for GitHub #7/#8
in `0.11.0`) — a Python-native `insert_rows(idx, count)`/`delete_rows(idx, count)` pair
would be nearly free glue on top of those exact functions, making this a strong P1
candidate once R1 ships. Row height/column width have no internal representation at all
today (not read, not stored, not written) — real new surface area, not glue, hence P2.
Grouping/outline levels have no evidence of demand and touch a genuinely unexplored part
of the schema — not planned absent a concrete request.

## 4. Formula / error / date handling

Already broadly comparable, no large gap. `set_cell_formula`/`set_cell_formula_batch`/
`recalculate()` cover formula writing; `ExcelError` (a typed object, `code` attribute)
round-trips real error cells since `0.11.0`'s `t="e"` work on `master` — arguably a
stronger contract than openpyxl's own (which surfaces an error cell's cached value as a
plain string, with no typed wrapper). Date handling is openpyxl's weaker spot too, not
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
| Read a sheet's merged ranges | 3 | 3 | 4 | 1 | 1 | yes (`fixture1`) | yes (`merged_ranges`) | **P1** |
| Create/remove a merge | 3 | 2 | 3 | 2 | 3 | yes | partial | **P1** |

`merged_ranges` (`src/vm/mod.rs:693`) is already fully populated from any loaded file and
already round-trips on save — there is simply no Python getter for it at all today (not
even read-only). A read-only `merged_cells(sheet=None) -> list[str]` is nearly free glue,
identical in spirit to R1's own `get_range`. Creating a *new* merge needs the writer to
correctly emit a `<mergeCell>` for a range that didn't have one in the source, which is
new writer surface — still low/contained risk (the shape is well understood, unlike
styles), but real, hence bundled with the read side into the same P1 slice rather than
split.

## 7. Defined names

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Read defined names | 3 | 2 | 4 | 1 | 1 | yes (`fixture4`) | partial | **P2** |
| Create/delete a defined name | 2 | 1 | 3 | 3 | 3 | yes | no | **P2** |

Defined names already survive round-trip (`0.10.0-C` slice 3, delete-gated passthrough) —
confirmed preserved, but there is no Python-visible representation of them at all (not
even internal to `Vm` as a queryable structure; they live purely as passthrough XML
today). Reading requires actually parsing `<definedNames>` into a real internal
structure for the first time — real, if modest, new work, hence P2 not P1.

## 8. Tables / AutoFilter / sorting

| Feature | Value | Freq | Fit | Impl risk | Corrupt risk | Fixture evid. | Primitive reuse | Milestone |
|---|---|---|---|---|---|---|---|---|
| Python-native `.sort_range(...)` | 2 | 2 | 3 | 1 | 1 | n/a | yes (`Vm`'s VBA `RangeSort` handler) | **P1** |
| Python-native AutoFilter (VM-effect only, matching `0.11.0`'s VBA scope) | 2 | 2 | 3 | 2 | 2 | n/a | yes | **P2** |
| Table (`<table>` part) creation | 2 | 2 | 2 | 4 | 4 | no (no fixture has a from-scratch table) | no | **P3 / user's own explicit deferral** |

Sort and AutoFilter already exist as VBA statements (`0.11.0`); a thin Python-native
wrapper calling the same internal handlers is cheap, matching this audit's general theme
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

Hyperlinks (both internal and external) already round-trip via `0.10.0-B`/`D`'s
relationship-backed restoration; there is still no Python-visible read of them. Comments
are VML-backed (`<legacyDrawing>`) — a genuinely more awkward OOXML shape than a plain
`r:id` hyperlink, and higher corruption risk to write, hence P3 over P2.

## 11. Charts / images / drawings

Read-only passthrough preservation exists (`0.10.0-D`, relationship-graph reachability);
zero creation/modification API of any kind. Explicitly out of scope for this round's
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

1. **R1 (this round)**: bulk worksheet range/row API — category 2, closes it almost
   entirely, zero writer changes, lowest risk in this document.
2. **P1**: sheet rename/move (category 1), Python-native row/col insert-delete wrapping
   the existing `0.11.0` VBA handlers (category 3), read-only merged-cell access
   (category 6) — all cheap glue over primitives that already exist and are already
   tested at the VBA layer.
3. **P2**: sheet copy/visibility (category 1), hidden row/col + width/height (category 3),
   number-format-only writing (category 5, the narrowest possible slice of the style
   engine), defined-name read/write (category 7), Python-native AutoFilter (category 8),
   hyperlink read/write (category 10).
4. **P3**: font/fill/border/alignment *read* (not write), comments, page-setup read, sheet
   protection exposure (each needs its own follow-up design decision, noted above).
5. **Not planned**: full style-engine writing, named styles, table/data-validation/
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
