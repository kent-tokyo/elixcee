# Resource limits

This file records the empirical basis for elixcee's own resource limits — each one added
only after measuring a real, confirmed cost on the real oracle (or, for `src/reader.rs`,
a crafted malicious input), never speculatively. See
[`docs/xlsx-security-model.md`](xlsx-security-model.md) for the full limits inventory and
threat model context; this file is where the *sizing* decision for each limit is worked
through and kept up to date as new measurements arrive.

## `packages/xlsx`: `MAX_RANGE_CELLS` (`ELIXCEE_RANGE_TOO_LARGE`)

`sheet_to_formulae`, `sheet_to_csv`, `sheet_to_txt` (which delegates to `sheet_to_csv`),
`sheet_to_json`, and `sheet_to_html` all walk every `(row, col)` pair inside a worksheet's
`!ref` rectangle regardless of sparsity. `packages/xlsx/src/internal/range-guard.cjs`
rejects ranges above `MAX_RANGE_CELLS = 5,000,000` cells before iterating them.

### Measurement (2026-08-16, one fixed pass — not repeated)

A sparse worksheet (only 2 populated cells: top-left and bottom-right) with `!ref` sized
to exactly the target cell count, measured in a fresh subprocess per case
(`process.memoryUsage().rss` sampled immediately after the call returns; wall time via
`Date.now()`):

| Range size | Function | Oracle: time | Oracle: RSS | elixcee: time | elixcee: RSS |
|---|---|---|---|---|---|
| 100,000 | `sheet_to_formulae` | 38 ms | 86 MB | 40 ms | 58 MB |
| 100,000 | `sheet_to_csv` | 39 ms | 89 MB | 67 ms | 60 MB |
| 100,000 | `sheet_to_txt` | 36 ms | 90 MB | 72 ms | 69 MB |
| 1,000,000 | `sheet_to_formulae` | 365 ms | 122 MB | 385 ms | 91 MB |
| 1,000,000 | `sheet_to_csv` | 385 ms | 126 MB | 404 ms | 97 MB |
| 1,000,000 | `sheet_to_txt` | 382 ms | 130 MB | 457 ms | 147 MB |
| 5,000,000 | `sheet_to_formulae` | 2,176 ms | 248 MB | not measured (blocked by the guard — see below) | |
| 5,000,000 | `sheet_to_csv` | 2,391 ms | 250 MB | not measured | |
| 5,000,000 | `sheet_to_txt` | 2,168 ms | 260 MB | not measured | |
| 10,000,000 | `sheet_to_formulae` | 4,897 ms | 229 MB | not measured | |
| 10,000,000 | `sheet_to_csv` | 5,950 ms | 329 MB | not measured | |
| 10,000,000 | `sheet_to_txt` | 5,277 ms | 346 MB | not measured | |

elixcee wasn't measured at 5M/10M directly (the guard itself blocks those calls through
the public API — measuring would require bypassing it). At 100K/1M, elixcee's wall time
and RSS are in the same order of magnitude as the oracle's (same O(rows × cols) walk
shape, same language), which is the basis for treating the oracle's 5M/10M numbers as a
reliable proxy for what elixcee's own cost would be at those sizes if unguarded.

### Decision: keep `MAX_RANGE_CELLS = 5,000,000`

At the threshold itself, the cost is ~2.2-2.4s and ~250MB RSS — noticeably slow for a
single synchronous call but not a severe hang, and the point beyond which cost keeps
climbing linearly with no natural ceiling (10M already reaches 5-6s / up to ~345MB, and
the original full-grid probe — `A1:XFD1048576`, ~17.18 billion cells — did not return
within 25s at all). 5,000,000 cells is also far beyond what any realistic populated
worksheet needs for these in-memory JS APIs — the number is chosen to reject
pathologically large ranges specifically, not to constrain normal use. No adjustment from
the original value.

## `src/reader.rs`: ZIP archive limits

The reader applies four fixed, conservative ZIP limits before consuming workbook XML:

| Limit | Value | Constant |
|---|---:|---|
| ZIP entry count | 10,000 | `ZIP_MAX_ENTRIES` |
| Per-entry decompressed size | 256 MiB | `ZIP_ENTRY_MAX_BYTES` |
| Total decompressed size | 1 GiB | `ZIP_MAX_TOTAL_BYTES` |
| Per-entry compression ratio | 1,000:1 | `ZIP_MAX_COMPRESSION_RATIO` |

After the ZIP checks, every XML part is subject to these document-level limits:

| Limit | Value | Constant |
|---|---:|---|
| Elements per document | 1,000,000 | `XML_MAX_ELEMENTS` |
| Attributes per document | 2,000,000 | `XML_MAX_ATTRIBUTES` |
| Attribute value length | 16 MiB | `XML_MAX_ATTRIBUTE_VALUE_BYTES` |
| Text node length | 64 MiB | `XML_MAX_TEXT_NODE_BYTES` |
| Nesting depth | 1,024 | `XML_MAX_DEPTH` |

The materialized workbook model also has these limits:

| Limit | Value | Constant |
|---|---:|---|
| Sheets per workbook | 4,096 | `WORKBOOK_MAX_SHEETS` |
| Cells per sheet | 5,000,000 | `SHEET_MAX_CELLS` |
| Merged ranges per sheet | 1,000,000 | `SHEET_MAX_MERGES` |
| Shared strings | 1,000,000 entries / 256 MiB | `SHARED_STRINGS_MAX_*` |

Formula parsing applies these limits before exposing an AST to the evaluator or reference
rewriter:

| Limit | Value | Constant |
|---|---:|---|
| Formula input | 1 MiB | `MAX_FORMULA_BYTES` |
| Formula references | 100,000 | `MAX_FORMULA_REFS` |
| Formula AST nodes | 200,000 | `MAX_FORMULA_NODES` |
| Formula nesting depth | 256 | `MAX_FORMULA_DEPTH` |

VBA execution applies a deterministic instruction budget of 10,000,000 statements or
loop iterations per run by default (`DEFAULT_MAX_VBA_INSTRUCTIONS`). It also limits nested
Sub/Function calls to 256 frames (`DEFAULT_MAX_VBA_CALL_DEPTH`). Trusted Rust callers can
explicitly set either `Vm::max_instructions` or `Vm::max_call_depth` to `None` to opt out.
Each retained VBA string is limited to 16 MiB (`DEFAULT_MAX_VBA_STRING_BYTES`), and each
runtime/VBA array is limited to 10,000,000 elements (`DEFAULT_MAX_VBA_ARRAY_ELEMENTS`).
These value budgets are checked on VBA assignments and cell writes; budget errors are not
swallowed by `On Error Resume Next`.
VBA-generated workbook state is also limited to 5,000,000 materialized cells across all
sheets (`DEFAULT_MAX_VBA_CELLS`).
Python callers can use `Vm.set_budgets()` to adjust the limits; omitted arguments use the
safe defaults, while an explicit `None` disables that individual limit.

VBA parsing applies these input limits before constructing a program AST:

| Limit | Value | Constant |
|---|---:|---|
| VBA source | 4 MiB | `MAX_VBA_SOURCE_BYTES` |
| VBA identifier | 1,024 characters | `MAX_VBA_IDENTIFIER_CHARS` |
| VBA tokens | 1,000,000 | `MAX_VBA_TOKENS` |

These are build-time safeguards, not claims that arbitrary hostile files are safe. The
limits are checked from ZIP metadata before part parsing; path traversal is rejected at
the same boundary, and DTD/ENTITY declarations are rejected. Numeric thresholds are
intentionally conservative until the planned normal-large-file and malicious-fixture
measurements establish a representative corpus.
