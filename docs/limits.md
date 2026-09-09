# Resource limits

This file lists enforced resource ceilings and their dated calibration evidence.
Some thresholds are conservative safeguards, not fully calibrated capacity guarantees.
See the [security model](xlsx-security-model.md) for threats and
[measurements](measurements/README.md) for platform-specific evidence.

Path-based workbook input is restricted to `.xlsx`, `.xlsm`, and `.ods` (case-insensitive)
before opening the file. Unsupported or missing extensions return a deterministic error;
the in-memory OOXML buffer reader is intentionally extension-independent.

## Python append-only writer (1.0.6 G1)

`create_stream` writes accepted rows to a temporary ZIP, retaining one row's values
and generated XML during `append`, not all accepted rows until `close`.
The existing `max_pending_bytes` default (64 MiB) remains a **cumulative admission
budget**, not an RSS cap. `pending_bytes` counts accepted estimates until close
and then resets to zero. `max_rows` and `max_columns` are optional positive limits.

G1 checks the row-count limit before opening the iterable, then checks column and
byte limits while consuming each item. Detecting excess width consumes at most
one extra item, without converting that item. The estimate includes a `Variant`
slot per cell plus UTF-8 payload; empty strings no longer have zero cost. This
is stricter than the prior 1.0.3 string accounting. Rejected rows do not change
the accepted counters or write partial row XML; valid rows may then be appended.

This does **not** bound caller-owned Python memory, execution time of arbitrary
Python iterators/conversions, Python UTF-8 conversion caches, allocator overhead,
or XML-escaping expansion to exactly that byte count. ZIP buffering and metadata
also consume memory. I/O-error abort/cleanup and row-count-independent RSS remain
separate [G5 work](../ROADMAP.md). Use process isolation for hostile Python code.
The materialized VM writer still holds the workbook and raw passthrough payloads;
neither path is advertised as fully constant-memory.

Regression: `python -I tests/python/test_stream_writer_limits.py` against an
installed development wheel. This is a small-fixture gate, not an RSS measurement.
The separate `create_stream_bounded` API remains bounded/experimental by contract,
but its current Rust/Python build and macOS large-process RSS evidence is recorded
in [the measurement index](measurements/README.md). Cross-platform and CI evidence
remain separate.
The reproducible measurement entry point is
`scripts/measure-stream-writer-memory.py`; it uses one child process per row-count
case and records peak RSS, wall time, output size, and rows. It intentionally does
not infer constant-memory from two cases.

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
| 5,000,000 | `sheet_to_formulae` | 2,176 ms | 248 MB | not measured (historical record; see note below) | |
| 5,000,000 | `sheet_to_csv` | 2,391 ms | 250 MB | not measured | |
| 5,000,000 | `sheet_to_txt` | 2,168 ms | 260 MB | not measured | |
| 10,000,000 | `sheet_to_formulae` | 4,897 ms | 229 MB | not measured | |
| 10,000,000 | `sheet_to_csv` | 5,950 ms | 329 MB | not measured | |
| 10,000,000 | `sheet_to_txt` | 5,277 ms | 346 MB | not measured | |

The historical probe did not measure elixcee at 5M/10M. Its original explanation
that the guard blocks both is incorrect for the current strict `>` comparison:
exactly 5,000,000 cells passes this guard; 10,000,000 does not. Oracle timings are
not measured elixcee timings and must not be substituted for them.

### Decision: keep `MAX_RANGE_CELLS = 5,000,000`

At the threshold itself, the cost is ~2.2-2.4s and ~250MB RSS — noticeably slow for a
single synchronous call but not a severe hang, and the point beyond which cost keeps
climbing linearly with no natural ceiling (10M already reaches 5-6s / up to ~345MB, and
the original full-grid probe — `A1:XFD1048576`, ~17.18 billion cells — did not return
within 25s at all). This ceiling also constrains legitimate large workloads. It is a safety/performance
tradeoff, not a claim about the maximum useful worksheet size. The value is unchanged.

The XML element ceiling can be reached before the model's cell ceiling. The
[1m-cell benchmark](benchmarks/workbook-large-speedup-2026-09-06.md) uses four sheets;
it does not establish support for one million cells on a single sheet.

## `src/reader.rs`: ZIP archive limits

The reader applies four fixed, conservative ZIP limits before consuming workbook XML:

| Limit | Value | Constant |
|---|---:|---|
| ZIP entry count | 10,000 | `ZIP_MAX_ENTRIES` |
| Per-entry decompressed size | 256 MiB | `ZIP_ENTRY_MAX_BYTES` |
| Total decompressed size | 1 GiB | `ZIP_MAX_TOTAL_BYTES` |
| Per-entry compression ratio | 1,000:1 | `ZIP_MAX_COMPRESSION_RATIO` |
| Overall read work budget | 2 GiB-equivalent units by default; declared entry bytes plus 4 KiB per entry | `DEFAULT_READ_MAX_WORK_UNITS`, `ReadOptions` |

After the ZIP checks, every XML part is subject to these document-level limits:

| Limit | Value | Constant |
|---|---:|---|
| Elements per document | 1,000,000 | `XML_MAX_ELEMENTS` |
| Attributes per document | 2,000,000 | `XML_MAX_ATTRIBUTES` |
| Attribute value length | 16 MiB | `XML_MAX_ATTRIBUTE_VALUE_BYTES` |
| Text node length | 64 MiB | `XML_MAX_TEXT_NODE_BYTES` |
| XML control characters | Parser rejects control characters except TAB/LF/CR | `validate_xml_budget` |
| Nesting depth | 1,024 | `XML_MAX_DEPTH` |

The materialized workbook model also has these limits:

| Limit | Value | Constant |
|---|---:|---|
| Sheets per workbook | 4,096 | `WORKBOOK_MAX_SHEETS` |
| Cells per sheet | 5,000,000 | `SHEET_MAX_CELLS` |
| Merged ranges per sheet | 1,000,000 | `SHEET_MAX_MERGES` |
| Shared strings | 1,000,000 entries / 256 MiB | `SHARED_STRINGS_MAX_*` |
| Defined names | 100,000 | `DEFINED_NAMES_MAX_COUNT` |
| Defined-name formula text | 1 MiB | `DEFINED_NAME_MAX_TEXT_BYTES` |

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
Blocked external effects are rejected at runtime with a `SECURITY:` error by default;
ordinary unsupported statements remain no-ops for compatibility.
Python callers can use `Vm.set_budgets()` to adjust the limits; omitted arguments use the
safe defaults, while an explicit `None` disables that individual limit.

VBA parsing applies these input limits before constructing a program AST:

| Limit | Value | Constant |
|---|---:|---|
| VBA source | 4 MiB | `MAX_VBA_SOURCE_BYTES` |
| VBA identifier | 1,024 characters | `MAX_VBA_IDENTIFIER_CHARS` |
| VBA tokens | 1,000,000 | `MAX_VBA_TOKENS` |

These are runtime-enforced safeguards with compiled defaults, not a claim that
arbitrary hostile files are safe. ZIP metadata limits are checked before part parsing;
XML/model/formula/VBA limits are checked in their respective processing stages; path traversal is rejected at
the same boundary, and DTD/ENTITY declarations are rejected. The reader also exposes a
total-work budget, deadline, and cooperative cancellation; native CLI cancellation can
be requested by SIGINT, with a cancel-file option on `snapshot`. Dated macOS large-input, cancellation, and
resource-reclamation measurements are stored in `docs/measurements/`. The thresholds
remain conservative because Linux/Windows and independent-oracle calibration is not yet
complete.
