# XLSX architecture

## Current implementation

The root `elixcee` crate owns the Rust workbook model, hand-written XML reader,
formula engine, VBA parser/VM, and ZIP-based writer. `elixcee-types` contains
shared value types. The JavaScript package uses `elixcee-wasm` for reading;
its public write APIs use a separate JavaScript OOXML/ZIP writer.

Runtime dependencies are intentionally small: `zip` is used for workbook
containers and PyO3 is optional for the Python feature. XML parsing and the VBA
parser are hand-written. This keeps the CLI, Python extension, and WASM build
on the same Rust core for native operations. The JS facade has its own API and writer boundaries.

## Data flow

Native: workbook → Rust reader → VM/workbook model → Rust writer → file.
CLI and Python wrap that native path.

JavaScript: bytes → Rust/WASM reader → JS workbook object → JS writer → bytes.
Node file APIs wrap these byte APIs; browsers use in-memory APIs only.

The VM uses 1-based row and column coordinates at its public boundaries, as
Excel/VBA does. Cell values and formulas are stored separately so formulas can
be recalculated when calculation mode permits. Sheet keys are resolved
case-insensitively for VBA behavior.

## Preservation policy

The writer regenerates the parts it models and passes through many unknown ZIP
parts. It currently models cell values, formulas, styles, merges, hidden
rows/columns, workbook metadata, and selected worksheet objects such as tables,
filters, and validations. A passthrough part is not enough by itself: its
relationship must also remain connected to the regenerated owner part.
For worksheet drawing and legacy-drawing owners, save now checks the owner
relationship id, the worksheet rels declaration, its normalized internal target,
and target-part survival before restoring the opaque owner fragment. The check
also walks surviving internal part relationships, so Drawing-to-Chart/image
targets are covered without fetching external URLs. Pivot cache validation
remains future work.

Macro-enabled workbooks preserve `xl/vbaProject.bin` and the macro-enabled
content type on supported save paths. Unmodeled worksheet objects can still be
lost or disconnected, so the README's compatibility warning takes precedence
over any assumption of lossless editing.

The [G0–G6 strengthening plan](../ROADMAP.md) separates connected OOXML
preservation from object editing/recalculation, workbook-wide formula evaluation,
and memory-bounded output. In particular, passing through a Pivot cache or external
link part currently does not restore the omitted `pivotCaches` or
`externalReferences` owner elements in regenerated workbook XML.

The native VM writer now retains worksheet/workbook/styles/rels XML and only
edited table XML for relationship and structural analysis, while copying other XML and non-XML
passthrough payloads directly from the source ZIP one at a time during output,
without a payload-sized intermediate buffer.
Read-only defined-name loading is narrower still: it validates the source ZIP and
reads only `xl/workbook.xml`, without indexing sibling worksheet or table payloads.
Unedited table parts follow the same deferred path; only a table with a pending
structural edit is read back for patching.
Python's
append-only writer already writes each accepted row to ZIP, but materializes one
row and its XML; its cumulative byte counter is not retained-memory telemetry.

Workbook formula evaluation now has two explicit paths: the existing fast
single-sheet evaluator for unqualified references, and a workbook slow path for
qualified references. The latter builds a cross-sheet formula order and maps
each sheet into a separate internal coordinate band before invoking the same
range/function evaluator. Its full-recalculation entry point is also exported
as `calculateWorkbook(bytes)` from the shared WASM bridge for Node/browser
consumers. `diagnoseWorkbook(bytes)` provides a deterministic preflight JSON
summary (sheet/formula counts, qualified formulas, and parse errors). Incremental
JS calculation and worker/async loading contracts remain future work.
The optional `@elixcee/xlsx/runtime` subpath exposes the same shared calculation
functions plus stateful `WorkbookEditor` in Node and browser package conditions.
The vendored WASM payload has an intentional baseline and a 10% growth gate in
`scripts/wasm-smoke.mjs`; a baseline update requires a reviewed implementation
change.
G1 tightens admission checks; deferred G5 addresses lazy passthrough, including
unchanged table parts, row/work
budget separation, failure cleanup, and independently measured memory scaling.

## JavaScript/WASM package

`packages/xlsx` provides synchronous `read`/`readFile`/`readFileSync` and
`write`/`writeFile`/`writeFileSync` APIs for XLSX. The browser entry point uses
embedded WASM and is intended for bundled applications. The package is kept
private and is not published yet.

The native Rust and Python readers expose the same cooperative read controls:
the default total work budget is 2 GiB-equivalent units, Python additionally
exposes `timeout_ms` and `ReadCancellation`, and CLI snapshot exposes work-budget/deadline/cancel-file controls.
Run-mode --file uses the default budget with signal cancellation. The WASM `readWorkbook(bytes)` export remains
synchronous and takes only the input buffer; it applies the reader's default
limits but cannot observe a JavaScript cancellation request during the call.
Applications that need hard cancellation should run the synchronous call in a
worker and terminate that worker, treating termination as distinct from the
reader's cooperative `READER_CANCELED` result.

## Security boundaries

Workbook files are untrusted input. The reader enforces per-entry, archive-wide,
XML, and total-work budgets before returning a workbook. The native/Python
cooperative cancellation boundary is checked at ZIP read chunks and parsed-part
boundaries; a blocking filesystem call cannot be forcibly interrupted. The
JavaScript compatibility layer deliberately rejects selected dangerous or
resource-exhausting inputs even when the reference package accepts them. See
[xlsx-security-model.md](xlsx-security-model.md).

## Verification

Rust unit/integration/property tests, compatibility fixtures, and JavaScript
differential tests run in CI. Real Excel-authored fixtures cover selected
round-trip paths. They do not establish complete Excel VBA semantic
compatibility or guarantee that every unmodeled OOXML feature survives a save.
