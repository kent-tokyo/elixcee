# XLSX architecture

## Current implementation

The root `elixcee` crate owns the Rust workbook model, hand-written XML reader,
formula engine, VBA parser/VM, and ZIP-based writer. `elixcee-types` contains
shared value types. `elixcee-wasm` exposes the reader/writer to the JavaScript
package in `packages/xlsx`.

Runtime dependencies are intentionally small: `zip` is used for workbook
containers and PyO3 is optional for the Python feature. XML parsing and the VBA
parser are hand-written. This keeps the CLI, Python extension, and WASM build
on the same core implementation.

## Data flow

```text
.xlsx/.xlsm/.ods
        │
        ▼
  Rust reader ──► Vm / workbook model ──► Rust writer ──► workbook file
                         │
                         ├── CLI
                         ├── Python (PyO3)
                         └── WASM → @elixcee/xlsx
```

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

Macro-enabled workbooks preserve `xl/vbaProject.bin` and the macro-enabled
content type on supported save paths. Unmodeled worksheet objects can still be
lost or disconnected, so the README's compatibility warning takes precedence
over any assumption of lossless editing.

## JavaScript/WASM package

`packages/xlsx` provides synchronous `read`/`readFile`/`readFileSync` and
`write`/`writeFile`/`writeFileSync` APIs for XLSX. The browser entry point uses
embedded WASM and is intended for bundled applications. The package is kept
private and is not published yet.

## Security boundaries

Workbook files are untrusted input. The reader enforces a 64 MiB decompressed
size limit per ZIP entry; broader archive and XML budgets remain planned. The
JavaScript compatibility layer deliberately rejects selected dangerous or
resource-exhausting inputs even when the reference package accepts them. See
[xlsx-security-model.md](xlsx-security-model.md).

## Verification

Rust unit/integration/property tests, compatibility fixtures, and JavaScript
differential tests run in CI. Real Excel-authored fixtures cover selected
round-trip paths. They do not establish complete Excel VBA semantic
compatibility or guarantee that every unmodeled OOXML feature survives a save.
