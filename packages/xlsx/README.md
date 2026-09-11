# @elixcee/xlsx

Experimental, private JavaScript workbook package (`0.0.0-development`); **not published to npm**.
It is one package surface of elixcee's broader headless workbook automation
project, not the native Rust/Python VBA runtime and not a complete Excel clone.
Targets the documented subset of `xlsx@0.18.5` behavior, not a complete drop-in replacement.
Reads use the Rust/WASM bridge; writes use a separate JavaScript OOXML/ZIP writer.
The parent Rust/Python version does not describe this package's publication status.

The browser Playground is a separate client-only application built on this
surface. It adds the Excel-like UI (range selection, sheet tabs, formatting,
filters, tables, chart preview, worksheet-backed Pivot summaries, and downloads);
those UI features are not additional exports from this package.

## Supported surface

| API | Scope |
|---|---|
| `utils.*` | All 33 runtime utility exports in the pinned oracle; intentional security/limit differences are documented |
| `SSF` | Number formatting through `ssf@0.11.2` |
| `read(data, opts)` | Synchronous WASM-backed input, with no asynchronous initialization step |
| `readFile` / `readFileSync` | The same Node-only function; normal filesystem errors propagate |
| `write` | `bookType:"xlsx"`, or `"xlsm"` with opaque `!vbaProject` bytes; output `type:"buffer" / "array" / "base64"` |
| `writeFile` / `writeFileSync` | The same Node-only synchronous function |
| TypeScript | Declarations tested with and without DOM libraries; [classification](../../docs/typescript-compatibility.md) |

Read coverage includes SheetNames, !ref, merges, hidden rows/columns, and cell
`t/v/f/w/z`, including the documented cellDates/cellStyles paths.
The Rust/WASM read bridge also exposes a read-only `!dataValidations` projection
with validation `type` and 1-based `sqref` ranges. It is structural metadata;
it does not evaluate formulas or validate cell values. Basic bar/line charts are
projected to `!charts` with source range, title, legend, and anchor dimensions;
unsupported chart features remain outside this projection.
Write coverage includes scalar values, dates, formulas, multiple sheets, merges,
visibility, hidden rows/columns, basic number formats, and limited bar/line chart
parts including multiple charts per worksheet.
When `!vbaProject` is present, XLSM output preserves `xl/vbaProject.bin` and the
macro-enabled package relationships. The browser treats those bytes as opaque:
it does not parse or execute VBA. This remains a bounded preservation path, not
lossless preservation of every OOXML part.

Browser file APIs are present but throw `ELIXCEE_UNSUPPORTED_IN_BROWSER`.
Use FileReader/fetch to supply bytes, and download the result of `write` yourself.
Unsupported output book types throw `ELIXCEE_UNSUPPORTED_BOOK_TYPE`.

The private `@elixcee/xlsx/runtime` entry also exposes `executeOperationPlan` and
`executeOperationPlanOnEditor`.
This is a data-only automation boundary: plans support `setNumber`, `setString`,
and `setBoolean` cell writes with 1-based coordinates, matching
`cell.write.number`, `cell.write.string`, or `cell.write.boolean` capabilities,
operation/JSON-size budgets, and dry-run by default. It never evaluates
caller-supplied code or performs external I/O; pass `apply: true` to apply an
already validated plan.
`executeOperationPlanOnEditor` applies the same plan to the WASM `WorkbookEditor`
inside one transaction; it remains a dry-run unless `apply: true` is explicit.

`createOperationPluginRegistry()` provides the corresponding restricted plugin
boundary. A plugin is only a named, immutable operation plan: it may contain
typed cell writes, declared capabilities, and finite operation/JSON budgets.
Registration rejects executable callbacks and unknown schema fields, execution
requires the declared capabilities, and dry-run remains the default. This is
deliberately not an arbitrary JavaScript plugin system and performs no external
I/O.

## Boundaries

- No `writeFileAsync` or streaming API. Rust/Python streaming does not imply JS streaming.
- No ODS output or other `write` book types; utility text exporters are separate.
- No complete Excel/VBA emulation through this JavaScript surface.
- Chart read/write is limited to the documented bar/line projection; it is not
  a lossless chart or Pivot cache editor.
- `sheet_to_html` escapes `cell.h` by default. Use `rawHtml:true` only for
  independently trusted markup; it is an intentional security extension.
- Whole-buffer synchronous reads apply default reader limits. Use a worker when
  hard cancellation is required; worker termination is not cooperative reader cancellation.

## Runtime and bundling

Node >=18 is required. Both WASM loaders embed their bytes; no separate asset-copy
step is needed. Browser consumption assumes a bundler, not a bare script import.

| Environment | Constraint |
|---|---|
| Native Node CJS/ESM imports | Synchronous reader and file APIs |
| esbuild Node CJS bundle | Reader, file APIs, and writer |
| esbuild Node ESM bundle | Keep this package external when using file APIs or write, so Node resolves fs/zlib |
| Browser bundle | In-memory read/write; ZIP output is STORED (uncompressed), unlike Node DEFLATE |
| Browser validation | Chromium smoke coverage; Safari is not verified |

## Verification

From this package directory, run `npm run typecheck`, `npm run typecheck:no-dom`,
`npm run wasm:smoke`, `npm run xlsm:smoke`, `npm run xlsm:reopen:smoke`,
`npm run comment:smoke`, `npm run pivot:smoke`, and
`npm run audit:pack`, and `npm run pack:consumer`. The XLSM reopen check uses
the repository's real VBA fixture and LibreOffice when installed; it is not an
Excel-oracle result.
`npm run browser:smoke` requires the browser environment used by its script.
These commands require installed development dependencies and generated WASM artifacts.

[Compatibility harness](../../compat/differential/),
[known differences](../../docs/compatibility-known-defects.md), and
[security model](../../docs/xlsx-security-model.md) define the comparison scope.
Historical test counts are not a guarantee about the current checkout.

## License and architecture

[MIT license](LICENSE); [third-party notices](THIRD_PARTY_NOTICES.md).
See [architecture](../../docs/xlsx-architecture.md) for the Rust/Python/JS boundaries.
