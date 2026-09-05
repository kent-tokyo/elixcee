# @elixcee/xlsx

Experimental, private JavaScript package (`0.0.0-development`); **not published to npm**.
Targets the documented subset of `xlsx@0.18.5` behavior, not a complete drop-in replacement.
Reads use the Rust/WASM bridge; writes use a separate JavaScript OOXML/ZIP writer.
The parent Rust/Python version does not describe this package's publication status.

## Supported surface

| API | Scope |
|---|---|
| `utils.*` | All 33 runtime utility exports in the pinned oracle; intentional security/limit differences are documented |
| `SSF` | Number formatting through `ssf@0.11.2` |
| `read(data, opts)` | Synchronous WASM-backed input, with no asynchronous initialization step |
| `readFile` / `readFileSync` | The same Node-only function; normal filesystem errors propagate |
| `write` | `bookType:"xlsx"` only; output `type:"buffer" / "array" / "base64"` |
| `writeFile` / `writeFileSync` | The same Node-only synchronous function |
| TypeScript | Declarations tested with and without DOM libraries; [classification](../../docs/typescript-compatibility.md) |

Read coverage includes SheetNames, !ref, merges, hidden rows/columns, and cell
`t/v/f/w/z`, including the documented cellDates/cellStyles paths.
Write coverage includes scalar values, dates, formulas, multiple sheets, merges,
visibility, hidden rows/columns, and basic number formats.
It is not the native writer's arbitrary-part/VBA preservation path.

Browser file APIs are present but throw `ELIXCEE_UNSUPPORTED_IN_BROWSER`.
Use FileReader/fetch to supply bytes, and download the result of `write` yourself.
Unsupported output book types throw `ELIXCEE_UNSUPPORTED_BOOK_TYPE`.

## Boundaries

- No `writeFileAsync` or streaming API. Rust/Python streaming does not imply JS streaming.
- No ODS output or other `write` book types; utility text exporters are separate.
- No complete Excel/VBA emulation through this JavaScript surface.
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
`npm run wasm:smoke`, `npm run audit:pack`, and `npm run pack:consumer`.
`npm run browser:smoke` requires the browser environment used by its script.
These commands require installed development dependencies and generated WASM artifacts.

[Compatibility harness](../../compat/differential/),
[known differences](../../docs/compatibility-known-defects.md), and
[security model](../../docs/xlsx-security-model.md) define the comparison scope.
Historical test counts are not a guarantee about the current checkout.

## License and architecture

[MIT license](LICENSE); [third-party notices](THIRD_PARTY_NOTICES.md).
See [architecture](../../docs/xlsx-architecture.md) for the Rust/Python/JS boundaries.
