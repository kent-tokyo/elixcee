# elixcee-wasm

`elixcee-wasm` is the Rust/WASM bridge for bounded, headless workbook
operations. It is designed for browser and Node.js integrations that need the
same Rust reader and formula/diagnostic behavior without installing Excel.

This is the reusable runtime behind the project's browser
[Playground](https://kent-tokyo.github.io/elixcee/playground/). The Playground
also contains a separate JavaScript XLSX writer and an Excel-like UI; those are
integration layers, not part of this crate.

## Install

Add the crate to a Rust/WASM application:

```toml
[dependencies]
elixcee-wasm = "0.1"
```

For a browser package, build the bindings with
[`wasm-pack`](https://rustwasm.github.io/wasm-pack/):

```bash
wasm-pack build --target web
```

## Current API

The `wasm-bindgen` exports provide:

- `readWorkbook(bytes)` for bounded XLSX/XLSM/ODS reading;
  simple workbook defined names are returned as `Workbook.Names` with their
  `Name`, `Ref`, and optional zero-based local `Sheet` scope;
- `calculateWorkbook(bytes)` for supported formula recalculation;
- `diagnoseWorkbook(bytes)` for deterministic workbook diagnostics;
- `WorkbookEditor` for bounded typed cell edits, recalculation, snapshots,
  transactions, and undo/redo.

Coordinates exposed by the editor are 1-based, matching Excel and VBA.
Malformed, oversized, or otherwise unsafe input is rejected by the shared Rust
reader limits. The editor also validates worksheet names and cell bounds.

### Minimal browser integration

The crate is intended to be consumed through the `wasm-bindgen` bindings it
generates. The generated package exposes the same API in Node and browsers:

```js
import init, { WorkbookEditor, readWorkbook } from "elixcee-wasm";

await init();
const summary = JSON.parse(readWorkbook(workbookBytes));
const editor = new WorkbookEditor(workbookBytes);
editor.setNumber("Sheet1", 1, 1, 42);
const calculated = JSON.parse(editor.recalculate());
```

`readWorkbook` and `WorkbookEditor` return JSON-shaped workbook data so an
application can provide its own grid, clipboard, accessibility model, and
serialization policy. The crate deliberately does not prescribe an Excel-like
UI. For a complete example, see the repository's
[Playground](https://kent-tokyo.github.io/elixcee/playground/) and its
[source](https://github.com/kent-tokyo/elixcee/tree/main/playground).

## Scope boundary

This crate is a reusable runtime bridge, not an Excel UI and not a complete XLSX writer.
It is the crate-level distribution target for the Playground's reusable
browser runtime; the Excel-like screen itself remains a separate static web
application. In other words, the Playground's read/calculation/diagnostics/
bounded typed-edit capability is reusable as a crate, while its grid and
JavaScript XLSX writer are separate integration layers.
The current JavaScript package owns XLSX serialization and the Playground owns
the spreadsheet UI. The Playground is an integration example, not part of this
crate's API. It does not execute arbitrary VBA, access the filesystem,
network, COM, Shell, `MsgBox`, or `UserForm`. Existing VBA project bytes may be
preserved by supported JavaScript/native round-trip paths, but preservation is
not arbitrary macro execution.

Use the repository's [support contract](https://github.com/kent-tokyo/elixcee/blob/main/docs/v1-support-contract.md) and
[XLSX architecture](https://github.com/kent-tokyo/elixcee/blob/main/docs/xlsx-architecture.md) for the exact compatibility
and safety boundaries. The crate is prepared for publication after the
 matching native `elixcee` release is available on crates.io. Its current
 checkout may use a newer native API than the last published `elixcee` version,
 so the release workflow publishes the native crate first and waits for index
 visibility before publishing this bridge. Publication remains an explicit
 release operation.

## Building

```bash
cargo check -p elixcee-wasm --target wasm32-unknown-unknown
wasm-pack build crates/elixcee-wasm --target web
```

The repository build scripts inline the generated WASM into the browser and
Node package entry points. Consumers should use those package artifacts until
the crate's JS API and generated package layout are declared stable.

See the [MIT license](https://github.com/kent-tokyo/elixcee/blob/main/LICENSE).
