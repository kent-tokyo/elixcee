# Crate and browser API boundary

elixcee has two reusable Rust artifacts and a separate browser application.
The Rust crates provide workbook runtime behavior; the Playground provides the
Excel-like screen. The screen is not part of either crate.

| Artifact | Reusable capability | Not included |
| --- | --- | --- |
| `elixcee` | Native Rust reader, formula calculation, workbook editing, VBA data-processing runtime, diagnostics, streaming I/O | Browser UI, arbitrary VBA compatibility, lossless OOXML editing |
| `elixcee-wasm` | wasm-bindgen bridge for bounded read, formula calculation, diagnostics, typed cell editing, transactions, undo/redo | Spreadsheet UI, XLSX writer, arbitrary VBA, filesystem and network access |
| `@elixcee/xlsx` | Browser/Node XLSX read/write package and WASM runtime integration | Full Excel OOXML fidelity and arbitrary VBA execution |
| `playground` | Excel-like grid, range selection, editing, chart preview, download, bounded VBA demonstration | A crate, Microsoft Excel, collaboration, arbitrary VBA execution |

The Playground combines `elixcee-wasm` with the package's JavaScript writer.
This keeps the crate useful to applications that already have their own UI or
writer, while making the demo discoverable without claiming that a Rust crate
contains a web spreadsheet application.

The machine-readable version of this contract is
[`crate-api-matrix.json`](crate-api-matrix.json).

## Distribution shape

The reusable browser runtime is released as the `elixcee-wasm` crate. The
Excel-like screen is intentionally released separately as a static web app;
it is not a Rust crate and does not become part of the crate's API contract.
Applications can therefore use the crate with their own UI, or use the
Playground as a reference integration with `@elixcee/xlsx`'s JavaScript
writer. The release order is: publish the matching `elixcee` version, publish
`elixcee-wasm` after the crates.io index resolves that version, then publish
the browser package and deploy the Playground.

## Minimal WASM editor usage

```rust
use elixcee_wasm::WorkbookEditor;

let mut editor = WorkbookEditor::new(&xlsx_bytes)?;
editor.set_number("Sheet1", 1, 1, 42.0)?;
let snapshot_json = editor.recalculate()?;
```

Coordinates are one-based. Input limits and rejected OOXML are inherited from
the shared reader. XLSX serialization remains an explicit integration boundary;
the current browser package performs it in JavaScript after validating output.

## Publishing boundary

The crate release unit for the browser runtime is `elixcee-wasm`. It is
versioned and published independently from the static Playground and from the
private `@elixcee/xlsx` package. A release must therefore describe three
separate artifacts:

1. `elixcee-wasm`: reusable Rust/WASM read, calculation, diagnostics, and
   bounded editor APIs.
2. `@elixcee/xlsx`: JavaScript read/write integration and generated WASM
   bindings used by Node and browsers.
3. `playground`: the Excel-like reference UI and tutorial surface.

This prevents a crate consumer from accidentally depending on the Playground's
UI or assuming that a crate call can serialize every OOXML feature. The
publish gate is native `elixcee` first, then matching `elixcee-wasm`, followed
by the JavaScript package and GitHub Pages deployment.
