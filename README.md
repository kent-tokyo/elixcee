# elixcee

[![CI](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml/badge.svg)](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml)
[![Docs](https://docs.rs/elixcee/badge.svg)](https://docs.rs/elixcee/1.0.12/elixcee/)
[![Version](https://img.shields.io/badge/version-1.0.12-blue.svg)](https://github.com/kent-tokyo/elixcee/releases/tag/v1.0.12)

Headless Excel workbook automation in Rust and Python. Edit `.xlsx`/`.xlsm`
files, recalculate supported formulas, and run or diagnose data-processing VBA
without Microsoft Excel. A CLI, reusable Rust/WASM runtime crates, and an
experimental JavaScript/WASM package are also included.

The current release is `1.0.12`. The browser playground demonstrates range
selection, cell editing, formatting, filters, tables, chart previews,
worksheet-backed Pivot summaries, sheet operations, formula recalculation, and
bounded VBA execution.

[English](README.md) | [日本語](README_ja.md) | [中文](README_zh.md)

## Start here

1. [Quick start](docs/quickstart.md) — install and run the first operation.
2. [Beginner tutorial](docs/tutorial-beginners.md) — edit cells, calculate formulas, and run VBA.
3. [Browser playground](https://kent-tokyo.github.io/elixcee/playground/) — try it without installing anything.

The playground switches between English, Japanese, and Simplified Chinese.
For local playground work, see [playground/README.md](playground/README.md).

## Install

```bash
pip install elixcee
```

CLI binaries are available from [GitHub Releases](https://github.com/kent-tokyo/elixcee/releases).

The reusable Rust crates are `elixcee` (native/Python runtime) and
`elixcee-wasm` (browser/Node bridge). They provide headless reading, supported
formula calculation, diagnostics, and bounded workbook editing. The
[elixcee-wasm API docs](https://docs.rs/elixcee-wasm) are available separately.
[Playground](https://kent-tokyo.github.io/elixcee/playground/) is a separate
static web app: it combines the WASM runtime with the package's JavaScript
XLSX writer and Excel-like UI. See the [crate/API boundary](docs/crate-api-boundary.md)
for the exact contract.

## Python

```python
import elixcee

vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, 10)                 # Excel-style 1-based coordinates
vm.set_cell_formula(1, 2, "=A1*2")
vm.recalculate()
vm.run("""Sub Macro()
    Cells(1, 3).Value = Cells(1, 2).Value + 1
End Sub""", "Macro")
vm.save_workbook("output.xlsx")
```

The Python API also provides ranges, sorting, merges, styles, tables,
validation, hidden rows/columns, undo/redo, events, streaming row I/O, and
bounded reader/execution controls. See [elixcee.pyi](elixcee.pyi) for the API.

## CLI

```text
elixcee <file.bas>... <MacroName> --file input.xlsx --output result.xlsx
elixcee check <file.bas>... [--entry MacroName] [--json]
elixcee snapshot <workbook.xlsx|ods> [--json]
elixcee diagnose <file.bas>... <MacroName> --file input.xlsx [--json]
elixcee test-workbook fixture.toml [--json]
```

Use `--json` for automation. The stable output contract is documented in
[docs/agent-contract.md](docs/agent-contract.md).

## Supported scope

The runtime covers common data-processing VBA (`If`, `For`, `Do`, `Select Case`,
`With`, arrays, multiple sheets, `Range`/`Cells`, and selected collections),
plus arithmetic, lookup, text, logical, date/time, statistical, financial,
reference, and dynamic-array formulas.

It is not a full Excel desktop replacement. UI operations are skipped,
modeled, or reported. Chart/Drawing/Pivot editing and formula compatibility
remain bounded features; unsupported OOXML parts may be preserved, lost, or
rejected depending on the operation. The browser playground rejects external
links, PivotTables/PivotCaches, embedded media, threaded comments, slicers,
and custom XML when lossless preservation is not available.

See the [v1 support contract](docs/v1-support-contract.md),
[formula list](FUNCTIONS.md), and [limits](docs/limits.md) before adoption.

## Development

See the [roadmap](ROADMAP.md) for current priorities and release gates.
Unreleased changes are listed in [CHANGELOG.md](CHANGELOG.md).

License: [MIT](docs/licensing.md). See [third-party notices](THIRD_PARTY_NOTICES.md).
