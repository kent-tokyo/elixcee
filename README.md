# elixcee

[![CI](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml/badge.svg)](https://github.com/kent-tokyo/elixcee/actions/workflows/ci.yml)
[![Docs](https://docs.rs/elixcee/badge.svg)](https://docs.rs/elixcee/1.0.11/elixcee/)
[![Version](https://img.shields.io/badge/version-1.0.11-blue.svg)](https://github.com/kent-tokyo/elixcee/releases/tag/v1.0.11)

Headless Excel workbook automation in Rust and Python. Edit `.xlsx`/`.xlsm`,
recalculate supported formulas, and run or diagnose data-processing VBA
without Microsoft Excel. CLI and experimental JavaScript/WASM surfaces are included.

[Quick start](docs/quickstart.md) · [Beginner tutorial](docs/tutorial-beginners.md) ·
[Browser playground](https://kent-tokyo.github.io/elixcee/playground/)

[English](README.md) | [日本語](README_ja.md) | [中文](README_zh.md)

## Install

```bash
pip install elixcee
```

## Python

```python
import elixcee

vm = elixcee.load_workbook("input.xlsx")
vm.set_cell(1, 1, 10)              # 1-based (row, column)
vm.set_cell_formula(1, 2, "=A1*2")
vm.recalculate()
vm.run("""Sub Macro()
    Cells(1, 3).Value = Cells(1, 2).Value + 1
End Sub""", "Macro")
vm.save_workbook("output.xlsx")
```

## CLI

```text
elixcee <file.bas>... <MacroName> --file input.xlsx --output result.xlsx
elixcee check <file.bas>... [--entry MacroName] [--json]
elixcee diagnose <file.bas>... <MacroName> --file input.xlsx [--json]
elixcee snapshot <workbook.xlsx|ods> [--json]
```

## Scope

The supported subset includes workbook editing, formula recalculation, and
data-processing VBA (`If`, `For`, `Do`, arrays, multiple sheets, `Range`/`Cells`).
UI effects and unsupported OOXML are bounded, modeled, preserved, rejected, or
reported depending on the operation. See [support contract](docs/v1-support-contract.md),
[function list](FUNCTIONS.md), and [limits](docs/limits.md).

[Roadmap](ROADMAP.md) · [Changelog](CHANGELOG.md) · [MIT license](docs/licensing.md)
