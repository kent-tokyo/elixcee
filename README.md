# elixcee

Run and test a practical subset of Excel VBA without Microsoft Excel. The core is
Rust, with a Python API (PyO3), a standalone CLI, and an experimental
`@elixcee/xlsx` JavaScript/WASM package.

Current release: **0.27.0**.

elixcee is intended for data-processing macros. It is not a replacement for the
Excel desktop application: UI features such as charts, dialogs, and screen
updates are skipped, modeled, or reported according to the operation.

## Install

```bash
pip install elixcee
```

Pre-built CLI binaries are published on the [GitHub Releases](https://github.com/kent-tokyo/elixcee/releases) page.

For a source build:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install maturin
maturin develop
```

## CLI

```text
elixcee <file.bas>... <MacroName> [--file input.xlsx] [--sheet Sheet1]
                         [--output result.xlsx] [--json]
elixcee check <file.bas>... [--entry MacroName] [--json]
elixcee snapshot <workbook.xlsx|ods> [--json]
elixcee test-workbook fixture.toml [--json] [--seed N] [--case N]
elixcee diagnose <file.bas>... <MacroName> --file input.xlsx [--json]
elixcee diagnose-workbook fixture.toml [--json] [--seed N] [--case N] [--cases N]
```

The run command accepts one or more standard VBA modules (`.bas`, `.vbs`, or
`.txt`). Use `Module.Sub` for a qualified entry point in a multi-module run.
`--json` emits one stable JSON object on stdout; see
[docs/agent-contract.md](docs/agent-contract.md) for the contract.

## Python quick start

```python
import elixcee

vm = elixcee.Vm()
vm.set_cell(1, 1, 10)          # coordinates are 1-based, like Excel
vm.run("""
Sub DoubleIt()
    Cells(1, 2).Value = Cells(1, 1).Value * 2
End Sub
""", "DoubleIt")
print(vm.get_cell(1, 2))       # 20

vm = elixcee.load_workbook("input.xlsx")
vm.run(vba_code, "ProcessData")
vm.save_workbook("output.xlsx")
```

The Python API also provides formula evaluation, ranges, sorting, merges,
hidden rows/columns, sheet management, styles, tables, data validation,
AutoFilter, defined-name inspection, pandas export, and `.xlsx`/`.xlsm`/`.ods`
workbook I/O. See [elixcee.pyi](elixcee.pyi) for signatures and behavior.

For large XLSX/XLSM files, `open_stream(path, sheet=None)` yields rows without
materializing the whole workbook. Set `include_row_numbers=True` to receive
`(row_number, values)` tuples. `create_stream(path)` provides an append-only
XLSX row writer.

`Vm(on_msgbox="skip")` is the default. Use `on_msgbox="error"` to make a
`MsgBox` call raise an error.

## Supported VBA and formulas

The interpreter supports common data-processing constructs including
`Sub`/`Function`, variables and arrays, `If`, `For`, `For Each`, `Do`,
`Select Case`, `With`, `On Error`, user-defined types, named ranges, multiple
sheets, and Excel-style `Range`/`Cells` operations. Formula support includes
arithmetic, comparisons, criteria functions, lookup functions, date/time,
text, statistical, financial, logical, and dynamic-array functions.

The maintained coverage list is [FUNCTIONS.md](FUNCTIONS.md). Unsupported or
intentional no-op behavior is documented there and in the diagnostic contract.

## Workbook compatibility

The Rust reader/writer preserves supported cell data, formulas, styles, merges,
hidden rows/columns, and many unknown OOXML parts. Macro projects in `.xlsm`
files are preserved during supported round trips. Features not modeled by the
writer can still be lost or disconnected; tables, drawings, comments,
hyperlinks, and other OOXML objects should be treated as compatibility gaps
unless covered by tests for the version in use.

The project runs Rust tests, property tests, compatibility fixtures, and
differential tests for the JavaScript package in CI. Compatibility with Excel's
VBA execution semantics is not claimed for every macro, and post-save macro
execution has not been fully validated against desktop Excel.

## Development

```bash
cargo test --workspace
cargo fmt --check
cargo clippy --all-targets --all-features -- -D warnings
```

The short-term plan is in [ROADMAP.md](ROADMAP.md). Public design and policy
documents are in [docs/](docs/).

License: MIT. See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
