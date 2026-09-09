# elixcee

Headless Excel workbook automation in Rust and Python: edit workbooks, recalculate
supported formulas, and run data-processing VBA without Microsoft Excel. The core
is Rust, with a Python API (PyO3), a standalone CLI, and an experimental
`@elixcee/xlsx` JavaScript/WASM package.

Version: **1.0.5**. See the [changelog](CHANGELOG.md) for versioned changes.
The experimental JS package remains private and is not published. [English](README.md) | [日本語](README_ja.md) | [中文](README_zh.md)

elixcee is a workbook automation runtime, not a VBA-only execution tool. Use the
same workbook model for direct data edits, supported formula recalculation, and
execution, diagnosis, or testing of supported data-processing VBA. It is not a
replacement for the Excel desktop application: UI features such as charts,
dialogs, and screen updates are skipped, modeled, or reported according to the
operation.

### Choose by workflow

| Need | Fit |
|---|---|
| Direct cell edits only | A general workbook library may be sufficient. |
| Edit a workbook and recalculate supported formulas headlessly | elixcee |
| Run or diagnose data-processing VBA in Linux, macOS, or CI | elixcee |
| Full Excel desktop object model, UI, or complete OOXML compatibility | Check Excel or a specialized library and the elixcee support boundaries. |

Choose elixcee when the workflow needs one or more of these without a desktop
Excel installation: read/edit/save `.xlsx` or `.xlsm`, calculate supported
formulas, run and diagnose supported VBA, or test workbook behavior in CI.
For a complete Excel desktop object model or full OOXML feature compatibility,
check the documented support boundaries before adopting it.

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

The run command accepts standard VBA modules (`.bas`, `.vbs`, `.txt`) and
exported class modules (`.cls`). Use `Module.Sub` for a qualified entry point in a multi-module run.
For valid invocations, `--json` emits a result/error JSON object on stdout; see
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
vm.set_cell(1, 1, 10)                              # edit a cell
vm.set_cell_formula(1, 2, "=A1*2")                # add a formula
vm.recalculate()                                  # recalculate formula cells
vm.run("Sub ProcessData()\n    Cells(1, 1).Value = 42\nEnd Sub", "ProcessData")
vm.save_workbook("output.xlsx")

# Cell/formula edits also support bounded undo/redo.
vm.set_cell(1, 1, 20)
vm.undo()

# Group a large edit into one undoable operation.
vm.begin_transaction()
vm.set_range("A1:C1000", [[row, row * 2, row * 3] for row in range(1, 1001)])
vm.commit_transaction()
vm.undo()  # restores the state before the whole transaction

# Optional reader resource controls (the cancellation check is cooperative).
cancel = elixcee.ReadCancellation()
vm = elixcee.load_workbook(
    "input.xlsx", max_work_units=100_000_000, timeout_ms=30_000, cancellation=cancel
)
```

The Python API provides headless formula calculation, cell/range editing,
bounded undo/redo, formula evaluation, ranges, sorting, merges,
hidden rows/columns, sheet management, styles, tables, data validation,
AutoFilter, defined-name inspection, pandas export, and `.xlsx`/`.xlsm`/`.ods`
workbook I/O. See [elixcee.pyi](elixcee.pyi) for signatures and behavior.
Explicit zero-argument workbook/worksheet handlers can be invoked with
`vm.run_event(...)`; `EnableEvents` is honored. `vm.run_worksheet_change(...)`
can explicitly bind an A1 target range to `Worksheet_Change(Target As Range)`.
`vm.run_with_events(...)` opts into `Workbook_Open` dispatch before the selected
macro, and VBA cell/range writes in that mode automatically dispatch a unique
`Worksheet_Change` handler with bounded event chaining. `Target.Value`,
`Target.Address`, `Target.Row`, `Target.Column`, `Target.Rows.Count`,
`Target.Columns.Count`, and `Target.Cells.Count` are available for the
supported single-area target model. `Target.Parent.Name` identifies the bound
worksheet. Ambiguous multiple handlers are rejected deterministically.
`Vm.tables()` and `Vm.data_validations()` return typed structural metadata
projections; they do not evaluate calculated-column or validation formulas.
For loaded XLSX/XLSM sheets, `Vm.sheet_id(name)` and
`Vm.sheet_name_for_id(sheet_id)` expose stable source identities independently
of tab order and rename operations; new or ODS sheets have no inferred ID.
Loaded workbooks also support bounded edits to selected existing Chart-series
formulas, caches, marker attributes, smooth flags, negative-value display and
series visibility flags, plus the first Chart
title text run; Chart creation and general object editing remain outside the
current contract.

For large XLSX/XLSM files, `open_stream(path, sheet=None)` yields rows without
materializing the whole workbook. Set `include_row_numbers=True` to receive
`(row_number, values)` tuples, or `max_rows=N`/`max_row_bytes=N`/`max_columns=N` to bound a read.
Set `timeout_ms=N` to bound how long each `next()` waits for another row.
`create_stream(path)` provides an append-only XLSX row writer. Set
`max_rows=N`, `max_columns=N`, and/or `max_pending_bytes=N` to bound accepted output.
The byte budget is cumulative until close, not retained RSS; this is not a
constant-memory guarantee. See [limits and Unreleased G1 changes](docs/limits.md).
For a separate per-row and total-work budget, use `create_stream_bounded(...)`.

`Vm(on_msgbox="skip")` is the default. Use `on_msgbox="error"` to make a
`MsgBox` call raise an error. Set `Vm(timeout_ms=N)` or pass `timeout_ms=N`
to `run_macro` to bound VBA execution time.

The read-only CLI snapshot accepts `--max-work-units N`, `--timeout-ms N`, and
`--cancel-file PATH`. Creating the cancel-file while a read is in progress
requests a cooperative stop; a blocking filesystem read cannot be forcibly
interrupted by this mechanism.
Repeated runs on one `Vm` reuse the parsed AST; `vm.fork()` creates an isolated batch copy.
`vm.snapshot()` returns detached sheet values, tab order, defined names, calculation mode,
visibility, merges, and hidden intervals. Set `include_formulas=True` for formula text.
This is richer than the CLI snapshot schema.
Use `diagnose_macro(vba_code, macro_name, workbook_path)` for structured
diagnostics matching the CLI `diagnose --json` contract.

## Supported VBA and formulas

The interpreter supports common data-processing constructs including
`Sub`/`Function`, variables and arrays, `If`, `For`, `For Each`, `Do`,
`Select Case`, `With`, `On Error`, user-defined types, named ranges, multiple
sheets, Excel-style `Range`/`Cells` operations, and the documented built-in VBA
`Collection`, in-memory Dictionary, and class-module subsets (see the documented limits). Formula support includes
arithmetic, comparisons, criteria functions, lookup functions, date/time,
text, statistical, financial, logical, and dynamic-array functions.

The maintained coverage list is [FUNCTIONS.md](FUNCTIONS.md). Unsupported or
intentional no-op behavior is documented there and in the diagnostic contract.

## Workbook compatibility

The Rust reader/writer preserves supported cell data, formulas, styles, merges,
hidden rows/columns, and many unknown OOXML parts. Macro projects in `.xlsm`
files are preserved during supported round trips. Features not modeled by the
writer can still be lost or disconnected. Existing Drawing/relationship chains
are preserved on tested paths, and bounded APIs can update selected Chart-series
formulas, caches, marker attributes, smooth flags, negative-value display and series visibility flags, Chart title text, two-cell
Drawing anchors, and worksheet-backed Pivot source fields. Drawing shape metadata and selected solid fill/line style
attributes (rotation, flips, RGB/ARGB fill and line color, line width, and
preset dash) are also available through bounded APIs. Chart creation, general
Drawing shape editing, Pivot cache recalculation, comments, hyperlinks, and
other OOXML objects remain compatibility gaps unless covered by tests for the
version in use.

The project runs Rust tests, property tests, compatibility fixtures, and
differential tests for the JavaScript package in CI. Compatibility with Excel's
VBA execution semantics is not claimed for every macro, and post-save macro
execution has not been fully validated against desktop Excel.

The v1 support boundaries are documented in
[docs/v1-support-contract.md](docs/v1-support-contract.md). Version 1.0 is a
stable contract for the documented data-processing subset and its safe failure
behavior; it is not a claim of complete Excel or VBA compatibility.

## Development

```bash
cargo test --workspace
cargo fmt --check
cargo clippy --all-targets --all-features -- -D warnings
cargo audit --no-fetch
cargo deny check --disable-fetch
python3 scripts/check-reader-measurements.py docs/measurements/*.json
python3 scripts/check-reader-measurements.py --self-test
bash scripts/check-measurement-boundary.sh
```

Offline dependency checks use the local advisory database and `deny.toml`.
A stale database does not establish the absence of newly published advisories.

The short-term plan is in [ROADMAP.md](ROADMAP.md). Public design and policy
documents are in [docs/](docs/).

License: [MIT licensing information](docs/licensing.md). See [third-party notices](THIRD_PARTY_NOTICES.md).
