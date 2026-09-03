# elixcee

Run and test a practical subset of Excel VBA without Microsoft Excel. The core is
Rust, with a Python API (PyO3), a standalone CLI, and an experimental
`@elixcee/xlsx` JavaScript/WASM package.

Current release: **1.0.0**.

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

# Optional reader resource controls (the cancellation check is cooperative).
cancel = elixcee.ReadCancellation()
vm = elixcee.load_workbook(
    "input.xlsx", max_work_units=100_000_000, timeout_ms=30_000, cancellation=cancel
)
```

The Python API also provides formula evaluation, ranges, sorting, merges,
hidden rows/columns, sheet management, styles, tables, data validation,
AutoFilter, defined-name inspection, pandas export, and `.xlsx`/`.xlsm`/`.ods`
workbook I/O. See [elixcee.pyi](elixcee.pyi) for signatures and behavior.

For large XLSX/XLSM files, `open_stream(path, sheet=None)` yields rows without
materializing the whole workbook. Set `include_row_numbers=True` to receive
`(row_number, values)` tuples, or `max_rows=N`/`max_row_bytes=N`/`max_columns=N` to bound a read.
Set `timeout_ms=N` to bound how long each `next()` waits for another row.
`create_stream(path)` provides an append-only XLSX row writer. Set
`max_rows=N`, `max_columns=N`, and/or `max_pending_bytes=N` to bound pending output.

`Vm(on_msgbox="skip")` is the default. Use `on_msgbox="error"` to make a
`MsgBox` call raise an error. Set `Vm(timeout_ms=N)` or pass `timeout_ms=N`
to `run_macro` to bound VBA execution time.

The read-only CLI snapshot accepts `--max-work-units N`, `--timeout-ms N`, and
`--cancel-file PATH`. Creating the cancel-file while a read is in progress
requests a cooperative stop; a blocking filesystem read cannot be forcibly
interrupted by this mechanism.
Repeated runs of the same source on one `Vm` reuse its parsed AST.
Use `vm.fork()` to create an isolated copy for batch execution.
Use `vm.snapshot()` to obtain a detached read-only view of all sheets.
Pass `include_formulas=True` to include stored formulas separately from values.
Snapshots also include the workbook's worksheet tab order.
Snapshots also include runtime defined names.
Snapshots include the current `calculation_mode` (`automatic` or `manual`).
Snapshots include per-sheet visibility states (`visible`, `hidden`, or `veryHidden`).
Snapshots include per-sheet merged ranges in A1 notation.
Snapshots include hidden row and column intervals.
Use `diagnose_macro(vba_code, macro_name, workbook_path)` for structured
diagnostics matching the CLI `diagnose --json` contract.

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
```

依存監査は、ネットワークアクセスを必要としないローカル検証として実行できます。
`cargo audit --no-fetch`と`cargo deny check --disable-fetch`は、手元にある advisory DB と
リポジトリ内の`deny.toml`を使用します。advisory DBが古い場合は、結果が最新の公開情報を
反映していない可能性があります。

The short-term plan is in [ROADMAP.md](ROADMAP.md). Public design and policy
documents are in [docs/](docs/).

License: MIT. See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
