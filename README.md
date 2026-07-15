# elixcee

**English** | [日本語](README_ja.md) | [中文](README_zh.md)

A library to emulate and execute Excel macro (VBA) data processing logic at high speed on Linux, macOS, and Windows — without installing Microsoft Excel.

The core engine is written in **Rust**; Python bindings are provided via **pyo3 + maturin**.

## Name

**elixcee** = **Excel** + **elixir** + **C**

An *elixir* that cures your Excel dependency — running at C-level speed via Rust.

---

## Comparison with similar tools

| Feature | **elixcee** | xlwings | LibreOffice UNO | openpyxl | xlcalculator |
|---------|:-----------:|:-------:|:---------------:|:--------:|:------------:|
| Runs VBA macros | Yes | Yes | Yes (subset) | No | No |
| Requires Excel | No | Yes | No | No | No |
| Requires LibreOffice | No | No | Yes | No | No |
| Evaluates formulas | Yes | Yes | Yes | No | Yes |
| macOS/Linux/Windows | Yes | partial | Yes | Yes | Yes |
| Simple Python API | Yes | Yes | No | Yes | Yes |
| Read .xlsx | Yes | Yes | Yes | Yes | Yes |
| Read .ods | Yes | Yes | Yes | No | No |
| Write .xlsx | Yes | Yes | Yes | Yes | No |
| Write .ods | Yes | Yes | Yes | No | No |
| Execution speed | Rust (native) | COM/IPC (slow) | IPC (slow) | — | Python |

**Notes:**
- **xlwings** requires Excel for Mac on macOS (via AppleScript) and Excel on Windows (via COM). Linux support requires a running Excel instance or a cloud bridge.
- **LibreOffice UNO** has a slow startup (≥ 1 s process launch) and a complex API. It runs VBA via LibreOffice's own interpreter, which may not match Excel's behavior exactly.
- **openpyxl** reads cached formula values from .xlsx files but does not re-evaluate formulas at runtime.
- **xlcalculator** re-evaluates Excel formulas in Python but has no VBA support.
- elixcee's VBA interpreter covers the subset of VBA used in typical data-processing macros (loops, conditionals, cell read/write, string/math functions, multi-sheet access). Excel-UI operations (charting, formatting, dialogs) are no-ops.

---

## Installation

```bash
pip install elixcee
```

Development build (from source):

```bash
python3 -m venv .venv && source .venv/bin/activate
maturin develop
```

---

## CLI (Windows / Linux / macOS)

Pre-built binaries are available on the [Releases](https://github.com/kent-tokyo/elixcee/releases) page — no Python required.

| Download | Platform |
|---|---|
| [elixcee-x86_64-windows.exe](https://github.com/kent-tokyo/elixcee/releases/latest/download/elixcee-x86_64-windows.exe) | Windows x64 |
| [elixcee-x86_64-linux](https://github.com/kent-tokyo/elixcee/releases/latest/download/elixcee-x86_64-linux) | Linux x64 |
| [elixcee-aarch64-macos](https://github.com/kent-tokyo/elixcee/releases/latest/download/elixcee-aarch64-macos) | macOS Apple Silicon |

### Usage

```
elixcee <vba_file>... <MacroName> [OPTIONS]

Arguments:
  <vba_file>...  One or more VBA source files (.vbs / .bas / .txt). With
                 more than one, use Module.Sub to disambiguate same-named
                 Subs/Functions across modules.
  <MacroName>    Name of the Sub to execute (last argument)

Options:
  --file <path>    Load cell data from spreadsheet (.xlsx / .xlsm / .ods)
  --sheet <name>   Active sheet name (default: first sheet in --file)
  --output <path>  Save result cells to spreadsheet (.xlsx / .ods)
  --json           Emit a single JSON object (result or error) instead of plain text
```

### Examples

Run a VBA file and print results to stdout:

```bat
elixcee macro.vbs ProcessData
```

Load data from an Excel file, run a macro, and save the output:

```bat
elixcee macro.vbs ProcessData --file input.xlsx --output result.xlsx
```

Output format — one line per non-empty cell, tab-separated address and value:

```
A1    Hello
B1    42
A2    3.14
```

`MsgBox` calls are printed to stdout.

### Multiple files (multi-module projects)

Pass more than one source file to run a project spanning several modules.
Sub/Function names are shared project-wide — use `Module.Sub` to pick a
specific one if the bare name exists in more than one module (module names
come from `Attribute VB_Name` if present, else the filename):

```bat
elixcee Helpers.bas Main.bas Main.ProcessData
```

There's no project manifest yet (see [docs/agent-contract.md](docs/agent-contract.md)
for exactly what is/isn't supported, including how cross-module name
collisions are handled).

### JSON output (for scripts / AI agents)

Add `--json` for a single machine-readable JSON object instead of plain text:

```bat
elixcee macro.vbs ProcessData --json
```

```json
{"schema_version":1,"ok":true,"entrypoint":"ProcessData","duration_ms":0.42,"cells":[{"sheet":"sheet1","address":"A1","value":42}],"messages":[]}
```

Full contract — error codes, exit codes, `messages` semantics: [docs/agent-contract.md](docs/agent-contract.md).

### Static analysis without running the macro

`elixcee check` inspects one or more `.bas` files without executing them: parse errors, whether the entrypoint macro exists, undefined Sub/Function calls anywhere in the body, and interactive `MsgBox` calls. Every positional argument is a file; the entrypoint (if any) is always `--entry`, never positional — so `elixcee check *.bas` checks every module in a project without asserting any particular entrypoint.

```bat
elixcee check macro.vbs --entry ProcessData --json
```

```json
{"schema_version":1,"ok":true,"diagnostics":[]}
```

### Workbook snapshot

`elixcee snapshot` reads a `.xlsx`/`.xlsm`/`.ods` file directly — no VBA
execution — and prints every sheet's non-empty cells as Markdown by default,
or JSON with `--json`:

```bat
elixcee snapshot Book1.xlsx --json
```

```json
{"schema_version":1,"ok":true,"file":"Book1.xlsx","sheets":[{"name":"Sheet1","sheet_id":"1","stable_id":"sheet1","cells":[{"address":"A1","value":42}]}]}
```

`stable_id` is derived from the file's own `sheetId` when available (else a
positional fallback) — it is **not** VBA's `CodeName` property. See
[docs/agent-contract.md](docs/agent-contract.md) for the full rationale.

### Property-based workbook testing

`elixcee test-workbook` reruns a macro against a starting workbook many
times with generated boundary-value inputs (blank, `0`, `1`, `-1`, near
overflow, empty/short/long strings), checking every run for panics,
runtime errors, timeouts, and Excel error values — each case starts from a
completely fresh workbook state:

```toml
# fixture.toml
name = "order calculation"
workbook = "orders.xlsx"
vba_files = ["Main.bas"]
macro = "Main.Process"
cases = 100
seed = 42

[[inputs]]
range = "Input!B2:B10"
strategy = "boundary_numeric"

[[assertions]]
range = "Result!A1:F100"
rule = "no_excel_errors"
```

```bat
elixcee test-workbook fixture.toml --json
```

A failing case reports its seed and case index so it can be reproduced
exactly: `elixcee test-workbook fixture.toml --seed 42 --case 17`. Full
schema, strategies, and assertion rules: [docs/agent-contract.md](docs/agent-contract.md).

### Build from source

```bash
cargo build --release --bin elixcee
# binary: target/release/elixcee  (or elixcee.exe on Windows)
```

---

## Quick Start

```python
import elixcee

# Run a VBA macro and get all resulting cells
cells = elixcee.run_macro("""
Sub FillSquares()
    For i = 1 To 5
        Cells(i, 1).Value = i * i
    Next i
End Sub
""", "FillSquares")
# cells == {(1,1): 1, (2,1): 4, (3,1): 9, (4,1): 16, (5,1): 25}

# Pre-populate cells from Python, then run a macro
vm = elixcee.Vm()
vm.set_cell(1, 1, 100)
vm.set_cell(2, 1, 200)
vm.run("""
Sub CalcTotal()
    total = Cells(1,1).Value + Cells(2,1).Value
    Cells(3,1).Value = total
End Sub
""", "CalcTotal")
print(vm.get_cell(3, 1))   # 300
print(vm.variables())       # {"total": 300}

# Load cell data from an existing Excel file, then run a macro
vm = elixcee.load_workbook("data.xlsx")
vm.run(vba_code, "ProcessData")
result_cells = vm.cells()   # {(row, col): value, ...}

# Store a worksheet formula on a cell and evaluate it
vm.set_cell_formula(4, 1, "=SUM(A1:A3)")
print(vm.get_cell(4, 1))   # sum of rows 1-3 in column A

# Control MsgBox behavior
vm = elixcee.Vm(on_msgbox="skip")   # silently ignore MsgBox calls (default)
vm = elixcee.Vm(on_msgbox="error")  # raise RuntimeError on MsgBox
```

---

## Python API

| Method | Description |
|---|---|
| `Vm(on_msgbox="skip")` | Create a new VM. `on_msgbox="error"` raises `RuntimeError` on `MsgBox`. |
| `vm.run(vba_code, macro_name)` | Parse and execute the named Sub. |
| `vm.set_cell(row, col, value)` | Write a value into a cell (1-based). |
| `vm.get_cell(row, col)` | Read a cell value. Returns `None` for empty cells. |
| `vm.cells()` | All non-empty cells as `{(row, col): value}`. |
| `vm.variables()` | All VBA variables as `{name: value}`. |
| `vm.set_cell_formula(row, col, formula)` | Store a formula (e.g. `"=SUM(A1:A3)"`) and evaluate it. |
| `vm.set_cell_formula_batch(formulas)` | Set multiple formulas at once: `{(row, col): formula_str}`. |
| `vm.recalculate()` | Re-evaluate all formula cells (useful after manual cell writes). |
| `vm.set_sheet(name)` | Switch the active sheet (creates it if absent). |
| `vm.active_sheet()` | Name of the currently active sheet. |
| `vm.sheet_names()` | List of all sheet names. |
| `vm.get_sheet(name)` | Cells of a named sheet as `{(row, col): value}`. |
| `vm.save_workbook(path)` | Save all sheets to `.xlsx` or `.ods`. |
| `vm.cells_df()` | Return the active sheet as a **pandas DataFrame** (requires pandas). |
| `elixcee.run_macro(vba, name)` | One-shot: run a macro and return `{(row, col): value}`. |
| `elixcee.load_workbook(path)` | Load an `.xlsx` or `.ods` file into a `Vm`. |

---

## Coverage

See **[FUNCTIONS.md](FUNCTIONS.md)** for the complete function and VBA syntax reference, including Excel version for each function.

**Highlights:**
- **Classic (Excel 2003-)**: SUM, VLOOKUP, IF, PMT, FV, PV, NPER, RATE, IPMT, PPMT, NPV, IRR, MIRR, XNPV, XIRR, DGET, DSUM, DAVERAGE, DCOUNT, DCOUNTA, DMAX, DMIN, and 100+ core functions
- **2007–2019**: IFERROR, COUNTIFS/SUMIFS, XOR, IFS, SWITCH, TEXTJOIN, MAXIFS/MINIFS
- **365/2021**: XLOOKUP, XMATCH, FILTER, SORT, UNIQUE, SEQUENCE, LET, LAMBDA, MAP, REDUCE
- **2024/365**: TEXTSPLIT, TEXTBEFORE, TEXTAFTER, VSTACK, HSTACK, TAKE, DROP, CHOOSECOLS, CHOOSEROWS
- **VBA**: For/If/While/With/On Error/Function/`Type...End Type`/Named Ranges/Array of UDT

### Named Ranges

Register a named range in VBA with `Range("A1:B5").Name = "MyData"`, then use the name anywhere a range address is accepted:

```vba
Range("MyData").Value = 0          ' write to all cells in the range
For Each cell In Range("MyData")   ' iterate over cells
    total = total + cell
Next cell
```

Named ranges are stored on `vm.named_ranges` (a `dict[str, str]` mapping lowercase name → address).

### Criteria Syntax (COUNTIF / SUMIF / SUMIFS / etc.)

| Criteria | Example | Meaning |
|---|---|---|
| Number | `10` | Exact numeric match |
| String | `"apple"` | Case-insensitive string match |
| Comparison | `">5"`, `"<=10"`, `"<>"` | Numeric comparison |
| Wildcard | `"a*"`, `"?bc"` | `*` = any chars, `?` = one char |

### Application Object

| Property / Method | Description | Behavior |
|---|---|---|
| `Application.Calculation = xlCalculationManual` | Disable auto-recalculation | **Active** |
| `Application.Calculation = xlCalculationAutomatic` | Enable auto-recalculation + re-evaluate all formula cells | **Active** |
| `Application.ScreenUpdating = False/True` | Suppress screen refresh | **No-op** (no screen) |
| `Application.EnableEvents = False/True` | Disable/enable event triggers | **No-op** (no events) |
| `Application.DisplayAlerts = False/True` | Suppress dialog boxes | **No-op** (no dialogs) |
| `Application.StatusBar = "..."` / `False` | Set/clear status bar text | **No-op** (no UI) |
| `Application.Cursor = xlWait` / `xlDefault` | Change cursor shape | **No-op** (no UI) |
| `Application.CutCopyMode = False` | Cancel clipboard mode | **No-op** (no clipboard) |

> **No-op** properties are parsed and accepted without error, but have no effect. This allows VBA macro performance patterns (e.g., `Application.ScreenUpdating = False` at the start of a macro) to run unchanged.

---

## Not Yet Supported

See **[FUNCTIONS.md — Not Yet Supported](FUNCTIONS.md#not-yet-supported)** for the full list.

Key gaps by category:
- **Statistical**: NORM.S.DIST, T.INV, F.DIST, CHISQ.DIST, and more
- **Text**: REPT, NUMBERVALUE, PHONETIC
- **Out of scope**: IMAGE (URL image fetch), GROUPBY (pivot aggregation), TRIMRANGE

---

## Status Legend

| Mark | Meaning |
|---|---|
| Done | Implemented and tested |
| TBD | Not yet scheduled |

---

## Development Phases

| Phase | Content | Status |
|---|---|---|
| Phase 1 | Rust project setup + pyo3 Python bindings | Done |
| Phase 2 | VBA parser MVP (Sub/End Sub, assignment, Cells) | Done |
| Phase 3 | Virtual Excel VM (variables, cell storage, interpreter) | Done |
| Phase 3.5 | Excel formula engine (SUM, IF, VLOOKUP, Application.Calculation, etc.) | Done |
| Phase 4 | Control flow (For loop, If/Else, arithmetic expressions) | Done |
| Phase 5 | Python interface (Vm class, run_macro, load_workbook, MsgBox) | Done |
| Phase 6 | Formula function expansion (100+ Excel functions, 118 tests) | Done |
| Phase 7 | Advanced VBA constructs (ElseIf, Exit, For Each, On Error, Function, arrays, While-Wend) | Done |
| Phase 8 | Range API (ClearContents, Offset, Sheets.Cells, WorksheetFunction, multi-sheet) | Done |
| Phase 9 | Multi-sheet support (Sheets HashMap, With Sheets, Python API, load_workbook all sheets) | Done |
| Phase 10 | Worksheet function expansion (math, trig, stats, array/spill, lambda functions) | Done |
| Phase 11 | User-defined types (`Type...End Type`), named ranges, `RANDARRAY`, pandas integration (`cells_df`), type stubs (`.pyi`) | Done |
| Phase D1 | Remove rust_xlsxwriter, hand-written XLSX via zip (dependencies: 5→4) | Done |
| Phase D2 | Remove pest/pest_derive, hand-written recursive descent VBA parser (dependencies: 4→3) | Done |
| Phase D3 | Remove calamine from runtime, hand-written XLSX/ODS reader (dependencies: 3→2) | Done |
| Perf R4 | SUM/AVERAGE/MIN/MAX fast path (skip `Vec<Variant>`), RangeWrite dirty-flag batching | Done |
| CLI | Standalone `elixcee` binary; pyo3 made optional; GitHub Actions release workflow | Done |
