# elixcee

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
- **Classic (Excel 2003-)**: SUM, VLOOKUP, IF, PMT, DGET, and 100+ core functions
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
- **Financial**: FV, PV, RATE, NPER, NPV, IRR, XNPV, XIRR, and more
- **Math**: FACT, PERMUT, GCD, LCM, SIGN, and more
- **Statistical**: NORM.DIST, CORREL, COVARIANCE.S, and more
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
