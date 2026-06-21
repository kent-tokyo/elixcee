# elixcee

A library to emulate and execute Excel macro (VBA) data processing logic at high speed on Linux, macOS, and Windows — without installing Microsoft Excel.

The core engine is written in **Rust**; Python bindings are provided via **pyo3 + maturin**.

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

## Coverage

### VBA Syntax

| Syntax | Example | Status |
|---|---|---|
| Sub / End Sub | `Sub MySub() ... End Sub` | Done |
| Variable assignment | `a = 10` | Done |
| Cell write | `Cells(1, 1).Value = a` | Done |
| Cell read | `x = Cells(1, 1).Value` | Done |
| Comment | `' comment` | Done |
| Application.Calculation | `Application.Calculation = xlCalculationAutomatic` | Done |
| Arithmetic expressions | `Cells(1, 1).Value = a + 1` | Done |
| For loop | `For i = 1 To N ... Next i` | Done |
| For loop (step) | `For i = 10 To 1 Step -2` | Done |
| If / Else | `If x > 0 Then ... Else ... End If` | Done |
| Do While loop | `Do While x > 0 ... Loop` | TBD |
| Select Case | `Select Case x ... End Select` | TBD |

### Excel Worksheet Functions (cell formulas)

#### Arithmetic & Statistical

| Function | Description |
|---|---|
| `SUM` | Sum of a range |
| `AVERAGE` | Average of a range |
| `AVERAGEIF` | Conditional average |
| `AVERAGEIFS` | Multi-criteria average |
| `MIN` / `MAX` | Minimum / Maximum value |
| `MINIFS` / `MAXIFS` | Conditional min / max |
| `COUNT` / `COUNTA` | Count of numeric / non-empty cells |
| `COUNTIF` / `COUNTIFS` | Conditional count |
| `SUMIF` / `SUMIFS` | Conditional sum |
| `SUMPRODUCT` | Sum of element-wise products |
| `PRODUCT` | Product of all values |
| `MEDIAN` | Median value |
| `MODE.MULT` | Most frequent value |
| `LARGE` / `SMALL` | Kth largest / smallest |
| `RANK` | Rank of a number |
| `PERCENTILE` / `PERCENTILE.INC` | Percentile (inclusive) |
| `PERCENTRANK` / `PERCENTRANK.INC` | Percent rank |
| `ROUND` / `ROUNDUP` / `ROUNDDOWN` | Rounding functions |
| `INT` | Floor (toward negative infinity) |
| `TRUNC` | Truncate toward zero |
| `MOD` | Modulo |
| `RAND` | Random float 0–1 |
| `RANDBETWEEN` | Random integer in range |
| `SUBTOTAL` | Aggregate with selectable function (1–6, 9, 101–106, 109) |
| `AGGREGATE` | Extended subtotal (1–6, 9, 12–16) |

#### Logical

| Function | Description |
|---|---|
| `IF` | Conditional value |
| `IFS` | Multi-condition branch |
| `SWITCH` | Switch/case |
| `AND` / `OR` / `NOT` | Logical operators |
| `XOR` | Exclusive OR |
| `IFERROR` | Fallback on error |

#### Text

| Function | Description |
|---|---|
| `LEFT` / `RIGHT` / `MID` | Extract characters |
| `LEFTB` / `RIGHTB` / `MIDB` | Extract by DBCS bytes |
| `LEN` / `LENB` | Character / byte count |
| `UPPER` / `LOWER` / `PROPER` | Case conversion |
| `TRIM` | Remove extra spaces |
| `FIND` | Case-sensitive position search |
| `SEARCH` | Case-insensitive wildcard search |
| `SUBSTITUTE` | Replace by value |
| `REPLACE` | Replace by position |
| `CONCATENATE` / `CONCAT` | Concatenate strings |
| `TEXTJOIN` | Join with delimiter |
| `TEXT` | Format number as string |
| `VALUE` | Parse string to number |
| `EXACT` | Case-sensitive equality |
| `CHAR` / `UNICHAR` | Character from code |
| `CODE` / `UNICODE` | Code point of first character |
| `ASC` | Full-width → half-width (DBCS) |
| `JIS` | Half-width → full-width (DBCS) |

#### Date & Time

| Function | Description |
|---|---|
| `DATE` | Create date serial (Excel epoch) |
| `TODAY` / `NOW` | Today's date / current datetime |
| `YEAR` / `MONTH` / `DAY` | Extract date parts |
| `WEEKDAY` | Day of week (1–3 return types) |
| `DAYS` | Days between two dates |
| `EDATE` | Date N months from start |
| `EOMONTH` | Last day of month N months from start |
| `DATEDIF` | Difference in Y / M / D / MD / YM / YD |
| `DATEVALUE` | Parse "YYYY/MM/DD" or "YYYY-MM-DD" |
| `TIME` | Create time serial |
| `TIMEVALUE` | Parse "HH:MM:SS" |
| `HOUR` / `MINUTE` / `SECOND` | Extract time parts |
| `NETWORKDAYS` | Workdays between dates (Sat+Sun weekend) |
| `NETWORKDAYS.INTL` | Workdays with custom weekend |
| `WORKDAY.INTL` | Date N workdays away with custom weekend |

#### Lookup & Reference

| Function | Description |
|---|---|
| `VLOOKUP` / `HLOOKUP` | Vertical / horizontal lookup |
| `XLOOKUP` | Flexible lookup (exact, next-larger, next-smaller) |
| `LOOKUP` | Sorted vector lookup |
| `INDEX` | Value at row/column offset |
| `MATCH` | Position of a value |
| `XMATCH` | Extended MATCH with mode/search options |
| `CHOOSE` | Choose from list by index |
| `ROW` / `COLUMN` | Row / column number of reference |

#### Information

| Function | Description |
|---|---|
| `ISBLANK` | Is the value empty? |
| `ISERROR` / `ISERR` | Is the value an error? |
| `ISNA` | Is the value #N/A? (always FALSE — no N/A type) |
| `ISNUMBER` | Is the value numeric? |
| `ISTEXT` | Is the value a string? |
| `ISLOGICAL` | Is the value a boolean? |
| `ISNONTEXT` | Is the value not a string? |

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
