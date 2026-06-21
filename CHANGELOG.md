# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

### Added

- **Financial functions**: `FV`, `PV`, `NPER`, `RATE` (Newton-Raphson), `IPMT`, `PPMT`, `NPV`, `IRR`, `MIRR`, `XNPV`, `XIRR` — all share the `annuity_fv` / `compute_pmt` helpers
- **Database functions**: `DSUM`, `DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DMAX`, `DMIN` — all take `(database, field, criteria)` and reuse the existing `db_row_matches_criteria` / `resolve_db_field` infrastructure from `DGET`
- **GitHub Actions CI/CD**: `.github/workflows/publish.yml` — builds wheels for Linux x86_64/aarch64, Windows x86_64, macOS universal2, and an sdist; publishes to PyPI via OIDC Trusted Publisher on `v*` tag push
- **README_zh.md**: Simplified Chinese translation of README

### Removed

- `FUNCTIONS_ja.md`: duplicate of `FUNCTIONS.md`; `README_ja.md` now links to the English reference

### Tests

323 unit tests (↑ from 299)

---

## [0.1.0] — Initial Release

### Added — VBA Parser & Interpreter

- **Sub / End Sub** with parameter passing
- **Variable assignment** and arithmetic expressions
- **Cell read/write** via `Cells(row, col).Value` and `Range("A1").Value`
- **For / Next** loops with optional `Step`
- **For Each** iteration over cell ranges
- **If / ElseIf / Else / End If** conditional branches
- **Do While / Loop** and **While / Wend** loops
- **Select Case** with value, range (`To`), and comparison (`Is`) patterns
- **Exit For**, **Exit Do**, **Exit Sub**, **Exit Function**
- **Function / End Function** with return values; **Call** statement
- **On Error Resume Next**, **On Error GoTo label**, **Resume**, **GoTo**
- **With / End With** blocks (plain and `With Sheets("name")`)
- **Const** declarations; `Option Explicit` / `Option Base` ignored
- **Dim** variable declarations; `Dim arr(n)` and `ReDim [Preserve]` arrays
- **Type ... End Type** user-defined types with typed field initialization
- **Public / Private / Friend / Static** modifiers on Sub/Function (modifier ignored)
- **Debug.Print** / **Debug.Assert** as no-ops
- **MsgBox** — configurable skip or RuntimeError
- Range operations: `ClearContents`, `Clear`, `Copy`, `Delete`, `Insert`, `Sort`, `Offset.Value`
- Sheet operations: `Sheets.Add`, `Sheets.Delete`, `Sheets("name").Cells`
- `Application.Calculation` (Manual / Automatic); `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `StatusBar`, `Cursor`, `CutCopyMode` as no-ops
- `WorksheetFunction.*` prefix forwarding to formula engine
- `Cells(Rows.Count, col).End(xlUp).Row` and related `.End(dir).Row/Column` — indexed with `BTreeSet` for O(log n) performance

### Added — Named Ranges

- `Range("A1:B5").Name = "MyName"` registers a workbook-level named range
- All Range operations (Read/Write/Clear/Delete/Insert/Sort/Copy/ForEach) transparently resolve named range strings

### Added — Formula Engine (200+ functions)

#### Arithmetic & Statistical
`SUM`, `AVERAGE`, `AVERAGEIF`, `AVERAGEIFS`, `MIN`, `MAX`, `MINIFS`, `MAXIFS`,
`COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS`, `COUNTBLANK`,
`SUMIF`, `SUMIFS`, `SUMPRODUCT`, `PRODUCT`, `MEDIAN`, `MODE.MULT`,
`LARGE`, `SMALL`, `RANK`, `PERCENTILE` / `PERCENTILE.INC`, `PERCENTRANK` / `PERCENTRANK.INC`,
`ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `TRUNC`, `MOD`,
`RAND`, `RANDBETWEEN`, `SUBTOTAL`, `AGGREGATE`

#### Statistical
`STDEV` / `STDEV.S`, `STDEVP` / `STDEV.P`, `VAR` / `VAR.S`, `VARP` / `VAR.P`

#### Math & Trigonometry
`ABS`, `SQRT`, `POWER`, `EXP`, `LN`, `LOG`, `LOG10`, `PI`,
`SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `DEGREES`, `RADIANS`,
`FLOOR` / `FLOOR.MATH`, `CEILING` / `CEILING.MATH`, `MROUND`

#### Logical
`IF`, `IFS`, `SWITCH`, `AND`, `OR`, `NOT`, `XOR`, `IFERROR`

#### Text
`LEFT`, `RIGHT`, `MID`, `LEFTB`, `RIGHTB`, `MIDB`,
`LEN`, `LENB`, `UPPER`, `LOWER`, `PROPER`, `TRIM`,
`FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`,
`CONCATENATE`, `CONCAT`, `TEXTJOIN`, `TEXT`, `VALUE`, `EXACT`,
`CHAR`, `UNICHAR`, `CODE`, `UNICODE`, `ASC`, `JIS`

#### Date & Time
`DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `DAYS`,
`EDATE`, `EOMONTH`, `DATEDIF`, `DATEVALUE`,
`TIME`, `TIMEVALUE`, `HOUR`, `MINUTE`, `SECOND`,
`NETWORKDAYS`, `NETWORKDAYS.INTL`, `WORKDAY.INTL`

#### Lookup & Reference
`VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `LOOKUP`,
`INDEX`, `MATCH`, `XMATCH`, `CHOOSE`,
`ROW`, `COLUMN`, `INDIRECT`, `OFFSET`, `ADDRESS`

#### Information
`ISBLANK`, `ISERROR`, `ISERR`, `ISNA`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, `ISNONTEXT`

#### Array / Spill
`FILTER`, `UNIQUE`, `SORT`, `SORTBY`, `SEQUENCE`, `TRANSPOSE`,
`TOCOL`, `TOROW`, `WRAPCOLS`, `WRAPROWS`, `RANDARRAY`

#### Lambda & Higher-Order
`LET`, `LAMBDA`, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`

### Added — Formula Engine Features

- **Topological sort** for formula recalculation (`topo_sort_formulas`): formulas are evaluated in dependency order; circular references fall back to best-effort ordering
- **Application.Calculation** mode — Manual suppresses recalc; switching to Automatic triggers full recalc
- A1-notation and R1C1-notation cell references; range references (`A1:B10`)
- DBCS byte semantics (`LENB`, `LEFTB`, `RIGHTB`, `MIDB`) matching Excel's 2-byte-per-CJK rule
- Excel 1900 leap-year bug compatibility in date serial arithmetic

### Added — Python API

| Method | Description |
|---|---|
| `Vm(on_msgbox=)` | Create a VM; `"skip"` or `"error"` on MsgBox |
| `vm.run(vba, name)` | Execute a Sub |
| `vm.set_cell(r, c, v)` / `get_cell(r, c)` | 1-based cell read/write |
| `vm.cells()` | All non-empty cells as `{(r, c): value}` |
| `vm.cells_df()` | Active sheet as pandas DataFrame (requires pandas) |
| `vm.variables()` | VBA variables as `{name: value}` |
| `vm.set_cell_formula(r, c, f)` | Set and evaluate a formula string |
| `vm.set_cell_formula_batch(d)` | Batch formula set: `{(r,c): formula}` |
| `vm.recalculate()` | Re-evaluate all formula cells |
| `vm.set_sheet(name)` / `active_sheet()` / `sheet_names()` | Sheet management |
| `vm.get_sheet(name)` | Cells of a named sheet |
| `vm.save_workbook(path)` | Save to `.xlsx` or `.ods` |
| `vm.named_ranges` | Dict of registered named ranges |
| `elixcee.run_macro(vba, name)` | One-shot macro runner |
| `elixcee.load_workbook(path)` | Load `.xlsx` / `.ods` into a `Vm` |

- `Variant::Date` → Python `datetime.date` conversion
- `Variant::Error` → Python `ExcelError` class with `.code` attribute (bidirectional)
- Type stubs `elixcee.pyi` for IDE completion

### Added — File I/O

- **Read**: `.xlsx`, `.xlsm`, `.ods` — hand-written XML parser (no calamine at runtime)
- **Write**: `.xlsx` — hand-written XML + zip; `.ods` — hand-written XML + zip
- Multi-sheet support: all sheets loaded on `load_workbook`; saved on `save_workbook`

### Performance

- `Cells.End` searches (`xlUp`, `xlDown`, `xlToLeft`, `xlToRight`) use a lazy `BTreeSet` index — O(log n) per query after O(n) rebuild on cell mutation
- Zero-copy formula parse caching via `recalculate_all` with topological ordering

### Dependencies (runtime)

| Library | Purpose |
|---|---|
| `pyo3` | Python bindings |
| `zip` | XLSX / ODS archive read-write |

`calamine` is kept as a `[dev-dependencies]` oracle for diff-testing the hand-written reader.

### Tests

299 unit tests covering parser, formula engine, VM interpreter, file round-trips, and diff tests against calamine.
