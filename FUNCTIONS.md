# elixcee — Function & VBA Coverage Reference

Coverage reference for **elixcee 1.0.6**. See [CHANGELOG](CHANGELOG.md) for versioned changes.
“Done”/“Active” describes the documented subset, not complete Excel equivalence or release status.
Each worksheet function shows the minimum Excel version in which it was introduced as a built-in.

The mode-rich subset has an executable engine contract in
[`compat/formula-contracts.json`](compat/formula-contracts.json). It records accepted
argument shapes, supported and rejected modes, and regression-test anchors. Its
`oracle_status` is deliberately `unverified`: this is an implementation contract,
not a claim of Excel-equivalent results.

---

## Version Legend

| Label | Minimum Excel Version |
|---|---|
| Classic | Excel 2003 and earlier (always built-in) |
| 2007 | Excel 2007+ |
| 2010 | Excel 2010+ |
| 2013 | Excel 2013+ |
| 2019 | Excel 2019+ |
| 365/2021 | Microsoft 365 / Excel 2021+ |
| 2024/365 | Excel 2024 / Microsoft 365 (latest channel) |

---

## VBA Syntax

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
| Single-line If | `If x > 0 Then y = 1 Else y = 2` | Done |
| Do While loop | `Do While x > 0 ... Loop` | Done |
| Select Case | `Select Case x ... End Select` | Done |
| While / Wend | `While x > 0 ... Wend` | Done |
| For Each | `For Each cell In Range("A1:A5")`, `For Each item In collection` | Done (Range / built-in Collection / Dictionary.Keys) |
| With block (sheet) | `With Sheets("Sheet1") ... End With` | Done |
| With block (UDT) | `With p ... .Field = val ... End With` | Done |
| On Error Resume Next | `On Error Resume Next` | Done |
| On Error GoTo label | `On Error GoTo ErrH` | Done |
| Err.Number / Err.Description | Read the number/description of the most recent caught error | Done |
| Err.Clear | Reset `Err.Number` to 0 and `Err.Description` to `""` | Done |
| Err.Raise | `Err.Raise Number[, Source][, Description]` — `Source` is parsed (so `Description` in the 3rd slot is never misread) but not modeled as a readable property | Done |
| Function / Call | `Function Foo() ... End Function` | Done |
| Exit For / Exit Sub | `Exit For`, `Exit Sub` | Done |
| Array / ReDim | `Dim arr(10)`, `Dim arr(2 To 8)`, `Dim arr()` (dynamic, sized later via `ReDim`), `ReDim arr(n)` | Done |
| Erase | `Erase arr` — resets a fixed-size array's elements to `Empty` in place; comma-separated `Erase a, b` isn't parsed | Done |
| Option Base | `Option Base 1` — sets the default array lower bound (for declarators without an explicit `lo To hi`) | Done |
| Multi-declarator Dim | `Dim a As Integer, b As Range` | Done |
| Const | `Const PI = 3.14` | Done |
| Type ... End Type | User-defined types (UDT) | Done |
| Nested UDT | Field of type UDT (`p.Addr.Street`) | Done |
| Array of UDT | `Dim arr(10) As MyType` | Done |
| Named ranges | `Range("A1:B3").Name = "MyData"` | Done |
| Debug.Print | `Debug.Print x` | Done (no-op) |
| Option Explicit | `Option Explicit` | Done (ignored) |

### Collection Object

| Property / Method | Behavior |
|---|---|
| `Dim c As New Collection` / `Set c = New Collection` | **Active** — creates an empty built-in Collection |
| `c.Add item [, key] [, before] [, after]` | **Active** — Variant values, case-insensitive string keys, and 1-based positioning; named arguments are accepted |
| `c.Add objectVariable` | **Active** — preserves Range/Worksheet/Workbook/Collection/class-instance identity |
| `Call c.Add(...)` / `Call c.Remove(...)` | **Active** — parenthesized `Call` spelling preserves the same key/order semantics |
| `c.Count` | **Active** |
| `c.Item(index_or_key)` / `c(index_or_key)` | **Active** — numeric indices are 1-based |
| `Set x = c.Item(index_or_key)` / `Set x = c(index_or_key)` | **Active** — retrieves an object-valued item without losing identity |
| `c.Remove index_or_key` | **Active** |
| `For Each item In c` | **Active** — iterates over a stable value/object snapshot in insertion order |
| `Set alias = c` | **Active** — aliases share Add/Remove mutations |
| `With c ... .Add/.Remove/.Count/.Item ... End With` | **Active** — object-valued `Item` may also be a direct With target |

Object-valued items must originate from a live Set-assigned object variable or
another Collection item, and must be retrieved with `Set`; plain scalar
assignment reports an explicit error. Collection size uses the same
configurable element budget as VBA arrays. Unreachable nested/self-referential
Collection graphs are reclaimed by VM reachability.

### In-memory Dictionary

`Scripting.Dictionary` is active in the **1.0.6 VM-local adapter**, not a COM object.

| Operation | Supported subset |
|---|---|
| `Dim d As New Scripting.Dictionary` / `Set d = New Scripting.Dictionary` | New VM-local dictionary |
| `d.Add key, item` / `d.Exists(key)` | Insert unique keys / query membership |
| `d.Item(key)` / `d(key)` | Read an existing item; object values require `Set` |
| `d.Remove key` / `d.RemoveAll` / `d.Count` | Delete / clear / size |
| `d.CompareMode = 0` or `1` | Case-sensitive default or lowercase-normalized text comparison; cannot change mode while entries remain |
| `For Each key In d.Keys` | Snapshot of normalized string keys in insertion order |

Keys are coerced with VBA string conversion: numeric and string keys with the same
text collide, empty keys are rejected, and text mode stores lowercased keys.
This is **not full Scripting.Dictionary key-type/locale compatibility**. Item assignment,
full Keys/Items array semantics, and unlisted members are not promised.
`CreateObject`, COM, filesystem, and network access remain blocked.

### Class Modules

| Construct | Behavior |
|---|---|
| Exported `.cls` / `VERSION 1.0 CLASS` | **Active** — registered outside the standard-module procedure namespace |
| `Dim x As New ClassName` / `Set x = New ClassName` | **Active** — creates a VM-local object identity and runs zero-argument `Class_Initialize` when present |
| Module-level instance fields | **Active** — scalar and object-valued state; names are case-insensitive and `Private` access is enforced |
| `Call x.Method(...)` / `x.Method args` | **Active** — invokes a class `Sub`; zero-argument calls use the explicit `Call x.Method` spelling |
| `value = x.Function(...)` / `Set value = x.Function(...)` | **Active** — invokes scalar- or object-returning class `Function`; typed scalar/object arguments are supported |
| `Property Get` / `Property Let` / `Property Set` | **Active** — scalar/object and indexed properties; object assignment and retrieval require `Set` |
| Object-valued parameters | **Active** — typed class/object parameters preserve reference identity in class Subs and Functions |
| `Class_Terminate` | **Active** — runs once when an unreachable instance is reclaimed; termination failures propagate |
| `Implements InterfaceName` | **Active** — validates Sub/Function/Property signatures; an interface-typed variable dispatches only its contract to VBA-style `Interface_Member` implementations, while a concrete-typed variable uses direct members |
| `Public` / `Private` / `Friend` | **Active** — private class members are class-only; Friend remains project-visible |
| `Dim objects(...) As ClassName` | **Active** — fixed/dynamic, multidimensional object arrays with `Set`, `ReDim [Preserve]`, and `Erase` |
| `With x` / `With collection.Item(...)` | **Active** — field and method access keeps the same instance identity |
| Collection storage / `For Each` | **Active** — class instances are object-valued Collection entries |

Property index arguments and class Function arguments preserve declared
scalar/object types. Interface static types are retained across assignment,
typed parameters, class fields, and `With`. Function calls require parentheses;
a zero-argument Sub uses explicit `Call x.Method`.

### Range Object

| Property / Method | Behavior |
|---|---|
| `Range("A1")` / `Cells(1, 1)` without `.Value` | **Active** — default `Value` read/write; a multi-cell read returns a Variant array |
| `Set r = Range("B2:D4")` then `r.Cells(row, col)` | **Active** — 1-based coordinates relative to the first area's top-left cell |
| `r.Range("A1:B2")` | **Active** — A1 address relative to the first area's top-left cell |
| `r(row, col)` / `r.Item(row, col)` | **Active** — default Range `Item`, returning a Range object in `Set` context or its default value in scalar context |
| `r(index)` / `r.Item(index)` | **Active** — 1-based row-major item within the first area; out-of-range access is rejected |
| `With r ... .Cells/.Range/.Value ... End With` | **Active** — all relative members share the same captured Range target |

Relative addressing currently anchors to the first area of a multi-area Range.
`SpecialCells(Type[, Value])` supports every `XlCellType` constant:
`xlCellTypeAllFormatConditions`, `xlCellTypeSameFormatConditions`,
`xlCellTypeAllValidation`, `xlCellTypeSameValidation`, `xlCellTypeComments`,
`xlCellTypeFormulas`, `xlCellTypeConstants`, `xlCellTypeBlanks`,
`xlCellTypeLastCell`, and `xlCellTypeVisible`. Constants and formulas accept
the additive `xlNumbers` / `xlTextValues` / `xlLogical` / `xlErrors` mask.
Legacy comments/notes are recognized; threaded comments are outside this model.
`SameValidation` and `SameFormatConditions` use the receiver's first cell as
their deterministic anchor. Conditional-format sameness means membership in
the same worksheet `<conditionalFormatting>` rule block(s), because rule
evaluation itself remains opaque.

### Worksheet / Workbook Objects

| Property / Method | Behavior |
|---|---|
| `Worksheets(key)` / `Sheets(key)` | **Active** — string name or 1-based index; usable directly or with `Set ws = ...` |
| `ws.Name` | **Active** — readable/writable; live Worksheet and Range references survive the rename |
| `ws.Index` | **Active** — read-only, 1-based workbook tab position |
| `ws.Visible` | **Active** — readable/writable with `xlSheetVisible`, `xlSheetHidden`, or `xlSheetVeryHidden`; hiding the active or last visible sheet is rejected |
| `ws.UsedRange` | **Active** — returns the occupied rectangle, or `A1` for an empty sheet |
| `ws.Activate` / `ws.Select` | **Active** — changes `ActiveSheet`; hidden sheets are rejected |
| `ThisWorkbook.Name` / `ActiveWorkbook.Name` / `wb.Name` | **Active** — current loaded workbook name |
| `ThisWorkbook.ActiveSheet` / `wb.ActiveSheet` | **Active** — returns the current Worksheet reference |
| `Worksheets.Count` / `Sheets.Count` / `wb.Worksheets.Count` | **Active** — current worksheet count |

This is an in-memory object-model subset. Save/SaveAs/Close, external links,
windows, chart creation/general editing, Pivot cache recalculation, and other
desktop or external-effect members are not part of the supported contract.
Loaded-workbook OOXML editing additionally exposes bounded Python methods for
selected existing Chart-series formulas, Chart title text, two-cell Drawing
anchor markers, and worksheet-backed Pivot source fields; these are not VBA
object-model members.

### Application Object

| Property / Method | Behavior |
|---|---|
| `Application.Calculation = xlCalculationManual` | **Active** — disables auto-recalculation |
| `Application.Calculation = xlCalculationAutomatic` | **Active** — re-evaluates all formula cells |
| `Application.ScreenUpdating = False/True` | **No-op** (no screen) |
| `Application.EnableEvents = False/True` | **Active** — controls explicit `Vm.run_event` / Python `Vm.run_event` and `run_worksheet_change` dispatch; ordinary cell writes do not auto-fire events |
| `Application.DisplayAlerts = False/True` | **No-op** (no dialogs) |
| `Application.StatusBar = "..."` / `False` | **No-op** (no UI) |
| `Application.Cursor = xlWait` / `xlDefault` | **No-op** (no UI) |
| `Application.CutCopyMode = False` | **No-op** (no clipboard) |

---

## Worksheet Functions

Function names below describe implemented subsets, not independently verified
Excel equivalence. The evaluator currently rejects sheet-qualified references
(including an explicitly qualified reference to the current sheet); parsing and
reference rewriting do not imply evaluation support. Workbook-wide evaluation,
per-function mode coverage, and independent oracle tests are tracked in
[G3–G4](ROADMAP.md).

### Arithmetic & Statistical

| Function | Description | Excel |
|---|---|---|
| `SUM` | Sum of a range | Classic |
| `AVERAGE` | Average of a range | Classic |
| `MIN` / `MAX` | Minimum / Maximum value | Classic |
| `COUNT` / `COUNTA` | Count of numeric / non-empty cells | Classic |
| `PRODUCT` | Product of all values | Classic |
| `MEDIAN` | Median value | Classic |
| `LARGE` / `SMALL` | Kth largest / smallest | Classic |
| `RANK` | Rank of a number | Classic |
| `ROUND` / `ROUNDUP` / `ROUNDDOWN` | Rounding functions | Classic |
| `INT` | Floor (toward negative infinity) | Classic |
| `TRUNC` | Truncate toward zero | Classic |
| `MOD` | Modulo | Classic |
| `RAND` | Random float 0–1 | Classic |
| `RANDBETWEEN` | Random integer in range | Classic |
| `RANK` / `RANK.EQ` / `RANK.AVG` | Rank a number in a list | Classic / 2010 / 2010 |
| `SUMPRODUCT` | Sum of element-wise products | Classic |
| `COMBIN` | Number of combinations C(n, k) | Classic |
| `COUNTIF` / `SUMIF` / `AVERAGEIF` | Conditional count / sum / average | 2007 |
| `COUNTIFS` / `SUMIFS` / `AVERAGEIFS` | Multi-criteria count / sum / average | 2007 |
| `SUBTOTAL` | Aggregate with selectable function (1–11, 101–111; hidden rows are not modeled) | Classic |
| `AGGREGATE` | Extended subtotal (1–21; array/reference option filtering remains bounded) | 2010 |
| `PERCENTILE` / `PERCENTILE.INC` | Percentile (inclusive) | Classic / 2010 |
| `PERCENTILE.EXC` | Percentile (exclusive) | 2010 |
| `PERCENTRANK` / `PERCENTRANK.INC` | Percent rank | Classic / 2010 |
| `PERCENTRANK.EXC` | Percent rank (exclusive) | 2010 |
| `MODE.MULT` | Most frequent value | 2010 |
| `MINIFS` / `MAXIFS` | Conditional min / max | 2019 |
| `SUMSQ` | Sum of squares | Classic |
| `GEOMEAN` | Geometric mean | Classic |
| `HARMEAN` | Harmonic mean | Classic |
| `DEVSQ` | Sum of squared deviations | Classic |
| `AVEDEV` | Average absolute deviation | Classic |
| `QUARTILE` / `QUARTILE.INC` | Inclusive quartile | Classic / 2010 |
| `QUARTILE.EXC` | Exclusive quartile | 2010 |

### Financial

| Function | Description | Excel |
|---|---|---|
| `PMT`     | Periodic payment for a loan (`rate, nper, pv, [fv], [type]`) | Classic |
| `FACT`    | Factorial n! | Classic |
| `PERMUT`  | Permutations P(n, k) | Classic |
| `GCD`     | Greatest common divisor (varargs) | Classic |
| `LCM`     | Least common multiple (varargs) | Classic |
| `QUOTIENT` | Integer division (truncated quotient) | Classic |
| `SIGN`    | Sign of a number (−1, 0, or 1) | Classic |
| `FV`   | Future value of an investment | Classic |
| `PV`   | Present value of an investment | Classic |
| `NPER` | Number of periods for an annuity | Classic |
| `RATE` | Interest rate per period (Newton-Raphson) | Classic |
| `IPMT` | Interest portion of a periodic payment | Classic |
| `ISPMT` | Interest payment for a straight-line principal schedule | Classic |
| `PPMT` | Principal portion of a periodic payment | Classic |
| `NPV`  | Net present value (regular cash flows) | Classic |
| `IRR`  | Internal rate of return (regular cash flows) | Classic |
| `MIRR` | Modified internal rate of return | Classic |
| `XNPV` | Net present value (irregular cash flows, date-weighted) | Classic |
| `XIRR` | Internal rate of return (irregular cash flows) | Classic |
| `SLN` | Straight-line depreciation | Classic |
| `SYD` | Sum-of-years' digits depreciation | Classic |
| `DB` | Fixed-declining balance depreciation | Classic |
| `DDB` | Double-declining balance depreciation | Classic |
| `VDB` | Variable declining-balance depreciation with optional straight-line switch | Classic |
| `AMORLINC` | Linear depreciation for an asset purchased during a period | Classic |
| `AMORDEGRC` | Declining-balance depreciation for an asset purchased during a period | Classic |
| `EFFECT` | Effective annual interest rate | Classic |
| `NOMINAL` | Nominal annual interest rate | Classic |
| `RRI` | Equivalent growth interest rate | 2013 |
| `CUMIPMT` / `CUMPRINC` | Cumulative interest / principal paid between periods | 2010 |
| `FVSCHEDULE` | Future value using a sequence of rates | Classic |
| `DOLLARDE` / `DOLLARFR` | Convert fractional dollar quotations | Classic |
| `PDURATION` | Number of periods for a growth investment | 2013 |
| `PRICEDISC` | Price of a discounted security | Classic |
| `DISC` | Discount rate of a security | Classic |
| `RECEIVED` | Amount received at maturity for a discounted security | Classic |
| `YIELDDISC` | Annual yield of a discounted security | Classic |
| `ACCRINT` | Accrued interest for a periodic coupon security | Classic |
| `ACCRINTM` | Accrued interest paid at maturity | Classic |
| `COUPDAYBS` | Days from coupon period start to settlement | Classic |
| `COUPDAYS` | Days in coupon period | Classic |
| `COUPDAYSNC` | Days from settlement to next coupon | Classic |
| `COUPNCD` / `COUPPCD` | Next / previous coupon date | Classic |
| `COUPNUM` | Number of coupons remaining | Classic |
| `INTRATE` | Interest rate for a fully invested security | Classic |
| `PRICEMAT` | Price of a security paying interest at maturity | Classic |
| `YIELDMAT` | Yield of a security paying interest at maturity | Classic |
| `DURATION` | Macaulay duration of a security | Classic |
| `MDURATION` | Modified duration of a security | 2010 |
| `PRICE` | Price of a coupon-bearing security | Classic |
| `YIELD` | Yield of a coupon-bearing security | Classic |
| `ODDFPRICE` / `ODDFYIELD` | Price / yield of a security with an odd first coupon period | Classic |
| `ODDLPRICE` / `ODDLYIELD` | Price / yield of a security with an odd last coupon period | Classic |
| `TBILLPRICE` / `TBILLYIELD` / `TBILLEQ` | Treasury bill price, yield, and equivalent yield | Classic |

### Statistical

| Function | Description | Excel |
|---|---|---|
| `STDEV` / `STDEV.S` | Sample standard deviation | Classic / 2010 |
| `STDEVA` | Sample standard deviation including logical values and text | Classic |
| `STDEVP` / `STDEV.P` | Population standard deviation | Classic / 2010 |
| `STDEVPA` | Population standard deviation including logical values and text | Classic |
| `VARA` | Sample variance including logical values and text | Classic |
| `AVERAGEA` | Average including logical values and text as zero | Classic |
| `MINA` / `MAXA` | Minimum / maximum including logical values and text as zero | Classic |
| `VARPA` | Population variance including logical values and text | Classic |
| `VAR` / `VAR.S` | Sample variance | Classic / 2010 |
| `VARP` / `VAR.P` | Population variance | Classic / 2010 |
| `CORREL` | Pearson correlation coefficient | Classic |
| `PEARSON` | Pearson correlation coefficient (alias) | Classic |
| `SLOPE` | Slope of a linear regression | Classic |
| `INTERCEPT` | Y-intercept of a linear regression | Classic |
| `RSQ` | R-squared of a linear regression | Classic |
| `FORECAST.LINEAR` / `FORECAST` | Linear regression forecast | 2016 / Classic |
| `FORECAST.ETS` | Bounded seasonal forecast for a regular numeric timeline | 2016 |
| `FORECAST.ETS.SEASONALITY` | Detect a bounded repeating seasonality length | 2016 |
| `FORECAST.ETS.CONFINT` | Bounded forecast interval from seasonal-naive residuals | 2016 |
| `FORECAST.ETS.STAT` | Bounded ETS diagnostic metrics (1-8) | 2016 |
| `TREND` | Linear regression predictions | Classic |
| `GROWTH` | Exponential regression predictions | Classic |
| `LINEST` | Linear regression coefficients and optional statistics | Classic |
| `LOGEST` | Exponential regression coefficients and optional statistics | Classic |
| `STEYX` | Standard error of predicted y values | Classic |
| `COVARIANCE.S` / `COVAR` | Sample covariance | Classic / 2010 |
| `COVARIANCE.P` | Population covariance | 2010 |
| `FTEST` | Two-sample F-test probability for equality of variances | Classic |
| `CHITEST` | Chi-square test probability for observed and expected data | Classic |
| `MODE` / `MODE.SNGL` / `MODE.MULT` | Most frequent numeric value | Classic / 2010 / 2010 |
| `FREQUENCY` | Frequency distribution over bins | Classic |
| `PROB` | Probability of values within inclusive limits | Classic |
| `TRIMMEAN` | Mean after symmetric outlier trimming | Classic |
| `SKEW` | Sample skewness | Classic |
| `SKEW.P` | Population skewness | 2010 |
| `KURT` | Excess sample kurtosis | Classic |
| `NORM.DIST` / `NORMDIST` | Normal distribution CDF or PDF | 2010 / Classic |
| `NORM.INV` / `NORMINV` | Inverse normal distribution | 2010 / Classic |
| `T.DIST` | Student's t-distribution CDF or PDF | 2010 |
| `TDIST` | Legacy one- or two-tailed Student's t right-tail probability | Classic |
| `T.DIST.2T` | Two-tailed Student's t-distribution probability | 2010 |
| `T.DIST.RT` | Right-tailed Student's t-distribution probability | 2010 |
| `T.INV` | Inverse Student's t-distribution | 2010 |
| `T.INV.2T` / `TINV` | Two-tailed inverse Student's t-distribution | 2010 / Classic |
| `TTEST` | Student's t-test probability for paired or two-sample data | Classic |
| `NORM.S.DIST` / `NORMSDIST` | Standard normal distribution CDF or PDF | 2010 / Classic |
| `NORM.S.INV` / `NORMSINV` | Inverse standard normal distribution | 2010 / Classic |
| `BINOM.DIST` / `BINOMDIST` | Binomial distribution PMF or CDF | 2010 / Classic |
| `BINOM.DIST.RANGE` | Binomial probability over an inclusive success range | 2010 |
| `BINOM.INV` / `CRITBINOM` | Inverse binomial cumulative probability | 2010 / Classic |
| `NEGBINOM.DIST` / `NEGBINOMDIST` | Negative binomial distribution PMF or CDF | 2010 / Classic |
| `HYPGEOM.DIST` / `HYPGEOMDIST` | Hypergeometric distribution PMF or CDF | 2010 / Classic |
| `POISSON.DIST` / `POISSON` | Poisson distribution PMF or CDF | 2010 / Classic |
| `GAMMA` | Gamma function | 2010 |
| `GAMMALN` / `GAMMALN.PRECISE` | Natural logarithm of gamma function | 2010 / 2013 |
| `BETA.DIST` / `BETADIST` | Beta distribution PDF or CDF | 2010 / Classic |
| `BETA.INV` / `BETAINV` | Inverse beta distribution | 2010 / Classic |
| `CHISQ.DIST` / `CHIDIST` | Chi-squared distribution PDF or CDF | 2010 / Classic |
| `F.DIST` / `FDIST` | F distribution PDF or CDF | 2010 / Classic |
| `F.DIST.RT` | Right-tailed F distribution probability | 2010 |
| `F.DIST.2T` | Two-tailed F distribution probability | 2010 |
| `F.INV` / `F.INV.RT` / `FINV` | Inverse F distribution, left/right tail | 2010 / Classic |
| `GAMMA.DIST` / `GAMMADIST` / `GAMMA.INV` / `GAMMAINV` | Gamma distribution and inverse | 2010 / Classic |
| `CHISQ.DIST.RT` | Right-tailed chi-squared probability | 2010 |
| `CHISQ.INV` / `CHISQ.INV.RT` / `CHIINV` | Inverse chi-squared distribution, left/right tail | 2010 / Classic |
| `WEIBULL.DIST` / `WEIBULL` | Weibull distribution PDF or CDF | 2010 / Classic |
| `EXPON.DIST` / `EXPONDIST` | Exponential distribution PDF or CDF | 2010 / Classic |
| `LOGNORM.DIST` / `LOGNORMDIST` | Lognormal distribution PDF or CDF | 2010 / Classic |
| `LOGNORM.INV` / `LOGINV` | Inverse lognormal distribution | 2010 / Classic |
| `FISHER` / `FISHERINV` | Fisher transformation and inverse | Classic |
| `STANDARDIZE` | Convert a value to a standardized z-score | Classic |
| `Z.TEST` / `ZTEST` | One-tailed z-test p-value | 2010 / Classic |
| `CONFIDENCE.NORM` / `CONFIDENCE` | Normal-distribution confidence interval half-width | 2010 / Classic |
| `CONFIDENCE.T` | Student's t confidence interval half-width | 2010 |
| `PHI` | Standard normal probability density | 2013 |
| `GAUSS` | Standard normal CDF minus 0.5 | 2013 |

### Math & Trigonometry

| Function | Description | Excel |
|---|---|---|
| `ABS` | Absolute value | Classic |
| `ISEVEN` / `ISODD` | Test the parity of the truncated integer value | Classic |
| `SQRT` | Square root | Classic |
| `SQRTPI` | Square root of a number multiplied by π | Classic |
| `POWER` | Base raised to an exponent | Classic |
| `EXP` | e raised to a power | Classic |
| `LN` | Natural logarithm | Classic |
| `LOG` / `LOG10` | Logarithm (any base / base 10) | Classic |
| `PI` | Value of π | Classic |
| `SIN` / `COS` / `TAN` | Sine / cosine / tangent | Classic |
| `ASIN` / `ACOS` / `ATAN` / `ATAN2` | Inverse trig functions | Classic |
| `DEGREES` / `RADIANS` | Angle conversion | Classic |
| `FLOOR` / `CEILING` | Round down / up to nearest integer | Classic |
| `FLOOR.MATH` / `CEILING.MATH` | Round down / up to nearest multiple | 2013 |
| `MROUND` | Round to nearest multiple | Classic |
| `EVEN` / `ODD` | Round away from zero to an even / odd integer | Classic |
| `CEILING.PRECISE` / `FLOOR.PRECISE` | Sign-independent multiple rounding | 2010 |
| `ISO.CEILING` / `ECMA.CEILING` / `ISO.FLOOR` | ISO sign-independent multiple rounding aliases | 2010 / Classic |
| `COMBINA` | Combinations with repetitions | 2013 |
| `FACTDOUBLE` | Double factorial | Classic |
| `MULTINOMIAL` | Multinomial coefficient | Classic |
| `PERMUTATIONA` | Permutations with repetitions | 2013 |
| `SUMX2MY2` / `SUMX2PY2` / `SUMXMY2` | Pairwise sums of squares and square differences | Classic |
| `SERIESSUM` | Sum a power series using coefficient values | Classic |
| `SINH` / `COSH` / `TANH` | Hyperbolic sine / cosine / tangent | Classic |
| `ASINH` / `ACOSH` / `ATANH` | Inverse hyperbolic functions | 2013 |
| `SEC` / `SECH` / `CSC` / `CSCH` | Secant / hyperbolic secant / cosecant variants | 2013 |
| `COT` / `COTH` | Cotangent / hyperbolic cotangent | 2013 |
| `ACOT` / `ACOTH` | Inverse cotangent / inverse hyperbolic cotangent | 2013 |

### Engineering

| Function | Description | Excel |
|---|---|---|
| `BESSELJ` | Bessel function of the first kind | Classic |
| `BESSELI` | Modified Bessel function of the first kind | Classic |
| `BESSELY` | Bessel function of the second kind | Classic |
| `BESSELK` | Modified Bessel function of the second kind | Classic |
| `BITAND` / `BITOR` / `BITXOR` | Bitwise AND / OR / XOR for non-negative 48-bit integers | 2013 |
| `BITLSHIFT` / `BITRSHIFT` | Bitwise left / right shift | 2013 |
| `DELTA` | Returns 1 when two numbers are equal | Classic |
| `GESTEP` | Returns 1 when a number is greater than or equal to a step | Classic |
| `ERF` / `ERF.PRECISE` | Error function over one point or an interval | 2010 / 2013 |
| `ERFC` / `ERFC.PRECISE` | Complementary error function | 2010 / 2013 |

| `DEC2BIN` / `DEC2HEX` / `DEC2OCT` | Convert decimal integers to binary / hexadecimal / octal text | Classic |
| `BIN2DEC` / `HEX2DEC` / `OCT2DEC` | Convert binary / hexadecimal / octal text to decimal integers | Classic |
| `BIN2HEX` / `BIN2OCT` | Convert binary text to hexadecimal / octal text | Classic |
| `HEX2BIN` / `HEX2OCT` | Convert hexadecimal text to binary / octal text | Classic |
| `OCT2BIN` / `OCT2HEX` | Convert octal text to binary / hexadecimal text | Classic |
| `CONVERT` | Convert between supported length, mass, time, area, volume, temperature, pressure, energy, power, speed, force, frequency, and information units | Classic |
| `ROMAN` / `ARABIC` | Convert between integers and canonical Roman numerals | Classic |
| `COMPLEX` | Construct a complex number string | Classic |
| `IMREAL` / `IMAGINARY` / `IMARGUMENT` | Extract real / imaginary part or argument of a complex number | Classic |
| `IMABS` | Absolute value of a complex number | Classic |
| `IMSUM` / `IMSUB` | Add / subtract complex numbers | Classic |
| `IMPRODUCT` / `IMDIV` | Multiply / divide complex numbers | Classic |
| `IMCONJUGATE` | Complex conjugate | Classic |
| `IMEXP` / `IMLN` / `IMLOG10` / `IMLOG2` | Complex exponential and logarithms | Classic |
| `IMSQRT` / `IMPOWER` | Complex square root / power | Classic |
| `IMSIN` / `IMCOS` / `IMTAN` | Complex sine / cosine / tangent | Classic |
| `IMSINH` / `IMCOSH` / `IMTANH` | Complex hyperbolic functions | Classic |
| `IMSEC` / `IMCSC` / `IMCOT` | Complex reciprocal trigonometric functions | Classic |
| `IMASIN` / `IMACOS` / `IMATAN` / `IMACOT` | Complex inverse trigonometric functions | Classic |
| `IMASINH` / `IMACOSH` / `IMATANH` | Complex inverse hyperbolic functions | Classic |
| `IMSECH` / `IMCSCH` / `IMCOTH` | Complex reciprocal hyperbolic functions | Classic |

### Logical

| Function | Description | Excel |
|---|---|---|
| `IF` | Conditional value | Classic |
| `AND` / `OR` / `NOT` | Logical operators | Classic |
| `IFERROR` | Fallback on any error | 2007 |
| `IFNA` | Fallback on `#N/A` only | 2013 |
| `XOR` | Exclusive OR | 2013 |
| `IFS` | Multi-condition branch | 2019 |
| `SWITCH` | Switch/case | 2019 |

### Text

| Function | Description | Excel |
|---|---|---|
| `LEFT` / `RIGHT` / `MID` | Extract characters | Classic |
| `LEFTB` / `RIGHTB` / `MIDB` | Extract by DBCS bytes | Classic |
| `LEN` / `LENB` | Character / byte count | Classic |
| `UPPER` / `LOWER` / `PROPER` | Case conversion | Classic |
| `PHONETIC` | Deterministic source-text fallback when OOXML phonetic runs are unavailable | Classic |
| `TRIM` | Remove extra spaces | Classic |
| `FIND` | Case-sensitive position search | Classic |
| `FINDB` | Case-sensitive byte-position search using DBCS widths | Classic |
| `SEARCH` | Case-insensitive wildcard search | Classic |
| `SEARCHB` | Case-insensitive wildcard search returning a DBCS byte position | Classic |
| `SUBSTITUTE` | Replace by value | Classic |
| `REPLACE` | Replace by position | Classic |
| `REPLACEB` | Replace by DBCS byte position and length | Classic |
| `CONCATENATE` | Concatenate strings (legacy) | Classic |
| `TEXT` | Format number as string | Classic |
| `VALUE` | Parse string to number | Classic |
| `EXACT` | Case-sensitive equality | Classic |
| `CHAR` | Character from code point | Classic |
| `CODE` | Code point of first character | Classic |
| `ASC` | Full-width → half-width (DBCS) | Classic |
| `JIS` | Half-width → full-width (DBCS) | Classic |
| `UNICHAR` | Character from Unicode code point | 2013 |
| `UNICODE` | Unicode code point of first character | 2013 |
| `CONCAT` | Concatenate strings / ranges | 2019 |
| `TEXTJOIN` | Join with delimiter | 2019 |
| `TEXTSPLIT` | Split text into a bounded row/column array with ignore, match, and pad modes | 2024/365 |
| `TEXTBEFORE` | Extract text before the Nth occurrence of a delimiter | 2024/365 |
| `TEXTAFTER` | Extract text after the Nth occurrence of a delimiter | 2024/365 |
| `VALUETOTEXT` | Convert any value to its text representation | 2024/365 |
| `ARRAYTOTEXT` | Convert a bounded array to deterministic text | 2024/365 |
| `ENCODEURL` | Percent-encode text as UTF-8 for a URL without performing network I/O | 2013 |
| `REGEXTEST` | Test text against a Rust regular expression | 2024/365 |
| `REGEXEXTRACT` | Extract the first, all, or first capture-group matches | 2024/365 |
| `REGEXREPLACE` | Replace all regular-expression matches without I/O | 2024/365 |
| `FILTERXML` | Extract values using a bounded local XML element/attribute XPath subset | 2013 |
| `HYPERLINK` | Return display text for a hyperlink without opening external resources | Classic |
| `REPT` | Repeat a text string N times | Classic |
| `NUMBERVALUE` | Parse number with locale-specific decimal/group separators | 2013 |
| `CLEAN` | Remove non-printing control characters | Classic |
| `T` | Return text or empty text for non-text values | Classic |
| `FIXED` | Format a number with fixed decimal places | Classic |
| `DOLLAR` | Format a number as currency text | Classic |
| `BAHTTEXT` | Format a number as Thai Baht text | Classic |
| `BASE` | Convert an integer to a string in a base from 2 to 36 | 2013 |
| `DECIMAL` | Convert a base-2-to-36 string to an integer | 2013 |

### Date & Time

| Function | Description | Excel |
|---|---|---|
| `DATE` | Create date serial (Excel epoch) | Classic |
| `TODAY` / `NOW` | Today's date / current datetime | Classic |
| `YEAR` / `MONTH` / `DAY` | Extract date parts | Classic |
| `WEEKDAY` | Day of week (1–3 return types) | Classic |
| `DATEDIF` | Difference in Y / M / D / MD / YM / YD | Classic |
| `DATEVALUE` | Parse "YYYY/MM/DD" or "YYYY-MM-DD" | Classic |
| `TIME` | Create time serial | Classic |
| `TIMEVALUE` | Parse "HH:MM:SS" | Classic |
| `HOUR` / `MINUTE` / `SECOND` | Extract time parts | Classic |
| `EOMONTH` | Last day of month N months from start | 2007 |
| `EDATE` | Date N months from start | 2007 |
| `NETWORKDAYS` | Workdays between dates (Sat+Sun weekend) | 2007 |
| `WORKDAY` | Date N workdays away (Sat+Sun weekend) | 2007 |
| `NETWORKDAYS.INTL` | Workdays with custom weekend | 2010 |
| `WORKDAY.INTL` | Date N workdays away with custom weekend | 2010 |
| `DAYS` | Days between two dates | 2013 |
| `DAYS360` | Days between dates using 360-day year conventions | Classic |
| `YEARFRAC` | Fraction of a year between two dates | Classic |
| `WEEKNUM` | Week number in a year (types 1/2/11-17/21) | Classic |
| `ISOWEEKNUM` | ISO 8601 week number | 2013 |

### Lookup & Reference

| Function | Description | Excel |
|---|---|---|
| `VLOOKUP` / `HLOOKUP` | Vertical / horizontal lookup | Classic |
| `INDEX` | Value or row/column array at an offset | Classic |
| `MATCH` | Position of a value (including exact-match wildcards) | Classic |
| `CHOOSE` | Choose from list by index | Classic |
| `INDIRECT` | Evaluate a cell reference from a string | Classic |
| `OFFSET` | Cell reference shifted by rows/cols | Classic |
| `ADDRESS` | Cell address as string (e.g. `"$A$1"`) | Classic |
| `COUNTBLANK` | Count blank cells in a range | Classic |
| `ROW` / `COLUMN` | Row / column number of reference | Classic |
| `ROWS` / `COLUMNS` | Number of rows / columns in a reference | Classic |
| `LOOKUP` | Sorted vector lookup | Classic |
| `TRANSPOSE` | Transpose rows and columns | Classic |
| `XLOOKUP` | Flexible lookup (exact, wildcard, next-larger, next-smaller; binary search on sorted numeric ranges) | 365/2021 |
| `XMATCH` | Extended MATCH with exact, wildcard, and binary search modes | 365/2021 |

### Information

| Function | Description | Excel |
|---|---|---|
| `ISBLANK` | Is the value empty? | Classic |
| `ISERROR` / `ISERR` | Is the value an error? | Classic |
| `ISNA` | Is the value #N/A? | Classic |
| `ISNUMBER` | Is the value numeric? | Classic |
| `ISTEXT` | Is the value a string? | Classic |
| `ISLOGICAL` | Is the value a boolean? | Classic |
| `ISNONTEXT` | Is the value not a string? | Classic |
| `ISREF` | Is the expression a reference? | Classic |
| `ISFORMULA` | Test whether a cell or bounded range contains a formula | 2013 |
| `AREAS` | Return the number of areas represented by a reference | Classic |
| `N` | Convert a value to a number | Classic |
| `NA` | Generate an #N/A error value | Classic |
| `TYPE` | Return numeric type code (1=num, 2=text, 4=bool, 16=error, 64=array) | Classic |
| `ERROR.TYPE` | Return numeric code of an error (#NULL!=1 … #N/A=7) | Classic |
| `FORMULATEXT` | Return formula of a cell as text | 2013 |
| `CELL` | Return cell metadata (address, row, col, type, contents, width, …) | Classic |
| `INFO` | Return bounded runtime information (`system`, `osversion`, `release`, `recalc`) | Classic |

### Array / Spill Functions

| Function | Description | Excel |
|---|---|---|
| `FILTER` | Filter array by boolean condition | 365/2021 |
| `UNIQUE` | Unique values from a range | 365/2021 |
| `SORT` | Sort a range or array | 365/2021 |
| `SORTBY` | Sort by one or more external arrays | 365/2021 |
| `SEQUENCE` | Generate a sequence of numbers | 365/2021 |
| `RANDARRAY` | Generate a random number array | 365/2021 |
| `TOCOL` | Convert range/array to a single column | 2024/365 |
| `TOROW` | Convert range/array to a single row | 2024/365 |
| `WRAPCOLS` | Wrap 1D array into multiple columns | 2024/365 |
| `WRAPROWS` | Wrap 1D array into multiple rows | 2024/365 |
| `TAKE` | Take first (or last) N elements from an array | 2024/365 |
| `DROP` | Drop first (or last) N elements from an array | 2024/365 |
| `VSTACK` | Stack arrays vertically (concatenate) | 2024/365 |
| `HSTACK` | Stack arrays horizontally (concatenate) | 2024/365 |
| `CHOOSECOLS` | Select specific columns by 1-based index | 2024/365 |
| `CHOOSEROWS` | Select specific rows by 1-based index | 2024/365 |
| `EXPAND` | Expand an array to a requested rectangle with a padding value | 2024/365 |
| `TRIMRANGE` | Remove blank outer rows and columns from an array | 2024/365 |
| `SINGLE` | Return the first value of a bounded array (implicit intersection) | 365/2021 |
| `GROUPBY` | Group a one-column array and aggregate values with a bounded reducer | 2024/365 |
| `PIVOTBY` | Cross-tabulate one-column row and column keys with a bounded reducer | 2024/365 |
| `MAKEARRAY` | Generate an array by calling a LAMBDA with row and column indices | 2024/365 |
| `MUNIT` | Identity matrix | Classic |
| `MMULT` | Matrix multiplication | Classic |
| `MDETERM` | Matrix determinant | Classic |
| `MINVERSE` | Matrix inverse | Classic |

### Lambda & Higher-Order Functions

| Function | Description | Excel |
|---|---|---|
| `LET` | Assign named variables within a formula | 365/2021 |
| `LAMBDA` | Define an anonymous function | 365/2021 |
| `MAP` | Apply LAMBDA to each element | 365/2021 |
| `REDUCE` | Reduce array to a single value via LAMBDA | 365/2021 |
| `SCAN` | Cumulative reduce (returns all partial results) | 365/2021 |
| `BYROW` | Apply LAMBDA to each row | 365/2021 |
| `BYCOL` | Apply LAMBDA to each column | 365/2021 |

### Database

| Function | Description | Excel |
|---|---|---|
| `DGET`     | Extract a single value from a database matching criteria | Classic |
| `DSUM`     | Sum values in a filtered database column | Classic |
| `DAVERAGE` | Average values in a filtered database column | Classic |
| `DCOUNT`   | Count numeric values in a filtered database column | Classic |
| `DCOUNTA`  | Count all values in a filtered database column | Classic |
| `DMAX`     | Maximum value in a filtered database column | Classic |
| `DMIN`     | Minimum value in a filtered database column | Classic |
| `DPRODUCT` | Product of numeric values in a filtered database column | Classic |
| `DSTDEV`   | Sample standard deviation of a filtered database column | Classic |
| `DSTDEVP`  | Population standard deviation of a filtered database column | Classic |
| `DVAR`     | Sample variance of a filtered database column | Classic |
| `DVARP`    | Population variance of a filtered database column | Classic |

---

## Criteria Syntax (COUNTIF / SUMIF / SUMIFS / etc.)

| Criteria | Example | Meaning |
|---|---|---|
| Number | `10` | Exact numeric match |
| String | `"apple"` | Case-insensitive string match |
| Comparison | `">5"`, `"<=10"`, `"<>"` | Numeric comparison |
| Wildcard | `"a*"`, `"?bc"` | `*` = any chars, `?` = one char |

---

## Not Yet Supported

### Excluded (out of scope)

| Function | Reason |
|---|---|
| `IMAGE(source, ...)` | Fetches images from URLs — not applicable in a headless VBA emulator |
