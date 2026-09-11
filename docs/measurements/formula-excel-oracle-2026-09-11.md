# Microsoft Excel formula oracle — 2026-09-11

Scope: Microsoft Excel for Mac 16.108 on the local macOS host, using a fresh
blank workbook created through the UI. The initial six formulas were entered through
AppleScript, read back after Excel's automatic calculation, and then evaluated
with the current source-built CPython 3.13 arm64 wheel in an isolated
environment.

| case | formula | Excel | elixcee | result |
|---|---|---:|---:|:---:|
| maxifs | `=MAXIFS(A1:A3,A1:A3,">1")` | 3 | 3 | match |
| minifs | `=MINIFS(A1:A3,A1:A3,">1")` | 2 | 2 | match |
| days | `=DAYS(DATE(2024,3,1),DATE(2024,2,28))` | 2 | 2 | match |
| ifna | `=IFNA(VLOOKUP(9,A1:A3,1,FALSE),"missing")` | missing | missing | match |
| xmatch | `=XMATCH(2,A1:A3)` | 2 | 2 | match |
| textjoin | `=TEXTJOIN(",",TRUE,"a","b")` | a,b | a,b | match |

The source cells were `A1:A3 = 1,2,3`. The Excel side was driven by
`osascript`; the elixcee side used `Vm.set_cell`, `Vm.set_cell_formula`,
`Vm.recalculate`, and `Vm.get_cell`. The result is six/six matching probes,
not a complete Excel compatibility claim. Locale, date-system variants,
coercion boundaries, arrays, errors, other Excel versions, and all remaining
functions still require separate oracle coverage.

## Boundary extension

Eight additional probes were run in the same workbook and environment:

| case | formula | Excel | elixcee | result |
|---|---|---|---|:---:|
| round-positive | `=ROUND(2.675,2)` | 2.68 | 2.68 | match |
| round-negative | `=ROUND(-2.675,2)` | -2.68 | -2.68 | match |
| date-start | `=DATE(1900,1,1)` | 1900-01-01 | 1900-01-01 | match |
| date-leap-boundary | `=DATE(1900,2,29)` | 1900-03-01 | 1900-03-01 | match |
| iferror | `=IFERROR(1/0,"fallback")` | fallback | fallback | match |
| isblank | `=ISBLANK(C1)` | TRUE | TRUE | match |
| rows | `=ROWS(A1:B2)` | 2 | 2 | match |
| columns | `=COLUMNS(A1:B2)` | 2 | 2 | match |

The date cells were compared by calendar value because AppleScript returns
Excel date results as localized date objects while the binding returns Python
`date` values. The combined result is 14/14 matching probes on one host; it
still does not cover the full type-conversion, 1904-date, array-spill, or
recalculation surface.

The boundary rerun also included `=AVERAGE("1",2)`: Excel and the rebuilt
wheel both returned `1.5`. This exposed and fixed a coercion bug in the
evaluator: direct text/logical arguments are included by `AVERAGE`, while
text/logical values inside a referenced range remain excluded. The corrected
source and wheel now match all 15 recorded probes.

The same boundary check found that Excel returns `1` for both
`=MIN("1",TRUE)` and `=MAX("1",TRUE)`. The evaluator now routes these direct
arguments through the existing coercion path; text and logical values in a
referenced range remain excluded. Focused regressions cover both functions.

Excel also returns `2` for `=PRODUCT("2",TRUE)`. `PRODUCT` now uses the same
direct-argument coercion path, with a regression for the observed result.

The direct-argument boundary also measured `COUNT("1") = 1`, `COUNT(TRUE) = 1`,
and `COUNT("x") = 0` in Excel. `COUNT` now counts numeric text and logical
scalar arguments while continuing to ignore those values inside a reference.

For the empty/error boundary, Excel returned `COUNTA(C1) = 0` for a blank
cell, `COUNTA(C2) = 1` for a formula returning `""`, `COUNTA(C3) = 1` for a
formula error, `COUNTBLANK(C1:C3) = 2`, and `COUNTA("") = 1`,
`COUNTA(1/0) = 1`, `COUNTA("x") = 1`, `COUNTA(TRUE) = 1`. The current VM
matches all eight results.

## Mixed-type aggregate range probes

Five additional range probes used `A1:B2 = {{1,1},{TRUE,"x"}}` in the same
workbook. Excel and the rebuilt wheel matched all five results: `AVERAGEA`
returned `0.75`, `MINA` returned `0`, `MAXA` returned `1`, `COUNT` returned
`2`, and `COUNTA` returned `4`. These cases confirm the current distinction
between numeric-only `COUNT`, non-empty `COUNTA`, and the logical/text
coercion used by the `*A` aggregate family for referenced ranges. They remain
small single-host probes, not complete type-coercion coverage.

## Dynamic-array shape probes

Excel returned `21` for `SUM(SEQUENCE(2,3))`, `6` for both
`COUNT(SEQUENCE(2,3))` and `COUNTA(SEQUENCE(2,3))`, `2` for
`ROWS(SEQUENCE(2,3))`, `3` for `COLUMNS(SEQUENCE(2,3))`, `4` for
`INDEX(SEQUENCE(2,2),2,2)`, and `27` for `SUM(SEQUENCE(2,3)+1)`. The current
Rust evaluator has matching source-level regressions for these 2D spill and
aggregation boundaries. These are additional Excel observations; they are
not a claim of complete spill-shape metadata or cross-version parity.

Excel returned `12` for the row/column broadcast
`SUM(SEQUENCE(2,1)+SEQUENCE(1,2))` and an error for the incompatible
`SUM(SEQUENCE(2)+SEQUENCE(3))`. The evaluator now uses the existing formula
shape resolver for row/column broadcasting and rejects incompatible axes.

Excel returned `-21` for `SUM(-SEQUENCE(2,3))` and `-15` for
`SUM(-SEQUENCE(2,3)+1)`. Unary minus now maps over bounded formula arrays in
the evaluator, including element Error preservation.

The previously failing `SUM(SEQUENCE(2,3)+1)` case is now supported by
element-wise scalar/array broadcasting with bounded equal-length validation;
array errors remain represented as per-element Excel errors.

## Automatic dependency recalculation

In a separate three-cell chain, `P1 = 1`, `P2 = P1+1`, and `P3 = P2*2`
produced `4` in Excel. After changing `P1` to `3`, Excel automatically
recalculated `P3` to `8`; the current VM produced the same `4 → 8` sequence
after `set_cell` and `recalculate`. This is a focused automatic-recalculation
probe, not evidence for every cross-sheet, volatile, cycle, or spill rule.

## Cross-sheet dependency recalculation

With `OracleInput!A1 = 1`, `OracleCalc!B1 = OracleInput!A1+1`, and
`OracleCalc!C1 = B1*2`, Excel returned `C1 = 4`. After changing the input to
`3` and issuing Excel's explicit full calculation, `C1 = 8`; the current VM
returned the same `4 → 8` sequence after its workbook recalculation. The
Excel probe is explicitly a full-calculation result because the AppleScript
multi-cell input sequence does not establish an automatic-update timing
guarantee.

Excel's Error/array probes returned an error for
`SUM(CHOOSE(SEQUENCE(2),1,1/0))` and `AVERAGE(CHOOSE(SEQUENCE(2),1,1/0))`,
`1` for `COUNT(CHOOSE(SEQUENCE(2),1,1/0))`, `2` for the corresponding
`COUNTA`, `2` for `SUM(IF(SEQUENCE(2),1,1/0))` (lazy branch), and an error for
`SUM(SEQUENCE(2,3)+1/0)`. The current evaluator now has regressions for these
six array Error/lazy-evaluation boundaries.
