# Microsoft Excel formula oracle — 2026-09-11

Scope: Microsoft Excel for Mac 16.108 on the local macOS host, using a fresh
blank workbook created through the UI. Six formulas were entered through
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
`Vm.recalculate`, and `Vm.get_cell`. The result is six/ six matching probes,
not a complete Excel compatibility claim. Locale, date-system variants,
coercion boundaries, arrays, errors, other Excel versions, and all remaining
functions still require separate oracle coverage.
