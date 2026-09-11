# G6 readiness record

Date: 2026-09-11 (Asia/Tokyo)

This record is the current completion audit for G6. A row is only marked
verified when the cited evidence covers the stated scope.

| Gate | Current evidence | Decision |
|---|---|---|
| G2 connection graph | `compat/ooxml-feature-matrix.json` records Chart/Drawing/Pivot relationship checks. | Partial: preservation is covered; Chart creation with bar or multiple charts still causes Excel recovery on macOS, including a data-row-only bar probe. |
| G3 workbook formulas | Rust regression suite plus workbook-level dirty-closure, cycle, rename, and event tests; 39 focused Microsoft Excel formula probes and a three-cell automatic dependency chain match the current evaluator. | Partial: implementation evidence exists, but complete Excel semantic oracle coverage is not established. |
| G4 independent oracle | LibreOffice 26.2.5 and a locally built 1.0.10 CPython 3.13 arm64 wheel evaluated 118 common probes, with 118/118 matches; separate Microsoft Excel probes now cover coercion, Empty/Error, dynamic arrays, and array errors. | Verified for these fixtures and host only; this is not cross-platform or full Excel evidence. |
| G5 memory | macOS arm64 append and normal-fresh measurements cover 100k/250k/1M rows and semantic output equality; the current 1.0.10 wheel passed a three-repetition 2026-09-11 append rerun at 100k/250k rows. | Partial: Linux/Windows matrix and full passthrough-laziness gate are not run. |
| G6 publication decision | Versioned matrix and dated measurement records are checked in. | Not complete: missing platform and Excel-reopen evidence prevents a full G6 sign-off. |

## Explicit remaining gates

1. Run the manual G5 matrix on Ubuntu, macOS, and Windows and validate every
   JSON artifact with `check-stream-writer-measurements.py`.
2. Resolve or formally bound the generated bar/multi-chart Excel recovery
   warning, then repeat the Excel reopen check.
3. Fix versions, calculation settings, and license conditions before any
   Excel／EPPlus／Aspose.Cells comparison; keep declared support separate from
   measured agreement.

The compatibility matrix intentionally keeps `excel_reopen` as `unverified`
for generated Chart/Drawing/Pivot features. A passing Rust or LibreOffice test
does not promote that status.
