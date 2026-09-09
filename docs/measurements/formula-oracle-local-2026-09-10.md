# Formula oracle local rerun

Date: 2026-09-10 (Asia/Tokyo)

This is a small, local LibreOffice oracle rerun. LibreOffice is an independent
calculation engine, not Microsoft Excel; the result must not be presented as
Excel compatibility evidence.

## Conditions and result

- Host: macOS arm64
- Backend: `/opt/homebrew/bin/soffice` (headless)
- Corpus: first 8 arithmetic scenarios (`arithmetic_0001`–`arithmetic_0008`)
- Command: `node run-libreoffice.mjs 8 0 _local20260910`
- Per-scenario timeout: 8 seconds (runner default)
- Completed: 0
- Timeout: 8
- Comparable cell outputs: 0

Raw output is [libreoffice-results_local20260910.json](../../compat/corpus/results/libreoffice-results_local20260910.json).
The runner classified all eight cases as `TIMEOUT`; no result was promoted to
`MATCH`, and no elixcee-versus-LibreOffice correctness claim can be made from
this run. The failure is consistent with the documented Range/Cells macro
invocation boundary in the corpus harness and does not establish an Excel
result.

## Reproduction

```sh
cd compat/corpus
node run-libreoffice.mjs 8 0 _local20260910
```

The full formula oracle and Microsoft Excel oracle remain open. In particular,
this record does not close the G4 type-conversion, rounding, date-system,
array-boundary, or Excel-reopen requirements.
