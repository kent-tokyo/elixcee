# Excel function coverage inventory

Date: 2026-09-10 (Asia/Tokyo)

The current Microsoft Excel alphabetical function table was downloaded from
the official support page and compared with the literal dispatch names in
`src/formula/eval.rs`.

## Result

- Official function-table names: 481
- elixcee dispatch literals: 534 (canonical names plus aliases and explicit external boundaries)
- Missing names after the explicit external-service boundary: 0
- External-service / host-service boundary names: 14

The boundary set is `CALL`, `CUBE*`, `IMAGE`, `REGISTER.ID`, `RTD`,
`STOCKHISTORY`, `TRANSLATE`, and `WEBSERVICE`. These functions require code
loading, a workbook service, a remote resource, or an external data source;
they are not counted as headless calculation functions. They must not be
replaced with fabricated local answers.

## Reproduction

Download the current page, then run the repository audit:

```bash
python3 scripts/check-excel-function-coverage.py \
  /path/to/excel-functions-alphabetical.html \
  --check
```

The audit compares only function-table cells (`<td>` entries whose first child
is an uppercase function link) with the quoted names in the `match name`
dispatch in `src/formula/eval.rs`. The result above was obtained on
2026-09-10.

This inventory proves coverage against the current Microsoft table, not
semantic equivalence. Argument coercion, error precedence, date systems,
dynamic-array spill behavior, numerical accuracy, and Excel reopen/oracle
validation remain separate gates. It also does not establish superiority over
Aspose.Cells or EPPlus; that requires fixed versions and identical fixtures.
