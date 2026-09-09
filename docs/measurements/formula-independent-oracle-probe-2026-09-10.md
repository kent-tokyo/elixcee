# Formula oracle follow-up probe

Date: 2026-09-10 (Asia/Tokyo)

This records the follow-up probes added to
`compat/oracle/run-libreoffice-formulas.py`. LibreOffice is an independent
calculation reference, not Microsoft Excel.

## Result

- total cases in the runner: 93
- LibreOffice-comparable cases: 85
- LibreOffice matches: 85/85
- elixcee-comparable cases: 85
- elixcee matches: 85/85
- build-specific skips: 8 (`DAYS`, `IFNA`, `MAXIFS`, `MINIFS`, `XMATCH`,
  `TEXTJOIN`, `ISOWEEKNUM`, `TYPE(TRUE)`)
- added probe coverage: `INT`, `TRUNC`, `SIGN`, `SQRT`, `ROWS`, `COLUMNS`,
  `ISLOGICAL`, boolean/text `N` coercion, `WEEKNUM`, `NETWORKDAYS`, `HOUR`,
  `MINUTE`, `SECOND`, `ISERR`, `ISNA`, `ISNONTEXT`, and `TYPE`

## Reproduction

```bash
python3 compat/oracle/run-libreoffice-formulas.py \
  --soffice /opt/homebrew/bin/soffice
```

The paired result was produced with a wheel built from the current source
(`elixcee` 1.0.5) and is stored at
`compat/corpus/results/formula-independent-20260910.json`. LibreOffice's
`TYPE(TRUE)` result is excluded because this build returns `1` where Excel's
logical type code is `4`. The probe must not be read as Microsoft Excel
compatibility evidence; Excel reopen and Excel oracle validation remain open.
