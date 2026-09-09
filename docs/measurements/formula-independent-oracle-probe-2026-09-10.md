# Formula oracle follow-up probe

Date: 2026-09-10 (Asia/Tokyo)

This records the follow-up probes added to
`compat/oracle/run-libreoffice-formulas.py`. LibreOffice is an independent
calculation reference, not Microsoft Excel.

## Result

- total cases in the runner: 86
- LibreOffice-comparable cases: 79
- LibreOffice matches: 79/79
- elixcee-comparable cases: 79
- elixcee matches: 79/79
- build-specific skips: 7 (`DAYS`, `IFNA`, `MAXIFS`, `MINIFS`, `XMATCH`,
  `TEXTJOIN`, `ISOWEEKNUM`)
- added probe coverage: `INT`, `TRUNC`, `SIGN`, `SQRT`, `ROWS`, `COLUMNS`,
  `ISLOGICAL`, boolean/text `N` coercion, `WEEKNUM`, `NETWORKDAYS`, `HOUR`,
  `MINUTE`, and `SECOND`

## Reproduction

```bash
python3 compat/oracle/run-libreoffice-formulas.py \
  --soffice /opt/homebrew/bin/soffice
```

The paired result was produced with a wheel built from the current source
(`elixcee` 1.0.5) and is stored at
`compat/corpus/results/formula-independent-20260910.json`. The probe must not
be read as Microsoft Excel compatibility evidence; Excel reopen and Excel
oracle validation remain open.
