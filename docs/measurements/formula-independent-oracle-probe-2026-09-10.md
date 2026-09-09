# Formula oracle follow-up probe

Date: 2026-09-10 (Asia/Tokyo)

This is a LibreOffice-only probe for the additional cases in
`compat/oracle/run-libreoffice-formulas.py`. It extends the existing
72-case artifact without replacing the paired elixcee/LibreOffice result.
LibreOffice is an independent calculation reference, not Microsoft Excel.

## Result

- total cases in the runner: 80
- LibreOffice-comparable cases: 74
- LibreOffice matches: 74/74
- build-specific skips: 6 (`DAYS`, `IFNA`, `MAXIFS`, `MINIFS`, `XMATCH`,
  `TEXTJOIN`)
- added probe coverage: `INT`, `TRUNC`, `SIGN`, `SQRT`, `ROWS`, `COLUMNS`,
  `ISLOGICAL`, and boolean/text `N` coercion

## Reproduction

```bash
python3 compat/oracle/run-libreoffice-formulas.py \
  --soffice /opt/homebrew/bin/soffice
```

The paired elixcee result for these nine cases remains open because the
current Python binding is not installed in this environment. The existing
paired 65/65 result remains at
`compat/corpus/results/formula-independent-20260910.json`; this probe must
not be read as an elixcee match rate or Excel compatibility evidence.
