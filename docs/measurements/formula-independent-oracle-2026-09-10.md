# Formula-only independent oracle

Date: 2026-09-10 (Asia/Tokyo)

The new formula-only harness generates an XLSX containing formulas, asks
headless LibreOffice to recalculate it, and reads cached results. It does not
invoke LibreOffice Basic, `Range`, or `Cells`, so it is independent of the
known VBA object-model timeout in the corpus runner. LibreOffice remains an
independent oracle, not Microsoft Excel.

## Result

- Host: macOS arm64
- Backend: `/opt/homebrew/bin/soffice`
- Cases: 15 (`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNTA`, `IF`, `IFERROR`, `ROUND`, `DATE`, `LEFT`, `MID`, `LEN`, `CONCATENATE`, `MATCH`, `VLOOKUP`)
- Comparable results: 15
- Matches: 15
- Mismatches: 0
- Command: `python3 compat/oracle/run-libreoffice-formulas.py --soffice /opt/homebrew/bin/soffice`

The harness exits with status 1 if any comparable case mismatches; this run
exited with status 0.

The `DATE` result was normalized from Python `datetime` to Excel serial
(`45351`) before comparison. This is a fixture-level type normalization, not a
claim that all date-system behavior is compatible.

## Reproduction

```sh
python3 compat/oracle/run-libreoffice-formulas.py \
  --soffice /opt/homebrew/bin/soffice
```

The script writes its temporary workbook and LibreOffice profile below a
temporary directory and removes them after the run. This record does not close
Excel oracle validation, 1904 date-system coverage, error/type coercion,
dynamic-array boundaries, or the EPPlus/Aspose comparison.
