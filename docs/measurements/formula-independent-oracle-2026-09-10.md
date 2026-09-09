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
- Cases: 34 (the original set plus logical, rounding, lookup, criteria, and text probes; `IFNA`, `XMATCH`, and `TEXTJOIN` are retained as oracle probes)
- Comparable results: 31
- Matches: 31
- Skipped: 3 (`IFNA`, `XMATCH`, and `TEXTJOIN`, because this LibreOffice build does not evaluate these probes)
- Mismatches: 0
- Command: `python3 compat/oracle/run-libreoffice-formulas.py --soffice /opt/homebrew/bin/soffice`

The harness exits with status 1 if any comparable case mismatches. Oracle
unsupported probes are reported in `skipped` and are not silently counted as
matches; this run exited with status 0.

The expanded run also covered logical operators, conditional aggregation,
rounding boundaries, INDEX/MATCH lookup, criteria aggregation, an unhandled
division error, error recovery, error inspection, and numeric/text mixed
ranges. Error text is
recorded as a cell result; it is not treated as a successful numeric coercion.

The `DATE` results were normalized from Python `datetime` to the 1900-system
Excel serial (`45351`) before comparison, including the workbook carrying
`workbookPr@date1904="1"`. This is a fixture-level type normalization, not a
claim that all date-system behavior is compatible; VM date-system metadata is
currently exposed separately and serial correction remains an open boundary.
The Rust regression also verifies that a loaded `date1904` flag survives a
VM save and reload; this checks metadata preservation only.

## Reproduction

```sh
python3 compat/oracle/run-libreoffice-formulas.py \
  --soffice /opt/homebrew/bin/soffice
```

The script writes its temporary workbook and LibreOffice profile below a
temporary directory and removes them after the run. This record does not close
Excel oracle validation, full 1904 date-system conversion, error/type coercion,
dynamic-array boundaries, or the EPPlus/Aspose comparison.
