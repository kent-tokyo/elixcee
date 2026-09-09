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
- Cases: 80 (the original set plus logical, rounding, lookup, criteria, text,
  numeric, date-boundary, and string-normalization probes; `DAYS`, `IFNA`,
  `MAXIFS`, `MINIFS`, `XMATCH`, and `TEXTJOIN`
  are retained as oracle probes)
- Comparable results: 74
- Matches: 74
- Skipped: 6 (`DAYS`, `IFNA`, `MAXIFS`, `MINIFS`, `XMATCH`, and `TEXTJOIN`, because this LibreOffice build does not evaluate these probes)
- Mismatches: 0
- Machine-readable result: `compat/corpus/results/formula-independent-20260910.json`
- Command: `python3 compat/oracle/run-libreoffice-formulas.py --soffice /opt/homebrew/bin/soffice`

The harness exits with status 1 if any comparable case mismatches. Oracle
unsupported probes are reported in `skipped` and are not silently counted as
matches; this run exited with status 0.

The expanded 80-case command was rerun after the later local changes using a
dedicated writable temporary directory (`TMPDIR=/private/tmp/elixcee-formula-oracle-run`)
because the default temporary directories were full or unavailable. It
reported 74/74 comparable matches, 6 explicit skips, and zero mismatches.

The same generated workbook was then recalculated by a wheel built from the
current source (`elixcee` 1.0.5) with `--with-elixcee`. After the documented
date-serial normalization, elixcee matched all 74 comparable probes (74/74);
the same six probes were
skipped because the LibreOffice oracle did not produce comparable values.
The wheel was elixcee 1.0.5 on CPython 3.13.6, macOS 26.5.2 arm64.
The binding-specific Error object was normalized to its displayed error code
for JSON output, while ordinary numeric and string values were compared without
coercion.

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

The nine newly added probes cover `INT`, `TRUNC`, `SIGN`, `SQRT`, `ROWS`,
`COLUMNS`, `ISLOGICAL`, and boolean/text `N` coercion. `ROWS` and `COLUMNS`
initially exposed a real `#NAME?` gap in elixcee; their implementation and
paired 74/74 result are included in the current artifact. See the
[follow-up probe](formula-independent-oracle-probe-2026-09-10.md).

## Reproduction

```sh
python3 compat/oracle/run-libreoffice-formulas.py \
  --soffice /opt/homebrew/bin/soffice
```

When the default temporary directory is unavailable, use:

```sh
TMPDIR=/private/tmp/elixcee-formula-oracle-run \
  python3 compat/oracle/run-libreoffice-formulas.py \
  --soffice /opt/homebrew/bin/soffice --output \
  compat/corpus/results/formula-independent-20260910.json
```

With the isolated elixcee wheel available on `PYTHONPATH`:

```sh
PYTHONPATH=/private/tmp/elixcee-test-venv-20260910/lib/python3.13/site-packages \
TMPDIR=/private/tmp/elixcee-formula-oracle-run \
  python3 compat/oracle/run-libreoffice-formulas.py \
  --soffice /opt/homebrew/bin/soffice --with-elixcee --output \
  compat/corpus/results/formula-independent-20260910.json
```

The script writes its temporary workbook and LibreOffice profile below a
temporary directory and removes them after the run. This record does not close
Excel oracle validation, full 1904 date-system conversion, error/type coercion,
dynamic-array boundaries, or the EPPlus/Aspose comparison.
