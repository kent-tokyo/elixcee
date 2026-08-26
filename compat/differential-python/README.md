# Python differential tests

Compares elixcee's Python-native worksheet APIs against `openpyxl`'s own read
of the same fixtures — see `docs/openpyxl-gap-audit.md`:

- `bulk_range_check.py`: the bulk worksheet range/row API (R1 + P1 remainder)
  — `get_range`/`set_range`/`append_row`/`iter_rows`/`iter_cols`/`max_row`/
  `max_column`/`calculate_dimension`.
- `sheet_ops_check.py`: sheet management (P1 core 3 + remainder + P2 hidden
  row/col) — `rename_sheet`/`move_sheet`/`merged_cells`/`merge_cells`/
  `unmerge_cells`/`hidden_rows`/`hidden_columns`/`set_row_hidden`/
  `set_column_hidden`, plus PyO3-layer bound-check pins for `sort_range`/
  `merge_cells` (no openpyxl comparison needed for those).

This is the one place in the repo with a genuine, disclosed new dependency:
`openpyxl` is a **test-only oracle**, never a runtime dependency of the
shipped `elixcee` package (`pyproject.toml` declares none, and this stays
true — nothing here is imported by anything under `src/`). It exists here for
the same reason `compat/differential/`'s JS harness exists against the `xlsx`
npm package: independent agreement with a second, differently-implemented
reader is a stronger signal than internal self-consistency alone.

## One-time setup

```
pip install openpyxl
maturin develop --release --features python
```

## Run

```
python3 compat/differential-python/bulk_range_check.py
python3 compat/differential-python/sheet_ops_check.py
```

Plain stdlib `unittest`, no test runner required. Exits non-zero on any
failure.

## Not wired into CI

No existing CI job runs `maturin develop` or any Python code today (see
`.github/workflows/ci.yml`). Adding one is a separate infrastructure decision
from this feature round — this script is a documented manual/local step for
now, matching this round's own regression checklist.
