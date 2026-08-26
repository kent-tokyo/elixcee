# Python differential tests

Compares elixcee's Python-native bulk worksheet range/row API (`get_range`/
`set_range`/`append_row`/`iter_rows`/`max_row`/`max_column`/
`calculate_dimension`, R1 — see `docs/openpyxl-gap-audit.md`) against
`openpyxl`'s own read of the same fixtures.

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
```

Plain stdlib `unittest`, no test runner required. Exits non-zero on any
failure.

## Not wired into CI

No existing CI job runs `maturin develop` or any Python code today (see
`.github/workflows/ci.yml`). Adding one is a separate infrastructure decision
from this feature round — this script is a documented manual/local step for
now, matching this round's own regression checklist.
