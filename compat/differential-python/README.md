# Python differential tests

Manual/local checks against openpyxl using the same worksheet fixtures.
openpyxl is a test-only oracle, not an elixcee runtime dependency.
API signatures are in [elixcee.pyi](../../elixcee.pyi); these scripts do not establish
complete openpyxl API or Microsoft Excel compatibility.

| Script | Scope |
|---|---|
| bulk_range_check.py | Bulk range read/write, append, row/column iteration, dimensions |
| sheet_ops_check.py | Sheet operations, metadata, merges, styles, validation, filtering, and related regression cases |

See each script's test methods for exact coverage. Some fixtures are openpyxl-authored,
others come from the Excel fixture set; do not label every case “Excel-authored.”
Some tests verify elixcee bounds/errors without an independent comparison.

## Setup and run

From the repository root, with an activated Python environment:

```sh
pip install maturin openpyxl
maturin develop --release --features python
python3 compat/differential-python/bulk_range_check.py
python3 compat/differential-python/sheet_ops_check.py
```

These are stdlib unittest scripts and exit nonzero on failure.
For performance comparisons use the separate
[benchmark protocol](../../docs/benchmarks/README.md), including its fixed versions.

## CI boundary

The [CI workflow](../../.github/workflows/ci.yml) runs Python-based checks and Rust
Python-feature compilation, but does not currently run these two differential scripts.
A compile check is not an installed-extension integration test.
