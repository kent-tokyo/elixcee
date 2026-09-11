# Density and multi-sheet reader calibration

Date: 2026-09-03 (Asia/Tokyo)

This local, offline measurement extends the large-input matrix with higher
column density and multiple worksheets. It is not a cross-machine performance
claim.

## Protocol

- Host: macOS/Darwin, `aarch64-apple-darwin`
- Rust: `rustc 1.97.0`
- Binary: `target/debug/elixcee`
- Fixture: deterministic XLSX, Stored ZIP, four style records
- Repetitions: 3 per profile, sequential
- Correctness: ZIP/XML row, cell, formula, non-default-style, and sheet counts
  are validated before execution; successful snapshot JSON must contain the
  expected sheet count and exact total cell count

Raw observations are in `reader-density-sheets-2026-09-03.json`.

## Summary

| profile | shape | sheets | cells | median wall | max RSS | outcome |
|---|---:|---:|---:|---:|---:|---|
| dense | 100,000 x 4 | 1 | 400,000 | 16.8 s | 73.3 MiB | 3/3 correct |
| dense-5col | 80,000 x 5 | 1 | 400,000 | 14.3 s | 79.5 MiB | 3/3 correct |
| sparse | 1,000,000 rows / stride 20 / 2 cols | 1 | 100,000 | 4.1 s | 24.7 MiB | 3/3 correct |
| style-heavy | 100,000 x 2 | 1 | 200,000 | 8.4 s | 42.8 MiB | 3/3 correct |
| formula-heavy | 80,000 x 3 | 1 | 240,000 | 10.1 s | 67.1 MiB | 3/3 correct |
| mixed | 80,000 x 4 | 1 | 320,000 | 16.3 s | 78.9 MiB | 3/3 correct |
| multi-sheet-4 | 25,000 x 4 x 4 sheets | 4 | 400,000 | 15.1 s | 67.6 MiB | 3/3 correct |

The expanded density and four-sheet cases remained within the active reader
limits. These values calibrate only the stated fixture shapes, host, debug
binary, and toolchain; they are not general throughput or Excel-compatibility
claims. Formula values are cached fixture values observed by `snapshot`, not a
formula-recalculation parity result. Release/cross-platform runs, repeated
process reclamation, and Excel/LibreOffice oracle comparison remain separate
gates.
