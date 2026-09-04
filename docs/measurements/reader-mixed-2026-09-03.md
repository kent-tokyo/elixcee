# Mixed large-input reader calibration

Date: 2026-09-03 (Asia/Tokyo)

This local, offline measurement covers large XLSX inputs with different cell
density and content populations. It is not a cross-machine performance claim.

## Protocol

- Host: macOS/Darwin, `aarch64-apple-darwin`
- Rust: `rustc 1.97.0`
- Binary: `target/debug/elixcee`
- Fixture: deterministic XLSX, Stored ZIP, one sheet, four style records
- Repetitions: 3 per profile, sequential
- Correctness: ZIP/XML row, cell, formula, and non-default-style counts are
  validated before execution; successful snapshot JSON must contain the exact
  expected cell count and `ok: true`

Raw observations are in `reader-mixed-2026-09-03.json`.

## Summary

| profile | configured shape | cells | formulas | styled cells | fixture size | median wall | max RSS | outcome |
|---|---:|---:|---:|---:|---:|---:|---:|---|
| dense | 100,000 x 4 | 400,000 | 0 | 0 | 16.6 MB | 14.9 s | 71.5 MiB | 3/3 correct |
| sparse | 1,000,000 rows / stride 20 / 2 cols | 100,000 | 0 | 0 | 4.9 MB | 5.0 s | 24.8 MiB | 3/3 correct |
| style-heavy | 100,000 x 2 | 200,000 | 0 | 150,000 | 9.3 MB | 10.3 s | 41.3 MiB | 3/3 correct |
| formula-heavy | 80,000 x 3 | 240,000 | 80,000 | 0 | 11.9 MB | 11.8 s | 70.9 MiB | 3/3 correct |
| mixed | 80,000 x 4 | 320,000 | 80,000 | 240,000 | 16.9 MB | 17.8 s | 89.6 MiB | 3/3 correct |

The mixed case is the highest observed resource profile in this matrix. These
values calibrate only the stated fixture shapes, host, debug binary, and
toolchain; they are not a general throughput or Excel-compatibility claim.
Formula values are cached fixture values observed by `snapshot`; this test
does not claim formula recalculation parity. Cross-platform/release runs,
repeated-process reclamation, and Excel/LibreOffice oracle comparison remain
separate gates.
