# Large reader calibration

Date: 2026-09-03 (Asia/Tokyo)

This local, offline measurement covers normal successful reads and the
documented XML safety boundary. It is not a cross-machine performance claim.

## Protocol

- Host: macOS/Darwin, `aarch64-apple-darwin`
- Rust: `rustc 1.97.0`
- Binary: `target/debug/elixcee`
- Fixture: deterministic XLSX, Stored ZIP, one numeric cell per row
- Sizes: 10,000, 100,000, 150,000, and 1,000,000 rows
- Repetitions: 3 per size, sequential
- Correctness: JSON `ok` plus exact observed row count for successful reads;
  deterministic error code/message for the rejected boundary

Raw observations are in `reader-large-2026-09-03.json`.

## Summary

| rows | fixture size | median wall | max wall | max RSS | outcome |
|---:|---:|---:|---:|---:|---|
| 10,000 | 477,297 B | 525.7 ms | 533.3 ms | 6.7 MiB | 3/3 correct |
| 100,000 | 5,067,300 B | 4,319.3 ms | 4,752.0 ms | 25.0 MiB | 3/3 correct |
| 150,000 | 7,667,300 B | 7,786.5 ms | 9,372.7 ms | 29.3 MiB | 3/3 correct |
| 1,000,000 | 53,667,303 B | 12,440.8 ms | 22,404.3 ms | 70.2 MiB | 3/3 safely rejected |

The 1,000,000-row fixture exceeds `XML_MAX_ELEMENTS` because each row has
multiple XML elements; it is therefore a rejection-boundary case, not a normal
success claim. The successful 150,000-row case provides the current large-input
baseline under the active limits.

These measurements calibrate only this fixture, host, debug binary, and
toolchain. Cross-platform values, release-binary values, repeated-process
resource reclamation, and broader dense/sparse/style/formula populations remain
separate measurement gates.
