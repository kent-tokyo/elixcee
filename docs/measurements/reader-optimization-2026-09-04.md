# Reader optimization calibration — 2026-09-04

This report records a local before/after calibration for the 1.0.1 reader optimization.
It is not a cross-platform benchmark or a comparison with Microsoft Excel or another
library.

## Conditions

- Host: macOS arm64
- Build: Cargo `release`
- Fixture: dense XLSX, 1 worksheet, 400,000 materialized cells
- Driver: `src/bin/measure_reader_vm_load.rs`
- Repetitions: 3 sequential reader-to-VM loads per binary
- Baseline: commit `3f62378` before the ownership/XML allocation optimizations
- Candidate: 1.0.1 working tree after ownership transfer, borrowed XML events, and fused
  worksheet validation/parsing

Both binaries reported one worksheet and 400,000 cells on every repetition.

## Result

| Metric | Baseline | 1.0.1 candidate | Change |
|---|---:|---:|---:|
| Median wall time per load | 1,386.496 ms | 650.549 ms | 53.1% lower; 2.13x throughput-equivalent speedup |
| Total user CPU time, 3 loads | 4.07 s | 1.92 s | 52.8% lower |
| Retired instructions, 3 loads | 46,950,617,712 | 19,766,613,706 | 57.9% lower |
| Maximum resident set size | 99,713,024 bytes | 75,972,608 bytes | 23.8% lower |

The optimization preserves XML DTD/control-character/structure/resource checks and
shared-string-index validation. Focused malicious-input tests assert that the fused
worksheet path returns the same rejection messages as the standalone validator.

## Reproduction boundary

Build `src/bin/measure_reader_vm_load.rs` from each revision and run:

```bash
cargo build --release --bin measure_reader_vm_load
/usr/bin/time -l target/release/measure_reader_vm_load FIXTURE.xlsx 3
```

Scheduler load can affect wall time. CPU instructions and correctness counts are included
to make the local result less dependent on wall-clock noise. Linux, Windows, other
allocators, compressed producer files, and independent-oracle comparisons remain open.
