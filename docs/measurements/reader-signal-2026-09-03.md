# Reader signal cancellation measurement

Date: 2026-09-03 (Asia/Tokyo)

This is a local, offline cancellation-safety measurement. It is not a general
throughput or memory benchmark.

## Fixture and protocol

- Host: macOS/Darwin, `aarch64-apple-darwin`
- Rust: `rustc 1.97.0 (2d8144b78 2026-07-07)`
- Command: `cargo test --offline --test cli_snapshot snapshot_exits_with_cancellation_after_sigint -- --nocapture`
- Fixture: synthetic XLSX with 1,000,000 worksheet rows, stored (uncompressed)
- Action: wait for the child CLI's signal-handler-ready marker, send real `SIGINT`,
  and assert JSON `READER_CANCELED` with no partial workbook
- Repetitions: 3, run concurrently only to shorten wall-clock collection time

## Observations

| run | wall time | peak RSS | result | signals received |
|---:|---:|---:|---|---:|
| 1 | 0.56 s | 102,400,000 bytes (97.7 MiB) | `READER_CANCELED` | 2 |
| 2 | 0.47 s | 102,432,768 bytes (97.7 MiB) | `READER_CANCELED` | 2 |
| 3 | 0.46 s | 102,416,384 bytes (97.7 MiB) | `READER_CANCELED` | 2 |

The three runs establish a reproducible macOS SIGINT cancellation smoke result
for this fixture. They do not calibrate default budgets across hardware, do not
measure uncanceled throughput, and do not validate Windows or Linux signal
delivery; those remain separate gates.
