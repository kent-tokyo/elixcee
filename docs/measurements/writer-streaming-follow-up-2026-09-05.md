# Writer streaming follow-up

## Scope

This is a local release-profile smoke/soak measurement after worksheet XML sink
output and passthrough payload clone reduction. It verifies saved output by
loading it again and checking the mutation. It is not a large-workbook RSS
benchmark or an Excel-oracle comparison.

## Run

```text
cargo run --release --bin measure_reader_write_inprocess --offline -- tests/fixtures/e2e/source.xlsx 3
```

Host: macOS arm64, 2026-09-05.

Observed wall times were `12.044 ms`, `6.327 ms`, and `7.711 ms`; all three
iterations round-tripped 17 cells successfully. The process allocator counters
reported `16,777,216` allocated bytes after each iteration, with in-use bytes
from `11,840` to `12,160` for this small fixture. These values are fixture- and
allocator-specific and do not establish large-workbook performance.
