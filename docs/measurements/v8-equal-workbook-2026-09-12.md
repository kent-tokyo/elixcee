# V8 equal-workbook benchmark — 2026-09-12

## Protocol

This is a local, VBA-free comparison of elixcee and openpyxl. Each arm uses
the same generated XLSX input and performs:

`load → mutate A1/B1 → durable save → reload → value/formula verification`

The elixcee arm is the release `examples/bench_workbook` binary. The benchmark
uses three rounds, five measured iterations per round, and alternates the arm
order by round. Two initial iterations per arm are discarded as explicit
warmups. The output is independently checked with openpyxl after each run.

Command:

```sh
python3 -B compat/benchmarks/equal_workbook_speed.py \
  --after target/release/examples/bench_workbook \
  --rounds 3 \
  --iterations 5 \
  --output /private/tmp/elixcee-v8-equal-workbook-2026-09-12.json
```

Host: macOS 26.5.2 arm64, Python 3.13.6. The elixcee binary was built from
the current working tree as version 1.0.12. This is not an Excel/xlflow
measurement and does not include VBA execution.

## Results

| Fixture | elixcee p50 (ms) | openpyxl p50 (ms) | p50 ratio (openpyxl / elixcee) | elixcee p95 (ms) | openpyxl p95 (ms) |
|---|---:|---:|---:|---:|---:|
| 17 cells | 4.950 | 10.095 | 2.04× | 5.931 | 14.221 |
| 1,000×10 | 12.773 | 84.045 | 6.58× | 14.362 | 116.448 |
| 10,000×10 | 92.976 | 918.499 | 9.88× | 103.783 | 974.685 |

All measured outputs passed the value/formula verification. The comparison is
valid only for this protocol, host, fixture family, versions, and durability
boundary. It is not evidence of a general speedup, VBA speedup, Excel
compatibility, or xlflow superiority.

