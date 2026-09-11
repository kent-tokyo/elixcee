# G5 paired large-workbook measurement

Date: 2026-09-10 (Asia/Tokyo)

This is a same-host before/after measurement of the current checkout against
the `v1.0.5` tag. It uses the repository's paired benchmark, which generates
the same fixture, applies the same A1/B1 edit, performs durable save and
reload, compares every decompressed ZIP member byte-for-byte, and validates
all cells/formulas with streaming openpyxl.

## Conditions

- Host: macOS arm64
- Before: `v1.0.5` tag, release `bench_workbook` binary
- After: current checkout `release/0.21.0`, release `bench_workbook` binary
- Build: offline release, `CARGO_INCREMENTAL=0`, debuginfo disabled
- Pairs: 20 per fixture, alternating execution order, 3 warmups per binary
- Target: at least 1.2x p50 total-time speedup
- Fixture: numeric 10-column workbook; 1M cells are split across four sheets

## Result

| Fixture | Cells | Before p50 total | After p50 total | p50 speedup | p95 before → after | Target |
|---|---:|---:|---:|---:|---:|---|
| 100k | 100,000 | 131.417 ms | 99.543 ms | 1.320x | 148.100 → 111.209 ms | met |
| 400k | 400,000 | 509.718 ms | 390.973 ms | 1.304x | 582.767 → 439.834 ms | met |
| 1m | 1,000,000 | 1,327.555 ms | 1,018.183 ms | 1.304x | 1,611.941 → 1,173.627 ms | met |

The 1.2x target was met for all three measured sizes. The paired harness
reported successful ZIP equality and independent output validation on every
verification checkpoint. This is a macOS benchmark result, not a claim about
Excel compatibility, other operating systems, or other libraries.

## Reproduction

```sh
python3 compat/benchmarks/large_workbook_speed.py \
  --before /path/to/v1.0.5/bench_workbook \
  --after target/release/examples/bench_workbook \
  --output /tmp/elixcee-g5-paired.json \
  --pairs 20 --cases 100k 400k 1m
```

The recorded run used the repository's 20-pair confirmation policy. RSS and
three-OS resource validation remain separate G5 gates.
