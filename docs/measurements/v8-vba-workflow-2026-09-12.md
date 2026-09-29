# V8 local VBA workflow measurement — 2026-09-12

## Scope

This is a local elixcee-only measurement. It covers the CLI process boundary:

`read → VBA → calculate → save → process exit`

It is not an Excel, xlflow, VM-only, or VBA-free comparison. The external
oracle remains `not_measured`.

Command:

```sh
python3 -B scripts/benchmark-vba-workflows.py \
  --binary target/release/elixcee \
  --repetitions 3 \
  --warmup 0 \
  --measure-rss \
  --output /private/tmp/elixcee-vba-workflow-benchmark-rss.json
```

Host: macOS 26.5.2, arm64, Python 3.13.6. Binary SHA-256:
`2fb4362c7214d404d788f0de64cb697bb620c11e1d10c7db0fc9d10475e014cb`.

## Results

| Case | wall p50 (ms) | wall p95 (ms) | RSS p50 (MiB) | RSS p95 (MiB) | output stable |
|---|---:|---:|---:|---:|---|
| headless-transfer-and-recalculate | 43.72 | 47.30 | 4.72 | 4.72 | yes |
| byref-alias-and-optional | 41.68 | 45.30 | 4.80 | 4.83 | yes |
| headless-table-sort-filter-find | 37.15 | 44.23 | 4.96 | 4.98 | yes |
| headless-object-aliases | 43.02 | 46.24 | 4.90 | 4.90 | yes |
| headless-range-value2-transfer | 35.70 | 42.74 | 4.98 | 5.14 | yes |

RSS uses `ps` polling at 1 ms intervals. The polling overhead is included in
the wall samples from this command, so these wall values must not be compared
with the non-RSS run. This is a baseline for the same host and binary, not a
cross-library speed claim.

## Warmup-excluded wall-time follow-up

After adding explicit warmup handling to the runner, five warmup runs followed
by five measured runs were collected without RSS polling. The following values
exclude warmups:

| Case | wall p50 (ms) | wall p95 (ms) | output stable |
|---|---:|---:|---|
| headless-transfer-and-recalculate | 37.78 | 38.74 | yes |
| byref-alias-and-optional | 43.52 | 46.11 | yes |
| headless-table-sort-filter-find | 39.29 | 43.49 | yes |
| headless-object-aliases | 44.51 | 44.68 | yes |
| headless-range-value2-transfer | 43.84 | 47.61 | yes |

Command:

```sh
python3 -B scripts/benchmark-vba-workflows.py \
  --binary target/release/elixcee \
  --warmup 5 \
  --repetitions 5 \
  --output /private/tmp/elixcee-vba-workflow-benchmark-warm.json
```

These measurements are still local elixcee-only results. They do not establish
an old-version regression result or an Excel/xlflow speedup.

The report also records the manifest hash, source/workbook inputs, output
hashes, raw samples, and `external_oracle_status` so later measurements can be
compared without losing the measurement boundary.

## VBA-free local baselines

The existing release-only measurement binaries were also run against
`tests/fixtures/e2e/source.xlsx` for five iterations. These are in-process
measurements and do not invoke VBA:

| Boundary | p50 (ms) | p95 (ms) | Notes |
|---|---:|---:|---|
| workbook file → VM load | 0.711 | 1.311 | 17 cells, one sheet |
| read → mutate → durable save → reload | 5.881 | 14.281 | `measure_reader_write_inprocess`, `fast=false` |

The load observations were `1.311, 0.631, 0.545, 0.760, 0.711` ms. The full
observations were `14.281, 5.881, 5.679, 7.152, 5.567` ms. The first
iteration is retained rather than silently discarded; a larger warmup and
alternating protocol is required before using this as a release gate.
