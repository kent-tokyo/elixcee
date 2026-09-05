# Formula dirty-propagation calibration

## Scope

This is a local microbenchmark for 100- and 1,000-formula direct-reference chains. It
compares a warm dirty recalculation with a forced formula-plan rebuild and
asserts equality at representative cells on every iteration. It is not an
end-to-end workbook or Excel-oracle measurement.

## Run

```text
cargo run --bin measure_formula_dirty --offline -- 5
```

Host: macOS arm64, debug profile, 2026-09-05.

```json
{"iterations":5,"cases":[{"formula_count":100,"dirty_p50_ms":0.513458,"dirty_p95_ms":0.531791,"full_p50_ms":0.707042,"full_p95_ms":0.720750},{"formula_count":1000,"dirty_p50_ms":5.386458,"dirty_p95_ms":5.487166,"full_p50_ms":7.443292,"full_p95_ms":11.125958}],"manual_to_automatic_p50_ms":5.348375,"manual_to_automatic_p95_ms":7.180000,"cycle_p50_ms":0.020000,"cycle_p95_ms":0.038292,"peak_rss_bytes":7766016,"user_cpu_us":213934,"system_cpu_us":34942,"resource_stats_supported":true,"wall_ms":166.302}
```

The binary checks representative input, near-root, middle-chain, and final
formula values against the forced full-rescan path. It additionally measures
the Manual-to-Automatic transition on the 1,000-formula template and a small
two-node cycle. On Unix it also emits process peak RSS and user/system CPU
counters via `getrusage`. The full roadmap item remains open until a larger
controlled before/after comparison and release-profile matrix are recorded.

## Release-profile follow-up

Command:

```text
cargo run --release --bin measure_formula_dirty --offline -- 30
```

The 30-iteration release run produced the following JSON. Resource counters are
process totals/peaks for the matrix, not per-case allocations.

```json
{"iterations":30,"cases":[{"formula_count":100,"dirty_p50_ms":0.040625,"dirty_p95_ms":0.071667,"full_p50_ms":0.042709,"full_p95_ms":0.094750},{"formula_count":1000,"dirty_p50_ms":0.441417,"dirty_p95_ms":0.534917,"full_p50_ms":0.455500,"full_p95_ms":0.506125}],"manual_to_automatic_p50_ms":0.426042,"manual_to_automatic_p95_ms":0.457833,"cycle_p50_ms":0.001208,"cycle_p95_ms":0.002209,"peak_rss_bytes":7421952,"user_cpu_us":109484,"system_cpu_us":20847,"resource_stats_supported":true,"wall_ms":72.414}
```

This is controlled local evidence for the current implementation only; it does
not establish Excel compatibility or cross-platform performance.
