# G5 large hot-path benchmark rerun

Date: 2026-09-10 (Asia/Tokyo)

This is a local release-mode Criterion measurement on macOS arm64. It covers
large in-memory VM hot paths only; it is not an end-to-end XLSX benchmark, a
cross-library comparison, or evidence for Excel compatibility.

## Conditions

- Branch: `release/0.21.0`
- Candidate: `e3d8e27` and its documentation-only descendants
- Harness: `cargo bench --bench vm_bench`, 10 samples, 2 seconds
- Build: Cargo `bench` profile
- Host: macOS arm64

## Results

| Case | Median estimate | 95% interval | Interpretation |
|---|---:|---:|---|
| `append_row_cached_5000` | 93.799 ms | 93.676–93.961 ms | cached next-row path |
| `append_row_reference_rescan_5000` | 105.81 ms | 105.68–105.94 ms | same-binary reference scan |
| `recalculate_dirty_chain_1000_single_input` | 1.0157 ms | 995.23 µs–1.0254 ms | tracked dirty closure |
| `recalculate_full_chain_1000_structure_rebuild` | 1.3437 ms | 1.3195–1.3564 ms | plan rebuild comparison |

The cached append path was approximately 1.128x faster than the reference
rescan path (`105.81 / 93.799`). The dirty-chain path was approximately 1.323x
faster than the structure-rebuild comparison (`1.3437 / 1.0157`). These ratios
are within-process microbenchmark observations, not general workbook speed
claims.

## Reproduction

```bash
cargo bench --bench vm_bench append_row_cached_5000 -- --sample-size 10 --measurement-time 2
cargo bench --bench vm_bench append_row_reference_rescan_5000 -- --sample-size 10 --measurement-time 2
cargo bench --bench vm_bench recalculate_dirty_chain_1000_single_input -- --sample-size 10 --measurement-time 2
cargo bench --bench vm_bench recalculate_full_chain_1000_structure_rebuild -- --sample-size 10 --measurement-time 2
```

The run completed successfully. A prior `--list | head` attempt produced only
a Criterion Broken pipe because `head` closed stdout; it was not used as a
measurement. The broader G5 item remains open for fixed 100k/400k/1M workbook
inputs, p50/p95 repetitions, RSS, durable save, and output verification.

## Baseline reconstruction boundary

The historical large-workbook comparison could not be rerun in this pass.
The optimization baseline is the parent of `2a0940b` (`3e68207`, v1.0.1),
but that checkout has no `bench_workbook` example, so
`cargo build --example bench_workbook --release --offline` fails with
`error: no example target named bench_workbook`. The saved
`compat/benchmarks/large-speedup.patch` also does not reverse-apply to the
current tree; it stops at `src/lib.rs:5926` because later changes have moved
that context. The temporary baseline worktree was removed after the check.

This is a reconstruction limitation, not a new performance result. The
existing 100k/400k/1M reports remain historical scoped evidence, while a new
same-input before/after comparison still requires reproducible baseline
artifacts plus RSS, durable-save, and output-verification capture.
