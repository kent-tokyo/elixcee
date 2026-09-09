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

## Follow-up BUILD boundary

The relationship graph and carry-over paths now parse UTF-8 XML through borrowed
views instead of cloning each relationship payload into a temporary `String`.
The targeted offline relationship test and the 54-test `xlsx_roundtrip`
integration suite passed after this change. A subsequent low-disk offline
workspace run also passed 1,681 library tests and every integration/benchmark
target. This is a small allocation reduction only; no speed or RSS improvement
is claimed without a fresh benchmark.

The follow-up also moves the writer-owned `xl/_rels/workbook.xml.rels` bytes
out of the raw map after carry-over analysis instead of cloning them. The
54-test `xlsx_roundtrip` suite and warnings-denied clippy passed after this
change. Its speed and RSS effect remain unmeasured.

The subsequent carry-over lookup changes the surviving-part membership check
from a linear `Vec` scan to a `HashSet` lookup. This preserves the emitted
relationships and connectivity decision while reducing the analysis cost as
the passthrough-part count grows. `cargo check --lib --offline` passed with
debuginfo disabled in a dedicated temporary target. The focused test binary
could not be linked in the available disk space, so no speed or RSS gain is
claimed and the full test gate remains open.

Reproduction for the workspace gate:

```bash
TMPDIR=/private/tmp CARGO_INCREMENTAL=0 RUSTFLAGS='-C debuginfo=0' \
  cargo test --offline --workspace --all-targets --quiet
```

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

## Relationship-ownership follow-up (negative measurement)

The ownership-only follow-up was built from the actual pre-change commit
`8dac236` and the current candidate `fc73748`, each with a separate clean
`CARGO_TARGET_DIR`, `CARGO_INCREMENTAL=0`, release profile, and
`-C debuginfo=0`. Both `bench_workbook` binaries produced the same SHA-256:
`8e2871b216756b9ae2cb0e5960c98e45bad995221e9933e007c0cfb5a3edfe95`.

Because the optimized executable was byte-identical, no before/after timing
or RSS number was collected. The change remains a source-level allocation
reduction verified by round-trip tests and strict Clippy; it is not evidence
of a runtime speedup. The result is intentionally retained as a negative
measurement so later benchmark claims do not reuse an invalid comparison.

The later same-day paired run in
[g5-paired-v1.0.5-2026-09-10.md](g5-paired-v1.0.5-2026-09-10.md) used the
available `v1.0.5` tag binary as a reproducible baseline and completed the
durable-save, ZIP-equality, and streaming output checks for 100k, 400k, and 1M
cells. It does not retroactively reconstruct the older historical baseline or
close the RSS, 20-pair confirmation, and cross-platform gates.
