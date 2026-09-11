# XLSX speed: equal-durability comparison (2026-09-06)

A later [three-library comparison including ClosedXML](workbook-closedxml-2026-09-06.md)
reruns the same cases with all three engines. Its samples are separate from this report.

## Result

On this host and these three fixtures, the optimized **standard** elixcee save
path beats openpyxl 3.1.2 for load → edit → durable save → reload. Both libraries
perform the same edits and the same file durability barrier and atomic rename.
`save_workbook_fast` is **not** used.

This is an **unreleased working-tree** measurement with manifest version 1.0.1,
not a benchmark of the published 1.0.1 package. It compares native Rust elixcee
API calls against Python openpyxl API calls; it does not measure PyO3 overhead.

Initial paired run, 30 samples per library per fixture (milliseconds):

| Input | elixcee before p50 | elixcee after p50 / p95 | openpyxl p50 / p95 | openpyxl p50 ÷ after p50 |
| --- | ---: | ---: | ---: | ---: |
| Small, 17 populated cells | 5.785 | 5.032 / 6.515 | 9.934 / 11.325 | 1.97× |
| Numeric, 1,000 × 10 | 53.416 | 35.867 / 39.887 | 84.266 / 116.484 | 2.35× |
| Numeric, 10,000 × 10 | 493.380 | 315.983 / 324.793 | 959.211 / 2,104.293 | 3.04× |

Each of the three round medians, not just the pooled median, favors optimized
elixcee on all fixtures. No outliers were removed. These are empirical latency
results, not confidence intervals or a universal speed guarantee.

Independent confirmation, 40 samples per library per fixture, with two-engine
batch order alternated across four rounds:

| Input | elixcee p50 / p95 | openpyxl p50 / p95 | p50 ratio |
| --- | ---: | ---: | ---: |
| Small | 5.021 / 6.603 ms | 9.538 / 11.922 ms | 1.90× |
| 1,000 × 10 | 35.591 / 36.868 ms | 84.496 / 115.417 ms | 2.37× |
| 10,000 × 10 | 317.622 / 328.594 ms | 947.311 / 1,060.125 ms | 2.98× |

All four round medians again favor elixcee on each fixture. Both runs are kept
separately, not selectively pooled. See the [40-sample raw confirmation](workbook-equal-durable-confirm-2026-09-06.json).

## Actual implementation improvement

The writer batches tiny XML fragments into a 64 KiB buffer **before compression**
and compressed output into a separate 64 KiB buffer before filesystem writes.
ZIP data descriptors avoid backward header-patching seeks that flush the output
buffer prematurely. Both buffers are drained with error propagation before the
existing `sync_all` and atomic rename. No compression-level, security-budget,
source-part preservation, or standard durability setting was weakened.

This adds bounded buffering (at most two 64 KiB buffers at once), not a full-sheet
XML allocation. Existing VM/source-part memory costs remain; peak RSS is not
measured here. The shared-string output uses the same bounded buffering.

| Input | Save before p50 | Save after p50 | openpyxl save p50 | Output bytes before / after / openpyxl |
| --- | ---: | ---: | ---: | ---: |
| Small | 5.214 ms | 4.482 ms | 6.114 ms | 5,567 / 5,720 / 4,786 |
| 1,000 × 10 | 36.283 ms | 18.564 ms | 27.680 ms | 44,620 / 36,846 / 36,462 |
| 10,000 × 10 | 324.167 ms | 147.492 ms | 231.727 ms | 387,758 / 306,676 / 322,967 |

Save latency decreased by about 49% and 55% on the two numeric fixtures. The
before/after comparison uses identical Rust operations and durability; this
improvement is distinct from correcting the openpyxl comparison below. The
corrected baseline already beats openpyxl in total latency on these fixtures.
Stage medians are computed separately and need not sum to the total median.

## Correction to earlier comparisons

On this installed Rust 1.97.0 toolchain, `std::fs::File::sync_all()` calls
`fcntl(fd, F_FULLFSYNC)` on macOS, verified in the local Rust standard-library
source (`library/std/src/sys/fs/unix.rs`, `File::fsync`). Python `os.fsync()` alone
does not match that barrier. The previous small-fixture comparison therefore
did not establish that openpyxl was faster under equal durability.

The new Python harness also calls `F_FULLFSYNC` on macOS, with **no fallback** to
a weaker barrier. On other platforms it uses `os.fsync`; those platforms have
not been measured in this report. Neither library syncs the containing directory,
so this is not a claim of power-loss-proof filename publication.

The old [smoke report](workbook-speed-2026-09-06.md) and exploratory
`workbook-equal-2026-09-06.json` / `workbook-equal-stream-2026-09-06.json` used the
old Python barrier. They are historical diagnostics only and are excluded from
the comparison tables above. The old opt-in no-sync result is also excluded.

## Reproduction and measurement contract

- Host: Apple M4, Mac16,12, 16 GiB RAM; macOS 26.5.2 arm64.
- Rust 1.97.0, standard release profile; Python 3.13.6, openpyxl 3.1.2 with lxml
  enabled (confirmed by the independent run's environment metadata).
- Working-tree HEAD: `3e68207eb6cdf798d1477f3a6ad788f14dd9ac87`, with pre-existing
  uncommitted development changes. It is not an exact release identifier.
- Input: existing `tests/fixtures/e2e/source.xlsx`, plus generated single-sheet
  numeric workbooks whose cell `(row, col)` is `row * 10 + col` (1-based).
  Fixture bytes are shared across engines within each run. ZIP timestamps can
  change generated-file hashes between independent runs; logical data is fixed.
- Operation: load, set A1 to integer 123, set B1 formula text to `=1+2`, save to
  an adjacent temporary file, finish/close ZIP, sync the file, close the file,
  atomically replace the destination, then reload the output.
- Two excluded warmups per engine/fixture. Warm filesystem cache; same temporary
  directory and storage. No intentionally parallel benchmark or build processes.
- Three rounds of ten iterations, with before/after/openpyxl batch order rotated
  each round. This balances batch position; it is not per-sample interleaving.
- Timers are inside each process and exclude startup and post-reload assertions.
  Python GC remains enabled. Disposing loaded workbook objects is outside the
  timer; incidental GC during timed calls is included. p50 is the median; p95
  is the nearest-rank 95th percentile.
- The Python side verifies all cell values and formula strings each iteration;
  the Rust worker checks changed cells and untouched values each iteration. An
  independent openpyxl read checks the entire output after every Rust batch.
  The final worker additionally checks untouched formula strings each iteration.
- Outputs need logical value/formula equality, not byte equality. This benchmark
  does not assert full OOXML/style equivalence or evaluate formula results in
  openpyxl (which does not calculate formulas). Any elixcee calculation performed
  by its standard setters is included in its time.

Run the current implementation against openpyxl:

```sh
cargo build --release --example bench_workbook --offline
python3 compat/benchmarks/equal_workbook_speed.py \
  --after target/release/examples/bench_workbook \
  --rounds 4 --iterations 10 --output /tmp/elixcee-workbook-comparison.json
```

Install openpyxl 3.1.2 in the chosen Python environment before running. The helper
is checkout-only and excluded from the Cargo registry package. For an actual
before/after experiment, build and preserve the pre-edit worker executable, then
pass it as `--before PATH`; do not substitute a published release binary or
relabel the optimized binary as the baseline. The original pre-buffering worker
was saved locally before making this turn's writer changes. Its hash is recorded
in the raw results; it is not a distributed benchmark binary.

Raw paired data: [30-sample run](workbook-equal-durable-2026-09-06.json). Raw files
record per-stage samples, batch order, fixture hashes, output sizes, versions,
and worker hashes. The independent confirmation also records current source
hashes and whether openpyxl uses lxml.

## Validation and boundaries

- `cargo test --all-targets --offline --quiet`: passed (1,509 library tests,
  51 XLSX round-trip tests, other integration tests and benchmark smoke checks).
- New regression: worksheet XML, shared strings, and compressed archive each
  exceed 64 KiB; Unicode/XML escaping, trailing buffered bytes, every cell,
  independent calamine reading, and repeated atomic replacement pass.
- `cargo clippy --all-targets --all-features --offline -- -D warnings` and
  `cargo fmt --all -- --check`: passed.
- Three Python harness tests verify the macOS barrier selection, sync-error
  propagation, and rejection of untouched-cell/formula loss.

Not established: Excel application/oracle interoperability of the new output,
cross-platform speed, cold-cache latency, Python-binding overhead, memory or
energy superiority, complex/style-heavy workbooks, macro execution comparisons,
or superiority to openpyxl versions other than the pinned 3.1.2. Small workloads
are particularly sensitive to filesystem synchronization and host activity.
No release, upload, or external publication was performed.
