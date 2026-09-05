# 1.2× workbook-speed target: current-tree before/after

Measured 2026-09-06; manifest version remains **1.0.1, unreleased working tree**.
The baseline is the exact elixcee executable used in the preceding ClosedXML
confirmation, verified by its SHA-256, not an older published release.

## Outcome: partial target completion

The numeric 10,000-cell and 100,000-cell cases exceed the requested **1.2×**
speed target. The additional mixed text/numeric case also exceeds it. The tiny
17-cell workbook does **not** reach 1.2×; the target is not marked achieved for
all workloads. No durability barrier was removed to improve that result.

Independent confirmation: 40 paired samples per engine/case. Times are for
load → edit → standard durable save → reload, in milliseconds:

| Input | Before p50 / p95 | After p50 / p95 | Ratio of p50s | Median paired ratio | 1.2× target |
| --- | ---: | ---: | ---: | ---: | --- |
| Small, 17 populated cells | 5.893 / 7.166 | 5.382 / 8.957 | 1.09× | 1.14× | Not reached |
| Numeric 1,000 × 10 | 65.900 / 78.876 | 32.937 / 45.135 | 2.00× | 1.95× | Reached locally |
| Numeric 10,000 × 10 | 685.854 / 1,228.679 | 341.296 / 541.033 | 2.01× | 2.12× | Reached locally |
| Mixed numeric/Unicode 1,000 × 10 | 77.088 / 87.637 | 38.038 / 44.536 | 2.03× | 2.00× | Reached locally |

A 1.2× speedup requires time to fall to at most 1/1.2 of baseline (a 16.7%
reduction). The three larger cases reduced total p50 by about 50%. Small-case
p95 was worse in this run; neither its median nor tail supports a 1.2× claim.

The initial 20-pair exploration yielded p50 ratios of 1.00× / 1.95× / 2.15× on
small / numeric-10k / numeric-100k. It predates final regression tests and minor
borrow cleanup; it is preserved separately, not pooled or used to select a
favorable result. No outlier samples were discarded.

### Previous ten-iteration batch protocol cross-check

The existing comparator was also rerun with before/after/openpyxl, ten timed
iterations per worker process, three rotating-order rounds (30 samples each).
This checks that the result is not specific to starting a worker for every pair.

| Input | Before p50 / p95 | After p50 / p95 | p50 speedup | openpyxl p50 / p95 |
| --- | ---: | ---: | ---: | ---: |
| Small | 5.209 / 6.259 ms | 4.719 / 6.091 ms | 1.10× | 11.318 / 13.043 ms |
| Numeric 1,000 × 10 | 78.534 / 110.031 ms | 32.380 / 43.044 ms | 2.43× | 226.645 / 444.352 ms |
| Numeric 10,000 × 10 | 539.294 / 629.231 ms | 243.469 / 271.659 ms | 2.22× | 1,573.384 / 2,388.347 ms |

The larger cases again exceed 1.2×; the small case again does not. Host load
varied, so do not combine absolute times across protocols/runs or multiply old
ClosedXML ratios by these speedups. The one-operation paired results are the
primary before/after evidence; this batch run is an additional cross-check.

## Changes

1. `find_next_open_tag` now borrows tag names from the XML source rather than
   allocating a `String` for every opening tag it scans.
2. Opaque-element extraction checks whether the literal local-name bytes occur
   at all before walking all cell tags. Absence is conclusive because XML names
   are literal, not entity-expanded. A hit is **only a prefilter**: the existing
   exact local-name and element-boundary checks still run, including prefixes.
3. Shared-string collection filters string-valued cells **before** sorting and
   clones only newly interned strings. Numeric cells no longer participate in
   that sort or the subsequent hash lookup. Coordinate order, deduplication,
   whitespace/escaping, and treatment of error cells remain unchanged.

These XML helpers serve both reader and writer metadata paths, so the benefit
is not limited to save time. The measurements are for the combined change; they
do not isolate a percentage contribution from each individual optimization.
ZIP compression configuration, bounded write buffers, read/security budgets,
the standard file sync, and atomic rename are unchanged.

The [scoped engine patch](../../compat/benchmarks/speedup120.patch) records only
this turn's changes to `src/lib.rs` and `src/reader.rs`, including scanner tests;
the mixed shared-string round-trip test is additionally in `tests/xlsx_roundtrip.rs`.
Existing unrelated working-tree edits were preserved.

## Correctness and checks

- Every pair was checked for **identical ZIP member names and byte-identical
  decompressed contents of every member**, not just A1 or cached formula values.
  Compression container metadata need not match; measured output sizes did match
  (5,720 / 36,846 / 306,676 / 53,673 bytes respectively).
- Each Rust operation validates changed cells, untouched values, and untouched
  formula strings. An independent openpyxl read verifies all cells/formula text
  after every ten pairs and at the end of each fixture.
- New scanner regressions cover borrowed Unicode/prefixed names, quoted `>`,
  repeated fragments, and false-positive names in text/attributes/longer tags.
- A new round-trip test checks deterministic shared-string indices/deduplication
  among 1,000 numeric cells, reverse insertion order, Unicode, XML escaping,
  leading/trailing whitespace, and all reloaded cell values.
- `cargo test --all-targets --offline --quiet`: passed, including 1,512 library
  tests and 52 XLSX round-trip tests, other integrations and benchmark smoke tests.
- `cargo clippy --all-targets --all-features --offline -- -D warnings`: passed.
- Seven Python harness tests passed, including formula/member loss and duplicate
  ZIP-member rejection in the new equality checker.

## Measurement contract and limits

- Same Apple M4 / macOS 26.5.2 ARM64 host, Rust 1.97.0 release profile. Python
  3.13.6 / openpyxl 3.1.2 generate/check fixtures, outside the Rust timers.
- Each before/after pair uses the exact same input bytes. A1 becomes integer
  123 and B1 becomes formula text `=1+2`. Both call the standard save API, with
  `sync_all` (`F_FULLFSYNC` on this Mac), file close and same-directory rename.
  Neither syncs the containing directory.
- Three excluded warmups per engine/fixture; then alternating AB / BA order on
  every pair to balance position and reduce slow host-load drift. All samples
  are retained. This host is **not dedicated or scheduling-isolated**.
- Each measurement invokes a fresh Rust worker for one timed operation. Worker
  startup, its untimed original-data preload, and post-reload assertions are
  excluded. This differs from a ten-iteration in-process batch; allocator/cache
  state and process lifetime can affect results even when startup is excluded.
- p50 is the median; p95 is nearest rank. `median_paired_speedup` is the median
  of corresponding before/after sample ratios; it is not the ratio of medians.
  Empirical percentiles are not confidence intervals.
- The mixed fixture uses numeric columns plus repeated and unique Unicode
  strings. Repeated strings deliberately include `sheetViews`, so the negative
  prefilter cannot simply bypass every metadata lookup on that input.
- Generated fixtures have fixed logical contents but may have different ZIP
  timestamps/hashes between runs. Input and executable hashes are in each JSON.
- Small after-save p50 is 4.587 ms out of 5.382 ms total. Fixed save costs remain
  substantial. This does not isolate the exact sync-only cost; no sync-only
  timing or universal lower bound is claimed. Small-workbook 1.2× remains open.
- Peak RSS, cold start, PyO3 overhead, other OSes, Excel-oracle behavior, macro
  execution, and a broad style-heavy matrix are not measured by this experiment.

## Reproduce / audit

Raw records: [20-pair exploration](workbook-speedup120-initial-2026-09-06.json),
[40-pair confirmation and mixed case](workbook-speedup120-confirm-2026-09-06.json).
The [ten-iteration batch cross-check](workbook-speedup120-batched-2026-09-06.json)
also records the separately rerun openpyxl results.

The before binary SHA-256 is
`f7cc1386dc589e2f13bbd833d4f3abe72b55fd1a9d8dfa70f8b6cac92769237a`;
it matches `elixcee_binary_sha256` in the preceding ClosedXML confirmation.
The local before binary/source backups are retained under the ignored
`target/benchmarks/speedup120-2026-09-06/` directory, not shipped in a package.

```sh
cargo build --release --example bench_workbook --offline
python3 compat/benchmarks/paired_speedup.py \
  --before target/benchmarks/speedup120-2026-09-06/before \
  --after target/release/examples/bench_workbook \
  --pairs 40 --mixed --output /tmp/elixcee-speedup120.json
```

For another checkout, preserve a genuinely pre-change executable before applying
the patch. To reconstruct the pre-change engine from the final working tree,
reverse-apply the scoped patch **in a separate scratch checkout**, then build
the same helper/profile/toolchain. The patch is against this development tree,
not a promise that it applies to the published 1.0.1 source. Never reverse user
changes in the active checkout merely to run a benchmark. Baseline source files
and artifact hashes permit checking what was actually compared locally.

Version remains fixed; no commit, push, tag, release, or external publication
was performed. These are local implementation and measurement artifacts.
