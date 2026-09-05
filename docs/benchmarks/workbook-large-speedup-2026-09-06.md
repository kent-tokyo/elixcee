# Large-workbook optimization: current-tree before/after

Measured 2026-09-06. Manifest version remains **1.0.1, unreleased working tree**.
This is an incremental comparison against the implementation at the start of
this optimization, already containing the preceding approximately 2× change.
It is not a comparison against the published 1.0.1 release.

## Results: improvement confirmed, 1.2× target only partly reached

Final implementation, 20 interleaved pairs per case. All times are milliseconds
for load → edit → standard durable save → reload.

| Numeric input | Before p50 / p95 | After p50 / p95 | Ratio of p50s | Median paired ratio | 1.2× target |
| --- | ---: | ---: | ---: | ---: | --- |
| 100,000 cells, one sheet | 238.636 / 330.404 | 202.181 / 248.713 | 1.180× | 1.229× | Not reached |
| 400,000 cells, one sheet | 804.461 / 1,300.440 | 610.141 / 830.182 | 1.318× | 1.305× | Reached locally |
| 1,000,000 cells, four sheets | 2,090.607 / 2,877.940 | 1,748.715 / 2,479.820 | 1.196× | 1.272× | Not reached |

The primary ratio uses the two total medians. In particular, **1.196× is below
1.2×**, even though rounding to two decimals would display 1.20. The secondary
paired-ratio median reflects adjacent AB/BA comparisons, but is not substituted
for the predeclared target metric. All three total p95 values improved in this
run; the 1m reload-stage p95 worsened (623.930 → 735.065 ms). This is not a
guarantee of stable tail latency.

Output sizes were identical before/after: 306,676 / 1,229,807 / 3,077,072 bytes.
Every pair passed all-member decompressed ZIP equality, with independent
all-sheet/all-cell/formula validation as described below.

### Smaller and mixed-data cross-check

A separate 20-pair run used the preceding fixture generator/protocol (normal
openpyxl Workbook, not write-only), plus the existing mixed Unicode/text case.
This checks correctness and regressions; it is not pooled with the large run.

| Input | Before total p50 / p95 | After total p50 / p95 | Ratio of p50s |
| --- | ---: | ---: | ---: |
| Small, 17 cells | 5.977 / 6.925 | 5.990 / 6.529 | 0.998× |
| Numeric 10,000 cells | 26.516 / 27.463 | 22.721 / 23.757 | 1.167× |
| Numeric 100,000 cells | 203.837 / 272.494 | 170.389 / 278.032 | 1.196× |
| Mixed numeric/Unicode 10,000 cells | 29.963 / 32.209 | 25.489 / 26.642 | 1.176× |

Small-case median was effectively unchanged (0.2% slower), not an improvement.
The numeric-100k p95 worsened in this cross-check; that result is retained.
No case in this secondary run reaches the strict incremental 1.2× p50 target.

### Retained intermediate evidence

- [Final large-case raw samples](workbook-large-speedup-confirm-2026-09-06.json):
  primary results above, with final binary/source hashes and stage timings.
- [Final small/mixed raw samples](workbook-large-speedup-regression-2026-09-06.json):
  the separate regression cross-check.
- [Intermediate 20-pair results](workbook-large-speedup-2026-09-06.json): before
  attribute-vector reuse, p50 ratios 1.198× / 1.225× / 1.323×. Its 1m total p95
  worsened (4,517.832 → 4,784.346 ms). These are **not** final implementation results.
- [Initial four-pair exploration](workbook-large-speedup-exploratory-2026-09-06.json):
  earlier still, before declaration-prefix screening, 100k/400k only. Too few
  samples for confirmation; retained to distinguish exploration from final evidence.

Host load changed materially between the intermediate and final runs. The
final run's 1/5/15-minute load averages started at 8.73 / 6.54 / 6.36 and ended
at 4.91 / 5.43 / 5.90. An apparent difference between separate runs does not
isolate the benefit of attribute reuse. No outliers were discarded or runs pooled.

## Implementation

- Worksheet output sorts one flat vector of coordinates, replacing per-cell
  BTreeMap insertion and per-row vector allocation/sorting. Empty hidden rows,
  height/style-only rows and style-only cells retain their original XML shape.
- Cell references use a reusable 17-byte stack buffer instead of allocating
  both a column string and an A1 address for each cell. The complete u32 input
  range is preserved, including the writer's pre-existing treatment of zero.
- XML tag closing delimiters are scanned as bytes. Quotes and Unicode
  attribute values retain the same boundaries and malformed-input handling.
- Duplicate-attribute checks compare directly for at most eight attributes;
  larger lists retain the hash-set path. Count, depth and value-length budgets
  and duplicate rejection are unchanged.
- Forbidden DTD/entity screening searches for the common `<!` prefix before
  case-insensitive comparison, replacing two per-byte window scans. It still
  rejects those literals inside comments, CDATA and attributes, as before.
- The worksheet reader reuses cell-type and attribute-vector storage. Owned
  entity-expanded attribute values are cleared after each consumed event;
  names continue borrowing immutable XML bytes.

No dependency, version, ZIP compression setting, read/security limit or save
durability policy was changed. Standard `sync_all` (macOS `F_FULLFSYNC`), close
and same-directory atomic rename remain timed. The fast/no-sync API is not used.
Individual changes were not measured in isolation; attribute a speedup only to
the combined implementation, not a particular bullet above.

## Measurement contract

- Apple M4, 16 GiB, macOS ARM64; exact platform/tool versions, source and binary
  SHA-256 values are stored in the JSON. Default Cargo release profile, Rust API.
- Inputs are openpyxl 3.1.2 write-only generated numeric workbooks, ten columns,
  values `sheet_index * rows_per_sheet * 10 + row * 10 + column`. Both binaries
  receive the exact same fixture bytes. Fixture creation is outside timing.
- 100k cells = 10,000 rows × 10 columns on one sheet; 400k = 40,000 × 10 on
  one sheet; 1m = four sheets of 25,000 × 10. These are distinct shapes.
- **1m cells on one sheet is not claimed.** Numeric cells require at least two
  XML elements each. A single million-cell worksheet would exceed the existing
  1,000,000-element-per-XML-part limit; this limit was not relaxed for the benchmark.
- Each operation loads all sheets into a fresh VM, changes only the first
  sheet's A1 to integer 123 and B1 to formula text `=1+2`, saves all sheets, then
  reloads the output. No bulk formula-recalculation workload is represented.
- Three excluded warmups per engine/fixture, then 20 pairs alternating AB/BA.
  Each worker performs one timed operation. Process startup, original-file
  preload, assertions and destruction are excluded; OS caches are warm.
- Every pair checks identical ZIP member names and **byte-identical decompressed
  contents of every member**. Every ten pairs and at the end, independent
  streaming openpyxl validation checks every cell/formula and row/sheet count.
  Worker assertions also check all untouched sheets, values and formula text.
- p50 is the median; p95 uses nearest rank. All raw samples, including outliers,
  are retained. Stage medians need not sum to total median. The target uses
  before-total-p50 / after-total-p50 ≥ 1.2 (at least 16.7% less elapsed time).
- The host was not isolated. Background load changed during runs; do not combine
  absolute timings across runs, pool intermediate and final results, multiply
  older competitor ratios by these speedups, or claim stable tail latency.
- Peak RSS, cold-cache speed, Python binding speed, other OSes, formula-heavy
  large files and complex large styled workbooks are not measured here.
  openpyxl is a fixture generator/independent validator, not a timed competitor
  in this report. ClosedXML was not remeasured.

## Validation

- `cargo test --all-targets --offline --quiet`: passed, including 1,518 library
  tests, 52 XLSX round-trip tests, other integration tests and benchmark smoke tests.
- `cargo clippy --all-targets --all-features --offline -- -D warnings`: passed.
- `cargo fmt --all -- --check` and `git diff --check`: passed.
- Nine Python benchmark-harness tests: passed. Existing openpyxl deprecation/
  resource warnings in the legacy tests are not claimed resolved.
- `bash scripts/check-measurement-boundary.sh`: passed; benchmark code remains
  excluded from the distributable package.
- New regressions cover A1 boundary values through u32::MAX, a golden sparse
  worksheet with overlapping empty-row metadata and style-only cells, both sides
  of the duplicate-attribute threshold, Unicode/quoted delimiters, declaration
  screening parity, and attribute-buffer reuse without stale names/values.

## Reproduction and baseline identity

The baseline was built **before engine edits**, after extending the benchmark
worker's out-of-timer assertions to support multiple sheets. Both sides use that
same worker source and timed operations. Preserved local files are under
`target/benchmarks/large-speedup-2026-09-06/`; these are not release artifacts.

| Artifact | SHA-256 |
| --- | --- |
| Baseline worker (`before-multisheet`) | `86deaaf3ab974aa6e02b2dc7f640be3f843d94f7911f8d229c8ec8280f087c5b` |
| Final worker (`after`) | `37090997e2d15df15f1c9961f03afb81cef09ad01271847475820b7f7ed44bfc` |
| Baseline `src/lib.rs` | `341e9c3f4d951b815a0718a2c18efed1bfca3a1b5c2dae00bd919a3768722c63` |
| Baseline `src/reader.rs` | `b0abd33e29eaa04200ee7394aea950eb2202e28754bae43d36d162a7b304b4f8` |

The [scoped final engine patch](../../compat/benchmarks/large-speedup.patch)
includes only this optimization's two engine files and new Rust unit tests.
Its reverse-apply check against the final tree passed. The
[intermediate patch](../../compat/benchmarks/large-speedup-intermediate.patch)
excludes the subsequent attribute-vector reuse. Existing unrelated working-tree
changes were preserved. Neither patch alone reconstructs the pre-existing dirty
tree from a published tag; reproduce from the same complete source snapshot.
Reverse patches only in an isolated scratch copy, never over a working checkout.

With the preserved baseline available, from the repository root:

```sh
cargo build --example bench_workbook --release --offline
python3 compat/benchmarks/large_workbook_speed.py \
  --before target/benchmarks/large-speedup-2026-09-06/before-multisheet \
  --after target/release/examples/bench_workbook \
  --pairs 20 --output target/benchmarks/large-rerun.json
```

Requires Python with openpyxl and the existing benchmark helpers; .NET is not
required for this before/after run. The command writes new results instead of
overwriting the dated evidence.
