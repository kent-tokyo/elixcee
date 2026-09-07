# VM hot-path optimization calibration — 2026-09-05

This report records local release-mode microbenchmarks for algorithmic
hot-path changes. It is not a cross-platform benchmark or a comparison with
Microsoft Excel or another library.

## Conditions

- Host: macOS arm64
- Rust: `rustc 1.97.0 (2d8144b78 2026-07-07)`
- Build: Cargo `bench` profile
- Harness: Criterion, 10 samples, 2 second measurement target
- Working version: 1.0.1 (version unchanged)

## Results

| Case | Before/reference median estimate | Candidate median estimate | Change |
|---|---:|---:|---:|
| 20,000-element array retained while 500 VBA loop iterations execute | 34.621 ms | 213.53 µs | 99.37% lower; 162.1x throughput-equivalent |
| 5,000 one-cell row appends | 45.691 ms | 1.9604 ms | 95.71% lower; 23.3x throughput-equivalent |
| Recalculate 50 formulas containing a 100,000-cell dependency range in an unselected IF branch | 1.8938 s | 237.25 µs | 99.988% lower; 7,982x throughput-equivalent |
| Recalculate 10,000 unchanged simple formulas | 13.045 ms | 10.079 ms | 22.7% lower; 1.29x throughput-equivalent |

The variable-budget baseline was measured immediately before replacing the
per-statement recursive scan. The candidate validates public/native values once
at execution start and validates values created by VBA at mutation boundaries.
Instruction, call-depth, string, array, Collection, and cell-count limits remain
enabled.

The append reference benchmark implements the former behavior in the same
binary: scan every non-Empty cell for the maximum row before every append. The
candidate caches the next row by sheet, invalidates it on general cell mutation,
and advances it after a non-Empty append. An all-Empty row deliberately does not
advance the cache, preserving the existing used-range contract.

The formula-dependency baseline expanded every referenced range coordinate into
a `HashSet`, even when only a few coordinates contained formulas. The candidate
indexes formula positions by row and column and emits only actual formula-to-formula
edges. Formula evaluation and circular-reference best-effort behavior are unchanged.

The formula-AST candidate caches successful and failed parses by sheet,
coordinate, and exact source text. Direct/native formula edits therefore refresh
the entry on the next recalculation, and deleted formula cells are removed from
the cache. The Criterion comparison reported a statistically significant
improvement (`p < 0.05`).

The subsequent local implementation slice also made `StreamWriter` write rows
directly into a temporary ZIP worksheet entry, removed procedure-body and
`DoLoop` body clones, batched object reachability collection at 32 mutations
with a procedure-boundary flush, and reserves capacity for bulk cell writes.
Those changes are correctness-tested here but require a separate before/after
benchmark matrix before claiming a numeric speedup.

## Rejected experiment

A borrowed range-value view for `SUMIF(S)`/`COUNTIF(S)` was also measured on a
100,000-row, two-criteria `SUMIFS`. The median estimate moved from 81.151 ms to
74.170 ms, but the result was not statistically significant (`p = 0.61`) and
the confidence interval included a regression. That implementation and its
benchmark were removed rather than presenting the result as an optimization.

## Reproduction

```bash
cargo bench --bench vm_bench variable_budget_large_array_20000x500 -- --sample-size 10 --measurement-time 2
cargo bench --bench vm_bench append_row_ -- --sample-size 10 --measurement-time 2
cargo bench --bench vm_bench recalculate_50_large_range_dependencies -- --sample-size 10 --measurement-time 2
cargo bench --bench vm_bench recalculate_10000_formula_parses -- --sample-size 10 --measurement-time 2
```

These are focused microbenchmarks. End-to-end `.xlsm` execution, Python-call
overhead, other CPU architectures, and different allocators remain unmeasured.

## Formula dirty-propagation run

On the same local macOS arm64 host, a release-mode Criterion run measured a
1,000-formula direct-reference chain. A single input change reached the whole
dependent chain in a median `1.3142 ms`; a no-op recalculation after the plan
was warm measured `1.4823 ms`; and the comparison case that invalidates and
rebuilds the formula plan measured `6.1787 ms`. The dirty case was therefore
The same benchmark matrix was rerun after adding range filtering and the
manual/cycle cases. The current medians were `3.0450 ms` for the single-input
chain, `948.55 us` for the no-op path, `2.5084 ms` for the structure-rebuild
comparison, and `1.2906 ms` when all 1,000 independent inputs changed. The
chain dirty case was therefore not faster than the full-rebuild case in this
run; this is a useful negative result showing that closure bookkeeping still
needs optimization when nearly the whole chain is affected. These are local
microbenchmarks with outliers, not end-to-end workbook claims. p95 and a
controlled before/after baseline remain open work.

A follow-up implementation then replaced the dirty-position full-plan scan with
a coordinate index and bit vector. Its post-change benchmark rerun is still
pending, so the values above remain the conservative recorded baseline.

## Formula dirty-propagation post-optimization run

After the coordinate-index and bit-vector change, the same release-mode
Criterion matrix was rerun on the same macOS arm64 host. The medians were
`739.78 us` for a single-input change through a 1,000-formula chain,
`586.72 us` for the warm no-op path, `1.6698 ms` for the structure-rebuild
comparison, and `559.90 us` when all 1,000 independent inputs changed. The
dirty chain was 55.7% lower than the structure-rebuild case in this run
(2.26x throughput-equivalent). This is workload-local evidence; p95,
controlled before/after runs, and end-to-end workbook measurements remain
open.

## Follow-up benchmark run

The same host rerun on 2026-09-05 used 10 samples and a one-second measurement
target. Current candidate medians were `vba_loop_10000` 2.6272 ms,
`set_cell_direct_10000` 679.68 us, and `append_row_cached_5000` 950.17 us.
The in-binary former-behavior reference for the append case was 20.585 ms,
so the cache path was 95.4% lower in this run. Criterion reported a significant
difference for both append variants (`p < 0.05`). These values are host-local
and do not establish an end-to-end workbook speedup.

The procedure-focused follow-up measured 1,000 user-Function calls at a
541.97 us median after immutable procedure definitions were changed to shared
`Arc` handles. A pre-change measurement for this newly introduced case is not
available, so this is a candidate baseline rather than a claimed improvement.
The unrelated 10,000-cell loop rerun was 3.0957 ms and was slower than the
previous 2.6272 ms observation; this difference is not attributed to the
procedure change and is not used as a regression claim without a controlled
repeat.

## Cell rectangle tile-cache run

The 128x128 dense `read_rect` benchmark was run on the same local host with
Criterion's 20-sample configuration after warming the candidate cache. The
canonical HashMap lookup reference measured 651.67 us median, while the tiled
path measured 437.28 us median, a 32.9% lower median (1.49x throughput-equivalent).
This is a repeated-read microbenchmark: the first cache fill, writes between
reads, sparse sheets, and end-to-end XLSX/Python overhead are separate cases.
The implementation therefore keeps the optimization limited to large rectangles
and invalidates the affected sheet on supported mutation accessors.

The bounded-cache pressure benchmark (`read_rect_tile_cache_lru_257`) measured
17.056 ms median for creating and reading 257 adjacent 32x32 tiles with a
10-sample Criterion run. This is a cache-pressure regression baseline, not an
end-to-end speed claim; it includes cache fill and eviction work and is intended
to detect accidental unbounded growth or incorrect eviction behavior.

## Sparse/Dense tile density calibration

The same-host release benchmark measured warmed repeated reads of one 32x32
tile at five populated-cell counts, using 10 Criterion samples: 1 cell,
47.301 us median; 32 cells, 49.030 us; 128 cells, 50.648 us; 129 cells,
33.056 us; and 1,024 cells, 39.537 us (the last case had a mild high outlier).
The 129-cell transition was materially faster on this run, so the implementation
keeps `DENSE_TILE_CELL_THRESHOLD = 128`. These are local microbenchmark results;
allocator, CPU, and larger workbook effects remain unmeasured.

The construction-versus-reuse run measured 1-cell tiles at 106.83 us for the
first build and 50.732 us warmed, while 1,024-cell tiles measured 276.13 us
for the first build and 36.980 us warmed. The overlay is therefore intended
for repeated reads; one-shot reads retain the uncached path where applicable.

For 100 single-cell writes into one already-cached 32x32 tile, followed by a
read after each write, the incremental update path measured 7.619 ms median.
The invalidation-and-rebuild reference measured 18.812 ms median, so the
incremental path reduced this local workload by 59.5% (2.47x throughput-
equivalent). This comparison includes VM setup and is limited to the repeated
single-tile workload; broad writes and end-to-end workbook writes remain
separate measurements.

For the next multi-tile slice, a 128x128 cached write uses adaptive targeted
invalidation rather than rebuilding all affected tiles during the write. The
local sample measured 43.692 ms median for the `write_rect` path versus 33.843
ms for the direct-mutation invalidation reference. This is not a broad-write
speedup claim: the benefit is preserving unaffected cached tiles and deferring
work until a tile is read, while the measured end-to-end write/read case was
29.1% slower on this host. The threshold is fixed at 512 written cells pending
further workload-specific calibration.

Threshold calibration on the same host used a cached 32x32 read and these
write shapes: 16x16 (256 cells), 32x16 (512), 33x16 (528), and 32x32 (1,024).
The adaptive/reference median pairs were 2.481/3.736 ms, 2.001/1.725 ms,
0.871/0.946 ms, and 1.342/0.711 ms respectively. The mixed result is
shape- and sample-sensitive: incremental maintenance helps at 256 cells,
while targeted invalidation is competitive or better at the larger shapes,
but the 1,024-cell case does not support a universal threshold change. The
512-cell cutoff therefore remains fixed; this is calibration evidence, not a
general performance claim.

## Current append-cache baseline

The existing append benchmark was rerun on 2026-09-07 with Criterion's
10-sample, two-second configuration on the current 1.0.4 tree. The reference
rescan path measured 106.88 ms median for 5,000 rows, while the cached path
measured 94.98 ms median. This is an approximately 11.1% local reduction for
this microbenchmark. The date, build, and Criterion settings differ from the
earlier follow-up above, so these values are a current baseline rather than a
replacement for the earlier result or evidence of the roadmap's 1.2x target.

## Current dirty-propagation baseline

The same release-mode Criterion configuration was rerun on 2026-09-07 for the
1,000-formula controlled matrix. Median times were 1.2825 ms for a single-input
dirty chain, 1.2363 ms for a warm no-op, 1.3090 ms for the structure-rebuild
comparison, and 1.1570 ms when all 1,000 independent inputs changed. The
single-input case was not materially faster than the rebuild case in this run,
so no speedup claim is made; closure bookkeeping remains a P0 optimization
candidate when the affected set is large. These are local microbenchmarks and
do not establish end-to-end workbook performance.

## Dirty-closure queue-index trial

On 2026-09-07, the dirty closure queue was changed to carry formula-plan
indices for discovered formula nodes, avoiding a coordinate-to-index lookup on
each visit. The same 10-sample, two-second Criterion run measured a median of
1.2538 ms for the single-input chain and 1.3223 ms for the structure-rebuild
comparison. Against the immediately preceding 1.2825 ms / 1.3090 ms baseline,
this is an approximately 2.2% local improvement for the chain, while the
rebuild case is unchanged within measurement noise. Range/cycle regression
tests passed. The independent-input case was not rerun in this slice, and no
end-to-end or universal 1.2x claim is made.
