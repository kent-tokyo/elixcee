# ClosedXML / elixcee / openpyxl workbook benchmark

Local measurement date: 2026-09-06. **Provisional, non-dedicated-host results.**
The same workload and file durability were used for each library, but scheduling
and host load were not isolated. No samples were removed.

This snapshot predates the [subsequent elixcee optimization](workbook-speedup120-2026-09-06.md).
ClosedXML was not rerun for that later change; these tables remain historical,
not a comparison of the newly optimized engine against ClosedXML.

## First run

Load → edit → durable save → reload; milliseconds, 30 samples per library/case:

| Input | elixcee p50 / p95 | ClosedXML p50 / p95 | openpyxl p50 / p95 | ClosedXML p50 ÷ elixcee p50 |
| --- | ---: | ---: | ---: | ---: |
| Small (17 populated cells) | 4.990 / 6.354 | 9.109 / 12.108 | 9.284 / 10.757 | 1.83× |
| Numeric 1,000 × 10 | 37.268 / 59.736 | 78.982 / 83.065 | 101.782 / 208.151 | 2.12× |
| Numeric 10,000 × 10 | 322.158 / 1,006.202 | 733.655 / 1,730.006 | 1,017.429 / 1,650.054 | 2.28× |

elixcee's pooled p50 and p95 were lower than ClosedXML's in all three cases.
This is **not** a per-round or universal superiority claim: in the third round
of the largest case, elixcee's median was 884.892 ms versus ClosedXML's 750.501 ms.
That reversal and the wide p95 tails must not be hidden by quoting only ratios.
ClosedXML also did not uniformly beat openpyxl's p95.

## Independent confirmation

Same worker binaries/protocol, fresh processes and temporary files, another
30 samples per library/case. These are retained separately, not pooled or
selected as replacements for the first run.

| Input | elixcee p50 / p95 | ClosedXML p50 / p95 | openpyxl p50 / p95 | ClosedXML p50 ÷ elixcee p50 |
| --- | ---: | ---: | ---: | ---: |
| Small | 5.972 / 10.068 | 11.528 / 19.682 | 14.874 / 18.760 | 1.93× |
| Numeric 1,000 × 10 | 36.476 / 89.413 | 101.802 / 199.932 | 96.702 / 286.817 | 2.79× |
| Numeric 10,000 × 10 | 317.351 / 328.016 | 743.599 / 5,049.466 | 938.926 / 993.568 | 2.34× |

The pooled-median direction repeats on all three cases. The changing ratios,
ClosedXML/openpyxl median reversal on the medium case, and ClosedXML's very
wide tail on the largest case underline why this remains a **provisional local
comparison**, not a reliable tail-latency or cross-library ranking guarantee.
Raw confirmation: [second run](workbook-closedxml-confirm-2026-09-06.json).

## Save stage and output sizes (first run)

Save includes ZIP completion, file barrier, close, and atomic rename.

| Input | elixcee save p50 | ClosedXML save p50 | openpyxl save p50 | Output bytes: elixcee / ClosedXML / openpyxl |
| --- | ---: | ---: | ---: | ---: |
| Small | 4.523 ms | 5.925 ms | 5.944 ms | 5,720 / 6,502 / 4,786 |
| 1,000 × 10 | 19.755 ms | 32.187 ms | 32.619 ms | 36,846 / 39,285 / 36,462 |
| 10,000 × 10 | 151.339 ms | 310.962 ms | 238.239 ms | 306,676 / 334,211 / 322,967 |

Stage medians are independently calculated and do not necessarily sum to the
total median. Different output bytes are permitted; the required equivalence
is cell values and formula text, not the whole ZIP or its compression ratio.

## Versions and environment

- elixcee: manifest 1.0.1, **unreleased local working tree**, native Rust API.
  Not a published 1.0.1 result; no PyO3 overhead is measured.
- [ClosedXML 0.105.1](https://www.nuget.org/packages/ClosedXML/0.105.1), assembly
  `0.105.1+b4ebe47bd3ecebf8480dea7422188d16b35b9d9a`. Direct dependency pinned
  exactly; transitive versions and content hashes in `packages.lock.json`.
- .NET SDK 10.0.400, runtime .NET 10.0.11, ARM64, Release build.
- openpyxl 3.1.2, Python 3.13.6, lxml enabled; Rust 1.97.0, release build.
- Apple M4 / Mac16,12 / 16 GiB RAM / macOS 26.5.2 ARM64.
- Repository HEAD `3e68207eb6cdf798d1477f3a6ad788f14dd9ac87` plus existing
  uncommitted development changes. Source and executable hashes are in the JSON.
- The machine was not reserved for benchmarking. During confirmation, one
  `uptime` snapshot showed 1/5/15-minute load averages of 9.89 / 6.94 / 5.68.
  This is evidence of a busy environment, not proof of the cause of any specific
  sample. Unrelated work was neither stopped nor modified.

## Workload and fairness contract

All engines receive the exact same input file bytes within a run:

1. Load the single-sheet workbook.
2. Set A1 to integer 123 and B1 formula text to `=1+2`.
3. Save into a newly created temporary file beside the destination.
4. Finish the ZIP, drain user-space buffers, call file `F_FULLFSYNC`, close, and
   atomically rename to the destination.
5. Reload the output through the respective library.

The smallest fixture is `tests/fixtures/e2e/source.xlsx`. Numeric fixtures are
generated using openpyxl with values `row * 10 + col` (1-based). Their logical
data is fixed; generated ZIP timestamps may differ between independent runs.

The C# worker uses ClosedXML's normal `SaveAs(Stream)` path. Formula
recalculation is not requested; this follows its documented default
([formula save options](https://docs.closedxml.io/en/latest/api/index.html)).
Validation reads `CachedValue` and `FormulaA1` without triggering calculation.
openpyxl likewise preserves formula text rather than calculating results.
Any calculation the standard elixcee setter performs is included in its time.
This is not a formula-engine performance comparison.

On this Mac, the Rust standard writer uses `File::sync_all()`, Python explicitly
calls `fcntl(F_FULLFSYNC)`, and the C# worker calls the same native primitive.
No path uses `save_workbook_fast` or silently falls back to ordinary `fsync`.
[Apple documents why ordinary fsync and F_FULLFSYNC differ](https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man2/fsync.2.html).
No engine syncs the parent directory; durable filename publication through power
loss is not claimed. C# currently refuses non-macOS platforms.

## Timing, JIT, and validation

- Five excluded warmups per engine and fixture. Warm filesystem cache.
- Three rounds of ten samples. Engine batch order rotates each round so all
  engines occupy each batch position once. Not per-sample interleaving.
- The ClosedXML process survives all batches in a run. Tiered compilation is
  disabled (`DOTNET_TieredCompilation=0`) to use optimized JIT code from first
  compilation, instead of mixing optimization tiers. This is an explicit
  benchmark runtime setting, **not the default .NET configuration**.
- Timers are internal to workers, excluding process startup. First-use work is
  warmed by the five operations; any JIT/GC inside subsequent timed calls still
  counts. No forced collections or GC disabling. Process creation/end-to-end
  CLI latency and workbook disposal are outside the timer.
- No benchmark batches or builds were deliberately run concurrently. This does
  not imply the host had no other active processes.
- All untouched/changed values and formula text are checked. Rust and C# workers
  assert each round trip; Python asserts each iteration and independently reads
  the entire output after every engine batch. Any mismatch aborts the run.
- Validation retains preloaded original data and makes additional untimed
  allocations (including a C# original workbook per batch). Runtime memory/GC
  effects are not isolated; this benchmark cannot support RSS superiority claims.
- p50 uses the median, p95 the nearest-rank 95th percentile. Thirty samples do
  not establish confidence intervals or a stable worst-case latency.

## Reproduce

The [worker instructions](../../compat/benchmarks/closedxml/README.md) contain
build, locked-restore, self-test, and driver commands. No Excel installation is
needed for this workload. Python 3.11+ is required by the new driver.

Primary raw samples: [first run](workbook-closedxml-2026-09-06.json). Each JSON
records fixture hashes, stage timings, engine order, output sizes, versions,
source hashes, and C# executable/dependency hashes. Old openpyxl measurements
were not mixed into this table; all three libraries were rerun together.

The benchmark harness, not the elixcee engine, was changed for this comparison.
C# dependencies/build products are excluded from Cargo publication. No project
version bump, push, registry release, or external document publication occurred.

Validation completed: locked NuGet restore; Release build with zero warnings or
errors; real invalid-file-descriptor sync-error self-test; five Python tests
(durability selection/error handling, value/formula mismatch rejection, median
and p95 calculation, invalid timing rejection); all 630 warmup/measured round
trips across the two runs and all 72 independent post-batch output checks.
Both raw files' summaries and current source/executable hashes were rechecked;
`cargo package --list --allow-dirty --offline` confirmed benchmark files are
excluded, and `git diff --check` passed.

Not measured: cold start, .NET default-tiering comparison, PyO3 overhead, peak
RSS, large text/style-heavy/multi-sheet workbooks, full OOXML equivalence,
Excel-oracle interoperability, macro execution, other versions, or other OSes.
Use a quiet/dedicated host and a wider fixture matrix before generalizing these
local results into a product-wide speed claim.
