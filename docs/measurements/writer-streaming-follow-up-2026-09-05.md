# Writer streaming follow-up

## Scope

This is a local release-profile smoke/soak measurement after worksheet XML sink
output and passthrough payload clone reduction. It verifies saved output by
loading it again and checking the mutation. It is not a large-workbook RSS
benchmark or an Excel-oracle comparison.

## Run

```text
cargo run --release --bin measure_reader_write_inprocess --offline -- tests/fixtures/e2e/source.xlsx 3
```

Host: macOS arm64, 2026-09-05.

Observed wall times were `12.044 ms`, `6.327 ms`, and `7.711 ms`; all three
iterations round-tripped 17 cells successfully. The process allocator counters
reported `16,777,216` allocated bytes after each iteration, with in-use bytes
from `11,840` to `12,160` for this small fixture. These values are fixture- and
allocator-specific and do not establish large-workbook performance.

## G5 bounded writer follow-up (2026-09-06)

The reproducible harness `scripts/measure-stream-writer-memory.py` measured the
installed development wheel in one child process per case. At 100,000, 250,000,
and 1,000,000 rows × 3 columns, peak RSS was 18.33, 18.45, and 18.42 MiB;
wall time was 325, 752, and 2,901 ms; output sizes were 1.43, 3.68, and 14.93 MiB.
Each ZIP passed CRC validation and contained the final row marker. This is evidence
for the append-only path on macOS arm64 / CPython 3.13 only. It is not a constant-memory
proof, a normal-VM writer result, an RSS cap, or a Linux/Windows/Excel result.

The same harness supports `--mode normal`. A one-column run measured 10,000 / 25,000 /
50,000 rows at peak RSS 22.01 / 26.06 / 33.63 MiB and wall time 37 / 75 / 133 ms;
all outputs passed ZIP, final-row, and reload checks. Larger cases currently hit the
Reader's 1,000,000-element safety limit before the save phase, so this is partial
normal-writer evidence rather than a full scaling result.

For a new in-memory VM, `--mode normal-fresh` measured 100,000 / 250,000 /
1,000,000 rows × 3 columns at peak RSS 109.25 / 200.42 / 748.54 MiB and wall
time 182 / 436 / 1,802 ms. ZIP and final-row checks passed. This intentionally
shows the current VM's all-cell retention cost; it is not a constant-memory claim.

The harness now accepts `--repetitions` and emits per-case p50/p95. A three-repeat
append run measured 100,000 / 250,000 rows × 3 columns at wall p50/p95
305.7/306.4 ms and 718.4/732.9 ms, with peak RSS p50/p95 18.58/18.66 MiB and
18.58/18.62 MiB. A two-repeat `normal-fresh` run measured wall p50/p95
185.3/189.8 ms and 448.6/452.3 ms, with peak RSS p50/p95 109.20/109.23 MiB
and 200.50/200.50 MiB. These are macOS arm64 / CPython 3.13 local results.

The harness also polls the isolated case directory at 10 ms intervals and reports
peak temporary-directory bytes; output files are removed after validation so
repetitions cannot accumulate disk usage.

Input-shape calibration is available through `--value-profile plain|escape|giant`:
ordinary text, XML-escape-heavy text, and a deterministic 1 MiB single string.
Small escape and giant append cases passed ZIP and final-row validation; these
profiles are calibration fixtures, not yet cross-OS or Excel evidence.
The harness now obtains peak RSS from `ru_maxrss` on macOS/Linux and the Win32
peak working set on Windows, so the same child-process measurement can be used
for the planned three-OS run.
Passing `--output FILE.json` additionally persists the schema-versioned report;
the manual CI matrix uploads one report per case and OS for later review before
promoting results into this document.
The checked-in `scripts/check-stream-writer-measurements.py` validates the schema,
percentile ordering, repetition counts, and ZIP/final-row success flags before
the artifact upload.
Its `--self-test` also covers an accepted report, schema corruption, percentile
inversion, and an unvalidated sample.

On 2026-09-06, three append repetitions on macOS arm64 / CPython 3.13 measured
the `escape` profile at 1,000 rows × 2 columns: peak RSS p50/p95 was 19.39/19.44 MiB,
peak temporary disk was 561.5 KiB, and output validation passed. The `giant` profile
at 10 rows × 2 columns measured peak RSS p50/p95 of 31.62/32.52 MiB and peak
temporary disk of 12.5 KiB; validation also passed. The giant value is highly
compressible, so its compressed output is small; the case is for input/allocator
calibration, not a constant-memory or cross-platform claim.

At 1,000,000 rows × 3 columns, three append repetitions produced wall p50/p95
2,718/2,935 ms and peak RSS p50/p95 18.55/18.59 MiB. Two normal-fresh
repetitions produced wall p50/p95 1,783/2,672 ms and peak RSS p50/p95
748.55/750.44 MiB. All repeated outputs passed ZIP and final-row checks.

With the 10 ms directory monitor enabled, the 1,000,000-row append run observed
14.24 MiB peak temporary-directory usage across three repetitions, and one
normal-fresh run observed 12.43 MiB. Both equal the final output size; monitored
wall time is not comparable with the unmonitored timing above.
