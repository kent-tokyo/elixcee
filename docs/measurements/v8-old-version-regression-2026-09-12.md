# V8 old-version regression check — 2026-09-12

## Protocol

The tagged `v1.0.11` and current `v1.0.12` release benchmark examples were
run with `compat/benchmarks/equal_workbook_speed.py`. Each arm used the same
generated fixture and performed:

`load → mutate A1/B1 → durable save → reload → value/formula verification`

The run used 10 alternating rounds and five measured iterations per round.
The harness also performs two excluded warmup iterations per arm and checks
the output after every run. This is a local macOS arm64 measurement; it is not
an Excel/xlflow or VBA comparison.

## Confirmation results

| Fixture | v1.0.11 p50 / p95 (ms) | current p50 / p95 (ms) | current p50 delta | current p95 delta |
|---|---:|---:|---:|---:|
| 17 cells | 4.514 / 5.866 | 4.986 / 7.181 | +10.5% | +22.4% |
| 1,000×10 | 15.328 / 29.372 | 14.647 / 20.277 | −4.4% | −31.0% |
| 10,000×10 | 86.494 / 101.887 | 91.214 / 103.727 | +5.5% | +1.8% |

All outputs passed the value/formula verification. The result is mixed, and
the 17-cell p95 exceeds the roadmap's 10% investigation threshold. Therefore
the V8 old-version regression gate is **not passed**. No general speedup or
release-performance claim should be derived from this run.

The current binary was built from the working tree at version 1.0.12; the
comparison binary was built from the immutable `v1.0.11` tag in an isolated
temporary worktree. Binary identity and raw samples are retained in the JSON
report used for this measurement.

