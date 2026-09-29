# V8 cancellation hot-path regression check — 2026-09-12

The save writer was changed so it borrows the optional cancellation flag once
when the writer is constructed, rather than looking up the VM cancellation
field at every XML write. The cancellation checks remain active whenever a
host-owned flag is attached.

The check compares the current working tree with a clean `v1.0.12` tag
worktree using `compat/benchmarks/equal_workbook_speed.py`: three alternating
rounds, five measured iterations per round, two excluded warmups per arm, and
value/formula verification after every output.

| Fixture | clean v1.0.12 p50 / p95 (ms) | current p50 / p95 (ms) | p50 delta | p95 delta |
|---|---:|---:|---:|---:|
| 17 cells | 4.810 / 5.962 | 4.796 / 5.306 | −0.3% | −11.0% |
| 1,000×10 | 13.278 / 17.998 | 13.613 / 16.878 | +2.5% | −6.2% |
| 10,000×10 | 102.454 / 109.283 | 103.212 / 113.055 | +0.7% | +3.5% |

All outputs passed value/formula verification. The result is below the 10%
case p95 investigation threshold in this sample, but it is not a formal release
gate: the run has only three rounds and does not establish a cross-platform
result. The earlier v1.0.11 comparison was performed before this hot-path
change and must not be used as the final post-fix result.

## Clean v1.0.12 confirmation after the hot-path change

The final boundary check used clean `v1.0.12` as the before arm and the same
working-tree binary as the after arm for 10 alternating rounds with five
iterations per round (50 measured samples per fixture). The comparison used
the same durable save and independent value/formula verification protocol.

| Fixture | clean v1.0.12 p50 / p95 (ms) | current p50 / p95 (ms) | p50 delta | p95 delta |
|---|---:|---:|---:|---:|
| 17 cells | 4.884 / 5.613 | 4.606 / 5.387 | −5.7% | −4.0% |
| 1,000×10 | 12.924 / 14.383 | 13.119 / 14.048 | +1.5% | −2.3% |
| 10,000×10 | 92.504 / 99.159 | 92.238 / 103.993 | −0.3% | +4.9% |

All outputs passed verification. The largest p95 delta was +4.9%, below the
10% investigation threshold. This is a local release-gate result only; it does
not establish a three-OS result, a VBA speedup, or Excel/xlflow superiority.
