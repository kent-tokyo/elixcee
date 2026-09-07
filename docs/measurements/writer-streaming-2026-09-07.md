# Writer streaming measurement — 2026-09-07

Scope: local macOS arm64, CPython 3.13, release wheel built from the current
1.0.4 tree, isolated child processes, three columns, plain text profile, two
repetitions. ZIP integrity, worksheet shape, and final-row checks passed for
every completed case.

| mode | rows | RSS p50 / p95 | wall p50 / p95 | output |
|---|---:|---:|---:|---:|
| append | 100,000 | 19.39 / 19.83 MiB | 353 / 367 ms | 1.36 MiB |
| append | 250,000 | 19.38 / 19.46 MiB | 717 / 756 ms | 3.51 MiB |

The normal-VM 100,000-row case was stopped after more than two minutes at
100% child CPU and is intentionally recorded as unmeasured, not as a failure
or a constant-memory result. The append data is one macOS run and does not
establish three-OS support or constant memory for the normal VM.
