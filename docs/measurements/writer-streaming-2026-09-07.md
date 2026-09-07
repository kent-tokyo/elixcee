# Writer streaming measurement — 2026-09-07

Scope: local macOS arm64, CPython 3.13, release wheel built from the current
1.0.4 tree, isolated child processes, three columns, plain text profile, two
repetitions. ZIP integrity, worksheet shape, and final-row checks passed for
every completed case.

| mode | rows | RSS p50 / p95 | wall p50 / p95 | output |
|---|---:|---:|---:|---:|
| append | 100,000 | 19.39 / 19.83 MiB | 353 / 367 ms | 1.36 MiB |
| append | 250,000 | 19.38 / 19.46 MiB | 717 / 756 ms | 3.51 MiB |

The corrected normal-VM run uses 4,096-row `set_range` batches to avoid
measuring one undo snapshot per input row. It is a storage/save observation,
not an append API throughput comparison:

| mode | rows | RSS p50 / p95 | wall p50 / p95 | output |
|---|---:|---:|---:|---:|
| normal-fresh | 100,000 | 612.58 / 623.22 MiB | 340 / 393 ms | 1.24 MiB |
| normal-fresh | 250,000 | 924.62 / 1,920.73 MiB | 1,068 / 1,178 ms | 3.11 MiB |

Both repetitions for both row counts passed ZIP, worksheet-shape, final-row,
and output validation. The 250,000-row RSS samples were 924.62 MiB and
1,920.73 MiB, so the high p95 is retained rather than normalized away. This
run does not support a constant-memory claim; 1,000,000 rows, three-OS
coverage, and Excel-oracle comparison remain open.

After changing transaction history to retain one pre-transaction snapshot and
making the normal-VM harness wrap all batches in one transaction, the same
measurement was rerun with the rebuilt 1.0.4 release wheel:

| mode | rows | RSS p50 / p95 | wall p50 / p95 | output |
|---|---:|---:|---:|---:|
| normal-fresh, one transaction | 100,000 | 116.92 / 118.30 MiB | 173 / 202 ms | 1.24 MiB |
| normal-fresh, one transaction | 250,000 | 203.56 / 203.61 MiB | 418 / 420 ms | 3.11 MiB |

All four samples passed the same output checks. This is a large-workbook
undo-history improvement, not proof that the normal VM is constant-memory;
the full 1,000,000-row and three-OS matrix remains open.

The initial normal-VM 100,000-row case, using one `append_row` call per row,
was stopped after more than two minutes at 100% child CPU. It is intentionally
recorded as unmeasured, not as a failure. The append data is one macOS run and
does not establish three-OS support or constant memory for the normal VM.
The harness now uses 4,096-row `set_range` batches for a separate normal-VM
storage/save measurement; its results must not be compared directly with
append API timings.
