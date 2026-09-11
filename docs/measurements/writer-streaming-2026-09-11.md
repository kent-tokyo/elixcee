# Writer streaming measurement — 2026-09-11

Scope: local macOS arm64, CPython 3.13, an isolated venv containing a wheel
built from the current source (`elixcee` 1.0.10), append mode, three columns,
plain values, two repetitions per case. Each child process output passed ZIP,
worksheet-shape, final-row, and streaming semantic-digest checks.

| rows | RSS p50 / p95 | child wall p50 / p95 | output | semantic |
|---:|---:|---:|---:|:---:|
| 100,000 | 19.88 / 19.97 MiB | 415 / 436 ms | 1.36 MiB | pass |
| 250,000 | 19.88 / 19.91 MiB | 1,019 / 1,071 ms | 3.43 MiB | pass |

The same wheel completed a one-million-row confirmation sample at 19.93 MiB
peak RSS and 4,498 ms child wall time. The output was 14.24 MiB and passed the
same semantic check. This is a single confirmation sample, not a p95 estimate.

The append writer's peak RSS stayed effectively flat across these two row
counts on this host. This is evidence for this bounded append path only; it
does not establish constant memory for the normal VM, Linux/Windows behavior,
Excel compatibility, or a complete G5 gate. The raw report was generated at
`/private/tmp/elixcee-stream-memory-20260911.json` and is intentionally kept
outside the repository because it contains host-specific temporary paths.

Command:

```text
python scripts/measure-stream-writer-memory.py --rows 100000 250000 \
  --columns 3 --mode append --value-profile plain --repetitions 2
```

## Three-repetition rerun

After reclaiming the local build cache, the same isolated wheel and fixture
were rerun three times per case. At 100,000 rows, RSS p50/p95 was 19.86/20.03
MiB and wall p50/p95 was 310/317 ms. At 250,000 rows, RSS p50/p95 was
19.90/20.00 MiB and wall p50/p95 was 719/720 ms. All six outputs passed ZIP,
worksheet-shape, final-row, and semantic-digest validation. This supersedes
neither the one-million-row confirmation nor the cross-platform gate; it is a
repeatability check on the same macOS host.
