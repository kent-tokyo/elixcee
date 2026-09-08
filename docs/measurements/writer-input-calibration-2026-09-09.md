# Writer input-memory calibration — 2026-09-09

Scope: local macOS arm64, CPython 3.13, the release wheel built from the
current 1.0.4 tree, isolated child processes, and three repetitions per case.
The reported peak RSS includes the Python input object, Rust allocator, XML
encoding, and ZIP writer. It is not an allocator breakdown and does not prove
constant memory.

| profile | rows | columns | RSS p50 / p95 | wall p50 / p95 | output | validation |
|---|---:|---:|---:|---:|---:|---|
| XML escape-heavy | 1,000 | 2 | 19.87 / 20.16 MiB | 322 / 325 ms | 561.5 KiB | ZIP / final row OK |
| 1 MiB input string | 10 | 2 | 30.53 / 30.61 MiB | 12 / 15 ms | 12.5 KiB | ZIP / final row OK |

The giant-string output is small because the repeated input compresses well;
the measurement still exercises the one-MiB Python string and the writer's
bounded row/work budgets. Every sample passed the worksheet-shape and final-row
checks in `scripts/measure-stream-writer-memory.py`, including the new
streaming semantic-value digest (`semantic_equal: true`). The resulting JSON
passed `scripts/check-stream-writer-measurements.py`.

Run with the release wheel installed in an isolated interpreter:

```text
python3 scripts/measure-stream-writer-memory.py --python <venv>/bin/python \
  --mode append --rows 1000 --columns 2 --repetitions 3 \
  --value-profile escape
python3 scripts/measure-stream-writer-memory.py --python <venv>/bin/python \
  --mode append --rows 10 --columns 2 --repetitions 3 \
  --value-profile giant
```

This is local calibration evidence only. It does not establish Linux or
Windows behavior, Excel compatibility, or a constant-memory guarantee for the
normal workbook VM.
