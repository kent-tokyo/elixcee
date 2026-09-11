# Append writer RSS rerun

Date: 2026-09-10 (Asia/Tokyo)

This rerun measures the installed candidate Python wheel in isolated child
processes. It covers the append-only writer, not the normal VM writer.

## Conditions and result

- Host: macOS arm64
- Python: CPython 3.13, isolated environment
- Columns: 3; value profile: plain
- Repetitions: 2 per row count
- Validation: ZIP, worksheet shape, generated-input semantic digest

| Rows | RSS p50 / p95 | Wall p50 / p95 | Output / peak temp | Validation |
|---:|---:|---:|---:|---|
| 100,000 | 21.63 / 21.72 MiB | 2,613 / 2,850 ms | 1.36 / 1.36 MiB | pass |
| 250,000 | 21.58 / 21.66 MiB | 5,897 / 5,911 ms | 3.51 / 3.51 MiB | pass |
| 1,000,000 | 21.58 / 21.83 MiB | 23,613 / 29,071 ms | 14.24 / 14.24 MiB | pass |

The append path's measured peak RSS stayed within about 0.15 MiB across the
three sizes in this run. This is a macOS append-path observation; it does not
establish constant memory for the normal VM, Excel compatibility, or support
on Linux and Windows.

## Reproduction

```sh
python3 scripts/measure-stream-writer-memory.py \
  --python /path/to/isolated/python \
  --rows 100000 250000 1000000 --columns 3 \
  --mode append --repetitions 2 \
  --output /tmp/elixcee-rss-append.json
python3 scripts/check-stream-writer-measurements.py /tmp/elixcee-rss-append.json
```
