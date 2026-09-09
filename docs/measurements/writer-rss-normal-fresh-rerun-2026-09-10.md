# Normal VM writer RSS rerun

Date: 2026-09-10 (Asia/Tokyo)

This local measurement quantifies the normal VM writer separately from the
append-only writer. The VM intentionally retains the workbook cell model, so
the result is not a constant-memory claim.

## Conditions and result

- Host: macOS arm64
- Python: CPython 3.13, isolated environment
- Columns: 3; value profile: plain
- Repetitions: 1 per row count
- Validation: ZIP, worksheet shape, generated-input semantic digest

| Rows | RSS | Wall | Output / peak temp | Validation |
|---:|---:|---:|---:|---|
| 100,000 | 120.31 MiB | 1,348 ms | 2.15 / 2.15 MiB | pass |
| 250,000 | 207.17 MiB | 3,307 ms | 5.40 / 5.40 MiB | pass |
| 1,000,000 | 754.95 MiB | 13,776 ms | 21.62 / 21.62 MiB | pass |

RSS increased with the retained VM model, while every output passed semantic
validation. This is a macOS single-sample boundary measurement; it does not
establish Excel compatibility or Linux/Windows resource behavior.

## Reproduction

```sh
python3 scripts/measure-stream-writer-memory.py \
  --python /path/to/isolated/python \
  --rows 100000 250000 1000000 --columns 3 \
  --mode normal-fresh --repetitions 1 \
  --output /tmp/elixcee-rss-normal-fresh.json
python3 scripts/check-stream-writer-measurements.py /tmp/elixcee-rss-normal-fresh.json
```
