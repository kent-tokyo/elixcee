# Workbook benchmark index

These are dated local measurements, not universal library rankings.
Each report records its own baseline, versions, timing/durability policy,
validation, raw samples, and limitations. The current source may be newer.

| Report (2026-09-06) | What it establishes |
|---|---|
| [Large-workbook optimization](workbook-large-speedup-2026-09-06.md) | Latest incremental comparison: 100k / 400k / 1m cells; strict 1.2× target met only at 400k |
| [Preceding 1.2× optimization](workbook-speedup120-2026-09-06.md) | Separate before/after baseline; tiny-workbook target unmet |
| [ClosedXML / openpyxl / elixcee](workbook-closedxml-2026-09-06.md) | Pinned three-library comparison, before the later optimizations; busy-host variability retained |
| [Equal-durability openpyxl comparison](workbook-equal-durable-2026-09-06.md) | Matched F_FULLFSYNC / atomic rename / reload conditions |
| [Initial workbook comparison](workbook-speed-2026-09-06.md) | Historical exploratory result; initial durability mismatch invalidated the fair-comparison claim |

Read the raw JSON linked by each report. A ratio of medians is not the median of
paired speedups; p95 and output equivalence matter alongside median time.
A rounded 1.196× is not a strict 1.2× success. One million cells split across four
sheets does not establish a one-million-cell single-sheet capacity.

The initial, corrected, intermediate, and regression samples are all retained.
Native durable-save results do not establish Python wrapper throughput, JS writer
performance, full Excel compatibility, or superiority on unmeasured workbooks.

[Internal performance/calibration records](../measurements/README.md) ·
[Roadmap](../../ROADMAP.md) · [Documentation map](../README.md)
