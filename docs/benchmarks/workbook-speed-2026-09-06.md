# Workbook processing speed benchmark

## Summary

**Superseded:** use the [equal-durability report](workbook-equal-durable-2026-09-06.md)
for performance comparisons. The historical numbers below do not match macOS
durability: Rust `File::sync_all()` uses `F_FULLFSYNC`, whereas this older Python
harness used only `os.fsync()`. The atomic-publication paths also differed.
The elixcee rows describe a local, unreleased working tree with manifest version
1.0.1, not the published 1.0.1 artifact. No speed ranking follows from this table.

This is an initial local benchmark for public reporting. It is a small-fixture
smoke measurement, not a claim of general library superiority.

| Library | Version | Operation | p50 | p95 |
| --- | --- | --- | ---: | ---: |
| elixcee | 1.0.1 | load, mutate A1/B1, save, reload/verify | 5.836 ms | 7.542 ms |
| elixcee fast | 1.0.1 | same, without final fsync | 3.117 ms | 6.028 ms |
| openpyxl | 3.1.2 | load, mutate A1/B1, save, reload/verify | 4.624 ms | 7.040 ms |
| LibreOffice | 26.2.5.2 | load/save conversion only | 1,583.594 ms | 1,595.950 ms |

## Reproduction

```text
python3 compat/benchmarks/workbook_speed.py \
  tests/fixtures/e2e/source.xlsx \
  --elixcee target/release/measure_reader_write_inprocess \
  --iterations 5
```

The harness and raw result are checked in as
`compat/benchmarks/workbook_speed.py` and
`docs/benchmarks/workbook-speed-2026-09-06.json`.

The fast elixcee result was obtained with the additional `--elixcee-fast`
option. It uses the new `save_workbook_fast` API and intentionally skips the
final filesystem durability barrier. Its p50 is lower than the fsync-matched
openpyxl run, but this is not a like-for-like durability comparison.

## Conditions and limits

- Host: macOS arm64; date: 2026-09-06.
- Fixture: `tests/fixtures/e2e/source.xlsx`, 17 populated cells.
- elixcee and openpyxl both mutate A1 and B1, save, reload, and verify A1.
- Both standard paths include a file durability barrier before reload, but
  the barriers were **not equivalent on macOS** (see correction above).
- The standard elixcee row includes its normal durability barrier. The fast row
  is an explicitly weaker, opt-in durability mode and must not be compared as
  the default API's result.
- LibreOffice was measured through headless XLSX conversion and did not mutate
  cells; its number is therefore not directly comparable to the first two.
- Only five iterations were run. No cross-platform, large-fixture, Excel
  oracle, peak-RSS, or CPU-per-library conclusion is established.
- SheetJS was not installed in the checkout; ClosedXML and Aspose.Cells were
  not available in this environment.
