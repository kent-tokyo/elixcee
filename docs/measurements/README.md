# Measurement index

Dated evidence records, not current API specifications. Each report's source,
fixture, host, and scope limit its conclusions. Open work belongs in the
[roadmap](../../ROADMAP.md); cross-library comparisons are in
[workbook benchmarks](../benchmarks/README.md).

## VM and writer

- [VM hot paths, 2026-09-05](vm-hotpath-optimization-2026-09-05.md): tile cache, write paths, threshold calibration.
- [Formula dirty propagation, 2026-09-05](formula-dirty-calibration-2026-09-05.md): dirty/full equivalence, matrix, p50/p95, resources.
- [Writer streaming follow-up, 2026-09-05](writer-streaming-follow-up-2026-09-05.md): worksheet sink and passthrough clone reduction; remaining whole-ZIP work.

## Reader and release validation

- [Optimization, 2026-09-04](reader-optimization-2026-09-04.md): reader→VM ownership and parsing changes.
- [Large inputs](reader-large-2026-09-03.md), [mixed density](reader-mixed-2026-09-03.md), [density / sheet count](reader-density-sheets-2026-09-03.md).
- [Resource reclamation](reader-reclamation-2026-09-03.md), [in-process soak](reader-inprocess-soak-2026-09-03.md), [mutation/write soak](reader-write-release-soak-2026-09-03.md).
- [Release-binary calibration](reader-release-2026-09-03.md), [signals / cancellation](reader-signal-2026-09-03.md).

The grouped reader records above are dated 2026-09-03. Raw JSON stays beside its
report; missing platform evidence is not implied by a successful macOS run.
[Documentation map](../README.md)
