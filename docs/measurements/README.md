# Measurement index

Dated evidence records, not current API specifications. Each report's source,
fixture, host, and scope limit its conclusions. Open work belongs in the
[roadmap](../../ROADMAP.md); cross-library comparisons are in
[workbook benchmarks](../benchmarks/README.md).

## VM and writer

- [VM hot paths, 2026-09-05](vm-hotpath-optimization-2026-09-05.md): tile cache, write paths, threshold calibration.
- [G5 large hot paths, 2026-09-10](g5-large-hotpath-2026-09-10.md): same-binary cached append and dirty-closure comparisons.
- [G5 paired v1.0.5 comparison, 2026-09-10](g5-paired-v1.0.5-2026-09-10.md): 100k/400k/1M same-fixture before/after measurements, all above the 1.2x target.
- [Formula oracle local rerun, 2026-09-10](formula-oracle-local-2026-09-10.md): eight arithmetic LibreOffice cases, all timeout and therefore not comparable.
- [Formula-only independent oracle, 2026-09-10](formula-independent-oracle-2026-09-10.md): 94 direct XLSX formula probes, 85/85 paired matches, and 9 explicitly skipped unsupported probes.
- [Formula oracle follow-up probe, 2026-09-10](formula-independent-oracle-probe-2026-09-10.md): `ROWS`/`COLUMNS` gap discovery and paired repair evidence.
- [Formula dirty propagation, 2026-09-05](formula-dirty-calibration-2026-09-05.md): dirty/full equivalence, matrix, p50/p95, resources.
- [Writer streaming follow-up, 2026-09-05](writer-streaming-follow-up-2026-09-05.md): worksheet sink and passthrough clone reduction; remaining whole-ZIP work.
- [Writer streaming, 2026-09-07](writer-streaming-2026-09-07.md): release-wheel append and transaction-batched normal-VM RSS rerun through one million rows.
- [Writer input calibration, 2026-09-09](writer-input-calibration-2026-09-09.md): XML escaping and one-MiB Python input RSS calibration.
- [Python date1904 contract, 2026-09-10](python-date1904-2026-09-10.md): isolated local wheel verification of metadata load and save preservation.
- [Append writer RSS rerun, 2026-09-10](writer-rss-rerun-2026-09-10.md): isolated 100k/250k/1M-row RSS and semantic-output verification.
- [Normal VM writer RSS rerun, 2026-09-10](writer-rss-normal-fresh-rerun-2026-09-10.md): retained-model RSS boundary at 100k/250k/1M rows.
- [G2d object editing, 2026-09-09](g2d-object-editing-2026-09-09.md): bounded Chart-series and worksheet-backed Pivot-source edit verification.
- [VBA runtime safety, 2026-09-09](vba-runtime-safety-2026-09-09.md): default rejection and structured diagnostics for Save/Close external effects.
- [VBA event dispatch, 2026-09-09](vba-event-dispatch-2026-09-09.md): explicit event execution, EnableEvents, and re-entry suppression.
- [VBA corpus local regression, 2026-09-10](vba-corpus-local-2026-09-10.md): 581 generated scenarios, all explained with no mismatch or unexplained outcome.

## Reader and release validation

- [Optimization, 2026-09-04](reader-optimization-2026-09-04.md): reader→VM ownership and parsing changes.
- [Large inputs](reader-large-2026-09-03.md), [mixed density](reader-mixed-2026-09-03.md), [density / sheet count](reader-density-sheets-2026-09-03.md).
- [Resource reclamation](reader-reclamation-2026-09-03.md), [in-process soak](reader-inprocess-soak-2026-09-03.md), [mutation/write soak](reader-write-release-soak-2026-09-03.md).
- [Release-binary calibration](reader-release-2026-09-03.md), [signals / cancellation](reader-signal-2026-09-03.md).

The grouped reader records above are dated 2026-09-03. Raw JSON stays beside its
report; missing platform evidence is not implied by a successful macOS run.
[Documentation map](../README.md)
