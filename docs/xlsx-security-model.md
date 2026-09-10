# XLSX security model

## Threat model

Workbooks and caller-supplied worksheet objects are untrusted input.
Compatibility with `xlsx@0.18.5` must not reproduce unsafe behavior.
The [compatibility goal](xlsx-compatibility-goal.md) separates normal-input
compatibility from intentional security and resource-limit differences.

This describes implemented safeguards, not immunity to unknown vulnerabilities.
Versioned changes and future Unreleased work are distinguished in
[CHANGELOG](../CHANGELOG.md).

## Existing limits (1.0.7)

The numeric inventory and calibration notes are maintained in **[Resource limits](limits.md)**.
Do not duplicate thresholds here: ZIP, XML, workbook model, formula, VBA parser,
instruction, call-depth, string, array, and cell limits apply at different layers.

- ZIP metadata is checked before parts are consumed: entry count, decompressed size,
  total size, compression ratio, work budget, and unsafe entry paths.
- The reader rejects absolute paths, parent components, and NUL in ZIP entry names.
  Shared ZIP checks cover file/buffer input, streaming, and save passthrough paths.
- XML parsing rejects forbidden declarations, malformed/incomplete documents,
  invalid controls, and over-budget structures. Worksheet validation and construction
  share one event pass; other parts are validated before their specific parser.
- Formula strings in loaded cells remain subject to XML limits; evaluator/parser
  limits apply when a formula is parsed. Do not confuse storage with evaluation.
- VBA budget failures and blocked external effects are not suppressed by
  On Error Resume Next. Unrelated unsupported statements can still be no-ops.
- Python `Vm.set_budgets()` uses safe defaults for omitted arguments; explicit
  None disables that individual VM limit. This does not disable all reader limits.

The XML iterator is nonrecursive, but flat parsing alone does not bound CPU or memory.
Multiple limits interact: a sheet may hit XML limits before its cell-count limit.

## packages/xlsx (JS) limits

The JS read path uses Rust/WASM validation. In-memory utilities have separate guards:

| Guard | Behavior |
|---|---|
| !ref rectangle | More than 5,000,000 cells fails before formulae/CSV/TXT/JSON/HTML iteration |
| Nonfinite column index | encode_col(+Infinity) is rejected |

The package now includes file APIs; the utility guards are not substitutes for
reader validation. Historical range-cost probes and the exact threshold boundary
are recorded in [limits](limits.md), not a guarantee of elapsed time on every host.

## Remaining validation

| Item | Implemented boundary | Remaining evidence |
|---|---|---|
| Native reader cancellation | Cooperative chunk/part checks; SIGINT, plus snapshot cancel-file | Linux/Windows real-signal calibration |
| Deadline | Cooperative checks; cannot preempt a blocking filesystem call | Hard process isolation and host-specific behavior |
| Synchronous WASM | Default limits; no JS cancellation during a running call | Worker termination or future async contract |
| Resources | Dated macOS large-file, signal, and reclamation tests | Cross-platform and long-duration CPU/RSS/fuzz coverage |

See [measurement records](measurements/README.md) and [roadmap](../ROADMAP.md).
Worker termination is distinct from a cooperative READER_CANCELED result.

## Prototype-pollution-safe key handling

Spreadsheet-derived strings such as __proto__, constructor, and prototype must
remain data, without changing an object's prototype.

- sheet_to_json writes row keys using Object.defineProperty, including explicit
  header arrays containing __proto__.
- book_append_sheet and table_to_book safely construct sheet-name maps.
- Default header inference retains the reference's renamed __proto___NaN text;
  that compatibility quirk is separate from explicit-header prototype injection.

Use own-data-property creation, a null-prototype object, or Map wherever an untrusted
key becomes a property. Do not replace this with unchecked bracket assignment.
See [known differences](compatibility-known-defects.md) and the
[classification registry](../compat/differential/classify.mjs) for fixtures.

## HTML-injection-safe attribute/URL handling

sheet_to_html applies three independent protections:

1. Attribute values are escaped, separately from text content. Text escaping may
   render line breaks as br tags and must not be used as an attribute escaper.
2. Hyperlink targets are allow-listed: http(s), mailto, tel, ftp, relative, and
   fragment targets. Leading/trailing whitespace, ASCII controls, and backslashes
   are rejected to avoid browser normalization ambiguity. Rejected links render as text.
3. cell.h markup is escaped by default. rawHtml:true is an explicit opt-in for
   independently trusted markup, not a sanitizer.

These are registered security divergences whether content came from a file or a caller.

## Safe input paths

The native path reader accepts .xlsx, .xlsm, and .ods, case-insensitively.
Unsupported/missing extensions are rejected before opening the file.
The extension-independent buffer API reads an in-memory OOXML ZIP.
The VM preserves the reader's path-minimizing unsupported-extension error;
this is not a promise that every diagnostic is free of file paths.

## Safe output paths

The native writer rejects unsupported extensions before creating output.
It checks existing destination/path components for symlinks and rejects unsafe
redirects; platform-managed temporary aliases such as macOS /tmp are allowed.
These path checks do not establish a race-free sandbox against concurrent filesystem
mutation by another actor.

XLSX/ODS output is serialized into a same-directory temporary file. Standard save
flushes and syncs it before rename; existing regular-file permissions are preserved
and read-only destinations rejected. Platform replacement fallback behavior is
not a universal atomic-replacement guarantee.

The explicitly selected save_workbook_fast path omits the final sync
and has weaker crash durability. It must not be compared with durable competitors
as if guarantees were identical. This native policy does not describe the separate
JS writer. See [equal-durability measurements](benchmarks/workbook-equal-durable-2026-09-06.md).

## Intentional non-compatibility policy

The differential harness classifies security and resource-limit divergences explicitly;
it does not count them as MATCH or hide the failed cases.
See [classify.mjs](../compat/differential/classify.mjs).

The pinned xlsx oracle is a development/test dependency, not this package's runtime
reader. Prior dependency checks identified
[prototype pollution](https://github.com/advisories/GHSA-4r6h-8v6p-xvw6) and
[ReDoS](https://github.com/advisories/GHSA-5pgg-2g8v-p4x9) advisories for that oracle.
These are historical references, not a fresh dependency audit. Keep test fixtures
isolated and recheck the advisory database before release.

## Open items

Cross-platform resource/cancellation calibration, long-duration fuzzing and isolation,
and any new user-configurable limit surface remain separate roadmap work.
Do not weaken a safeguard simply to match a reference package.
