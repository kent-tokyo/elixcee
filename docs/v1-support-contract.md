# elixcee v1 support contract

This document defines the v1 support policy for a headless Excel workbook
automation runtime, not a promise to emulate the Excel desktop application.
The runtime has three related surfaces on its workbook model: direct workbook
editing, supported formula recalculation, and execution/diagnosis of the
documented data-processing VBA subset. Current coverage is documented for **1.0.7**.
Use [CHANGELOG](../CHANGELOG.md) to identify changes by version.

## Supported contract

- Rust, Python, and CLI share the native workbook/diagnostic core. The private,
  unpublished JavaScript package uses the WASM reader but has its own workbook
  facade and writer; native preservation and VBA APIs do not automatically apply.
- Workbook coordinates are 1-based in the VBA and Python-facing APIs.
- The supported input formats are `.xlsx`, `.xlsm`, and `.ods`, subject to the
  reader and model limits in [docs/limits.md](limits.md).
- Data-processing VBA constructs, formulas, ranges, multiple worksheets,
  the documented built-in `Collection` and class-module subsets, structured
  diagnostics, and documented workbook editing operations are
  supported only to the extent listed in [FUNCTIONS.md](../FUNCTIONS.md) and
  the public API signatures.
- Formula references produced by `INDIRECT` and `OFFSET` are supported within
  the bounded contract: A1 ranges and absolute R1C1 references for
  `INDIRECT`, and bounded source ranges with `height`/`width` for `OFFSET`.
  Relative R1C1, sheet-qualified or external references, and Excel-oracle
  equivalence remain unverified or unsupported.
- Default safety behavior rejects blocked external effects, malformed or
  over-budget input, unsafe paths, and unsafe output conditions with an error.
  A rejected input does not produce a partially trusted workbook.
- Workbook external links are never fetched or executed. The Python loader
  accepts `external_links="preserve"` (default, round-trip only),
  `external_links="reject"` to fail closed, or `external_links="drop"` to
  remove external-link parts on save. The CLI exposes the same choice as
  `--external-links preserve|reject|drop` for `run` and `snapshot`.
- `.xlsm` VBA project bytes are preserved by supported round-trip paths;
  preservation does not mean every macro is executable by the VM.
- Structural edits on a workbook containing chart/drawing or pivot owners are
  rejected when their references cannot be rewritten safely. Loaded-workbook sheet
  rename is the narrow exception: qualified chart formulas and Pivot
  `worksheetSource@sheet` are updated. In addition, existing chart-series
  category/value formulas and worksheet-backed Pivot `worksheetSource@sheet`／A1
`ref` can be explicitly edited through the bounded APIs. Existing Chart title
text, indexed Drawing shape text runs, and two-cell Drawing anchor markers are
also supported by bounded APIs. Chart creation, cache
regeneration/recalculation, general shape editing, and
stale-anchor updates after row/column changes remain outside the contract.

## Explicit non-goals

The v1 contract does not claim:

- complete Excel function, VBA object model, or desktop UI compatibility;
- compatibility with charts, pivots, drawings, external links, or other
  OOXML objects unless the current compatibility documentation and tests cover
  that exact operation;
- that every workbook can be edited and reopened by Microsoft Excel without
  a warning or logical difference;
- that arbitrary VBA is safe to execute, or that the VM supports external
  files, Shell, COM, ActiveX, UserForms, or network effects;
- a performance advantage over ClosedXML, Aspose.Cells, openpyxl, SheetJS,
  LibreOffice, or Excel without a dated, reproducible measurement;
- unbounded input size, unbounded execution time, or complete protection from
  unknown future vulnerabilities.

## Compatibility vocabulary

Public documentation uses these separate states:

| State | Meaning |
|---|---|
| `supported` | The operation is implemented and covered by the relevant tests. |
| `preserved` | The data is retained through the supported read/write path, without claiming semantic editing support. |
| `warned` | The input or operation is recognized but has an explicit compatibility limitation. |
| `rejected` | Continuing would be ambiguous, unsafe, or outside the documented contract. |
| `unverified` | An external oracle, platform, or measurement required by the roadmap was unavailable. |

`unverified` is not counted as a compatibility success. Security divergence
from an oracle is intentional when reproducing the oracle would create an
injection, path traversal, or resource-exhaustion risk.

## Release evidence

Each release records implementation tests and static checks separately from
measurements requiring Microsoft Excel, large fixtures, or another independent
oracle. For v1.0.7, the local gate covers the Rust workspace/all-target tests,
strict clippy/Rustdoc, fresh dependency audit, feature compilation, packaged
crate/wheel/sdist, JavaScript differential/type checks, and the checked-in
reader-measurement contract. Dated macOS reader measurements are
available under `docs/measurements/`; Excel-oracle and cross-platform results
remain separate evidence and must not be inferred from the local gate.

The roadmap remains the source of truth for expanding this contract:
[ROADMAP.md](../ROADMAP.md).
