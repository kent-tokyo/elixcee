# elixcee v1 support contract

This document defines what `elixcee 1.0` supports and what it deliberately does
not claim. It is the public release boundary for the documented data-processing
subset, not a promise to emulate the Excel desktop application.

## Supported contract

- Rust, Python, CLI, and experimental JavaScript/WASM interfaces use the same
  documented workbook and diagnostic model where the interface is available.
- Workbook coordinates are 1-based in the VBA and Python-facing APIs.
- The supported input formats are `.xlsx`, `.xlsm`, and `.ods`, subject to the
  reader and model limits in [docs/limits.md](limits.md).
- Data-processing VBA constructs, formulas, ranges, multiple worksheets,
  structured diagnostics, and documented workbook editing operations are
  supported only to the extent listed in [FUNCTIONS.md](../FUNCTIONS.md) and
  the public API signatures.
- Default safety behavior rejects blocked external effects, malformed or
  over-budget input, unsafe paths, and unsafe output conditions with an error.
  A rejected input does not produce a partially trusted workbook.
- `.xlsm` VBA project bytes are preserved by supported round-trip paths;
  preservation does not mean every macro is executable by the VM.

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
oracle. For v1.0.1, the local offline gate covers the checked-in Rust workspace,
property/integration tests, clippy, dependency policy, feature compilation, and
the checked-in reader-measurement contract. Dated macOS reader measurements are
available under `docs/measurements/`; Excel-oracle and cross-platform results
remain separate evidence and must not be inferred from the local gate.

The roadmap remains the source of truth for expanding this contract:
[ROADMAP.md](../ROADMAP.md).
