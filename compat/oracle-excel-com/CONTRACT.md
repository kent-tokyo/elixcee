# Excel COM adapter contract — INTERFACE ONLY, NOT IMPLEMENTED, NOT VERIFIED

**Status: contract/interface definition only.** Nothing in this directory has been run
against Microsoft Excel. No file here should be read as "done," "working," or "verified"
— see `UNVERIFIED.md` for the itemized list of what a real run would need to confirm, and
`RunScenario.ps1` / `RunScenario.vbs`, both marked `UNVERIFIED` in their own header
comments, for untested scaffolding only.

This machine has no Windows and no licensed Excel install, and none is reachable from
this environment. This document exists so a future Windows+Excel session can implement
a real adapter without redesigning the pipeline — see `../corpus/README.md` for
step-by-step instructions addressed to that future session.

## Why this exists

`../corpus/run-libreoffice.mjs` is the actual, executed backend for this milestone.
LibreOffice is explicitly **not** an Excel oracle (see `../corpus/README.md`'s framing
note) — its VBA support is its own compatibility layer. The only way to know what real
Excel does is to run it. This contract defines the shape a Windows+Excel runner would
need to implement to plug into the exact same corpus/normalize/classify pipeline that
already exists and already runs against LibreOffice, so that work is a matter of
implementing one adapter, not rebuilding the harness.

## Interface shape

A conforming adapter is any program that, given one scenario from
`../corpus/scenarios.json` (see `../corpus/SCHEMA.md` for the field shapes), produces one
result record in this shape:

```jsonc
{
  "id": "arithmetic_0007",              // must equal the scenario's own id — the join key
  "category": "arithmetic",             // copied from the scenario, for grouping
  "oracle": "microsoft_excel",          // fixed literal — never "excel", never omitted
  "excel_version": "Microsoft 365, Version 2XXX (Build XXXXX)",  // from Application.Version
                                         // and Application.Build, or equivalent — record
                                         // the ACTUAL running version, never a guess
  "ok": true,                           // false if the macro raised an unhandled VBA error
  "status": "DONE",                     // "DONE" | "TIMEOUT" | "ERROR" | "NO_OUTPUT" —
                                         // same vocabulary run-libreoffice.mjs already uses,
                                         // so classify.mjs's ORACLE_UNAVAILABLE branch needs
                                         // no changes to accept Excel COM results too
  "cells": [
    { "address": "A1", "type": "number", "value": 42 }
  ],                                    // same {address, type, value} shape
                                         // run-libreoffice.mjs's harness already emits —
                                         // type is "number" | "string" | "formula_number",
                                         // decided by Cell.Type/HasFormula in COM terms,
                                         // NOT by a bare .Value read (a text cell's .Value2
                                         // read as a number would silently manufacture a
                                         // false MATCH — the same trap flagged in
                                         // ../corpus/normalize.mjs's doc comment)
  "error": null                         // {message, source} shape if ok is false
}
```

This is intentionally almost identical to what `run-libreoffice.mjs` already writes to
`results/libreoffice-results*.json`, with `oracle` and `excel_version` swapped in. A
result file that matches this shape drops into `../corpus/run-classify.mjs` by adding one
more `results/*-results*.json` glob pattern (or renaming — the current implementation
globs `libreoffice-results*.json` specifically; a `microsoft-excel-results*.json` pattern
is a one-line addition, not a redesign).

## What the future runner needs to do, per scenario

1. Read `../corpus/scenarios.json`. For each scenario:
2. Open the base workbook (`../corpus/workbooks/<workbook>.xlsx`, or a blank workbook if
   `workbook` is `null`) via `Excel.Application` COM automation — real Excel, not a
   compatibility layer.
3. Inject `vbaSource` as a new standard module in the workbook's `VBProject`
   (`Workbook.VBProject.VBComponents.Add(vbext_ct_StdModule)`, then
   `.CodeModule.AddFromString(vbaSource)`) — this requires "Trust access to the VBA
   project object model" enabled in Excel's Trust Center, a one-time manual setting on
   the Windows machine running this.
4. Run the scenario's `entrypoint` Sub (`Application.Run "Module1." & entrypoint`).
5. Walk `Worksheet.UsedRange`, and for every non-empty cell record `{address, type,
   value}` using the same number/string/formula branching described above.
6. Write the result record (schema above) to the shared results directory.
7. Close the workbook without saving; do not leave a modified copy on disk.

## What this contract deliberately does NOT specify

- Which exact Excel version/build to run against (see `UNVERIFIED.md` — that choice, and
  its effect on results, is itself something only a real run can pin down).
- Timeout handling specifics (COM automation hangs differently than the LibreOffice UNO
  hang documented in `../corpus/README.md`; a future implementer needs to determine this
  empirically on Windows, not inherit the 8s LibreOffice figure blindly).
- Any performance/concurrency requirements (the LibreOffice runner shards across
  processes to make ~580 timeouts tractable in wall-clock time; a real Excel run may not
  need this at all if scenarios actually complete instead of hanging).

## Never merge with LibreOffice's numbers

Per this project's explicit instruction: a future Excel COM run's classify-results are
their own report, tagged `oracle: "microsoft_excel"`, on their own summary table. They are
never combined with `oracle: "libreoffice"` results into one blended "compatibility
percentage" — see `../corpus/README.md`'s framing note for why.
