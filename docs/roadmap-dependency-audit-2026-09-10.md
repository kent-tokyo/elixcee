# Roadmap dependency audit

Date: 2026-09-10 (Asia/Tokyo)

This audit classifies every unchecked roadmap item. Classification is about
what can be completed in the current local workspace; it is not a completion
claim.

## Local work still actionable

These items can be advanced with repository code, local fixtures, and local
tooling, but are not complete yet:

- G2d general Chart/Drawing creation and editing, and general Pivot source/cache
  editing.
- G4 independent oracle expansion and runtime-macro regression coverage. The
  local formula-only oracle now has 31 comparable matches out of 34 probes,
  and the 581-scenario VBA corpus has 0 MISMATCH/UNEXPLAINED outcomes; Excel
  semantic coverage and real production macro fixtures remain open.
- G5 confirmation-grade large-workbook measurements, RSS, and resource
  scaling. A new `v1.0.5`-tag versus current paired run now covers 100k/400k/1M
  cells with durable save and output verification; the historical baseline,
  20-pair confirmation, and RSS/resource gates remain open.

## Requires an external engine, host, or service

These remain open and must not be marked complete from local evidence alone:

- Excel-authored fixture comparison, Excel repair-warning/reopen checks, and
  Excel semantic oracle validation.
- EPPlus/Aspose.Cells and LogiSheets fixed-version comparisons where the
  comparison packages and licensing conditions must be fixed and executed.
- Linux/Windows resource, clean-install, and distribution checks.
- Fresh external advisory/review and any registry or formal release checks.

## Intentionally held behind a specification or explicit release decision

- L5 general plugin execution remains held until the capability/resource
  sandbox contract is fixed; arbitrary code and external I/O are not enabled
  as a workaround.
- Tag, registry, workflow, formal Release, and publication state are separate
  release operations. Local commits and tests do not close them.

## Evidence policy

The roadmap remains authoritative for status. A local synthetic fixture,
LibreOffice result, or passing local gate can establish only the scope it
actually exercises. It cannot substitute for Excel, another OS, an external
review, or a published artifact. See the dated records in
`docs/measurements/` for the current local evidence.
