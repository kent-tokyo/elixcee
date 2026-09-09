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
- G4 independent oracle expansion and runtime-macro regression coverage.
- G5 reproducible large-workbook before/after measurements, RSS, durable save,
  and output verification. The current historical baseline cannot be rebuilt
  from the available artifacts.

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
