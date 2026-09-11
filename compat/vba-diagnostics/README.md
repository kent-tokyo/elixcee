# Redistributable VBA diagnostic artifacts

This directory contains small synthetic fixtures for CI and regression
reproduction. Each case has an `input.xlsx`, a `Run.bas` source file, and an
`expected.json` contract. The files are generated deterministically by:

```sh
python3 scripts/generate-vba-diagnostic-artifacts.py
python3 scripts/run-vba-diagnostic-corpus.py --json
```

The manifest is the authority for case IDs, seeds, and expected machine-readable
codes. These fixtures exercise elixcee's diagnostic contract; they do not prove
Microsoft Excel equivalence, VBA language completeness, or Excel-reopen safety.
They are MIT-licensed synthetic inputs and contain no customer workbook data.
