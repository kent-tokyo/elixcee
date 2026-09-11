# CI VBA diagnostics

elixcee can run and diagnose the documented data-processing VBA subset without
Microsoft Excel. The intended CI entry point is `diagnose-workbook`: it runs a
fixture with deterministic generated inputs and emits one machine-readable JSON
result. It is a workbook diagnostic contract, not a complete VBA compiler or
Excel object-model replacement.

## Minimal GitHub Actions step

Keep the fixture, exported VBA modules, and workbook under version control:

```yaml
- name: Install elixcee
  run: python -m pip install elixcee

- name: Diagnose workbook behavior
  run: |
    mkdir -p artifacts
    elixcee diagnose-workbook tests/fixtures/your-workbook/diagnose_fixture.toml \
      --json > artifacts/elixcee-diagnose.json
```

Replace the example path with a fixture committed by your project. The command
exits non-zero when the fixture fails, an input cannot be resolved,
the VBA cannot be parsed, or the run times out. Upload
`artifacts/elixcee-diagnose.json` on failure so the seed, case index, inputs,
failure detail, and classified root cause remain available to a CI reviewer.

## Fixture contract

`diagnose-workbook` uses the same bounded TOML fixture as `test-workbook`:

```toml
name = "order calculation"
workbook = "fixtures/orders.xlsx"
vba_files = ["Main.bas"]
macro = "Process"
cases = 100
seed = 42
timeout_secs = 10

[[inputs]]
range = "Input!B2:B10"
strategy = "boundary_numeric"

[[assertions]]
range = "Result!A1:F100"
rule = "no_excel_errors"
```

The fixture path is the reproducibility boundary. Pin the workbook, exported
VBA, case count, seed, and elixcee version together. A failure can be replayed
with `--seed` and `--case`; `--cases` is useful for a quick local reproduction,
but it does not replace the full CI run.

## JSON triage

Successful runs contain `schema_version`, `ok`, `seed`, and `cases_run`.
Failures additionally contain `case_index`, `inputs`, `failure`, and
`root_causes`. `root_causes` is empty for failures that are not yet classified;
the raw `failure` fields remain authoritative. Hidden-row/column findings are
reported separately as `observations` when present.

Do not turn an empty `root_causes` array into a pass condition. CI should use the
process exit code and `ok`, then retain the JSON for diagnosis. The output is
deterministic for the same fixture, seed, case selection, and runtime version;
it is not an Excel oracle.

## Recommended diagnostic corpus

The checked-in manifest can be executed without external data:

```bash
python3 scripts/run-vba-diagnostic-corpus.py --json
```

Use `--case <id>` for a focused replay. The runner has an allowlisted mapping
from corpus class to a Cargo test target; manifest text cannot supply shell
commands. The tests construct synthetic inputs or use checked-in fixtures, and
their assertions verify the expected root cause or observation.

The first public corpus should use input/output workbook pairs and classify
each task by the failure mechanism it is meant to catch:

| Class | Required evidence |
|---|---|
| Range and Paste | source/target addresses, shape, and classified mismatch |
| Hidden rows/columns | range plus observed visibility intervals |
| Protection | sheet protection state and rejected write |
| Formula result | formula text, calculated value, and expected value source |
| VBA resolution | module/entry point and parser or name-resolution diagnostic |

Store only synthetic or redistributable workbooks. Do not publish customer
workbooks, embedded secrets, or unreviewed Excel files. External benchmark
datasets and Excel-generated expected values require separate license and
oracle review before inclusion.

## Current limits

The diagnostic runner does not execute embedded VBA automatically, does not
emulate Excel UI, and does not claim complete Excel formula or OOXML
compatibility. A passing local fixture is evidence for that fixture and runtime
version only; it is not evidence of Excel parity, cross-platform parity, or
coverage of unclassified runtime errors.
