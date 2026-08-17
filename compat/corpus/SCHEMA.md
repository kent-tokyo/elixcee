# Corpus scenario schema

A **scenario** is one VBA macro run against one starting workbook, described in a way
that doesn't assume which engine executes it. `generate-scenarios.mjs` produces
`scenarios.json` (an array of these); `run-elixcee.mjs` and `run-libreoffice.mjs` each
read the same file and drive their own backend against it — neither script's shape
depends on the other's existence, per this project's "oracle-agnostic corpus" requirement
(see the milestone brief this was built under).

## Fields

```jsonc
{
  "id": "arith_add_int_0007",       // stable, unique, used as the join key across
                                     // elixcee-results.json / libreoffice-results.json /
                                     // classify-results.json — never regenerate IDs on a
                                     // corpus re-run without a reason, or history breaks.
  "category": "arithmetic",         // one of CATEGORIES in generate-scenarios.mjs
  "description": "A1 = 2 + 3 (integer literals)",
  "vbaSource": "Sub Scenario()\n  Range(\"A1\").Value = 2 + 3\nEnd Sub",
  "entrypoint": "Scenario",         // Sub name every scenario uses, by convention
  "workbook": "empty"               // key into workbooks/*.xlsx (see below); null means
                                     // "run with no --file at all" (elixcee's own default
                                     // single implicit sheet)
}
```

## Example

```json
{
  "id": "range_read_write_0012",
  "category": "range_readwrite",
  "description": "Read B2 from the numeric_grid fixture, double it into C2",
  "vbaSource": "Sub Scenario()\n  Range(\"C2\").Value = Range(\"B2\").Value * 2\nEnd Sub",
  "entrypoint": "Scenario",
  "workbook": "numeric_grid"
}
```

## What "observe" means

There is deliberately no `observe`/`expected` field. Both runners are required to dump
**every non-empty cell** on the sheet the macro left active, in `{sheet, address, value}`
form, sorted by `(row, column)` — this is exactly `elixcee --json`'s own `cells` contract
(see `docs/agent-contract.md`), and `run-libreoffice.mjs`'s harness macro is written to
emit the same shape. The normalizer/classifier then diff the two full dumps. This avoids
a second place (the scenario file) encoding an assumption about which cells matter — the
comparison is exhaustive by construction, and a bug that clobbers an *unexpected* cell is
still caught.

## What "expected Excel output" is NOT here

No field anywhere in this schema encodes what real Microsoft Excel would output. Per the
scoping decision for this milestone, nobody has run these scenarios against Excel, so
there is nothing honest to record. A future Excel COM run (see
`../oracle-excel-com/CONTRACT.md`) produces its own result file in the same shape as
`libreoffice-results.json`, tagged `"oracle": "microsoft_excel"` — never merged into this
schema as a baked-in "expected" value.

## Base workbooks (`workbooks/*.xlsx`)

Generated once by `workbooks/generate-workbooks.mjs` (uses the `xlsx` npm package already
present in `compat/package.json` as a devDependency for the unrelated xlsx-oracle work —
no new dependency added). A handful of named fixtures (`empty`, `numeric_grid`,
`mixed_types`, `with_text`, `with_negatives`) are reused across many scenarios rather than
one xlsx per scenario, to keep the corpus small and the fixtures auditable by hand.
