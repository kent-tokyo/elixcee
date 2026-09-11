# VBA corpus and oracle runners

The scenario corpus compares elixcee with LibreOffice, and separately checks
elixcee's registered pass/fail expectations. LibreOffice is **not Microsoft Excel**.
The [Excel COM adapter](../oracle-excel-com/UNVERIFIED.md) remains unverified scaffolding.

## Layout and comparison axes

| File | Purpose |
|---|---|
| [SCHEMA.md](SCHEMA.md), [scenarios.json](scenarios.json) | Scenario contract and generated corpus |
| generate-scenarios.mjs | Deterministic case generation |
| workbooks/generate-workbooks.mjs | Generate five base workbooks using the test-only xlsx dependency |
| run-elixcee.mjs | Run the native CLI and record outcomes |
| run-libreoffice.mjs | Invoke headless LibreOffice in isolated temporary profiles |
| normalize.mjs / classify.mjs / run-classify.mjs | Compare logical cell values with an independent engine |
| expected-outcomes.json / classify-elixcee-outcomes.mjs | Check elixcee's declared behavior, not Excel equivalence |

The corpus contains 581 scenario definitions. The latest local macOS run with the
current checkout's debug CLI on 2026-09-10 produced 572 PASS, 8
EXPECTED_RUNTIME_ERROR, and 1 NONDETERMINISTIC, with zero MISMATCH/UNEXPLAINED.
Self-expectation checks fail on MISMATCH or UNEXPLAINED;
they must not turn unknown failures into expected outcomes.
Oracle comparisons require actual comparable output: unavailable engines are not MATCH.

## Run from the repository root

```sh
cd compat
npm ci
cd ..
cargo build --release --bin elixcee
cd compat/corpus
node workbooks/generate-workbooks.mjs
node generate-scenarios.mjs
node run-elixcee.mjs
node classify-elixcee-outcomes.mjs
```

Regenerate scenarios/workbooks only when needed and review changed fixtures.
Self-checks: `node normalize.mjs` and `node classify.mjs`.

The optional LibreOffice leg requires a working soffice installation:

```sh
node run-libreoffice.mjs
node run-classify.mjs
```

Read the limitation below first: an 8-second timeout per scenario makes a full serial
run potentially exceed an hour. The runner accepts `[count] [startIndex] [outSuffix]`
for disjoint shards, each with its own temporary profile.
The classifier globs results/libreoffice-results*.json; avoid mixing stale shards
or different engine versions when interpreting a new run.

## Recorded LibreOffice invocation limitation

Historical results in [classify-results.json](results/classify-results.json) contain
581 records: 578 ORACLE_UNAVAILABLE, one MATCH, and two NONDETERMINISTIC.
The earlier prose total of 580 was inconsistent with these records.
These are retained results, not a rerun on the current host or implementation.

The harness smoke, which avoids Range/Cells, completed module insertion, invocation,
and cell dumping. The recorded object-model scenarios instead hung at Range/Cells
access through getScriptProvider().getScript(...).invoke(...).

Observed attempts included in-memory versus saved/reopened documents, visible versus
hidden loading, and Range versus Cells access. A document-load Auto_Open path did
not show the expected macro side effect. A lock/VBA-initialization explanation was
only a hypothesis: no confirming stack sample was obtained.

Timeouts are classified ORACLE_UNAVAILABLE, not silent skips or matches.
“Zero silent wrong results detected” here refers to **one comparable scenario**,
not a clean bill of health across the corpus.

## Retained harness debugging notes

| Observation in the earlier environment | Implication |
|---|---|
| Bundled LibreOffice Python was killed on startup | CLI macro invocation was used instead; this is not a current platform-wide claim |
| System/bundled Python ABI mismatch | A socket pyuno bridge was not validated |
| Repeated Environ() calls crashed the observed process | Bake parameters into the generated macro instead |
| Unescaped ampersands made .xba modules malformed | XML-escape the complete macro source before writing it |

Cell normalization reads getType(), not only getValue(): the latter can turn text
into zero and manufacture a false match. Boolean-versus-numeric ambiguity remains
a documented limitation in normalize.mjs; do not weaken equality to hide it.

## Current support versus historical failures

Old corpus runs exposed parser gaps, several of which were subsequently implemented.
Use [FUNCTIONS](../../FUNCTIONS.md) for current coverage and the current runner
output for execution status, not a historical count of parse/runtime failures.
Selected [Excel workbook round trips](../oracle-excel-com/results/0.9.0-A_summary.md)
do not close this corpus's missing Excel execution evidence.
