# VBA corpus local regression

Date: 2026-09-10 (Asia/Tokyo)

This is a local regression run of the generated VBA corpus against the
repository's current debug CLI. It checks declared elixcee behavior and
scenario stability; it is not an Excel-equivalence result.

## Result

- Host: macOS arm64
- CLI: `target/debug/elixcee` from the current checkout
- Scenarios: 581
- Successful scenarios: 572
- Declared runtime errors: 8 (`E1099`)
- Declared nondeterministic cases: 1 (`E1002`)
- MISMATCH: 0
- UNEXPLAINED: 0
- Classifier exit status: 0

All 581 scenarios were explained by the checked-in expected-outcomes policy.
The eight runtime errors remain expected behavior, and the single
nondeterministic scenario remains classified as such; neither category is
silently promoted to PASS.

## Reproduction

```sh
cd compat/corpus
TMPDIR=/private/tmp node run-elixcee.mjs ../../target/debug/elixcee
TMPDIR=/private/tmp node classify-elixcee-outcomes.mjs
```

The raw per-scenario output is
[elixcee-results.json](../../compat/corpus/results/elixcee-results.json).
The run used a debug binary because the host was below the available disk
budget for rebuilding the release binary. This does not measure release-binary
performance or close the Excel, Linux, or Windows validation gaps.
