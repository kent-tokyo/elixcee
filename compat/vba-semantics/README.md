# compat/vba-semantics/ — VBA value-correctness suite

Answers a different question from `compat/corpus/`: not "does elixcee run without
erroring" (that's `compat/corpus/`'s own PASS/FAIL axis, and separately its
elixcee-vs-LibreOffice/Excel oracle axis), but **is the VALUE elixcee produces the one
real, documented VBA semantics says it should be**. A function that runs without error
and returns a plausible-but-wrong number is invisible to `compat/corpus/`'s classifiers —
that was exactly the failure mode `Round()`'s negative-digits handling, `CInt`/`CLng`'s
rounding mode, `IsNumeric`'s string handling, `Str()` vs `CStr()`, `Val()`'s whole-string
parsing, and `Dim`'s Empty-variable registration all were before this suite existed (the
last of those was found by this suite's own first run, not by hand-auditing source).

Needs no oracle at all — the "expected" value for each case is computed from
`reference/*.mjs`, small, independently-checkable pure-JS reference implementations of
documented real VBA semantics (banker's rounding, `Str()`'s leading-space quirk,
`Val()`'s leading-numeric-prefix parsing, `And`/`Or`/`Xor`/`Not`'s logical-vs-bitwise
split, ...), not hand-typed one value at a time.

## Layout

- `reference/*.mjs` — the ground-truth computations. If one of these is wrong, every case
  built from it is wrong the same way — read these first when investigating a surprising
  `BUG` verdict.
- `generate-cases.mjs` → `cases.json` (VBA scenario definitions) + `expected-results.json`
  (documented-real-VBA expected outcome per case, keyed by id). Generated from templates
  parameterized over value ranges, not hand-typed — same "generate, don't hand-author, but
  commit the result" precedent as `../corpus/generate-scenarios.mjs`. Re-run this and
  commit the output when adding cases; don't hand-edit `cases.json`/`expected-results.json`.
- `run-elixcee.mjs` — drives `cases.json` against the real `elixcee` CLI binary (same
  shape as `../corpus/run-elixcee.mjs`) → `results/elixcee-results.json`.
- `report.mjs` — joins `elixcee-results.json` against `expected-results.json`, classifies
  every case, writes `results/report.json`. **This is the CI gate**: exits non-zero if any
  case is `BUG` or `UNCLASSIFIED`.

## How to re-run everything

```sh
cd compat/vba-semantics
node generate-cases.mjs      # only needed after editing generate-cases.mjs itself
cd ../..
cargo build --release --bin elixcee
cd compat/vba-semantics
node run-elixcee.mjs
node report.mjs
```

## Verdicts

| Verdict | Meaning |
|---|---|
| `MATCH_DOCUMENTED_SEMANTICS` | Actual value matches documented real-VBA semantics exactly. |
| `EXPECTED_ERROR` | Expected an error, got exactly that error message. |
| `NONDETERMINISTIC` | No fixed expected value is meaningful (e.g. `Now()`'s sub-second component) — only checked for running without erroring. |
| `KNOWN_LIMITATION` | Actual diverges from documented real-VBA semantics, but this exact case is registered (`knownLimitation` field in `expected-results.json`, always with a written reason — never inferred) as an already-disclosed gap. Doesn't gate CI. |
| `BUG` | Actual diverges from documented real-VBA semantics and is **not** registered as known. This is what the suite exists to catch. Must be 0. |
| `UNCLASSIFIED` | Something structurally wrong with the suite itself (no result recorded, no cell found at the expected address, an unrecognized `expected.kind`) — never "explained away"; a bug in the suite, not a verdict on elixcee. Must be 0. |

**Anti-laundering rule, mirroring `../differential/classify.mjs`'s `UNSUPPORTED_ALLOWLIST`
and `../corpus/expected-outcomes.json`'s own discipline**: `expected-results.json`'s
`value`/`errorMessage` fields always hold the *documented real-VBA* answer, even for a
case elixcee is known to get wrong — never elixcee's own (possibly wrong) output laundered
into looking like "the spec". A mismatch either gets a `knownLimitation` reason (written
by a human who looked at the actual divergence) or it's `BUG`. Nothing silently downgrades
a real bug into a passing case by weakening what's expected.

## Current state

208 cases across 12 categories (numeric conversion/rounding, negative `\`/`Mod`,
logical/bitwise, `Str`/`CStr`/`Val`, `IsNumeric`, `TypeName`/`VarType`, `Date`/`Time`/`Now`,
`Empty`/`Null`/error values, string boundaries, array indices, Range values, error kind).
Not padded to hit a round number — coverage depth varies by category based on how much
real semantic subtlety each one has (numeric rounding has the most tie-breaking/edge-case
richness; Range value round-tripping has the least). 0 `BUG`, 0 `UNCLASSIFIED`, 1 disclosed
`KNOWN_LIMITATION` (array-out-of-bounds error message text doesn't match real VBA's exact
wording, though the error *condition* is correct).
