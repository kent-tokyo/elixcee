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

## Recorded coverage and current limitations

The checked-in [report](results/report.json) contains 386 cases:
347 MATCH_DOCUMENTED_SEMANTICS, 23 EXPECTED_ERROR, 2 NONDETERMINISTIC,
14 KNOWN_LIMITATION, and zero BUG/UNCLASSIFIED.
Rechecked locally on macOS on 2026-09-06 with the **1.0.2 release CLI**.
This is a documented-semantics reference suite, not live Microsoft Excel execution.

Current expected answers and disclosed limitations are in
[expected-results.json](expected-results.json). Re-run the suite after implementation
changes; remove a knownLimitation only when the documented expected answer genuinely
matches. Do not preserve an old count by weakening expected values.
The [coverage reference](../../FUNCTIONS.md) supersedes historical parser/Range gap lists.

One deliberate non-coverage decision: **`Select Case` with a `Null` test expression.**
Microsoft's `Select Case` reference documents only that `testexpression` is "matched"
against each `expressionlist` and says nothing about `Null`. Deriving an answer from that
would be a guess, and this suite doesn't encode guesses — so it's left uncovered rather
than covered wrongly.
