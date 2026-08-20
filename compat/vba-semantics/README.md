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

386 cases across 25 categories: the original 12 (numeric conversion/rounding, negative
`\`/`Mod`, logical/bitwise, `Str`/`CStr`/`Val`, `IsNumeric`, `TypeName`/`VarType`,
`Date`/`Time`/`Now`, `Empty`/`Null`/error values, string boundaries, array indices, Range
values, error kind), plus division by zero, invalid procedure arguments, overflow,
single-line-If control transfer, `Exit` Sub/Function/For/Do, object-Nothing access,
`+`-vs-`&` operator coercion, comparison-operator coercion, `Select Case` matching, `With`
block resolution, and array bounds — and, added for the VBA-structural-semantics round,
**`null_propagation`** (38), **`colon_statement_separator`** (19), plus large expansions of
`with_block_resolution` (6 → 23) and `object_nothing_access` (2 → 12).
Not padded to hit a round number — coverage depth varies by category based on how much
real semantic subtlety each one has (numeric rounding has the most tie-breaking/edge-case
richness; `Select Case` matching, being unambiguous control flow with no type-coercion
question, has none of its 9 cases end up as a disclosed gap).

0 `BUG`, 0 `UNCLASSIFIED`. **16 `KNOWN_LIMITATION`, down from 28** — thirteen were genuinely
fixed across two rounds (a fixed divergence isn't `KNOWN_LIMITATION` by definition — it
becomes `MATCH_DOCUMENTED_SEMANTICS`/`EXPECTED_ERROR`, and its `knownLimitation` annotation
is *removed*, not weakened): the three Null-propagation ones, the two object-variable
unset/Nothing ones, the two `With`-target ones, the `Type mismatch` error-message one, and
the missing `Array()` builtin (structural-semantics round); `Dim arr(lo To hi)`, `Dim arr()`
(empty parens), `Option Base 1`, and `Erase` on a fixed-size array (a later round, adding
array lower-bound tracking). See CHANGELOG.md for what each fix actually changed. One
divergence was newly disclosed in that same later round, not fixed: writing to two array
elements that share dimension 1's index but differ in dimension 2 silently collides on the
same underlying element (`two_dimensional_array_second_index_is_silently_dropped` —
elixcee's array storage is genuinely 1-D; a previous, differently-shaped version of this
case passed by coincidence and had been miscited as evidence 2-D storage worked). The
remaining 16, grouped by root cause rather than by count:

- **No declared/runtime type-width tracking** (12): `CInt`/`CLng` silently truncate instead
  of raising `Overflow` on out-of-range values (5); a `Left`/`Right`/`Mid`/`Chr`/`InStr`
  call with an out-of-domain argument (negative length, zero start, out-of-range char code)
  silently clamps instead of raising `Invalid procedure call or argument` (7).
- **Array storage is 1-D only, not truly multi-dimensional** (2): `Dim arr(3, 2)` allocates
  only dimension 1's elements, silently discarding dimension 2's size; every array
  write/read indexes using only the first index expression, so a second (or later) index
  is dropped rather than addressing a distinct element, and `UBound(arr, dimension)` can't
  honor its dimension argument since there's nothing per-dimension to report. Needs real
  shape metadata and stride arithmetic — deliberately deferred as comparable in scope to
  this project's other deferred Variant-surface work, not attempted alongside the smaller,
  independent lower-bound-tracking fixes in the same round.
- **No per-Variant stored-type tag distinguishing "string that looks numeric" from
  "genuine number"** (1): `+` between two Variants that both hold strings numeric-adds
  instead of concatenating, even though real VBA's own documented rule concatenates
  whenever *both* sides are string-typed, independent of content.
- **A numeric Variant compared to a string Variant isn't unconditionally "less than"** (1):
  real VBA's documented rule for `<`/`>` between a numeric-typed and string-typed Variant
  ignores magnitude entirely; elixcee still numeric-compares when the string looks numeric.
  Deliberately not "fixed" — the current behavior is far more useful for the overwhelmingly
  more common real-world case (numeric-string-vs-number threshold checks), and the fix
  would need to invert it for every caller, not just this one case.

Two divergences found alongside the structural-semantics work are *not* in this suite,
because the shapes this suite can express don't distinguish them; both are recorded in
`ROADMAP.md`'s known-defects list instead: `Range.Range(...)`/`Range.Cells(...)` inside a
`With <range>` body resolve absolutely rather than relative to the base range, and a line
skipped wholesale as an unrecognized *block header* still swallows its trailing
`:`-separated statements.

One deliberate non-coverage decision: **`Select Case` with a `Null` test expression.**
Microsoft's `Select Case` reference documents only that `testexpression` is "matched"
against each `expressionlist` and says nothing about `Null`. Deriving an answer from that
would be a guess, and this suite doesn't encode guesses — so it's left uncovered rather
than covered wrongly.
