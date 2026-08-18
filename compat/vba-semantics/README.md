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

301 cases across 18 categories: the original 12 (numeric conversion/rounding, negative
`\`/`Mod`, logical/bitwise, `Str`/`CStr`/`Val`, `IsNumeric`, `TypeName`/`VarType`,
`Date`/`Time`/`Now`, `Empty`/`Null`/error values, string boundaries, array indices, Range
values, error kind), plus six added to reach the 300+ target: division by zero, invalid
procedure arguments, overflow, single-line-If control transfer, `Exit`
Sub/Function/For/Do, object-Nothing access, `+`-vs-`&` operator coercion, comparison-
operator coercion, `Select Case` matching, `With` block resolution, and array bounds.
Not padded to hit a round number — coverage depth varies by category based on how much
real semantic subtlety each one has (numeric rounding has the most tie-breaking/edge-case
richness; `Select Case` matching, being unambiguous control flow with no type-coercion
question, has none of its 9 cases end up as a disclosed gap).

0 `BUG`, 0 `UNCLASSIFIED`. 28 `KNOWN_LIMITATION` — every one a divergence found *by*
building this suite (verified against Microsoft's own VBA language reference where the
answer wasn't already common knowledge), not previously disclosed anywhere else in the
project, and none fixed this round (a fixed divergence isn't `KNOWN_LIMITATION` by
definition — it's `MATCH_DOCUMENTED_SEMANTICS`; several *other* divergences found while
building this same suite, e.g. the Boolean-arithmetic and Empty-equality bugs, were fixed
and so don't appear here — see CHANGELOG.md). Grouped by root cause rather than by count:

- **No declared/runtime type-width tracking** (12): `CInt`/`CLng` silently truncate instead
  of raising `Overflow` on out-of-range values (5); a `Left`/`Right`/`Mid`/`Chr`/`InStr`
  call with an out-of-domain argument (negative length, zero start, out-of-range char code)
  silently clamps instead of raising `Invalid procedure call or argument` (7).
- **Array declaration/resize gaps** (6): `Dim arr(lo To hi)` and `Dim arr()` (empty parens,
  for a later `ReDim`) both fail to parse; `Option Base 1` is parsed but not honored;
  `UBound(arr, dimension)` ignores its dimension argument; `Erase` on a fixed-size array
  doesn't reset elements to their type default; the `Array(...)` builtin isn't implemented
  at all.
- **No Null-propagation semantics anywhere in the VM** (3): `+`, `&` (when *both* operands
  are Null — one-Null-side already correctly degrades to `""`), and every comparison
  operator coerce Null to 0/`""` instead of producing Null.
- **No object-variable unset/Nothing state** (2): a never-`Set` object variable's member
  access silently no-ops instead of raising "Object variable ... not set"; `Set x = Nothing`
  silently no-ops instead of actually clearing the reference.
- **`With`-target resolution is a parse-time literal-string rewrite, not a runtime-resolved
  stack** (2): can't target a computed expression like `Cells(r, c)`, and a bare `.member`
  nested inside another block construct (`If`/`For`/`Do`/`Select Case`) within the body
  isn't recognized.
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
- **Error message text only, condition already correct** (1): a numeric Variant plus a
  non-numeric-string Variant does correctly raise a runtime error (not a silent wrong
  value), but with elixcee's own coercion-failure wording instead of real VBA's "Type
  mismatch". Not fixed this round: the message comes from a single helper shared by ~54
  call sites across the VM, and renaming it without auditing every other caller's own
  correct wording risks introducing a wrong message elsewhere.

Two more parser gaps were found alongside this work but are *not* in this suite (each is a
"does it parse at all" question, closer to `compat/corpus/`'s own scope than a value-
correctness one): the `:` multi-statement-per-line separator doesn't parse at all
(`a = 1: b = 2` fails), and the two `With`-target gaps above are confirmed only for the
specific shapes tested (`Cells(...)` as a target, one level of `If` nesting) — not
exhaustively characterized across every possible target expression or nesting depth. See
`ROADMAP.md`/`CHANGELOG.md` for the full disclosure.
