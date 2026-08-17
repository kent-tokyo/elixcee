# What remains unverified against real Microsoft Excel

Everything. No scenario in `../corpus/scenarios.json` has ever been run against real
Excel. This is the complete, explicit list of what a future Windows+Excel session would
need to confirm — framed as exactly that, not as a residual gap in an otherwise-verified
result.

## Why nothing here is verified

This session runs on macOS with no Windows and no licensed Excel install, and neither is
reachable from this environment. That was an explicit, upfront scoping decision (not
discovered partway through) — see the milestone brief this work was built under. The
corpus, both runners, the normalizer, and the classifier were all built to be
oracle-agnostic specifically so this gap has a clear, mechanical way to close later
(`CONTRACT.md` + `RunScenario.ps1`), rather than requiring a redesign.

## Itemized

1. **Every one of the ~580 scenarios in `../corpus/scenarios.json`** — arithmetic,
   boolean logic, string functions, range read/write, control flow (For/Do
   While/If/Select Case), arrays, type conversion, worksheet functions
   (`Application.WorksheetFunction.*`), nested Sub/Function calls, error handling
   (`On Error`), the deliberately-not-yet-implemented functions in the
   `unsupported_functions` category, and the deliberately time-dependent
   `nondeterministic` category. See `../corpus/results/classify-results.json` for the
   full list with LibreOffice's (not Excel's) measured outcome per scenario.

2. **Whether elixcee's `[array]`/`[record]` CLI serialization placeholders correspond to
   anything comparable in real Excel's output at all** — flagged in
   `../corpus/normalize.mjs`'s doc comment as a case that can never resolve to MATCH
   under the current CLI contract; unverified whether a real Excel comparison would need
   a richer elixcee output mode instead of working around it on the comparison side.

3. **Whether the `\` integer-division operator, infix `And`/`Or`/`Not`/`Xor` in a value
   expression, typed `Function` parameter/return annotations, comma-separated
   multi-variable `Dim`, and `With` over a `Range` object** — all valid VBA syntax that
   elixcee's parser rejected during this milestone's corpus run (see
   `../corpus/results/elixcee-results.json` for the exact `E2001`/`E1002` errors) — are
   gaps worth prioritizing. A real Excel run doesn't change whether elixcee supports
   these, but confirms whether real-world VBA actually relies on them enough to matter.

4. **LibreOffice's own measured MATCH rate is not a proxy for an Excel MATCH rate.**
   LibreOffice's VBA-compatibility layer has its own, independently-implemented semantics
   (see `../corpus/README.md`) — a MATCH between elixcee and LibreOffice says nothing
   about whether either agrees with Excel on that scenario, and a divergence between
   elixcee and LibreOffice says nothing about which one (if either) matches Excel.

5. **Every scenario that timed out under LibreOffice** (see `../corpus/README.md`'s
   "Known, reproducible limitation" section — expected to be the large majority of the
   corpus, since nearly every scenario exercises `Range`/`Cells`) has **no measured data
   point at all**, from either oracle, for whether elixcee's output is correct. A real
   Excel run is the only way to learn anything about elixcee's correctness on these,
   since LibreOffice could not execute them either.

6. **Which exact Excel version/build to standardize on.** Excel's own worksheet-function
   surface has changed across versions (see `FUNCTIONS.md`'s version-legend column, e.g.
   `XLOOKUP` requiring 365/2021+) — a scenario that's UNSUPPORTED against Excel 2016 might
   be a real MATCH-or-BUG question against Microsoft 365. `CONTRACT.md`'s result schema
   requires recording the actual running version specifically so this doesn't get
   silently averaged away.

## What would NOT still be unverified after a real run

Running `RunScenario.ps1`/`.vbs` (once implemented — they are untested scaffolding right
now, see their own header comments) against the existing `scenarios.json` and feeding the
output into `../corpus/run-classify.mjs` (with the one-line results-glob addition noted in
`CONTRACT.md`) would close all six items above in a single pass, using the exact same
scenario corpus, normalizer, and classifier already built and already exercised against
LibreOffice — no redesign needed, only the adapter itself.
