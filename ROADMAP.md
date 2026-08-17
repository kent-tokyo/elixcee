# Roadmap

This tracks where `elixcee`/`@elixcee/xlsx` stands and what's next. For what already
shipped, see `CHANGELOG.md` — this file only restates completed work when needed to explain
what's left. Historical phase-by-phase implementation notes (Japanese) live in
`tasks/todo.md`.

## Current state (0.3.0, released)

- **VBA object model**: `Range`/`Set`/`Union`/`Areas`/`SpecialCells`, multi-area Copy/Paste
  (matching-shape only), `ActiveSheet`/`ThisWorkbook`/`ActiveWorkbook`, `With Range(...)`,
  typed `Function` params/return, `Mod`/`\`/`^`/`And`/`Or`/`Xor`/`Not` at real VBA precedence,
  comma-separated multi-declarator `Dim`, single-line `If cond Then stmt [Else stmt]`,
  `Fix`/`Sgn`/`Round`/`CBool`/`Date`/`Time`/`Now` built-in functions (0.3.0).
- **`@elixcee/xlsx`**: all 33 `utils.*` exports differential-tested against the real
  `xlsx@0.18.5` oracle (512 MATCH + 14 disclosed intentional divergences), `SSF` number
  formatting backed by the real `ssf` engine, six real security fixes ported from oracle
  defects. `XLSX.read()` is a working sync WASM bridge (Node + browser), 19/19 MATCH against
  the oracle. `read`/`readFile`/`write*` beyond `read()` are not implemented; npm publish of
  `packages/xlsx` has not happened (`0.0.0-development`).
- Published: `elixcee` 0.3.0 (crates.io, PyPI), `elixcee-types` 0.1.0 (crates.io, unchanged
  since 0.2.0), CLI binaries (GitHub Release).
- Self-assessed at 87-89/100 against the project's own scoring framework — not claimed as
  90+ because the VBA-vs-Microsoft-Excel axis has never been exercised (see below).

## Known gaps (from CHANGELOG's "Known limitations", not re-litigated here)

1. **No Microsoft Excel validation, at all.** Every VBA differential result to date is
   against LibreOffice, not Excel — and LibreOffice's own VBA layer is not a verified proxy
   for Excel's. No Windows/Excel environment has ever been available in this project's
   toolchain. This is the single largest gap blocking a 90+ claim.
2. **LibreOffice headless oracle is broken for most of the VBA corpus.** 578/581 scenarios
   are `ORACLE_UNAVAILABLE` — headless UNO hangs on any `Range`/`Cells` access. Root-caused,
   not fixed (explicitly out of scope for 2B/2C: fixing it doesn't raise elixcee's own
   product value, only this one oracle's usability).
3. ~~Comma-separated multi-declarator `Dim`~~ — **fixed** (Unreleased): `parse_dim` now loops
   over every comma-separated declarator instead of returning after the first non-built-in
   one. Corpus's own parse-error count: 8 → 4 (see item 3b, a bug this fix unmasked).
3b. ~~Single-line `If cond Then stmt` (no `End If`) doesn't parse at all~~ — **fixed**
   (Unreleased): discovered while verifying item 3's fix (the 4 corpus parse errors left
   after fixing comma-`Dim` were all this, not further `Dim` cases). Identifier-led inline
   statements are recognized (covers 100% of what the corpus actually uses here), plus
   `Exit For|Do|Sub|Function`/`GoTo <label>` handled explicitly — an early version routed
   those through the generic identifier-statement parser too, which silently turned
   `If done Then Exit Sub` into a no-op instead of exiting (caught in review, fixed before
   shipping — see CHANGELOG). Corpus's own parse-error count is 0/581, verified by
   rerunning the corpus, not just unit tests.
4. ~~`Not` is boolean-truthy, not bitwise~~ — **fixed** (Unreleased): `Not` now splits
   logical-vs-bitwise the same way `And`/`Or`/`Xor` already did — a genuine `Boolean` gets
   logical negation, anything else gets a real bitwise complement. `Not 5 And 3` now matches
   real VBA's `2`.
5. **Multi-area Paste** only executes for the matching-shape case; every other combination
   (count/shape mismatch, single↔multi either direction) stays diagnose-only.
6. **`XLSX.read()`** covers cell values/formulas/dates/dimension/hidden rows-cols/formatting
   display strings, but not `read`/`readFile` (file-path/stream entry points), `write*`, or
   non-Node browser dispatch beyond the bundled-consumption case (its shared code still has a
   CJS `require('ssf')`).
7. ~~581-scenario corpus's 41 non-parse-error failures were uncategorized~~ — **root-caused**
   (Unreleased): a pass over every failure's actual error message/category (not the vague
   "probably intentional negative scenarios" guess from the previous round) found 28/41
   were three missing built-in VBA functions (`Sgn` ×13, `Fix` ×12, `Round` ×3), 2 more were
   a related `CBool` type bug — all now fixed (see CHANGELOG). The remaining 11 are correctly
   left as failures: 8 genuine `Division by zero` (matches real VBA), 2 explicitly-named
   `unsupported_functions` scenarios (deliberate negative tests), 1 `Timer()` (nondeterministic
   category — low value to implement, arguably intentional to leave out of a deterministic
   engine). Corpus is now 570/581 elixcee-side, with every remaining failure understood and
   correct, not just uninvestigated. Done via an ad-hoc analysis pass, not a new committed
   script — the corpus is small and fixed-size enough that this didn't seem worth a permanent
   tool; revisit if the corpus grows or this needs repeating regularly.
8. ~~Two small things found while implementing item 7~~ — **fixed** (Unreleased):
   - `Round(number, negativeDigits)` now errors ("Invalid procedure call or argument"),
     matching real VBA, instead of silently returning a plausible answer.
   - `Now`/`Date`/`Time` no longer return a Rust debug-formatted `SystemTime{...}` string.
     `Date()` returns a real `Variant::Date` matching the system clock; `Time()`/`Now()`
     return a numerically correct `Variant::Float` rather than `Variant::Date`, since
     `Variant::Date` is whole-day-only (`i64`) and can't carry a sub-day component without
     a shared-type change (`TypeName(Time())`/`TypeName(Now())` report `"Double"`, not real
     VBA's `"Date"` — disclosed, not silent).
9. ~~Bare no-parens zero-arg VBA function calls (`Date` without `()`) didn't parse~~ —
   **fixed** (Unreleased): a bare identifier now falls back to calling `Date`/`Now`/`Time`
   as zero-arg functions only after every other variable/constant lookup fails — the only
   three `eval_vba_func` entries that accept zero arguments, so this doesn't generalize to
   "any unrecognized identifier might be a function call" and a genuine variable-name typo
   still errors exactly as before (pinned by a regression test).

## Next candidates, roughly by leverage

Not committed to a specific order — pick based on what the next release is trying to prove.

- **Microsoft Excel validation** (item 1) — blocked on getting a Windows+Excel environment,
  not on engineering effort; the `compat/oracle-excel-com/CONTRACT.md` adapter is already
  written and waiting. Highest-value item on this list once an environment exists.
- **LibreOffice headless hang** (item 2) — would unblock 578 currently-dead corpus scenarios,
  but was explicitly ruled out twice already as not raising elixcee's own product value.
  Worth revisiting only if the corpus itself becomes the bottleneck rather than VBA coverage.
- **`XLSX.read`/`readFile`/`write*`** — extends `@elixcee/xlsx` from "can read what B7/2C's
  read() covers" toward actual drop-in file I/O parity with SheetJS.
- **General multi-area Paste** (item 5) — object-model completeness beyond what B7c shipped.
- **`packages/xlsx` npm publish** — currently `0.0.0-development`; would need a version/scope
  decision (see `CHANGELOG.md`'s Phase 0 note on `@elixcee` npm-scope ownership being
  unconfirmed) before it's a real release candidate.

## Non-goals (still, per existing ADRs)

- No new Rust runtime dependencies beyond what's already justified in `Cargo.toml`'s
  comments — matches this codebase's long-running dependency-minimization direction
  (`docs/xlsx-architecture.md`).
- `packages/xlsx` never depends on the real `xlsx` package at runtime (ADR, same doc).
- No byte-for-byte compatibility claims where SheetJS itself is non-deterministic (embedded
  timestamps, etc.) — compatibility is measured on parsed logical shape.
