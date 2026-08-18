# Roadmap

This tracks where `elixcee`/`@elixcee/xlsx` stands and what's next. For what already
shipped, see `CHANGELOG.md` — this file only restates completed work when needed to explain
what's left. Historical phase-by-phase implementation notes (Japanese) live in
`tasks/todo.md`.

## Current state (0.3.0 released; unreleased work since, not yet version-bumped)

- **VBA object model**: `Range`/`Set`/`Union`/`Areas`/`SpecialCells`, matching-shape
  multi-area Copy/Paste, `ActiveSheet`/`ThisWorkbook`/`ActiveWorkbook`, `With Range(...)`,
  typed `Function` params/return, `Mod`/`\`/`^`/`And`/`Or`/`Xor`/`Not` at real VBA precedence
  (real bitwise semantics on non-Boolean operands), comma-separated multi-declarator `Dim`,
  single-line `If cond Then stmt [Else stmt]`. `Dim x` now registers a real `Empty`-valued
  variable (was a complete no-op before this round — see "Recently fixed" below).
- **Built-in functions**: `Fix`/`Sgn`/`Round`(banker's rounding, rejects negative digits)/
  `CBool`/`CInt`/`CLng`(also banker's rounding)/`IsNumeric`(numeric strings)/`Str`(leading-
  space quirk, distinct from `CStr`)/`Val`(leading-numeric-prefix parsing)/`Date`/`Time`/
  `Now` (real values, callable with or without parens).
- **Test infrastructure**: two new committed, oracle-independent classifiers, distinct from
  the existing LibreOffice/Excel oracle-comparison axis (`compat/corpus/classify.mjs`):
  `compat/corpus/classify-elixcee-outcomes.mjs` explains elixcee's own pass/fail outcome
  for all 581 corpus scenarios by exact scenario ID (0 `UNEXPLAINED`, 0 `MISMATCH`); the new
  `compat/vba-semantics/` suite (208 cases) checks VALUE correctness against documented real
  VBA semantics, not just pass/fail (0 `BUG`, 0 `UNCLASSIFIED`, 0 `KNOWN_LIMITATION` — its
  first run found one disclosed gap, fixed in the same round rather than left registered) —
  see each directory's own README for what it measures and doesn't.
  CI now also runs `packages/xlsx`'s TypeScript typecheck and all four `compat/differential/`
  suites on a Node 20/22 matrix (`.github/workflows/ci.yml`'s new `node-js` job) — previously
  none of this ran anywhere except a developer's own machine.
- **`@elixcee/xlsx`**: all 33 `utils.*` exports differential-tested against the real
  `xlsx@0.18.5` oracle (512 MATCH + 14 disclosed intentional divergences), `SSF` number
  formatting backed by the real `ssf` engine, six real security fixes ported from oracle
  defects. `XLSX.read()` is a working sync WASM bridge (Node + browser), 19/19 MATCH against
  the oracle. `read`/`readFile`/`write*` beyond `read()` are not implemented; npm publish of
  `packages/xlsx` has not happened (`0.0.0-development`, currently **not publishable as-is**
  — see "npm/JS/WASM findings" below).
- Published: `elixcee` 0.3.0 (crates.io, PyPI), `elixcee-types` 0.1.0 (crates.io, unchanged
  since 0.2.0), CLI binaries (GitHub Release).
- Self-assessed at 87-89/100 against the project's own scoring framework as of 0.3.0's
  release — not re-scored here (this file doesn't set that number; see CHANGELOG.md history
  for how it's been assigned each round) — not claimed as 90+ because the VBA-vs-Microsoft-
  Excel axis has never been exercised (see "Known gaps" below).

## Recently fixed (this round — full detail and evidence in CHANGELOG.md's `[Unreleased]`)

Comma-separated `Dim`; single-line `If`/`Else` (plus two safety gaps a review caught before
shipping: `Exit`/`GoTo` inside it, and comma-`Dim`'s trailing-syntax tolerance); `Not`
bitwise semantics; `Fix`/`Sgn`/`Round`/`CBool` (root-caused via an automated pass over the
581-scenario corpus's non-parse-error failures); `Round`'s negative-digit rejection;
`Date`/`Time`/`Now`'s real values and no-parens calling; `CInt`/`CLng` banker's rounding;
`IsNumeric` numeric-string recognition; `Str()` vs `CStr()`'s leading-space distinction;
`Val()`'s leading-numeric-prefix parsing; `Dim`'s `Empty`-variable registration (found by
the new `compat/vba-semantics/` suite's first run). The last several were found by two
different systematic methods — an `eval_vba_func` source-code audit, and (once it existed)
the new suite itself — rather than one-off bug reports, and that audit is itself now
recorded as exhausted: no further candidates were found by either method as of this round.

## Known gaps

1. **No Microsoft Excel validation, at all.** Every VBA differential result to date is
   against LibreOffice, not Excel — and LibreOffice's own VBA layer is not a verified proxy
   for Excel's. No Windows/Excel environment has ever been available in this project's
   toolchain. This is the single largest gap blocking a 90+ claim. The
   `compat/oracle-excel-com/CONTRACT.md` adapter is written and waiting for one.
2. **LibreOffice headless oracle is broken for most of the VBA corpus.** 578/581 scenarios
   are `ORACLE_UNAVAILABLE` — headless UNO hangs on any `Range`/`Cells` access. Root-caused,
   not fixed (explicitly ruled out twice already: fixing it doesn't raise elixcee's own
   product value, only this one oracle's usability — revisit only if the corpus itself
   becomes the bottleneck rather than VBA coverage).
3. **Multi-area Paste** only executes for the matching-shape case; every other combination
   (count/shape mismatch, single↔multi either direction) stays diagnose-only. Extending this
   correctly needs a real oracle to verify against (LibreOffice's is broken, Excel's doesn't
   exist here) — implementing more without one risks guessing at real Excel Paste semantics,
   against this project's own stated epistemics.
4. ~~`Array` out-of-bounds error message text didn't match real VBA's exact wording~~ —
   **fixed** (Unreleased): was `"Array 'arr': index N out of bounds (len=N)"`, now real
   VBA's own `"Subscript out of range"` verbatim, matching this codebase's existing
   convention for other runtime error messages (e.g. `"Division by zero"` carries no extra
   detail either). Safe to change: confirmed via `docs/agent-contract.md`'s own documented
   policy that `message` is free text, not a stable/matchable field (`code`/`kind` are), and
   that `diagnose`/`diagnose-workbook` already read the rich per-failure detail (array name,
   index, bounds) from a structured `ResolutionFailureKind` side channel, not by parsing
   this string. `compat/vba-semantics/`'s one `KNOWN_LIMITATION` is now 0.
5. **`Time()`/`Now()` report `TypeName` `"Double"`, not real VBA's `"Date"`.** `Variant::Date`
   is whole-day-only (`i64`) in this codebase and can't carry a sub-day component without a
   structural, shared-type change (`elixcee-types`' public enum, semver-relevant). Design
   completed, not yet implemented — see `docs/date-time-runtime-model-adr.md` and "Date/Time
   runtime model" below.
6. **`XLSX.read()`** covers cell values/formulas/dates/dimension/hidden rows-cols/formatting
   display strings, but not `read`/`readFile` (file-path/stream entry points), `write*`, or
   non-Node browser dispatch beyond the bundled-consumption case (its shared code still has a
   CJS `require('ssf')`). No Rust writer exists at all yet, for either XLSX or ODS format.
7. **`packages/xlsx` is not currently publishable, even as an alpha** — three concrete,
   verified blockers, not a vague "needs polish". One is now fixed (Unreleased):
   ~~there was no package-level `README.md`, so `npm`'s registry page would show only the
   `description` field, which opened with "Drop-in replacement for xlsx" without disclosing
   that `write*`/`readFile` are unimplemented~~ — `packages/xlsx/README.md` now exists
   (confirmed via `npm pack --dry-run` that it's actually included in the tarball, npm
   includes it automatically regardless of the `files` array), stating current scope
   honestly, and `description` no longer opens with an unqualified "drop-in replacement"
   claim. **Two blockers remain, both deliberately left alone — they're a real "should this
   become publishable" policy stance, not a mechanical fix**: `package.json`'s
   `"private": true` hard-blocks `npm publish` outright; first publish of a scoped package
   also needs `--access public` or `publishConfig.access: "public"`, neither set. See
   "npm/JS/WASM findings" below for the full investigation.
8. ~~No Node/WASM/JS testing wired into CI at all~~ — **fixed** (Unreleased):
   `.github/workflows/ci.yml` gained a `node-js` job (Node 20/22 matrix) running
   `packages/xlsx`'s TypeScript typecheck (with and without the DOM lib), all four of
   `compat/`'s differential suites (`utils`/`ssf-format`/`read`/`metadata`), and
   `packages/xlsx/scripts/audit-pack-contents.mjs` (asserts every file `npm pack --dry-run`
   would actually publish: required files present, nothing forbidden, nothing unexpected
   under `src/internal/`); and a separate `wasm` job that builds `crates/elixcee-wasm`
   fresh (both `wasm-pack --target nodejs`/`--target web`, not just trusting the
   already-vendored copy the `node-js` job consumes) and runs
   `packages/xlsx/scripts/wasm-smoke.mjs` — Node sync `read()`, the `"browser"` export
   condition actually resolving *and running* (via `node --conditions=browser`, self-
   referencing the package by name — genuinely more than a resolution check, but this is
   still Node simulating the condition, not a real browser; no real browser executes
   anywhere in this project's CI, and Safari support is not claimed anywhere), a minimal
   `esbuild` bundle with `XLSX.read()` called from inside it, and the current WASM binary
   size logged (263,204 bytes as of this round — recorded, not gated against any
   threshold; no prior baseline exists to compare against, and gating without one would be
   exactly the kind of unjustified threshold this project avoids elsewhere). Every command
   verified working live before wiring either job in, not assumed from CHANGELOG.md's
   claimed numbers. One real consumer caveat found while writing the bundle check, not
   previously disclosed: the Node/CJS WASM loader (`elixcee_wasm.node.cjs`, wasm-pack's own
   generated code) locates its `.wasm` file via a `__dirname`-relative path, which is
   bundle-output-relative once bundled — a consumer bundling this package's Node entry must
   bundle to CJS (ESM bundle output has no `__dirname` at all — a `ReferenceError`, not a
   silent failure) and copy `elixcee_wasm_bg.wasm` next to their bundle output, or
   externalize the loader instead. Not fixed this round (would mean patching wasm-pack's
   own generated boilerplate); `wasm-smoke.mjs`'s header comment has the full detail.
   Genuinely still not built (a separate, bigger undertaking — see "npm/JS/WASM findings"
   below): real in-browser execution in CI (would need a headless-browser runner, e.g.
   Playwright — not added, matching "no large browser-test framework"); a WASM size
   *regression* check (the size is recorded now, but nothing fails CI if it grows).
9. **`@elixcee` npm scope ownership is unconfirmed** — cannot be resolved from this
   environment (`npm whoami` returns 401; no working publish credential exists locally, no
   analogous GitHub Actions secret exists yet either, unlike `CARGO_REGISTRY_TOKEN` for
   crates.io). Only the human maintainer can check this (`npm login` then `npm org ls
   elixcee`, or the npmjs.com web UI). (Corrects a stale, dangling citation this file
   previously had, pointing at a CHANGELOG.md "Phase 0 scope-ownership note" that doesn't
   actually exist in CHANGELOG.md's text — found and fixed this round.)
10. **28 `compat/vba-semantics/` `KNOWN_LIMITATION` cases** (Unreleased — suite grew from
    208 to 301 cases; full per-case list and root-cause grouping in
    `compat/vba-semantics/README.md`'s "Current state" section, raw detail in
    `compat/vba-semantics/results/report.json`), none fixed this round. By root cause:
    no declared/runtime type-width tracking (12 — `CInt`/`CLng` overflow, `Left`/`Right`/
    `Mid`/`Chr`/`InStr` out-of-domain arguments); array declaration/resize gaps (6 —
    `Dim arr(lo To hi)`/`Dim arr()` don't parse, `Option Base 1` ignored, `UBound`'s
    dimension argument ignored, `Erase` doesn't reset elements, `Array()` unimplemented); no
    Null-propagation semantics in the VM (3 — `+`/`&`-both-Null/comparisons all coerce Null
    to 0 or `""` instead of producing Null); no object-variable unset/Nothing state (2 — a
    never-`Set` variable's member access silently no-ops, `Set x = Nothing` doesn't clear
    the reference); `With`-target resolution is parse-time-textual, not a runtime stack (2 —
    can't target `Cells(...)`, can't reach a `.member` nested inside another block
    construct); no per-Variant stored-type tag (1 — `+` between two string-typed Variants
    numeric-adds instead of concatenating per VBA's own documented rule); a numeric-vs-
    string Variant comparison isn't unconditionally "numeric side is less" per VBA's
    documented rule (1 — deliberately not fixed, would invert the far more common
    numeric-string-vs-number magnitude comparison for every caller); one shared numeric-
    coercion helper's error message text doesn't match VBA's "Type mismatch" wording, though
    the error condition itself is already correct (1 — not fixed, the helper has ~54 call
    sites and a blind rename risks wrong wording elsewhere).
11. **Two parser gaps found alongside the vba-semantics work, not in that suite** (each is a
    "does it parse at all" question, not a value-correctness one): the `:` multi-statement-
    per-line separator doesn't parse at all (`a = 1: b = 2` fails with "expected newline, got
    Colon") — real VBA allows an arbitrary number of statements per line separated by `:`;
    and the two `With`-target gaps above (item 10) are confirmed only for the specific shapes
    tested (`Cells(...)` as a target, one level of `If` nesting), not exhaustively
    characterized across every possible target expression or nesting construct.

## npm/JS/WASM findings (from a dedicated investigation this round — see git history for the
full report; this is a summary)

Investigated, not implemented or published: CI coverage, npm scope ownership, and
`packages/xlsx` alpha-release readiness.

- **Now wired into `.github/workflows/ci.yml`'s `node-js` job** (each verified live before
  wiring, not assumed): `compat/differential/`'s utils (512 MATCH + 14 divergences)/
  SSF (1831/1831)/read (19/19)/metadata (34/34) suites; `packages/xlsx`'s TypeScript
  typecheck, both with and without the DOM lib present; a new `npm pack` content-audit
  script (`packages/xlsx/scripts/audit-pack-contents.mjs` — required files present, nothing
  forbidden, nothing unexpected under `src/internal/`, checked against `npm pack --dry-run
  --json`'s own file list, currently 17 files / 338.8 kB packed). CJS↔ESM export identity is
  asserted as part of the metadata suite (`metadata.test.mjs`), so it's covered too.
- **Also now wired, in a separate new `wasm` job** (both `wasm-pack` targets built fresh,
  esbuild added as `packages/xlsx`'s one new devDependency): a Node sync `read()` smoke
  test, the `"browser"` export condition resolving *and running* under
  `node --conditions=browser` (Node simulating the condition — no real browser executes in
  this project's CI, no Safari claim anywhere), a minimal esbuild bundle with an in-bundle
  `XLSX.read()` call, and the WASM binary size logged (263.2 kB as of this round, not
  gated). See item 8 above for the full detail, including a real consumer caveat found
  while building the bundle check (Node/CJS WASM loader's `__dirname`-relative `.wasm`
  lookup needs the file copied alongside a bundled consumer's own output).
- **Still doesn't exist at all**: real in-browser execution in CI (would need a headless-
  browser runner — not added, matching "no large browser-test framework"); a WASM binary
  size *regression* check (the size is recorded now; nothing fails CI if it grows) — a real
  policy call (what threshold, what to do when a legitimate feature grows it), deliberately
  not attempted this round.
- **0.2.0-alpha.1 (read+write) scope, if ever pursued**: `readFile` is near-free (pure
  WASM-bridge wiring onto the already-working `read_workbook_from_bytes`). `write`/
  `writeFile`/`writeFileSync` need a genuinely new Rust writer module — none exists for
  either XLSX or ODS today (confirmed by grep: no `write_workbook`/`writer.rs` anywhere in
  `src/`). One concrete risk to check *before* building, not assume: whether the `zip` crate
  feature set already trimmed for `wasm32-unknown-unknown` compatibility (`deflate`-only, no
  `zstd`) even supports *writing* under that target, not just reading.
- **`check-versions.sh` has no awareness of `packages/xlsx/package.json`'s own version** —
  only reconciles root `Cargo.toml` vs `pyproject.toml`. `0.0.0-development` could drift
  silently relative to a real release version with no CI signal.

## Date/Time runtime model — designed, not implemented

`Variant::Date(i64)` is whole-day-only, a structural reason `Time()`/`Now()` can't report
`TypeName` `"Date"` (see "Known gaps" #5). Fixing this properly touches `elixcee-types`'
public enum (semver-relevant: would be `elixcee-types` 0.2.0, `elixcee` 0.4.0-shaped, not a
patch). Full comparison, grounded in verified facts about the current codebase (not just
the option list), lives in `docs/date-time-runtime-model-adr.md` — three options compared
(A: change `Variant::Date(i64)` to `Variant::DateSerial(f64)`, breaking; B: keep `Date`,
add an additive `Variant::DateTime(f64)`; C: an internal-only representation never exposed
through `Variant` — shown not to actually solve the problem, since `Now()`'s return value
must be a real `Variant` to be assignable to a VBA variable at all). **Recommendation: B**
— same `elixcee-types` 0.2.0 version-bump cost as A, but far less code churn and zero
observable-behavior change for any value that's `Variant::Date` today. Not implemented —
this is a design document awaiting a decision, not a task in progress.

## Non-goals (still, per existing ADRs)

- No new Rust runtime dependencies beyond what's already justified in `Cargo.toml`'s
  comments — matches this codebase's long-running dependency-minimization direction
  (`docs/xlsx-architecture.md`).
- `packages/xlsx` never depends on the real `xlsx` package at runtime (ADR, same doc).
- No byte-for-byte compatibility claims where SheetJS itself is non-deterministic (embedded
  timestamps, etc.) — compatibility is measured on parsed logical shape.
