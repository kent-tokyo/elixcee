# Roadmap

This tracks where `elixcee`/`@elixcee/xlsx` stands and what's next. For what already
shipped, see `CHANGELOG.md` — this file only restates completed work when needed to explain
what's left. Historical phase-by-phase implementation notes (Japanese) live in
`tasks/todo.md`.

## Current state (0.4.0 released; unreleased work since — VBA structural semantics +
`@elixcee/xlsx` consumer/browser validation — not yet version-bumped)

- **VBA object model**: `Range`/`Set`/`Union`/`Areas`/`SpecialCells`, matching-shape
  multi-area Copy/Paste, `ActiveSheet`/`ThisWorkbook`/`ActiveWorkbook`, a **runtime** `With`
  stack (any target expression, including `With Cells(r, c)`; `.member` resolves correctly
  at any nesting depth inside `If`/`For`/`Do`/`Select Case`), real object-variable
  unset/`Nothing` state (alias-safe: `Set r2 = r` survives `Set r = Nothing`), `Variant::Null`
  with documented Null-propagation through `+`/`&`/comparisons, the `:` multi-statement
  separator, typed `Function` params/return, `Mod`/`\`/`^`/`And`/`Or`/`Xor`/`Not` at real VBA
  precedence (real bitwise semantics on non-Boolean operands), comma-separated
  multi-declarator `Dim`, single-line `If cond Then stmt [Else stmt]`. `Dim x` registers a
  real `Empty`-valued variable. See CHANGELOG.md's `[Unreleased]` "VBA structural semantics"
  section for the full detail on the four newest additions.
- **Built-in functions**: `Fix`/`Sgn`/`Round`(banker's rounding, rejects negative digits)/
  `CBool`/`CInt`/`CLng`(also banker's rounding)/`IsNumeric`(numeric strings)/`Str`(leading-
  space quirk, distinct from `CStr`)/`Val`(leading-numeric-prefix parsing)/`Date`/`Time`/
  `Now` (real values, callable with or without parens)/`Array(...)`.
- **Test infrastructure**: two committed, oracle-independent classifiers, distinct from the
  existing LibreOffice/Excel oracle-comparison axis (`compat/corpus/classify.mjs`):
  `compat/corpus/classify-elixcee-outcomes.mjs` explains elixcee's own pass/fail outcome
  for all 581 corpus scenarios by exact scenario ID (0 `UNEXPLAINED`, 0 `MISMATCH`); the
  `compat/vba-semantics/` suite (**386 cases**, up from 208 at 0.3.0) checks VALUE
  correctness against documented real VBA semantics, not just pass/fail (0 `BUG`,
  0 `UNCLASSIFIED`, 19 `KNOWN_LIMITATION` — see item 10 below for the full breakdown) — see
  each directory's own README for what it measures and doesn't. CI runs `packages/xlsx`'s
  TypeScript typecheck, all four `compat/differential/` suites plus their own self-checks,
  a real packed-npm-tarball consumer smoke, and a `wasm` job that builds
  `crates/elixcee-wasm` fresh and runs both a Node/browser-condition smoke and a **real
  headless-Chrome** smoke — see item 8 below.
- **`@elixcee/xlsx`**: all 33 `utils.*` exports differential-tested against the real
  `xlsx@0.18.5` oracle (512 MATCH + 14 disclosed intentional divergences), `SSF` number
  formatting backed by the real `ssf` engine, six real security fixes ported from oracle
  defects. `XLSX.read()`/`readFile()`/`readFileSync()` are a working sync WASM bridge (Node
  + browser), 30 MATCH + 3 disclosed (one root cause, `src/reader.rs`'s `xml:space`
  handling — see CHANGELOG.md) against the oracle. `write*` remains unimplemented; npm
  publish of `packages/xlsx` has not happened (`0.0.0-development`, currently **not
  publishable as-is** — see "npm/JS/WASM findings" below).
- Published: `elixcee` 0.5.0 and `elixcee-types` 0.2.0 on crates.io (the enum-variant-
  addition semver bump `Variant::Null` required — see CHANGELOG.md's "`elixcee-types`
  0.2.0" section). **PyPI (`elixcee` 0.5.0) and the CLI GitHub Release (`bin-v0.5.0`) are
  not yet done as of this note** — crates.io publish and PyPI/CLI release are separate,
  independently-approved steps in this project's process; PyPI still serves `elixcee`
  0.4.0, the CLI GitHub Release still serves `bin-v0.3.0` (0.4.0's CLI binaries were never
  released at all — a pre-existing gap, not caused by this round, not yet resolved).
- Not re-scored in this file (see CHANGELOG.md history for how the project's own scoring
  framework has been applied each round) — not claimed as validated against Microsoft Excel
  itself anywhere, because the VBA-vs-Excel axis has never been exercised (see "Known gaps"
  below); this round's own work is explicit that it doesn't change that.

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
6. **`XLSX.read()`/`readFile()`/`readFileSync()`** cover cell values/formulas/dates/
   dimension/hidden rows-cols/formatting display strings, but not `write*`, or non-Node
   browser dispatch beyond the bundled-consumption case (its shared code still has a CJS
   `require('ssf')`; `readFile`/`readFileSync` are Node-only by nature and throw
   `ELIXCEE_UNSUPPORTED_IN_BROWSER` from the browser entry point rather than faking a
   filesystem). No Rust writer exists at all yet, for either XLSX or ODS format. Also: all
   three read entry points share `src/reader.rs`'s defect trimming `xml:space="preserve"`
   whitespace (see CHANGELOG.md's `[Unreleased]`) — disclosed via
   `compat/differential/classify.mjs`'s `UNSUPPORTED_ALLOWLIST`, not fixed.
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
8. ~~No Node/WASM/JS testing wired into CI at all~~ — **fixed**, including real-browser
   coverage as of the structural-semantics/consumer-validation round (Unreleased):
   `.github/workflows/ci.yml`'s `node-js` job (Node 20/22 matrix) runs `packages/xlsx`'s
   TypeScript typecheck (with and without the DOM lib), all four of `compat/`'s
   differential suites (`utils`/`ssf-format`/`read`+`readFile`/`metadata`),
   `packages/xlsx/scripts/audit-pack-contents.mjs` (asserts every file `npm pack --dry-run`
   would actually publish), `compat/differential/`'s `classify.mjs`/`normalize.mjs`
   self-checks (existing scripts, never wired into CI before now), and a real
   packed-tarball consumer smoke (`npm run pack:consumer` — a genuine `npm pack` + `npm
   install` into a throwaway `node_modules`, not a relative-path shortcut). The `wasm` job
   builds `crates/elixcee-wasm` fresh (both `wasm-pack --target nodejs`/`--target web`) and
   runs `packages/xlsx/scripts/wasm-smoke.mjs` (Node sync `read()`; the `"browser"` export
   condition resolving *and running* under `node --conditions=browser` — still Node
   simulating the condition, not a real browser, and labelled as such everywhere; CJS *and*
   ESM esbuild bundles, each with `XLSX.read()` called from inside; WASM size logged, not
   gated) plus `packages/xlsx/scripts/browser-smoke.mjs` — **a real headless Chrome/Chromium
   process**, launched via Chrome's own `--dump-dom` (no browser-driver dependency added),
   serving an esbuild bundle over real HTTP and reading `XLSX.read()`'s result back out of
   the page's own DOM. This is genuinely distinct from the `--conditions=browser` check
   above and is never described using that check's language. Safari is not covered and not
   claimed anywhere. Every command verified working live before wiring either job in.

   **The `__dirname`-relative `.wasm`-lookup consumer caveat (disclosed above as "not fixed
   this round") is now fixed**: the Node/CJS WASM loader inlines its compiled WASM as
   base64, mirroring the technique the browser loader already used
   (`crates/elixcee-wasm/build-node-inline.mjs`, generated by `build.sh`, never hand-patched
   — a fresh rebuild reproduces the committed artifact byte-for-byte). No `.wasm`-copy step
   is required for CJS or ESM bundling anymore; browser bundling, previously broken outright
   (`esbuild --platform=browser` failed resolving `fs`), now works too. Package-size impact
   versus 0.4.0, measured not guessed: packed tarball 339,098 → 380,005 bytes (+12.1%),
   unpacked +12.7%, WASM payload itself unchanged at 263 KB (only its base64 containers
   grew — no `.wasm` file is vendored raw anymore, avoiding double-shipping the same bytes).
   See CHANGELOG.md's `[Unreleased]` for the full writeup, including why options B/C/D
   (a stable wrapper, a `bundler` export condition, an externalize-and-document approach)
   were considered and rejected in favor of inlining.

   Still genuinely not built: a WASM size *regression* check (the size is recorded now, but
   nothing fails CI if it grows — a real policy call on what threshold and what to do when a
   legitimate feature grows it, deliberately not attempted).

   **Also new: a `fuzz` CI job**, wired in after `fuzz/`'s 4 libFuzzer targets — which had
   no CI signal at all, and whose `fuzz/Cargo.lock` had silently gone stale since elixcee
   v0.1.2 — were actually run for the first time and immediately found a real crash (an
   i64-overflow panic in the VBA tokenizer, fixed; see `tasks/todo.md`'s `2026-08-20`
   session entry and `CHANGELOG.md`). Runs each target for a fixed 30s smoke budget per
   push/PR, not a fuzzing campaign; does not persist a corpus across runs (a real design
   question — where it would live, how it'd be curated — left open, not assumed).
9. **`@elixcee` npm scope ownership is unconfirmed** — cannot be resolved from this
   environment (`npm whoami` returns 401; no working publish credential exists locally, no
   analogous GitHub Actions secret exists yet either, unlike `CARGO_REGISTRY_TOKEN` for
   crates.io). Only the human maintainer can check this (`npm login` then `npm org ls
   elixcee`, or the npmjs.com web UI). (Corrects a stale, dangling citation this file
   previously had, pointing at a CHANGELOG.md "Phase 0 scope-ownership note" that doesn't
   actually exist in CHANGELOG.md's text — found and fixed this round.)
10. **19 `compat/vba-semantics/` `KNOWN_LIMITATION` cases** (Unreleased — suite grew from
    208 to 301 to **386** cases; full per-case list and root-cause grouping in
    `compat/vba-semantics/README.md`'s "Current state" section, raw detail in
    `compat/vba-semantics/results/report.json`). Down from 28: **nine were genuinely fixed
    this round** (see CHANGELOG.md) — the three Null-propagation ones, the two
    object-variable unset/Nothing ones, the two `With`-target ones, the `Type mismatch`
    error-message one, and the missing `Array()` builtin. The remaining 19, by root cause:
    no declared/runtime type-width tracking (12 — `CInt`/`CLng` overflow, `Left`/`Right`/
    `Mid`/`Chr`/`InStr` out-of-domain arguments); array declaration/resize gaps (5 —
    `Dim arr(lo To hi)`/`Dim arr()` don't parse, `Option Base 1` ignored, `UBound`'s
    dimension argument ignored, `Erase` doesn't reset elements); no per-Variant stored-type
    tag (1 — `+` between two string-typed Variants numeric-adds instead of concatenating per
    VBA's own documented rule); a numeric-vs-string Variant comparison isn't unconditionally
    "numeric side is less" per VBA's documented rule (1 — deliberately not fixed, would
    invert the far more common numeric-string-vs-number magnitude comparison for every
    caller).
11. **`Range.Range(...)`/`Range.Cells(...)` are not relative to the base range.** Inside a
    `With <range>` body (and through a `Set`-assigned Range variable), a `.Range("A1")`/
    `.Cells(r, c)` qualifier resolves as an independent, absolute reference on the active
    sheet. Real VBA resolves both relative to the base range's upper-left corner. Pre-
    existing behavior, pinned by `with_range_nested_range_reference_still_works`;
    deliberately left unchanged by the runtime-With-stack work (which was about *where* a
    `.member` can appear and *when* the target is evaluated, not about re-anchoring this
    qualifier). Not covered by `compat/vba-semantics/` — the test shapes there all happen to
    make relative and absolute agree.
12. **The `:` statement separator's interaction with unparsed lines.** `:` is now a real
    statement separator everywhere (see CHANGELOG.md), but a line elixcee skips wholesale as
    an unrecognized *block header* (`skip_to_eol` — a `With <unmodeled>` header, an `Option`
    line) still swallows the rest of that physical line, colons included. Statement-level
    skipping (`skip_to_stmt_end`) correctly stops at a `:`. No known real-world macro hits
    the difference; recorded so a future reader doesn't rediscover it as a surprise.

## npm/JS/WASM findings (from a dedicated investigation this round — see git history for the
full report; this is a summary)

Investigated, not implemented or published: CI coverage, npm scope ownership, and
`packages/xlsx` alpha-release readiness.

- **Wired into `.github/workflows/ci.yml`'s `node-js` job** (each verified live before
  wiring, not assumed): `compat/differential/`'s utils (512 MATCH + 14 divergences)/
  SSF (1831/1831)/read+readFile (30 MATCH + 3 disclosed)/metadata (36/36) suites, plus
  `classify.mjs`/`normalize.mjs`'s own self-checks; `packages/xlsx`'s TypeScript typecheck,
  both with and without the DOM lib present; the `npm pack` content-audit script
  (`packages/xlsx/scripts/audit-pack-contents.mjs`); and a **real packed-tarball consumer
  smoke** (`npm run pack:consumer` — genuine `npm pack` + `npm install` into a throwaway
  `node_modules` under `os.tmpdir()`, exercising `require`/`import`/TypeScript/`read()`/
  the `browser` condition entirely from inside that install, never a relative-path
  shortcut back into this repo). CJS↔ESM export identity is asserted as part of the
  metadata suite, so it's covered too.
- **Wired into a separate `wasm` job** (both `wasm-pack` targets built fresh; `esbuild`
  remains `packages/xlsx`'s only added devDependency across both rounds, zero new
  dependencies added for the browser-smoke work below): a Node sync `read()` smoke test,
  the `"browser"` export condition resolving *and running* under `node --conditions=browser`
  (Node simulating the condition — still labelled as such, not claimed as browser coverage
  on its own), CJS *and* ESM esbuild bundles each with an in-bundle `XLSX.read()` call, the
  WASM binary size logged (not gated), and — new this round — **a real headless
  Chrome/Chromium smoke test** (`npm run browser:smoke`, via Chrome's own `--dump-dom`, no
  browser-driver dependency): an actual browser process loads a real bundle over real HTTP
  and the result is read back out of the page's own DOM. This is genuinely distinct from
  the `--conditions=browser` check and is never described using its language anywhere. No
  Safari claim anywhere. The `__dirname`-relative `.wasm`-lookup caveat mentioned in earlier
  drafts of this file is now fixed (base64-inlined into the Node loader too, mirroring the
  browser loader) — see item 8 above and CHANGELOG.md for the full writeup and the measured
  package-size impact (+12.1% packed, WASM payload itself unchanged).
- **Still doesn't exist at all**: a WASM binary size *regression* check (the size is
  recorded now; nothing fails CI if it grows) — a real policy call (what threshold, what to
  do when a legitimate feature grows it), deliberately not attempted this round.
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
