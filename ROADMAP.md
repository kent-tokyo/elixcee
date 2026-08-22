# Roadmap

This tracks where `elixcee`/`@elixcee/xlsx` stands and what's next. For what already
shipped, see `CHANGELOG.md` — this file only restates completed work when needed to explain
what's left. Historical phase-by-phase implementation notes (Japanese) live in
`tasks/todo.md`.

## Current state (0.6.0 released; unreleased work since — real multi-dimensional VBA
arrays, call-frame-scoped `On Error`/full `Err` object, compile-time undefined-call/
argument-count/label checks, and `@elixcee/xlsx` `write()`/`writeFile()`/`writeFileSync()`
— all committed locally (7 commits, `07b4def`..`f9cc239`), but not yet version-bumped,
pushed, tagged, or published; see CHANGELOG.md's `[Unreleased]` and `tasks/todo.md`'s
"Phase C"/"`@elixcee/xlsx` 0.1.0-alpha.1 準備" sections for the full detail)

- **VBA object model**: `Range`/`Set`/`Union`/`Areas`/`SpecialCells`, matching-shape
  multi-area Copy/Paste, `ActiveSheet`/`ThisWorkbook`/`ActiveWorkbook`, a **runtime** `With`
  stack (any target expression, including `With Cells(r, c)`; `.member` resolves correctly
  at any nesting depth inside `If`/`For`/`Do`/`Select Case`), real object-variable
  unset/`Nothing` state (alias-safe: `Set r2 = r` survives `Set r = Nothing`), `Variant::Null`
  with documented Null-propagation through `+`/`&`/comparisons, the `:` multi-statement
  separator, typed `Function` params/return, `Mod`/`\`/`^`/`And`/`Or`/`Xor`/`Not` at real VBA
  precedence (real bitwise semantics on non-Boolean operands), comma-separated
  multi-declarator `Dim`, single-line `If cond Then stmt [Else stmt]`. `Dim x` registers a
  real `Empty`-valued variable. Since 0.6.0, committed locally (not yet released): **real multi-dimensional
  arrays** (`Variant::VbaArray`, per-dimension bounds and row-major storage — `Dim arr(3,2)`
  no longer aliases `arr(1,1)`/`arr(1,2)`, `UBound(arr, dimension)` honors its argument for
  real, `ReDim Preserve` enforces VBA's actual last-dimension-only rule) and
  **call-frame-scoped `On Error`** (`Err.Source`/`Err.HelpFile`/`Err.HelpContext`, full
  5-argument `Err.Raise`; fixes the previously-disclosed bug where `On Error GoTo <label>`
  never fired for an error inside a handler-less callee).
- **Built-in functions**: `Fix`/`Sgn`/`Round`(banker's rounding, rejects negative digits)/
  `CBool`/`CInt`/`CLng`(also banker's rounding)/`IsNumeric`(numeric strings)/`Str`(leading-
  space quirk, distinct from `CStr`)/`Val`(leading-numeric-prefix parsing)/`Date`/`Time`/
  `Now` (real values, callable with or without parens)/`Array(...)`.
- **Static checking**: `elixcee check` and `Vm::run_sub`/`run_sub_multi`'s own pre-flight
  pass now also catch undefined-procedure calls, argument-count mismatches (`E1008`), and
  undefined `GoTo`/`On Error GoTo` labels (`E1009`) — whole-project scope, uncatchable by
  `On Error` (committed locally, not yet released).
- **Test infrastructure**: two committed, oracle-independent classifiers, distinct from the
  existing LibreOffice/Excel oracle-comparison axis (`compat/corpus/classify.mjs`):
  `compat/corpus/classify-elixcee-outcomes.mjs` explains elixcee's own pass/fail outcome
  for all 581 corpus scenarios by exact scenario ID (0 `UNEXPLAINED`, 0 `MISMATCH`); the
  `compat/vba-semantics/` suite (**386 cases**, up from 208 at 0.3.0) checks VALUE
  correctness against documented real VBA semantics, not just pass/fail (0 `BUG`,
  0 `UNCLASSIFIED`, **14** `KNOWN_LIMITATION` as of the (committed locally, not yet
  released) multi-dimensional-array work — down from 16 — see item 10 below for the full
  breakdown) — see each
  directory's own README for what it measures and doesn't. CI runs `packages/xlsx`'s
  TypeScript typecheck, all four `compat/differential/` suites plus their own self-checks,
  a real packed-npm-tarball consumer smoke, and a `wasm` job that builds
  `crates/elixcee-wasm` fresh and runs both a Node/browser-condition smoke and a **real
  headless-Chrome** smoke — see item 8 below.
- **`@elixcee/xlsx`**: all 33 `utils.*` exports differential-tested against the real
  `xlsx@0.18.5` oracle (512 MATCH + 14 disclosed intentional divergences), `SSF` number
  formatting backed by the real `ssf` engine, six real security fixes ported from oracle
  defects. `XLSX.read()`/`readFile()`/`readFileSync()` are a working sync WASM bridge (Node
  + browser), 33/33 MATCH against the oracle — the one disclosed defect (`src/reader.rs`
  trimming a `t="str"` cell's `xml:space="preserve"` text unconditionally) is fixed as of
  a later round (see CHANGELOG.md). **`write()`/`writeFile()`/`writeFileSync()` now exist
  too** (`bookType: "xlsx"` only, committed locally — pure JS/XML/ZIP generation, no Rust
  writer needed; see item 6 below), 36 MATCH + 1 disclosed (`bookType: "ods"`) in
  `compat/differential/xlsx-write.test.mjs`. `package.json`'s `description` was updated to
  match, but **`version`/`private`/`publishConfig` were deliberately left untouched**
  (still `0.0.0-development`/`private: true`/no `publishConfig` — a version bump, and the
  `private: false` + `publishConfig.access: "public"` a first publish needs, were both
  explicitly out of scope for this round) — **npm publish has not happened** (confirmed
  live: `registry.npmjs.org/@elixcee/xlsx` 404s), correctly blocked by `private: true`
  alone even before considering unconfirmed `@elixcee` scope ownership (item 9 below).
- Published: `elixcee` **0.6.0** on both crates.io and PyPI, `elixcee-types` 0.2.0 on
  crates.io (the enum-variant-addition semver bump `Variant::Null` required — see
  CHANGELOG.md's "`elixcee-types` 0.2.0" section) — all confirmed live via each registry's
  own API, not assumed from local files. The CLI GitHub Release is at `bin-v0.5.0`
  (confirmed via `gh release list`) — **0.6.0's CLI binaries have not been released**, a
  gap that has shrunk (0.4.0's binaries were never released at all; 0.5.0's were) but not
  closed. crates.io/PyPI publish and the CLI GitHub Release remain separate,
  independently-approved steps in this project's process.
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
   **Also applies to item 13's "safe round-trip" work below**: its own tests are
   structural/synthetic-fixture-only, for the same reason — no real Excel-authored file
   or working Excel/LibreOffice oracle exists to verify against yet.
2. **LibreOffice headless oracle is broken for most of the VBA corpus.** 578/581 scenarios
   are `ORACLE_UNAVAILABLE` — headless UNO hangs on any `Range`/`Cells` access. Root-caused,
   not fixed (explicitly ruled out twice already: fixing it doesn't raise elixcee's own
   product value, only this one oracle's usability — revisit only if the corpus itself
   becomes the bottleneck rather than VBA coverage). Consistent with, not contradicted by,
   item 13's own structural-only verification below.
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
   dimension/hidden rows-cols/formatting display strings; non-Node browser dispatch beyond
   the bundled-consumption case is unchanged (its shared code still has a CJS
   `require('ssf')`; `readFile`/`readFileSync` are Node-only by nature and throw
   `ELIXCEE_UNSUPPORTED_IN_BROWSER` from the browser entry point rather than faking a
   filesystem). The `src/reader.rs` defect that used to trim a `t="str"` cell's
   `xml:space="preserve"` text unconditionally (all three read entry points shared it) is
   fixed as of a later round — see CHANGELOG.md. ~~No Rust writer exists at all yet, for
   either XLSX or ODS format~~ — **`write()`/`writeFile()`/`writeFileSync()` exist now**
   (a later, committed-locally round — see CHANGELOG.md's `[Unreleased]` and
   `tasks/todo.md`'s "`@elixcee/xlsx` 0.1.0-alpha.1 準備" section), `bookType: "xlsx"`
   only, no ODS. Turned out not to need a Rust writer at all — pure JS/XML/ZIP generation,
   verified against `src/reader.rs`'s own parsing so "own write -> own read" is a
   meaningful round trip. **This finding is scoped strictly to `@elixcee/xlsx` — a
   separate, independently-versioned npm package.** It does NOT apply to the root
   `elixcee` crate's own writer (`save_xlsx_impl`, `src/lib.rs`, wired to CLI `--output`
   and PyO3's `save_workbook()`), which — as of the "safe round-trip" milestone (item 13
   below) — turned out to need real work: until then it silently discarded every original
   ZIP part it didn't parse on every save (`xl/vbaProject.bin` included), and `.xlsm`
   output declared the wrong (non-macro-enabled) content type outright. This makes the "npm/JS/WASM findings" section's speculation below
   ("write/writeFile/writeFileSync need a genuinely new Rust writer module... whether the
   `zip` crate... supports writing under wasm32") moot for the actually-chosen approach —
   left in place below as a record of what was considered, not corrected in place.
7. **`packages/xlsx` is not currently publishable, even as an alpha** — three concrete,
   verified blockers, not a vague "needs polish". One is now fixed, two remain, by
   deliberate choice this round: ~~there was no package-level `README.md`, so `npm`'s
   registry page would show only the `description` field, which opened with "Drop-in
   replacement for xlsx" without disclosing that `write*`/`readFile` are unimplemented~~ —
   **fixed**: `packages/xlsx/README.md` now exists (confirmed via `npm pack --dry-run`
   that it's actually included in the tarball, npm includes it automatically regardless of
   the `files` array), stating current scope honestly, and `description` no longer opens
   with an unqualified "drop-in replacement" claim. **`package.json`'s `"private": true`
   still hard-blocks `npm publish` outright, and `publishConfig.access: "public"` (a
   scoped package's first publish needs it, or `--access public` at publish time) is still
   unset** — both were left exactly as committed this round on purpose, per this session's
   own stop-condition discipline (no version bump/publish-prep metadata change without
   explicit approval); flipping them is a one-line-each, separate, still-pending decision.
   **No actual `npm publish` has been run** — correctly blocked by `private: true` alone,
   before even considering gap #9 below (`@elixcee` scope ownership is unconfirmed and
   unresolvable from this environment). See "npm/JS/WASM findings" below for the full
   investigation.
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

   **Also new (0.6.0 phase): a `compat-vba` CI job**, running `compat/corpus/` (581
   scenarios) and `compat/vba-semantics/` (386 cases) — previously runnable only locally
   (`compat/README.md`'s own "CI" section said so explicitly) because both need a release
   build of the elixcee CLI binary first, which `node-js` deliberately doesn't do. Verified
   live before wiring in: `cargo build --release --bin elixcee` then both suites' gate
   scripts, 0 UNEXPLAINED/0 MISMATCH (corpus) and 0 BUG/0 UNCLASSIFIED (vba-semantics, 19
   disclosed non-gating `KNOWN_LIMITATION` — see item 10 below).
9. **`@elixcee` npm scope ownership is unconfirmed** — cannot be resolved from this
   environment (`npm whoami` returns 401; no working publish credential exists locally, no
   analogous GitHub Actions secret exists yet either, unlike `CARGO_REGISTRY_TOKEN` for
   crates.io). Only the human maintainer can check this (`npm login` then `npm org ls
   elixcee`, or the npmjs.com web UI). (Corrects a stale, dangling citation this file
   previously had, pointing at a CHANGELOG.md "Phase 0 scope-ownership note" that doesn't
   actually exist in CHANGELOG.md's text — found and fixed this round.)
10. **14 `compat/vba-semantics/` `KNOWN_LIMITATION` cases** (Unreleased — suite grew from
    208 to 301 to **386** cases; full per-case list and root-cause grouping in
    `compat/vba-semantics/README.md`'s "Current state" section, raw detail in
    `compat/vba-semantics/results/report.json`). Down from 28: nine were genuinely fixed in
    the structural-semantics round — the three Null-propagation ones, the two object-variable
    unset/Nothing ones, the two `With`-target ones, the `Type mismatch` error-message one,
    and the missing `Array()` builtin — **four more in the 0.6.0 array-bounds round**
    (see CHANGELOG.md): `Dim arr(lo To hi)`, `Dim arr()` (empty parens), `Option Base 1`, and
    `Erase` on a fixed-size array, all fixed by adding a per-variable array lower-bound side
    table to the VM — and **the last two, real multi-dimensional array support, in a later
    round**: `Variant::VbaArray` (a distinct type from the existing `Variant::Array`, which
    stays exactly what it was for Range-value reads/formula-array results/record arrays) now
    carries real per-dimension bounds and row-major element storage, so `Dim arr(3, 2)`,
    `arr(2,0)`/`arr(2,1)` no longer collide, and `UBound(arr, dimension)` honors its argument
    for real. The remaining 14, by root cause: no declared/runtime type-width tracking (12 —
    `CInt`/`CLng` overflow, `Left`/`Right`/`Mid`/`Chr`/`InStr` out-of-domain arguments); no
    per-Variant stored-type tag (1 — `+` between two string-typed Variants numeric-adds
    instead of concatenating per VBA's own documented rule); a numeric-vs-string Variant
    comparison isn't unconditionally "numeric side is less" per VBA's documented rule (1 —
    deliberately not fixed, would invert the far more common numeric-string-vs-number
    magnitude comparison for every caller).
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
13. **Root-crate `save_xlsx_impl` regenerated every `.xlsx`/`.xlsm` save entirely from
    scratch — fixed for a first, deliberately narrow slice ("safe round-trip" milestone,
    see `docs/xlsx-architecture.md`'s "Root-crate writer: regenerate vs. preserve-and-merge"
    section).** Two things only: (a) general unknown-OOXML-part passthrough — any ZIP entry
    in the loaded source that this writer doesn't itself regenerate is now copied through
    byte-for-byte; (b) `xl/vbaProject.bin` preservation specifically, including a correctly
    carried-over macro-enabled `[Content_Types].xml`/`.rels` declaration (previously
    `--output foo.xlsm` silently produced a non-macro-enabled `.xlsx`-shaped file). Verified
    via `tests/xlsx_roundtrip.rs` (3 tests, hand-built synthetic fixtures — no real `.xlsm`
    exists in this repo yet, see `tests/fixtures/xlsm_roundtrip/README.md`) plus a manual CLI
    smoke test (in-place `--file foo.xlsm --output foo.xlsm` overwrite, inspected by hand).
    **Slice 2 (same item, same docs section): per-cell style-index (`s="N"`) preservation +
    `xl/styles.xml` conditional passthrough.** Passing through `xl/styles.xml` alone would
    have been pointless — the writer never emitted a cell's `s="N"` attribute at all, so
    every cell's font/fill/border formatting was lost on every save regardless of whether the
    style *definitions* survived. Both fixed together: a cell's original style index is now
    captured on read (`WorkbookSheet::raw_style_indices`, independent of the existing
    numFmtId resolution) and re-emitted unchanged on write; `xl/styles.xml` itself is now the
    source's own bytes when available, not the hardcoded minimal stylesheet. Always safe: no
    VBA statement in this VM ever mutates a cell's style (`Range.Interior.Color =`/
    `.NumberFormat =` are explicit no-ops, confirmed by existing tests of those names). Same
    3 tests in `tests/xlsx_roundtrip.rs` extended to cover this (edited-cell style survives,
    untouched-cell style survives, brand-new-cell doesn't spuriously inherit one,
    `xl/styles.xml` byte-identical) rather than new tests added.

    **Slice 3 (same item): merged ranges and hidden rows/columns now written back.** No new
    reader work — `Vm::merged_ranges`/`Vm::sheet_visibility` were already populated from
    `WorkbookSheet::merged_ranges`/`hidden_rows`/`hidden_columns` and used elsewhere in the VM,
    but `build_xlsx_sheet` never emitted `<mergeCells>` or a `<row>`/`<col>` `hidden="1"`
    attribute at all (confirmed live: grepping `src/lib.rs` for `mergeCells`/`hidden` found
    zero matches before this slice) — a pure writer-completeness gap, present independent of
    any unknown-part-passthrough concern. Both fields promoted `pub(crate)` so the writer can
    read them directly; a hidden row with no cell data now gets a synthesized empty `<row
    hidden="1"/>` (hidden-ness lives on the element itself, so an absent `<row>` reads as
    visible). Same 3 tests extended again (merge and hidden-column/row assertions added to the
    flagship test, a merge-survival assertion added to the in-place-overwrite test) rather than
    new tests added.

    Still genuinely out of scope, not a rearchitecture blocker for any of it later: named
    ranges, tables/hyperlinks/comments/data-validation/freeze-panes/print-and-page-setup
    embedded inside worksheet XML (sheets are always fully regenerated, never diffed against
    the original — a stated simplification; merges and hidden rows/columns are the two
    exceptions carved out by slice 3 above), *authoring or changing* styles from VBA (this VM
    has no such capability at all — only *preserving* an existing cell's style survived slice
    2), charts/images/external-link consistency after a structural sheet change,
    streaming/large-file handling, `.ods` passthrough, and `@elixcee/xlsx`/
    `crates/elixcee-wasm` (both untouched by this milestone, by design — see item 6 above for
    why they're a separate, unrelated codepath).

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
