# Roadmap

This tracks where `elixcee`/`@elixcee/xlsx` stands and what's next. For what already
shipped, see `CHANGELOG.md` — this file only restates completed work when needed to explain
what's left. Historical phase-by-phase implementation notes (Japanese) live in
`tasks/todo.md`.

## Current state

`elixcee` **0.8.0** (Rust crate + Python package), `elixcee-types` **0.3.0**, `elixcee-wasm`
**0.1.0** (never published to crates.io by design — `publish = false`) — all confirmed live
via crates.io's/PyPI's own APIs, not assumed from local files. `bin-v0.8.0`'s GitHub Release
carries all three CLI platform binaries (macOS aarch64, Windows x86_64, Linux x86_64).
`@elixcee/xlsx` is unchanged: `read()`/`readFile()`/`readFileSync()` and
`write()`/`writeFile()`/`writeFileSync()` are both implemented and differential-tested, but
the package is still `0.0.0-development`/`private: true`/unpublished — no `npm publish` has
happened (confirmed live: `registry.npmjs.org/@elixcee/xlsx` 404s), and `@elixcee` scope
ownership itself is unconfirmed (item 9 below).

**0.7.0** shipped three VBA-runtime items: real multi-dimensional arrays (`Variant::VbaArray`,
per-dimension bounds and row-major storage — `Dim arr(3,2)` no longer aliases `arr(1,1)`/
`arr(1,2)`, `UBound(arr, dimension)` honors its argument for real, `ReDim Preserve` enforces
VBA's actual last-dimension-only rule), call-frame-scoped `On Error` with a full `Err` object
(`Err.Source`/`Err.HelpFile`/`Err.HelpContext`, 5-argument `Err.Raise`), and compile-time
undefined-procedure-call/argument-count/`GoTo`-label checks (`E1008`/`E1009`, uncatchable by
`On Error`, whole-project scope). Full detail in `CHANGELOG.md`'s `[0.7.0]`.

**0.8.0** shipped the first three slices of a new direction: "safe round-trip" — read an
existing workbook, run/modify it, and write it back without destroying what elixcee doesn't
understand. Root-crate `save_xlsx_impl` (CLI `--output`, PyO3 `save_workbook()`) used to
discard the entire original file and regenerate a brand-new minimal workbook from scratch on
every save; `--output foo.xlsm` silently produced a non-macro-enabled `.xlsx`-shaped file.
Fixed in three slices: (1) general unknown-OOXML-part passthrough plus `xl/vbaProject.bin`
preservation with a correctly carried-over macro-enabled `[Content_Types].xml`/`.rels`
declaration; (2) per-cell style-index (`s="N"`) preservation plus `xl/styles.xml` conditional
passthrough, so a cell's font/fill/border/number-format formatting survives a value edit; (3)
merged ranges and hidden rows/columns, previously captured on read but never written back at
all, now correctly re-emitted. Full detail, including the exact writer-owned-vs-passthrough
split and what's still explicitly out of scope, in `docs/xlsx-architecture.md`'s "Root-crate
writer: regenerate vs. preserve-and-merge" section and `CHANGELOG.md`'s `[0.8.0]`.

Test suite as of `0.8.0`: `cargo test --workspace` 955/955 (up from 872 at `0.6.0`),
`compat/vba-semantics` 386 cases (0 `BUG`, 0 `UNCLASSIFIED`, 14 `KNOWN_LIMITATION` — see item
10 below), `compat/corpus` 581 scenarios (0 `UNEXPLAINED`, 0 `MISMATCH`), every GitHub Actions
job green on `master`.

**`0.9.0`'s real-Excel validation is underway (`0.9.0-A`, in progress — not yet released as
`0.9.0`)**: 5 real Microsoft-Excel-for-Mac-authored `.xlsm` fixtures (values/formula/style/
merge/hidden rows-cols; VBA project + macro; table/data validation/conditional formatting;
hyperlink/comment/defined name; chart/image/print area), each edited via elixcee, saved both
ways (save-as and in-place), and reopened in real Excel — 0 repair warnings, 0 `vbaProject`
loss, 0 relationship breakage, 0 in-place-save failures, across all 5. Found and fixed three
real bugs the synthetic fixtures never exercised (formula flattening, orphaned relationships,
wrong `.xlsm` content type for a non-macro workbook — the last one made Excel refuse to open
the file outright). See `CHANGELOG.md`'s `[Unreleased]` and
`compat/oracle-excel-com/results/0.9.0-A_summary.md` for full detail, including two open items
neither fixed nor newly discovered: worksheet-embedded features (tables/validation/
conditional-formatting/hyperlinks/defined-names/charts/images) are silently dropped on any
save — already disclosed in `0.8.0`'s own Non-goals, confirmed live here, in `0.10.0`'s scope
to fix, not `0.9.0`'s — and macro *execution* verification is blocked by a Mac Excel/VBA
environment issue unrelated to elixcee. Not yet done: the 10-consecutive-cycle exit criterion
(only 1 cycle per fixture per save mode so far) and the VBA-semantics-vs-Excel axis (`0.9.0-B`,
paused — see the roadmap below).

## Known gaps

1. **No Microsoft Excel validation, at all.** Every VBA differential result to date is
   against LibreOffice, not Excel — and LibreOffice's own VBA layer is not a verified proxy
   for Excel's. No Windows/Excel environment has ever been available in this project's
   toolchain. This is the single largest gap blocking a 90+ claim. The
   `compat/oracle-excel-com/CONTRACT.md` adapter is written and waiting for one.
   **Also applies to item 13's "safe round-trip" work (shipped as `0.8.0`)**: its own tests
   are structural/synthetic-fixture-only, for the same reason — no real Excel-authored file
   or working Excel/LibreOffice oracle exists to verify against yet. Closing this gap for
   real, with a real Windows+Excel environment and real Excel-authored fixtures, is `0.9.0`'s
   entire purpose — see the roadmap below.

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

    **Shipped as `elixcee` `0.8.0`** (confirmed live on crates.io/PyPI). Validating it
    against a real Microsoft-Excel-authored `.xlsm` — not just the synthetic fixtures in
    `tests/xlsx_roundtrip.rs` — is `0.9.0`'s job; see the roadmap below.

## npm/JS/WASM: still-open gaps

CI wiring, browser/WASM smoke coverage, and package-size measurement are all done (see
`CHANGELOG.md`'s history for the full investigation and writeup) — the `@elixcee/xlsx`
roadmap below covers what's still needed for an actual publish far more concretely than the
narrative version of this section used to. One narrow gap that roadmap doesn't separately
call out, still real:

- **No WASM binary size *regression* check.** The size is measured and logged every CI run,
  but nothing fails CI if it grows — a real policy call (what threshold, what to do when a
  legitimate feature grows it) that hasn't been made.

~~`check-versions.sh` has no awareness of `packages/xlsx/package.json`'s own version~~ —
**fixed**: it now guards the one concrete failure mode this actually risked — `"private":
false` (publish-ready) paired with `"version": "0.0.0-development"` (the placeholder nobody
meant to actually publish) — without cross-checking its version against `Cargo.toml`, since
`@elixcee/xlsx` versions independently by design.

## Roadmap: 0.9.0 → 1.0.0

Basic policy going forward: not "add more VBA syntax," but three pillars, in order —
(1) prove compatibility against real Microsoft Excel, not just LibreOffice/synthetic
fixtures; (2) preserve-and-merge existing workbooks rather than regenerate-from-scratch,
extended past the `0.8.0` slice; (3) a documented, stable `1.0` support profile with real
external usage behind it. 100/100 is not the goal — chasing full Excel feature parity would
make this project unbounded. elixcee's winning position is a lightweight Rust runtime that
inspects, executes, and safely updates existing `.xlsm` workbooks without installing Excel at
all — not a SheetJS/calamine/openpyxl/xlwings/LibreOffice replacement across the board, but
better than each of them at this specific intersection: writes more than calamine, runs VBA
unlike SheetJS, is more Rust/WASM/diagnostics-native than openpyxl, doesn't need Excel unlike
xlwings, and is lighter to embed than LibreOffice.

### 0.9.0 — Excel-Validated Round Trip

**Goal**: move from "looks correct against synthetic fixtures" to "confirmed not to break in
real Microsoft Excel." This is the shortest path from 95 to 96.

**Split into two independent tracks, `0.9.0-A` and `0.9.0-B`** (decided mid-milestone, once
it became clear file-preservation and VBA-semantics-vs-Excel needed very different
verification strategies and failure-handling): `0.9.0-A` is **file-preservation round-trip
only** (open in Excel → elixcee edits a value/formula, never intentionally-erroring VBA →
save → reopen in Excel → compare) — no dynamic code injection, no VBA execution required at
all beyond a pre-existing macro someone already wrote. `0.9.0-B` is the VBA-semantics-vs-Excel
oracle (rerunning macros, including error scenarios) and stays paused.

**`0.9.0-A` status: in progress, real progress made, not released.** Using Mac Excel (16.108),
not Windows — the originally-planned Windows+Excel environment (item 1 below) was never set
up; Mac was available and sufficient for pure file-preservation checks (no COM, but
AppleScript covers open/save/reopen/cell-read/`has vb project` cleanly). 5 real Excel-authored
`.xlsm` fixtures (not yet 10), each round-tripped through elixcee (save-as and in-place) and
reopened in Excel: 0 repair warnings, 0 `vbaProject` loss, 0 relationship breakage, 0
in-place-save failures. Found and fixed 3 real bugs (formula flattening, orphaned
relationships, wrong `.xlsm` content type — see `CHANGELOG.md`). Confirmed, not newly
discovered: worksheet-embedded features (tables/validation/conditional-formatting/
hyperlinks/defined-names/charts/images) are silently dropped on every save — this is `0.10.0`'s
job, not `0.9.0`'s, and was already disclosed as a `0.8.0` Non-goal. Not done: the
10-consecutive-cycle requirement (item 5 below, only 1 cycle so far), and macro-*execution*
verification specifically (blocked by a Mac Excel/VBA "license information not found" error
that reproduces on an untouched file from Excel's own UI — an environment issue on this
machine, not an elixcee defect, and not chased further this round). Full results:
`compat/oracle-excel-com/results/0.9.0-A_{results.json,summary.md}`.

**The earlier same-day live spike into Mac Excel AppleScript automation** (VBA's own
`VBComponents.Add`/`CodeModule.AddFromString` self-modification trick, triggered via
AppleScript's `run VB macro` with a string argument, to dynamically inject and run arbitrary
VBA source) — see `compat/oracle-excel-com/MACOS_APPLESCRIPT_EXPLORATION.md` — remains paused
and unresolved (VBE hangs on an injected runtime error; a `-50` parameter error that didn't
un-break on revert). `0.9.0-A` deliberately did **not** need this mechanism at all: no dynamic
injection, no intentionally-failing VBA, only reading/writing cell values, comparing
`has vb project`, and running a macro that already existed in the saved file (a materially
easier, and so far reliable, case). Also newly confirmed empirically this round: Mac Excel's
AppleScript dictionary *documents* `make new list object`/`add data validation`/`make new
format condition`/`make new chart object`, but none of them actually work (`-50` parameter
errors against a live instance) — table/data-validation/conditional-formatting/hyperlink/chart
content for fixtures 3–5 had to be authored manually in Excel's UI, not automated.

1. **A real Windows+Excel verification environment.** Actually run
   `compat/oracle-excel-com`, not just keep its `CONTRACT.md` waiting. Record, per run: Excel
   version, 32/64-bit, Windows version, locale, workbook calculation mode, macro security
   setting, and the run's own timestamp. *Not done — Mac Excel was used instead for `0.9.0-A`;
   still open for `0.9.0-B`'s VBA-semantics-vs-Excel work.*
2. **Real Excel-authored fixtures — at least 10.** Suggested mix: 5+ `.xlsm`, 3+ `.xlsx`, 2+
   with a chart/image/table, 2+ with data validation/conditional formatting, 2+ with
   comments/hyperlinks/defined names. All self-authored, containing no personal or
   confidential data. *5 done (all `.xlsm`, all self-authored, no personal/confidential data;
   see `0.9.0-A`'s fixture list above) — not yet 10.*
3. **Real round-trip procedure**, automated or semi-automated, per fixture: create in Excel →
   run the VBA to record its initial result → edit a cell value or formula via elixcee → save
   in the same format → reopen in Excel → check for a repair-warning dialog → rerun the VBA →
   compare the edited cell, the untouched cells, and overall workbook structure. *Done for
   value/formula edits and repair-warning checks, all 5 fixtures, both save modes; VBA rerun
   only exercised for the one fixture with a pre-existing macro, and blocked there by the
   environment issue noted above.*
4. **Classify results, not just pass/fail**: `EXACT_MATCH`, `SEMANTIC_MATCH`,
   `EXPECTED_REWRITE`, `UNSUPPORTED_PRESERVED`, `ELIXCEE_DATA_LOSS`,
   `ELIXCEE_RELATIONSHIP_BREAK`, `EXCEL_REPAIR_REQUIRED`, `ORACLE_FAILURE`,
   `NONDETERMINISTIC` — matching this project's existing verdict-enum discipline elsewhere
   (`compat/differential/classify.mjs`).
5. **Hard gates, all zero**: Excel repair-warning dialogs; `xl/vbaProject.bin` loss; silent
   loss of any property elixcee claims to support; a changed result on VBA rerun; a wrong
   value in an edited cell; loss of any unknown ZIP part. *All zero across the 5 fixtures
   done so far — including "loss of any unknown ZIP part," now that bugs 2/3 above are
   fixed. Worksheet-embedded features (tables/validation/etc.) don't count against this gate:
   elixcee has never claimed to support preserving those (see `0.8.0`'s Non-goals), so their
   loss isn't a broken claim — it's `0.10.0`'s open scope.*

**Explicitly not this round**: a large batch of new VBA language features, an ODS writer,
new chart generation, full `PivotTable` support, or `@elixcee/xlsx` stable npm publish.

**Exit criteria**: 10+ real Excel-authored fixtures (5+ `.xlsm`), 0 repair warnings, macro
rerun succeeds on every fixture, every failure gets a reproduction fixture and a real fix
(not a downgraded gate), results recorded as machine-readable JSON, README states the
"Microsoft Excel validated" scope precisely (which fixtures, which properties) rather than
as a blanket claim. *Not yet met: 5 of 10 fixtures, 1 of 10+ cycles per fixture/mode, macro
rerun verified on only 1 of 5 fixtures (blocked there, see above). README not yet updated —
premature before the fixture count and cycle count are both met.*

### 0.10.0 — Lossless Worksheet Preservation

**Goal**: `0.8.0` already preserves unknown ZIP parts, `xl/vbaProject.bin`, style
definitions, merges, and hidden rows/columns — but worksheet XML itself is still always
fully regenerated (`build_xlsx_sheet`), so anything elixcee doesn't understand that lives
*inside* a `<worksheet>` element (not a separate part) is still lost. `0.10.0` closes that.

**In priority order**: defined names, tables, hyperlinks, comments/notes, data validation,
conditional formatting, freeze panes, autofilter, print area/print titles, page margins/page
setup, row/column dimensions, richer workbook properties.

**Architecture — recommended: preserve-and-merge.** Read the original worksheet XML, update
only the elements elixcee itself owns, keep unknown elements/attributes verbatim, keep
relationship IDs stable, and only remove something when a change explicitly calls for
removing it. This needs to be namespace- and OOXML-element-order-aware — not a blind string
substitution — matching `0.8.0`'s own schema-ordering discipline for `<cols>`/`<sheetData>`/
`<mergeCells>`.

**Relationship-graph validation.** Model and check `worksheet → table`, `worksheet →
drawing`, `drawing → image`, `worksheet → comments`, `worksheet → hyperlink`, `workbook →
worksheet`, and `content-types → part` as a graph: every reference resolves, no orphan parts,
no duplicate IDs, no path traversal.

**Exit criteria**: 20+ Excel-authored fixtures covering these features, every untouched
unsupported XML node preserved byte- or semantically-equivalent, 0 Excel repair warnings, 0
loss of tables/validation/comments/etc., 0 broken chart/image relationships, successful
in-place save.

### 0.11.0 — VBA Semantic Closure

**Goal**: structurally close out the VBA semantic gaps that remain. `0.7.0` already fixed
multi-dimensional arrays and call-frame error handling; `Date`/`Time` and type-width tracking
are what's left.

1. **DateTime runtime model.** `Variant::Date(i64)` is whole-day-only today — a structural
   reason `Time()`/`Now()` report `TypeName` `"Double"` instead of real VBA's `"Date"` (item 5
   above). This has already been designed, not yet implemented: `docs/date-time-runtime-model-adr.md`
   compares three options (A: change `Variant::Date(i64)` to a breaking `Variant::DateSerial(f64)`;
   B: keep `Date`, add an additive `Variant::DateTime(f64)`; C: an internal-only
   representation, shown not to actually work since `Now()`'s return value must be a real
   `Variant` to be assignable to a VBA variable at all). **Recommendation: B** — same
   `elixcee-types` minor-version cost as A, far less code churn, zero observable-behavior
   change for any value that's already `Variant::Date`. Scope for `0.11.0`: `Date`, `Time`,
   `Now`, `CDate`, `DateSerial`, `TimeSerial`, date/time arithmetic and comparison, `TypeName
   == "Date"`, the Python/JSON/WASM representations, `date1904`, Excel serial-60 handling.
2. **Separate declared type from runtime value**, at least for `Integer`/`Long`/`Double`/
   `Boolean`/`String`/`Date`/`Variant`/`Object`.
3. **Type width and overflow**: 16-bit `Integer`, 32-bit `Long`, conversion/assignment/
   arithmetic overflow, and how each interacts with `On Error`.
4. **`Variant`'s own stored-type tag** — correctly distinguish `"1" + "2"` from `1 + 2` from
   `CStr(1) + CStr(2)`, and handle a numeric-string-vs-number-Variant comparison explicitly
   rather than by accident.

**Exit criteria**: `compat/vba-semantics` suite grows to 500–600+ cases, 0 `BUG`, 0
`UNCLASSIFIED`, `KNOWN_LIMITATION` down from 14 to 5 or fewer, `Date`/`Time`'s `TypeName`
matches real VBA, Python/WASM/JSON round-trip verified, real-Excel differential agreement on
supported cases at 95%+.

### 0.12.0 — Practical Workbook Mutation

**Goal**: `0.8.0`–`0.10.0` are about *preserving* existing state; `0.12.0` is about safely
*changing* more of it — style edits, not just style preservation.

- **Style editing**: `Range.NumberFormat`, `Range.Interior.Color`, `Range.Font.Bold`,
  `Range.Font.Color`, borders, alignment, wrap text — de-duplicating against the existing
  style table when adding a new style rather than growing it unboundedly.
- **Worksheet operations**: add/delete/rename/reorder sheets, visible/hidden/very-hidden,
  changing the active sheet.
- **Workbook structure**: add/change/delete defined names, hyperlinks, comments, data
  validation, autofilter, minimal table updates, and a policy for discarding vs. regenerating
  the calculation chain.

**Exit criteria**: 0 Excel repair warnings on reopen, newly-applied styles render correctly,
relationship integrity holds, sheet-rename updates every reference to it, 0 silent no-ops for
a claimed-supported mutation, any genuinely unsupported property fails with an explicit error
rather than a silent no-op.

### 0.13.0 — Scale, Security, and Distribution

**Goal**: not just features — safely handling real-world-sized files in production.

- **Performance**: 10MB/50MB/100MB workbooks, 100K/1M cells, 100-workbook batches, cold
  start, peak RSS, write latency, Python call overhead, WASM payload size — all as continuous
  regression gates, not one-off measurements.
- **Security**: ZIP-bomb protection, oversized-XML limits, entry-count limits,
  decompression-ratio limits, path traversal, XML entity expansion, explicit
  formula-injection handling, unsafe external relationships, malformed/cyclic relationship
  graphs.
- **Fuzz**: a persisted corpus, automatic promotion of crash-producing inputs into fixtures,
  round-trip fuzzing across parser/reader/writer, keeping the existing 30-second CI gate
  while running any longer fuzzing campaign as a separate scheduled workflow, and classifying
  panic/OOM/hang outcomes distinctly.
- **Distribution**: SBOM, build provenance, reproducible builds, checksums, signed releases,
  dependency license audit, vulnerability scanning.

### 1.0.0 — Stable Supported Profile

**What 1.0 means here**: not "full Microsoft Excel feature parity." Defined instead as the
**elixcee Supported VBA and Workbook Profile 1.0** — within that documented scope, no silent
corruption, and a stable, guaranteed API and behavior contract.

**Required**:
- *VBA*: 95%+ agreement with real Excel on supported semantic cases; 0 silently-wrong
  results; unsupported syntax rejected explicitly at parse/check time; stable runtime error
  numbers and metadata; `DateTime` and type-width support in place; 750+ semantic-suite cases.
- *Workbook*: 30+ Excel-authored fixtures (10+ `.xlsm`); 0 repair warnings; 0 loss of any
  supported property; VBA project preserved; styles both preserved and editable;
  tables/validation/comments/etc. preserved; chart/image relationships intact.
- *API*: stable Rust API, stable Python API, a fixed CLI JSON schema with real schema
  versioning, a fixed WASM API, a documented deprecation policy and migration guide.
- *Distribution and track record*: consistent crates.io/PyPI/GitHub-Release publishing, a
  published npm package, 3–5 real external usage examples, a security policy, a support
  matrix, reproducible releases, and a documented rollback/yank policy.

**Explicitly still out of scope, even at 1.0**: the Excel UI itself, `UserForm`, ActiveX, COM
add-ins, Power Query, full `PivotTable` compatibility, full chart-generation compatibility,
the complete VBA event model, the VBA IDE, and "replaces Excel entirely" as a claim.

### Score trajectory

| State | Score |
|---|---|
| `0.7.0` | 94 |
| `0.8.0` | 95 |
| Real-Excel round trip succeeds (`0.9.0`) | 96 |
| Preserve-and-merge extended (`0.10.0`) | 96–97 |
| Known VBA semantic gaps down to 5 or fewer (`0.11.0`) | 97 |
| npm alpha + real external usage | 97 |
| `1.0.0` Supported Profile | 97–98 |

100/100 isn't the target — chasing full Microsoft Excel feature/compatibility parity would
let this project grow without bound. The current highest-priority work is not a new feature
but building real evidence: a genuine Microsoft-Excel-authored `.xlsm` round trip, in
`0.9.0`. Clearing that is what makes the 95 → 96 move concretely defensible.

### `@elixcee/xlsx` — independent roadmap

Versioned independently of the root crate, same as today (`0.0.0-development`/
`private: true`, `read`/`write` already implemented, still unpublished).

**`0.1.0-alpha.1` publish conditions**: `@elixcee` npm scope ownership confirmed, npm publish
credentials available, the writer differential suite wired into regular CI, a real
`npm pack` tarball consumer smoke green on both Node 20 and 22, a real-Chrome smoke green,
CJS and ESM both verified, an accurate package README, and a documented supported/unsupported
matrix. Target `package.json` shape at that point:

```json
{
  "version": "0.1.0-alpha.1",
  "private": false,
  "publishConfig": {
    "access": "public",
    "tag": "alpha"
  }
}
```

**What alpha guarantees**: `read`/`readFile`/`readFileSync`, `write`/`writeFile`/
`writeFileSync`, `bookType: "xlsx"` only, Node and browser, `Buffer`/`Uint8Array`/base64
output, the documented `utils` subset.

**What alpha does not guarantee**: an ODS writer, `.xls`/`.xlsb`, encrypted workbooks,
`PivotTable`, chart creation, full SheetJS option coverage, or API stability.

**`0.1.0-alpha.2`+**: fixes from real npm users' issues, bundler-compatibility follow-ups,
Deno/Bun verification, webpack/Vite/Rollup consumer tests, TypeScript type-compatibility
checks, package-size reduction.

**`0.1.0-beta.1` exit criteria**: 3+ external users, 0 significant silent-corruption reports,
a frozen supported-API surface, a migration guide, a documented semver policy, and a
documented browser/Node support matrix. Publishing to npm alone doesn't move elixcee's
overall score to 96 — but real external usage is a real prerequisite for it.

## Non-goals (still, per existing ADRs)

- No new Rust runtime dependencies beyond what's already justified in `Cargo.toml`'s
  comments — matches this codebase's long-running dependency-minimization direction
  (`docs/xlsx-architecture.md`).
- `packages/xlsx` never depends on the real `xlsx` package at runtime (ADR, same doc).
- No byte-for-byte compatibility claims where SheetJS itself is non-deterministic (embedded
  timestamps, etc.) — compatibility is measured on parsed logical shape.
- Out of scope even at `1.0.0` (see the roadmap above): the Excel UI itself, `UserForm`,
  ActiveX, COM add-ins, Power Query, full `PivotTable` compatibility, full chart-generation
  compatibility, the complete VBA event model, the VBA IDE, and "replaces Excel entirely" as
  a claim. elixcee's scope is a lightweight runtime for inspecting, executing, and safely
  updating existing workbooks without Excel installed — not a full Excel reimplementation.
