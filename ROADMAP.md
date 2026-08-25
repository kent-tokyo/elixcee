# Roadmap

This tracks where `elixcee`/`@elixcee/xlsx` stands and what's next. For what already
shipped, see `CHANGELOG.md` — this file only restates completed work when needed to explain
what's left. Historical phase-by-phase implementation notes (Japanese) live in
`tasks/todo.md`.

## Current state

`elixcee` **0.10.1** (Rust crate + Python package), `elixcee-types` **0.3.0** (unchanged),
`elixcee-wasm` **0.1.0** (never published to crates.io by design — `publish = false`) — all
confirmed live via crates.io's/PyPI's own APIs, not assumed from local files. `0.10.1` is a
single targeted patch over `0.10.0`, fixing a critical unbound-`r:`-namespace-prefix
regression (see `CHANGELOG.md`'s `[0.10.1]`) — released from the same commit across PyPI,
crates.io, and a GitHub Release (`bin-v0.10.1`, 3 platform binaries), each independently
re-verified against the live published artifact, not just a local build. The downloaded
macOS CLI binary confirmed runnable and reports `elixcee 0.10.1` via `--version`.
`@elixcee/xlsx` is unchanged:
`read()`/`readFile()`/`readFileSync()` and `write()`/`writeFile()`/`writeFileSync()` are both
implemented and differential-tested, but the package is still
`0.0.0-development`/`private: true`/unpublished — no `npm publish` has happened (confirmed
live: `registry.npmjs.org/@elixcee/xlsx` 404s), and `@elixcee` scope ownership itself is
unconfirmed (item 9 below).

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

**`0.9.0` shipped `0.9.0-A`: the first real-Excel-validated round trip.** 5 real
Microsoft-Excel-for-Mac-authored `.xlsm` fixtures (values/formula/style/merge/hidden
rows-cols; VBA project + macro; table/data validation/conditional formatting;
hyperlink/comment/defined name; chart/image/print area), each edited via elixcee, saved both
ways (save-as and in-place), and reopened in real Excel — 0 repair warnings, 0 `vbaProject`
loss, 0 relationship breakage, 0 in-place-save failures, across all 5. Found and fixed three
real bugs the synthetic fixtures never exercised (formula flattening, orphaned relationships,
wrong `.xlsm` content type for a non-macro workbook — the last one made Excel refuse to open
the file outright). See `CHANGELOG.md`'s `[0.9.0]` and
`compat/oracle-excel-com/results/0.9.0-A_summary.md` for full detail, including two open items
neither fixed nor newly discovered: worksheet-embedded features (tables/validation/
conditional-formatting/hyperlinks/defined-names/charts/images) are silently dropped on any
save — already disclosed in `0.8.0`'s own Non-goals, confirmed live here, in `0.10.0`'s scope
to fix, not `0.9.0`'s. **Excel-authored `.xlsm` save/reopen/structural preservation is
verified against real Excel. Macro re-execution after a save is NOT verified** — a Mac
Excel environment VBA license error blocks running any macro at all, reproduced on the
untouched original file before elixcee ever touched it; this is neither a pass nor a fail
for elixcee's own round-trip, it's simply unevaluated. Do not describe this release as having
verified macro compatibility or confirmed VBA still works post-save — `README.md`/
`README_ja.md`/`README_zh.md`'s own "Microsoft Excel round-trip validation" section states
this precisely. The 10-consecutive-cycle exit criterion is superseded, not met: a 5-cycle
chained in-place stress test on the same file (harder than 5 independent cycles — any
accumulating corruption would compound) stayed clean through real Excel reopen, judged
sufficient evidence in place of the full 10. The VBA-semantics-vs-Excel axis (`0.9.0-B`)
stays paused — see the roadmap below.

`cargo test --workspace` 961/961 (up from 955 at `0.8.0`), `compat/vba-semantics` 386 cases
(0 `BUG`, 0 `UNCLASSIFIED`, 14 `KNOWN_LIMITATION`, unchanged from `0.8.0`), `compat/corpus`
581 scenarios (0 `UNEXPLAINED`, 0 `MISMATCH`), every GitHub Actions job green on `master`
before this release.

**`0.10.0` shipped the first three slices of Lossless Worksheet Preservation** (design in
`docs/xlsx-worksheet-preservation-0.10.0-design.md`). `0.10.0-A` (foundation — `WorksheetOrigin`/
`sheetId` preservation), `0.10.0-B` (inline worksheet elements: freeze panes/selection,
sheetPr/sheetFormatPr/phoneticPr/dataValidations, pageMargins, internal hyperlinks minus
`<autoFilter>`/row-col style), and `0.10.0-C` (workbook-level: workbookPr/bookViews/calcPr/
extLst/definedNames) are all done, mechanical-check-verified, and real-Excel
reopen-verified (0 repair warnings; `fixture4`'s defined name and `fixture5`'s print area
both confirmed byte-for-byte in Excel's own Name Manager/print preview). Also fixed: three
independent, pre-existing correctness bugs affecting every released version before `0.10.0`
— a save's sheet tab order silently followed an alphabetical sort instead of the source
order, a sheet's display-name letter case was silently lowercased on every save, and
`Sheets.Add` could silently no-op (no new sheet, no error) whenever the sheet set had a
numbering gap. This closed
[GitHub issue #1](https://github.com/kent-tokyo/elixcee/issues/1) (the display-name-case and
spurious-extra-sheet bugs, reported against `0.9.0`, already fixed in these same commits
before the issue was filed).

A separate, more severe regression was then reported against the published `0.10.0` wheel: a
source workbook binding the OOXML relationships namespace to a non-`r:` prefix (valid OOXML —
binding is about the URI, not the prefix spelling) round-tripped into a file with `r:` used
but never bound, rejected outright by any strict XML consumer. Reproduced exactly, root-caused
(the writer always hardcodes the literal `r:` prefix regardless of what the source used), and
fixed via `reader::ensure_r_prefix_bound()` — full detail in `CHANGELOG.md`'s `[0.10.1]`.
Released as `elixcee` `0.10.1` (PyPI/crates.io/GitHub Release, same commit), verified against
the published `0.10.1` wheel (standard prefix, alternate prefix, already-correct `xmlns:r`,
`xmlns:r` bound to a wrong URI, save-as/in-place/two-consecutive-saves — all reopening
cleanly). The original reporter independently re-verified this fix plus both of issue #1's
fixes against the published wheel and closed issue #1 themselves — no action needed on this
repo's side.

`0.10.0-D` (relationship-backed features, the actual fix for `SOURCE_REFERENCE_LOSS`) is
**not released yet** — see `CHANGELOG.md`'s `[Unreleased]`. It has a decided design
(origin-based worksheet part naming — see the roadmap entry below); `D1` (the
`WorksheetOutputPlan` output plan itself) is done, and every relationship-backed element
with real fixture evidence — `<tableParts>`, `<drawing>`, `<legacyDrawing>`, `<hyperlinks>`
(including r:id-backed ones, a rewrite of `0.10.0-B4`'s prior relationship-free-only scope)
— is now restored. All 7 real fixtures report `CLEAN` across every `mechanical_check.py`
category, including `source_references`: `SOURCE_REFERENCE_LOSS` is eliminated from the
entire current fixture set. `D4` (reachability-based cleanup of a deleted sheet's
exclusively-reachable parts) is also done, closing Known gaps item 15; sheet rename/reorder
are marked N/A rather than implemented (this `Vm` has no such primitive, and adding one
purely to fill a test-table row would invert this milestone's own hard gate). Plain
(relationship-free) `<pageSetup>` is also restored — `fixture5`'s real shape, previously
silently lost and uncaught by any checker. `r:id`-backed `<pageSetup>` (a `printerSettings`
relationship) remains the only genuinely open item, blocked on the project's own hard gate:
no fixture with that shape exists in the repo. All of `0.10.0-D` above is
mechanical-check-verified but not yet real-Excel reopen-verified — that verification, plus
whatever `<pageSetup r:id>` needs, gates the next release.

**Next round: release qualification, not new development — version not decided yet.**
`0.10.0-D` and the `t="e"` error-cell fix (Known gaps item 14, above) are both merged to
`master` and both correctly excluded from `0.10.1`'s score (see the `0.10.1` scorecard) —
neither ships until each clears its own real-Excel check, done by hand against real Excel
(this `Vm` has no way to drive that itself):

- `0.10.0-D`, per element: a table survives; an external hyperlink still works; a
  drawing/chart/image still displays; a comment/note survives; plain `<pageSetup>` is
  still applied; deleting a sheet leaves no orphaned part triggering a repair warning; both
  save-as and in-place produce zero repair warnings.
- `t="e"`: the cell is genuinely error-typed in Excel, not just displaying error-looking
  text — a formula referencing it and `ISERROR()` both need to see it as an error; and the
  type survives a save → reopen → save cycle, not just the first save.

Whether the next release is `0.10.2` (if `0.10.0-D`'s real-Excel pass turns up only small
fixes) or `0.11.0` (if it lands as a real new capability, table/drawing/hyperlink
preservation formally added) is an open question until that verification is done — not
decided in advance.

## Known gaps

1. **VBA semantic differential results are still validated against LibreOffice only, not
   Excel — narrower than it used to be, not closed.** Every VBA differential result (the
   386-case `compat/vba-semantics` suite, the 581-scenario `compat/corpus`) is still checked
   against LibreOffice, not Excel, and LibreOffice's own VBA layer is not a verified proxy
   for Excel's — no Windows/Excel environment has ever been available in this project's
   toolchain, and the `compat/oracle-excel-com/CONTRACT.md` adapter is still written and
   waiting for one. **What changed as of `0.9.0`**: the separate "does elixcee's own
   round-trip save path corrupt a real Excel-authored file" question — item 13's "safe
   round-trip" work (shipped as `0.8.0`, structural/synthetic-fixture-only at the time) — is
   now validated against real Microsoft Excel (5 fixtures, `0.9.0-A`, see the roadmap below
   and `CHANGELOG.md`'s `[0.9.0]`). Macro *re-execution* after a save is a further, separate,
   still-unvalidated question (blocked by a Mac Excel environment issue, see below) — do not
   conflate "file-preservation validated" with "VBA-vs-Excel semantics validated" or "macro
   execution after save validated." Closing the semantic-differential gap for real needs a
   real Windows+Excel environment; that work is `0.9.0-B`, paused — see the roadmap below.

2. **LibreOffice headless oracle is broken for most of the VBA corpus.** 578/581 scenarios
   are `ORACLE_UNAVAILABLE` — headless UNO hangs on any `Range`/`Cells` access. Root-caused,
   not fixed (explicitly ruled out twice already: fixing it doesn't raise elixcee's own
   product value, only this one oracle's usability — revisit only if the corpus itself
   becomes the bottleneck rather than VBA coverage). Unrelated to item 13's own real-Excel
   round-trip validation below (this gap is about VBA *semantic* differential testing, a
   different axis from file-preservation round-tripping).
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

    **Shipped as `elixcee` `0.8.0`** (confirmed live on crates.io/PyPI). **Validated against
    5 real Microsoft-Excel-authored `.xlsm` fixtures as of `0.9.0`** (`0.9.0-A`, see the
    roadmap below) — not just the synthetic fixtures in `tests/xlsx_roundtrip.rs`, which
    remain as a fast regression guard alongside the real fixtures now under
    `compat/oracle-excel-com/fixtures/pristine/`. Three real bugs this real-Excel validation
    found and fixed (formula flattening, orphaned relationships, wrong `.xlsm` content type)
    are in `CHANGELOG.md`'s `[0.9.0]`.

14. ~~A cell holding a real Excel error value (`t="e"` in the source XML, e.g. `#VALUE!`)
    round-trips as a plain text string, not an error.~~ — **fixed**: `src/reader.rs`'s
    `SheetCell` enum gained a real `Error(ExcelError)` variant (`elixcee_types::ExcelError`
    now has a `FromStr` impl, the `as_str()` inverse), threaded through `Vm` and the writer
    the same way `Variant::Error` already was at the VBA-runtime level. `xlsx_cell_xml` now
    emits `t="e"` with the literal error text in `<v>`, never shared-string indexed —
    confirmed against real Excel's own output, which never puts e.g. `"#VALUE!"` in
    `xl/sharedStrings.xml` either. `@elixcee/xlsx`'s `read()` (`crates/elixcee-wasm`) got
    the matching fix: error cells now come back as `{t:"e", v:<BIFF code>, w:<string>}`,
    the real `xlsx` oracle's own shape. Verified against fixture5's real `D8` cell (both
    the Rust round-trip and a real CLI save) and a new differential case. See
    `CHANGELOG.md`'s `[Unreleased]` for the full account.

15. ~~`mechanical_check.py`'s structural checks don't catch an orphaned worksheet `.rels`
    file when the sheet it belongs to is the one deleted. `check_roundtrip`'s orphan check
    (#2b) only flags a *worksheet part* that survives unreferenced; it never asks "does a
    surviving `.rels` file still have a sibling part it implicitly belongs to"~~ — **fixed**
    by `0.10.0-D4`: `check_deleted_sheet_cleanup()` (a dedicated package-reachability check,
    written and self-test-verified first) now catches this shape, and the writer's
    `deleted_sheet_prunable_parts` prunes the orphan itself, transitively, rather than
    leaving it behind. Verified against the exact real scenario that first exposed this gap
    (a fixture with a relationship-bearing sheet deleted) — the orphaned `.rels` is now
    genuinely absent. See the `0.10.0-D4` roadmap entry above for the full account.

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

### 0.9.0 — Excel-Validated Round Trip — **shipped**

**Goal**: move from "looks correct against synthetic fixtures" to "confirmed not to break in
real Microsoft Excel." This is the shortest path from 95 to 96.

**Released.** `0.9.0-A`'s scope (file-preservation round trip only, detailed below) shipped
as `elixcee` `0.9.0` — confirmed live on crates.io and PyPI, `bin-v0.9.0` CLI binaries
published, a fresh venv/fresh Cargo project/downloaded binary all independently re-verified
post-publish. Shipped as a **deliberately partial** win, same pattern as `0.8.0`: the
10-fixture and 10-cycle targets below were not literally met (superseded by alternate
evidence — see the item-by-item status notes), and macro-execution verification stayed
open, not silently absorbed into a broader "Excel compatibility validated" claim. `0.9.0-B`
(VBA-semantics-vs-Excel) remains a separate, still-paused track.

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
job, not `0.9.0`'s, and was already disclosed as a `0.8.0` Non-goal. The 10-consecutive-cycle
requirement (item 5 below) is **superseded, not met**: a 5-cycle chained in-place stress test
on fixture 1 (the same file edited+saved 5 times in a row, not 5 fresh copies) stayed clean
through a real Excel reopen after cycle 5, judged sufficient in place of the full 10 rather
than run to completion. **Macro *execution* verification is a separate, still-open item, not
done and not classified either way**: running fixture 2's macro fails with a Mac Excel VBA
"license information not found" error that reproduces identically on the untouched original
file from Excel's own UI, before elixcee ever touched it — this doesn't confirm elixcee broke
macro execution, and it doesn't confirm elixcee's round-trip preserves working macros either;
it's simply unevaluated, and should not be described as verified in either direction. Full
results: `compat/oracle-excel-com/results/0.9.0-A_{results.json,summary.md}`.

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
   value in an edited cell; loss of any unknown ZIP part. *Confirmed zero, across the 5
   fixtures done so far: repair warnings, `vbaProject.bin` loss, loss of any unknown ZIP
   part (now that bugs 2/3 above are fixed), wrong edited-cell values. Worksheet-embedded
   features (tables/validation/etc.) don't count against this gate: elixcee has never
   claimed to support preserving those (see `0.8.0`'s Non-goals), so their loss isn't a
   broken claim — it's `0.10.0`'s open scope. **"A changed result on VBA rerun" is NOT
   confirmed zero — it is unevaluated.** The one fixture with an intentional macro
   (fixture 2) never got a rerun at all, blocked by a Mac Excel VBA environment error that
   also reproduces on the untouched original file. This is a gap in the evidence, not a
   passing result — do not report it as satisfied.*

**Explicitly not this round**: a large batch of new VBA language features, an ODS writer,
new chart generation, full `PivotTable` support, or `@elixcee/xlsx` stable npm publish.

**Exit criteria**: 10+ real Excel-authored fixtures (5+ `.xlsm`), 0 repair warnings, macro
rerun succeeds on every fixture, every failure gets a reproduction fixture and a real fix
(not a downgraded gate), results recorded as machine-readable JSON, README states the
"Microsoft Excel validated" scope precisely (which fixtures, which properties) rather than
as a blanket claim. *Status: 5 of 10 fixtures. Cycle count superseded by the 5-cycle
chained in-place stress test (see above), not pursued to the literal 10. Macro rerun: not
verified on any fixture — attempted once (fixture 2, the only one with an intentional
macro), blocked by the environment error above; this is an open gap, not a passed check.
Results recorded as machine-readable JSON:
`compat/oracle-excel-com/results/0.9.0-A_results.json`. README not yet updated to state a
"Microsoft Excel validated" scope — premature while macro rerun is unverified and the
fixture count is short of 10; this is a separate, later review, not part of this round.*

### 0.10.0 — Lossless Worksheet Preservation — **A/B/C shipped as `elixcee` 0.10.0, D in progress**

**Goal**: `0.8.0` already preserves unknown ZIP parts, `xl/vbaProject.bin`, style
definitions, merges, and hidden rows/columns — but worksheet XML itself is still always
fully regenerated (`build_xlsx_sheet`), so anything elixcee doesn't understand that lives
*inside* a `<worksheet>` element (not a separate part) is still lost. `0.10.0` closes that.

**Design done, split into 4 milestones (A/B/C/D), not this flat priority list.** The
architecture below (preserve-and-merge, relationship-graph validation) was the right
direction but not the actual implementation plan — see
`docs/xlsx-worksheet-preservation-0.10.0-design.md` for the real one, produced from grounded
evidence (real fixture inspection + one empirical elixcee-save-and-check run, not
assumption), reviewed and revised twice. Key departures from this section's original sketch:
milestones are split by "does it touch the relationship graph" (not by feature name), a new
`WorksheetOrigin` identity struct threads a sheet's original `sheetId`/`workbook.xml`
`r:id`/part-name through save so cross-save identity is stable, and a hard gate governs every
element — **no writer code until a real Excel-authored fixture demonstrates it, its XSD
sequence is confirmed against the actual ECMA-376 schema (not memory), and
`mechanical_check.py` has a negative test for its loss.**

**0.10.0-A (foundation, done)** — zero new writer features. `WorksheetOrigin` implemented
end-to-end (closes a real, previously-disclosed gap: `snapshot.rs`'s `stable_id` wasn't
actually stable across a save+resave cycle before this). `mechanical_check.py` gained
`check_source_references()` and a new `SOURCE_REFERENCE_LOSS` violation category — a
worksheet-level relationship's `.rels`/target part survive a save byte-identical, but the
regenerated worksheet XML no longer references the `r:id` that activates it; confirmed
systemic across every fixture with a worksheet-level relationship at all (not yet fixed —
that's `0.10.0-D`'s job). Two new real Excel-authored fixtures added (internal hyperlink,
freeze pane — neither existed in the repo before). One real, independent bug found and fixed
along the way: a workbook whose sheets are never literally named "Sheet1" gained a spurious
extra empty sheet on every save (`Vm::new()`'s default sheet leaking past load).

**0.10.0-B (inline worksheet preservation, functionally done)** — relationship-free elements
*inside* `<worksheet>`, via an opaque-fragment mechanism (capture the source's raw XML for an
element, splice it back at the correct schema position, don't parse/reconstruct it). 4 slices
shipped, each independently fixture-verified and 2 real-Excel-reopened (repair warnings: 0):
`<sheetViews>` (freeze panes, selection), `<sheetPr>`/`<sheetFormatPr>`/`<phoneticPr>`/
`<dataValidations>`, `<pageMargins>`, and internal (`location=`) hyperlinks — the last one
structurally different from the rest, since a `<hyperlinks>` container can mix
relationship-free children with `r:id`-backed ones that stay out of scope until `0.10.0-D`,
so it's reconstructed from filtered children rather than byte-copied whole. Deliberately left
out (not blocking): `<autoFilter>` (no fixture has it as a standalone worksheet element yet —
the only real example lives inside a table part) and row/column style properties beyond
hidden state (real fixture evidence exists but needs its own design pass).

**0.10.0-C (workbook-level preservation, done)** — the same opaque-fragment mechanism as B,
applied to `xl/workbook.xml`'s own direct children, split into 3 slices by position-
dependence: C1 (`<workbookPr>`/`<calcPr>`/`<extLst>`, plus the root tag's own namespace
declarations — position-independent, plain verbatim copy), C2 (`<bookViews>` — its
`<workbookView>` can in principle carry sheet-position `activeTab`/`firstSheet` attributes,
but all 7 real fixtures were checked and none sets either, so this ships as plain verbatim
too rather than building unvalidated gating logic for a hazard with zero fixture evidence),
and C3 (`<definedNames>`, including print area/titles — `localSheetId` DOES have real
fixture evidence, so it's carried verbatim only when no sheet has been deleted since load,
and dropped entirely otherwise, rather than risk a stale reference; per-name `localSheetId`
remapping for the delete case is left as documented future work). New `WORKBOOK_ELEMENT_LOSS`
category (plus a dedicated `check_defined_names()` for C3's delete-dependent correctness) in
`mechanical_check.py`, all 7 fixtures confirmed CLEAN. **Real-Excel reopen verification
done**: `fixture4` (defined name `test`, workbook-scope, `=Sheet1!$F$5`, comment
`test desu!!!`) and `fixture5` (`_xlnm.Print_Area`, `Sheet1!$E$3`) both confirmed live in
Mac Excel — Name Manager shows every field byte-for-byte matching the source, Print_Area's
print preview shows exactly the (empty) `E3` cell rather than the sheet's real data table (a
positive control: had `Print_Area` broken and fallen back to "print everything," the data
table would have appeared instead), 0 repair warnings across all 3 output files
(`fixture4` save-as, `fixture4` in-place, `fixture5` save-as).

**0.10.0-D (relationship-backed features, including the actual fix for
`SOURCE_REFERENCE_LOSS`)**: design decided; `D1` done, every relationship-backed element
with real fixture evidence (`<tableParts>`, `<drawing>`, `<legacyDrawing>`, `<hyperlinks>`)
restored — all 7 real fixtures now `CLEAN` on `source_references`. `<pageSetup r:id>` (no
fixture has one yet) and `D4` not started. Worksheet
parts were previously named `sheet{i+1}.xml` by output position, while a worksheet's
`.rels` file (and whatever it points at — tables, drawings, comments) survived keyed by
the *original* part path. That was self-consistent only because nothing carried
worksheet-level relationships forward yet; `0.10.0-D` reconnects the relationship graph.
**Decided: origin-based part naming** — an existing sheet's output part name stays
`WorksheetOrigin.original_part_name` regardless of save-time position, a new sheet gets
`max(existing sheetN) + 1` (never reusing a freed number), and a deleted sheet's exclusively-
reachable target parts (not shared with a surviving sheet) are removed via package
reachability rather than blind deletion. Positional renumbering with `.rels` remapped at
write time was considered and rejected — Open XML discovers parts via the relationship
graph, not sequential naming, so leaving existing part URIs untouched is safer than
rewriting every reference. Full design (the `WorksheetOutputPlan` sketch, D1–D4 commit
breakdown, required test case table) in
`docs/xlsx-worksheet-preservation-0.10.0-design.md` §10.

**`D1` (output plan + tests, done).** `WorksheetOutputPlan` + `plan_worksheet_output` now
drive every writer call site that used to independently re-derive `sheet{i+1}.xml`/
`sheetId`/`r:id`. Confirmed as a real bug fix (not just a rename) via a git-stash A/B
comparison: a surviving sheet's content used to land at a position-derived part name while
its passthrough `.rels` file stayed at the original name, so deleting an earlier sheet
could orphan a later sheet's relationship file. Deliberately does not yet restore
relationship-backed references (`SOURCE_REFERENCE_LOSS` itself, `D2`/`D3`'s job). One
known checker gap surfaced while verifying D1 (Known gaps item 15): when the sheet with
its own `.rels` file is the one deleted, the orphaned `_rels/sheetN.xml.rels` that
resulted wasn't flagged by any `mechanical_check.py` category that existed at the time —
harmless to Excel in practice (nothing references it), pre-existing under a different
stale filename before D1. Closed by `D4`'s reachability-based cleanup, below.

**`<tableParts>` restored (first relationship-backed element, done).** D1's own doc
comment described a separate `D2` slice ("carry a surviving sheet's `.rels` through at
the original part name with `r:id`s unchanged") as the precondition for restoring any
reference. Verified that precondition before writing any splice code: running
`check_source_references()` against `fixture3` with the D1 binary showed every violation
was "the sheet doesn't reference the rId", never "the `.rels` differs from the original"
— the generic passthrough loop has copied worksheet `.rels` byte-identical since `0.9.0`,
and D1 already made it land at the correct co-located part name. There was no separate D2
implementation needed; the whole remaining gap was the splice itself. Restored via the
same opaque-fragment mechanism `0.10.0-B` established, spliced right after
`<pageMargins>` (confirmed schema position: nothing between `pageMargins` and
`tableParts` is emitted yet). Gated on `rels_survived` (`is_existing` AND the sheet's own
`.rels` actually present in this save's passthrough set) — restoring a reference whose
`.rels` didn't survive would emit a dangling `r:id`, a real Excel repair warning strictly
worse than the prior silent inertness; a dedicated negative test confirms the gate
actually fires (verified by temporarily removing it and watching the test fail).
`fixture3`'s `check_source_references()` verdict is now `CLEAN`; `fixture4`/`fixture5`
still report `SOURCE_REFERENCE_LOSS` for hyperlink/vmlDrawing/drawing `r:id`s, unaffected
— next slices, in fixture-evidence order: `<drawing>`/`<legacyDrawing>` next,
`<hyperlinks>` last (a rewrite of `0.10.0-B4`'s existing filtering, not a new addition),
`<pageSetup r:id>` only if a fixture ever has one. Real-Excel reopen verification not yet
done for this slice.

**`<drawing>`/`<legacyDrawing>` restored (done).** Same opaque-fragment mechanism and
`rels_survived` gate as `<tableParts>`, spliced right after it in the same
`pageMargins`-then-`drawing`-then-`legacyDrawing`-then-`tableParts` block (schema order
confirmed against §8's real-XSD sequence — `pageSetup` through `smartTags` aren't emitted
yet, so `drawing`'s correct position is still simply "right after `pageMargins`").
`fixture5_chart_image_freeze_print.xlsm`'s only worksheet-level relationship is its
`<drawing r:id>`, so `check_source_references()` now reports `CLEAN` for this fixture too
— the second real fixture (after `fixture3`) to reach that state.
`fixture4_hyperlink_comment_name.xlsm`'s `<legacyDrawing r:id>` (VML comment shapes) is
restored as well, though that fixture's `.rels` also carries an r:id-backed hyperlink,
left unrestored at this point — `SOURCE_REFERENCE_LOSS` remained on `fixture4` until the
next change.

**`<hyperlinks>` r:id children restored (done) — all 7 real fixtures now `CLEAN`, the
last `SOURCE_REFERENCE_LOSS` gap closed.** A rewrite of `0.10.0-B4`'s shipped behavior,
not a new addition: B4's `extract_relationship_free_hyperlinks` unconditionally excluded
every r:id-bearing `<hyperlink>` child, since restoring one required the relationship-graph
reconnection that didn't exist at the time. Renamed to `extract_hyperlinks(xml,
include_relationship_backed: bool)`: location-only children are always kept (unchanged
from B4); r:id-bearing children are kept only when `save_xlsx_impl` passes
`rels_survived` as the flag. `fixture4`'s hyperlink is r:id-backed (external URL,
`TargetMode="External"`) — exactly the shape B4's own negative test asserted must NOT
survive; that test (`real_excel_external_only_hyperlink_omits_the_hyperlinks_container_entirely`)
is rewritten to `real_excel_external_hyperlink_survives_a_save`, asserting the opposite.
`fixture4` is now fully `CLEAN`, its last violation cleared. `SOURCE_REFERENCE_LOSS` is
eliminated from the entire current fixture set: all 7 real fixtures report `CLEAN` across
every `mechanical_check.py` category. Real-Excel reopen verification not yet done for
`tableParts`/`drawing`/`legacyDrawing`/`hyperlinks`.

**`D4` (deleted-sheet reachability cleanup, done) — closes Known gaps item 15.** A
deleted sheet's own worksheet-level `.rels` (and whatever it exclusively pointed at)
used to survive a save as an orphan — byte-identical, but unreferenced by anything in the
output, invisible to `check_roundtrip()`'s structural checks alone. Fixture→checker→writer,
this milestone's own hard gate: `check_deleted_sheet_cleanup()` (a package-reachability
BFS over the source's own relationship graph, independent of what the writer actually
does) was written and self-test-verified first; the Rust writer's
`deleted_sheet_prunable_parts` is a direct port of the same algorithm, wired into the
existing passthrough-building loop so a prunable part never enters `passthrough` — no
separate cleanup pass needed, since `carried_overrides` and both `carry_over_rels` calls
already only keep what's still present after that. A part shared between the deleted
sheet and anything else (a surviving sheet, or a workbook-level relationship) is correctly
kept; a part exclusively reachable from the deleted sheet is correctly pruned,
transitively. Two real bugs in the reachability computation were found and fixed on the
Python checker side before the Rust port even started: naively using `xl/workbook.xml`
itself as a "reachable elsewhere" root walks its own unfiltered `.rels`, which in the
source still lists the deleted sheet, silently reintroducing everything reachable from it
— fixed by threading an `exclude` set through every hop of the BFS, not just the initial
roots. Verified against the exact real scenario that exposed item 15 earlier this
session: the orphaned `.rels` is now genuinely absent, both `check_roundtrip()` (no false
positive) and the new checker report `CLEAN`. A dedicated shared-vs-exclusive-target
scenario was also run end to end through the CLI, confirming the exclusive target is
pruned while the shared one survives.

Two rows in the design doc's required test-case table (sheet rename, sheet reorder) are
deliberately marked N/A rather than implemented: this `Vm` has no rename/reorder
primitive, and adding VBA statement support purely to make a test-table row reachable
would invert this project's own hard gate.

**Plain (relationship-free) `<pageSetup>` restored (done).** Checked all 7 real fixtures
for the `<pageSetup r:id>` shape first — none has it. `fixture5`'s own `<pageSetup
paperSize="9" orientation="portrait" horizontalDpi="0" verticalDpi="0"/>` has no `r:id`,
and no fixture's `.rels` declares a `printerSettings` relationship, so the `r:id`-backed
case remains genuinely blocked on this milestone's own hard gate. But `fixture5`'s plain
`pageSetup` turned out to be a real, previously-uncaught bug in its own right: `pageSetup`
was never added to `_INLINE_WORKSHEET_ELEMENTS`, so it was silently lost on every save,
invisible to every existing checker category. New `check_page_setup()` — kept out of
`check_inline_worksheet_elements()`'s blanket list for the same reason
`check_internal_hyperlinks()` already is (the `<hyperlinks>` precedent): `CT_PageSetup`
genuinely CAN carry an `r:id`, so a blanket check would false-positive the day a real
`r:id`-backed fixture shows up. New `reader::root_tag_has_rid()` gates the writer
symmetrically: a plain `pageSetup` restores unconditionally (same mechanism as
`pageMargins`, no relationship dependency at all); one with `r:id` is dropped entirely
rather than risk a dangling reference with no `rels_survived` gate wired up for it.

`0.10.0-D`'s only remaining open item: `<pageSetup r:id>` itself (still no fixture with
that shape). Real-Excel reopen verification remains not done for `tableParts`/`drawing`/
`legacyDrawing`/`hyperlinks`/`D4`/plain `pageSetup`.

**Exit criteria** (unchanged from the original sketch, still the target): every untouched
unsupported XML node preserved byte- or semantically-equivalent, 0 Excel repair warnings, 0
loss of tables/validation/comments/etc., 0 broken chart/image relationships, successful
in-place save. The original "20+ fixtures" figure was aspirational, not load-bearing — actual
progress is gated by real fixture availability per element (7 fixtures total as of
`0.10.0-B` slice 4), consistent with `0.9.0-A`'s own precedent of shipping partial, honestly-scoped
wins rather than blocking on a headline number.

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
