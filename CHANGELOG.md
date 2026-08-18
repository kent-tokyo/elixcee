# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

### Added

- **`compat/vba-semantics/`, a new VBA value-correctness suite** — a genuinely different
  question from `compat/corpus/`'s own "does elixcee run without erroring": is the VALUE
  elixcee produces the one real, documented VBA semantics says it should be. Needs no
  oracle — `reference/*.mjs` are small, independently-checkable pure-JS reference
  implementations of documented real VBA semantics (banker's rounding, `Str()`'s
  leading-space quirk, `Val()`'s prefix parsing, `And`/`Or`/`Xor`/`Not`'s logical-vs-bitwise
  split, ...), used to compute 208 generated cases' expected outcomes programmatically. Six-
  verdict classification (`MATCH_DOCUMENTED_SEMANTICS`/`EXPECTED_ERROR`/`NONDETERMINISTIC`/
  `KNOWN_LIMITATION`/`BUG`/`UNCLASSIFIED`); `BUG`/`UNCLASSIFIED` both gate at 0. Current
  state: 198 `MATCH_DOCUMENTED_SEMANTICS` + 8 `EXPECTED_ERROR` + 2 `NONDETERMINISTIC` = 208,
  0 `BUG`, 0 `UNCLASSIFIED`, 0 `KNOWN_LIMITATION` (the suite's first run found one disclosed
  gap — array-out-of-bounds error message text — fixed in the same round; see "Fixed" below).
  See `compat/vba-semantics/README.md`.
- **CI now runs `@elixcee/xlsx`'s own tests.** `.github/workflows/ci.yml` gained a `node-js`
  job (Node 20/22 matrix): `packages/xlsx`'s TypeScript typecheck (with and without the DOM
  lib present) and all four `compat/differential/` suites (`utils`/`ssf-format`/`read`/
  `metadata`). Previously none of this ran anywhere except a developer's own machine, despite
  every command already working — verified live before wiring each one in, not assumed from
  this file's own previously-claimed numbers.
- **`packages/xlsx/scripts/audit-pack-contents.mjs`**, also wired into the new `node-js` CI
  job — asserts what `npm pack` would actually publish (every required file present —
  `LICENSE`, `README.md`, `THIRD_PARTY_NOTICES.md`, the four public entry points; nothing
  forbidden — `node_modules/`, `test/`, `.git`, `tsconfig*`; nothing unexpected under
  `src/internal/`), checked against `npm pack --dry-run --json`'s own real file list, not a
  reimplementation of npm's inclusion rules. Didn't exist at all before — a manual dry-run
  was clean (17 files, 338.8 kB), but nothing asserted this in CI.
- **`packages/xlsx/README.md`** — a package-level README, previously absent (npm's registry
  page would have shown only the `description` field, which opened with an unqualified
  "Drop-in replacement for xlsx" and never disclosed `write*`/`readFile` are unimplemented).
  States current scope honestly: what's implemented (all 33 `utils.*` exports, `SSF`,
  `XLSX.read()`, each with its own differential-testing numbers), what isn't
  (`write*`/`readFile`), and points to `THIRD_PARTY_NOTICES.md`/`docs/compatibility-known-
  defects.md` for licensing and disclosed divergences. `description` in `package.json`
  updated to match (no longer opens with an unqualified drop-in-replacement claim).
  Confirmed via `npm pack --dry-run` that the new README is actually included in the
  tarball (npm does this automatically regardless of the `files` array). This closes one of
  three concrete `packages/xlsx` alpha-publish blockers found this round — the other two
  (`"private": true`, missing `publishConfig.access`) are a real publishability policy
  decision, deliberately left alone here, not a mechanical fix.

### Fixed

- Array out-of-bounds errors used elixcee's own diagnostic wording (`"Array 'arr': index N
  out of bounds (len=N)"`) instead of real VBA's actual runtime error 9 message,
  `"Subscript out of range"` — found and disclosed as a `compat/vba-semantics/`
  `KNOWN_LIMITATION` when that suite first ran, fixed in the same round rather than left
  registered. Safe to change: `docs/agent-contract.md` already documents `message` as free
  text, not a stable/matchable field (`code`/`kind` are); `diagnose`/`diagnose-workbook`
  already read the rich per-failure detail (array name, index, bounds) from a structured
  `ResolutionFailureKind` side channel set alongside this string, not by parsing it — so
  nothing that actually depends on the old wording broke.
- All 3 READMEs' "XLSX.read()" section still claimed the browser-target WASM artifact
  "isn't wired into the package's public API yet" — true as of Phase 2B, but Phase 2C
  (already shipped in 0.3.0) added exactly that wiring (a `"browser"` export condition).
  Found while writing `packages/xlsx/README.md` (see "Added" above) and cross-checked
  against this file's own Phase 2C entry, which already correctly described the fix — only
  the top-level READMEs had gone stale. Corrected to state the real remaining caveat: the
  browser entry point assumes bundled consumption (its shared code has a CJS
  `require('ssf')`), not that it's unwired.
- `Dim x` (and `Dim x As <builtin type>`) was a complete no-op — the variable name was
  never recorded at all, so `IsEmpty(x)`/`x + 5`/any read before assignment hit "Undefined
  variable" instead of real VBA's `Empty`. An extremely common real-world VBA idiom
  (`Dim x`, then `If IsEmpty(x) Then ...` before ever assigning it), found on the very
  first run of the new `compat/vba-semantics/` suite (see "Added" above), not by
  source-code audit. `x` now registers as a real `Empty`-valued variable when `Dim`'d.
- `Val()` required its argument to parse as a number in its *entirety* — `Val("123abc")`
  was `0`, not real VBA's `123`. Real VBA's `Val()` parses a leading numeric prefix and
  stops at the first character that doesn't fit, only returning `0` when there's no valid
  numeric prefix at all. Found while designing the new `compat/vba-semantics/` value-
  correctness test suite — the same "never independently verified against documented
  semantics" bug class
  as `IsNumeric`. Scoped to the core grammar (optional sign, digits, one decimal point);
  real VBA's documented embedded-whitespace-stripping inside the numeric prefix
  (`Val("1 2 3")` == `123`) isn't attempted — no evidence it's needed.
- `Str()` was grouped with `CStr()` and shared its implementation — but real VBA's `Str()`
  reserves a leading space for the sign position on a non-negative number (`Str(459)` is
  `" 459"`, not `"459"`), a real, documented behavior difference from `CStr(459)` == `"459"`,
  not an alias of it. Found in the same systematic pass as `IsNumeric` below. Now its own
  arm, scoped to numeric inputs (the only case `Str()` is documented for); anything else
  falls back to the same plain formatting `CStr` uses. Previously untested; now covered.
- `IsNumeric` only checked whether its argument was already an `Integer`/`Float` Variant —
  `IsNumeric("123")` was `False`, missing real VBA's numeric-string recognition entirely.
  Found by a systematic pass over `eval_vba_func` for the same "grouped/never independently
  tested" bug class as `CBool`/`CInt`/`CLng` above. Now also accepts a string that parses as
  a plain decimal/scientific-notation number (after trimming whitespace) and `Empty`
  (coerces to 0 in a numeric context, matching real VBA). Deliberately not chasing VBA's
  fuller numeric-string grammar (currency symbols, locale-specific decimal separators,
  parenthesized negatives) — no evidence any of that is needed, and this project doesn't
  guess at locale-specific parsing rules. Previously entirely untested; now covered.
- `CInt`/`CLng` used Rust's default round-half-away-from-zero (`f64::round()`) instead of
  real VBA's banker's rounding (round-half-to-even) — `CInt(0.5)` was `1`, not `0`. Found by
  auditing for the same bug class the `Round()` fix (below) had already been fixed for:
  `to_i64_rounded` (used by `\`/`Mod` operand coercion) already documented "the same
  round-half-to-even ... that CLng/Round use," but `CInt`/`CLng`'s own arm never actually
  used it. Now reuses that exact existing helper — the `test_vba_clng` test had silently
  computed a tie-case value (`CLng(-2.5)`) without ever asserting on it, which is likely how
  this went unnoticed; that assertion is filled in now, plus dedicated tie-case coverage.
- `Round(number, negativeDigits)` (e.g. `Round(1234.5, -2)`) silently returned a plausible
  answer instead of erroring — real VBA's own `Round()` raises "Invalid procedure call or
  argument" for a negative digit count (unlike `WorksheetFunction.Round`/Excel's `ROUND()`
  formula, which both accept negative digits to round left of the decimal point). Found and
  disclosed, not fixed, in the 0.3.0 round; fixed now.
- `Now`/`Date`/`Time` returned a Rust debug-formatted `SystemTime { tv_sec: ..., tv_nsec:
  ... }` string regardless of which of the three was called — visibly wrong if ever
  displayed or compared, not just imprecise. `Date()` now returns a real `Variant::Date`
  matching the actual system clock (Excel-serial epoch math, same `25569` constant the
  formula engine's own `NOW()` already uses); `Time()`/`Now()` return a numerically correct
  `Variant::Float` (0.0-1.0 for `Time()`, serial-plus-fraction for `Now()`) rather than a
  `Variant::Date`, since `Variant::Date` is whole-day-only (`i64`) in this codebase and
  can't carry a sub-day component without a shared-type change — so `TypeName(Time())`/
  `TypeName(Now())` report `"Double"`, not real VBA's `"Date"`. A disclosed, narrower gap
  than the debug-string bug, not a silent one.
- The bare no-parens form (`Date` without `()`, real VBA allows omitting `()` on zero-arg
  functions) didn't parse as a function call at all — `Expr::Var("date")` always hit
  "Undefined variable". Found alongside the fix above, fixed in the same round: a bare
  identifier now falls back to calling `Date`/`Now`/`Time` as zero-arg functions only after
  every other variable/constant lookup fails — scoped to exactly these three names (the only
  `eval_vba_func` entries that accept zero arguments) rather than a general "any unrecognized
  identifier might be a function call" rule, so a genuine variable-name typo still errors the
  same way it always did (verified with a regression test).

## [0.3.0]

Root `elixcee` (Rust crate + Python package) only — `elixcee-types` stays `0.1.0`
(unchanged, no source changes this release) and `@elixcee/xlsx` stays
`0.0.0-development`/unpublished. Verified via a fresh-venv install of a locally-built
wheel, not just `cargo test` (see `Fix`/`Sgn`/`Round`/`CBool` below, all re-checked
through the real Python API after install).

### Added

- **`Fix`, `Sgn`, and `Round` VBA functions.** Root-caused via an automated pass over the
  581-scenario corpus's remaining non-parse-error failures (see below), not guessed at:
  28 of 41 turned out to be `Unknown VBA function` for these three, previously-missing,
  ordinary built-in functions (13 `Sgn`, 12 `Fix`, 3 `Round`) — not deliberate negative
  tests (2 more of the 41 were the related `CBool` bug below, and 1 was a low-value
  `Timer()` left unimplemented; see "Known gaps" in `ROADMAP.md` for the full 41-way
  breakdown). `Fix` truncates toward zero (unlike `Int`, which floors — `Fix(-3.9)` is `-3`, not
  `-4`). `Sgn` returns -1/0/1. `Round` uses real VBA's own banker's rounding
  (round-half-to-even), which is a genuinely different function from
  `WorksheetFunction.Round`/Excel's `ROUND()` formula (round-half-away-from-zero) — `Round`
  does *not* alias or share an implementation with the pre-existing `WorksheetFunction.Round`
  arm; verified both give different, each individually correct, answers on the same tied
  input (`Round(2.5)` is `2`; `WorksheetFunction.Round(2.5)` is `3`).

### Fixed

- `CBool` was grouped with `CLng`/`CInt` and returned a numeric `Variant::Integer` via the
  same numeric-coercion path they use — so `CBool(5)` returned `5` typed `Long`, not `True`
  typed `Boolean` (`TypeName` confirms this live), and `CBool("True")`/`CBool("False")`
  errored outright trying to parse the literal string as a number. Found while implementing
  `Fix`/`Sgn`/`Round` above (same corpus failures also involved `CBool`), not something the
  corpus itself flagged directly — no scenario happened to check `CBool`'s return *type*, only
  whether the string-literal call errored. Now its own arm: a genuine `"true"`/`"false"`
  string (case-insensitive) converts directly, anything else numeric-coerces to boolean, and
  the result is always a real `Variant::Boolean`.
- Single-line `If cond Then stmt [Else stmt]` (no `End If`) now parses — previously
  unsupported at all (`parse_if` unconditionally required a newline right after `Then`).
  Identifier-led statements (assignment, sub call, array/field write — whatever
  `parse_ident_stmt` already covers) are recognized inline, and `Exit For|Do|Sub|Function`
  / `GoTo <label>` are handled explicitly rather than routed through `parse_ident_stmt` —
  the first implementation didn't do this and silently turned `If done Then Exit Sub` into
  a no-op that let execution fall through instead of exiting, caught in review before this
  ever shipped (verified live: `y = 99` after `If x > 0 Then Exit Sub` no longer runs).
  Anything still unrecognized degrades to `Stmt::Unsupported`, same precedent as
  `parse_set`'s unmodeled-target fallback and the identical fallback an ordinary
  unparenthesized bare sub call already hits in block-form VBA — not a new risk this adds.
  This was discovered, not hunted for: it's what the comma-`Dim` fix below unmasked on the
  4 corpus scenarios that fix's own parse-error count didn't reach 0 on. With this fix,
  the corpus's parse-error count is genuinely 0/581 (verified by rerunning
  `compat/corpus/run-elixcee.mjs` — the *committed* `compat/corpus/results/` snapshot
  still shows the pre-fix numbers, since that file is regenerated on demand, not on every
  commit; don't read it as current without rerunning it).
- Comma-separated `Dim`'s built-in/bare-declarator branch (below) lost its old tolerance
  for trailing per-declarator syntax it doesn't model (e.g. `Dim s As String * 10`'s
  fixed-length-string suffix) when the comma loop was added — the first implementation
  returned immediately instead of consuming up to the next comma, so that syntax now
  hard-failed `eat_eol()` instead of being silently skipped like it always was. Caught in
  review before shipping; both the fixed-length-string case and its combination with a
  second comma-declarator (`Dim s As String * 10, i As Integer`) are covered by tests.
- `Not` now does a real bitwise complement on numeric operands (`Not 5` is `-6`, matching real
  VBA), instead of coercing the operand to truthy/falsy first. Only a genuine `Boolean`
  operand still gets logical negation — the same logical-vs-bitwise split `And`/`Or`/`Xor`
  (Phase 2C) already used, `Not` just hadn't been reconciled with it yet (see CHANGELOG's own
  0.2.0 Known limitations: `Not 5 And 3` used to diverge from real VBA's `2`; it now matches).
- Comma-separated multi-declarator `Dim` (`Dim a As Integer, b As Range`) — 0.2.0's last
  documented gap on the 581-scenario VBA corpus's own parse-error surface. Previously, a
  declarator with a non-built-in type (e.g. `b As Range`) returned from `parse_dim` as soon
  as it finished, leaving `, nextDecl` unconsumed; the statement dispatcher's `eat_eol()`
  then hit the stray comma and hard-failed the whole macro. `parse_dim` now loops over every
  comma-separated declarator (`parse_dim_declarator`), wrapping 2+ into a new `Stmt::DimMulti`
  that the VM/`check`/name-resolution passes execute or inspect by replaying each inner
  declarator through the exact same code path a single-declarator `Dim` already used — no
  new semantics, just no longer losing the rest of the line. Verified against the real
  corpus, not just new unit tests: elixcee's own parse-error count on the 581 scenarios goes
  from 8 to 4.
Everything below was previously listed under `[Unreleased]`; this release closes that
section rather than adding new scope. Developed in two internal phases (2B, then 2C after
an integration review found real gaps in 2B's first pass): 2B added the VBA object model
and a working `XLSX.read()` MVP; 2C closed the parser-level and `read()`-completeness gaps
that review surfaced. See "Compatibility" and "Known limitations" below before reading
this as a finished 90-point milestone — the VBA-macro-vs-Microsoft-Excel axis was never
attempted this release (no Windows/Excel environment available), and that gap alone means
a full compatibility claim can't be made on this release's evidence, however solid the
other axes are.

### Added

- **`@elixcee/xlsx` compatibility groundwork (Phase 0)**: investigation and scaffolding for a planned npm package that would be a drop-in replacement for `xlsx@0.18.5` (SheetJS) —
  - `docs/xlsx-compatibility-goal.md`, `docs/xlsx-architecture.md` (ADR: target crate/npm workspace shape, and concrete resolutions for the `formula`↔`vm` circular type dependency and `reader.rs`'s path-only I/O — neither executed yet), `docs/xlsx-security-model.md` (resource-limit design, prototype-pollution-safe key handling), and `docs/licensing.md` (elixcee is MIT; the `xlsx` package and its 7 transitive SheetJS dependencies are all Apache-2.0)
  - New `compat/` Node.js project (not part of the Rust build): `compat/oracle/generate-manifest.mjs` installs and introspects the real `xlsx@0.18.5` at runtime (both its CJS and ESM entrypoints) rather than guessing from documentation, producing the committed `compat/oracle/api-manifest.json`; `compat/differential/classify.mjs` defines the six-value compatibility verdict (`MATCH`/`INTENTIONAL_SECURITY_DIVERGENCE`/`UNSUPPORTED`/`BUG`/`ORACLE_AMBIGUITY`/`NONDETERMINISTIC`) future differential tests will use, with a `run-demo.mjs` proving the plumbing
  - No `elixcee` Rust source, Python binding, CLI, or test behavior changed by this milestone
- **`@elixcee/xlsx` Phase 1A-1C**: `packages/xlsx` now implements every one of the real oracle's 33 `utils.*` runtime exports — `Object.keys(XLSX.utils)` matches the oracle exactly, both content and insertion order — differential-tested against the real oracle throughout (550+ permanent public-API fixtures across `compat/differential/`, plus a separate 1831-case internal SSF-backend conformance suite) —
  - Address/workbook utilities (`encode_*`/`decode_*`, `book_new`, `book_append_sheet`, `book_set_sheet_visibility`), worksheet mutation (`aoa_to_sheet`/`sheet_add_aoa`, `json_to_sheet`/`sheet_add_json`, `sheet_add_dom`), cell lookup (`sheet_get_cell`), JSON extraction (`sheet_to_json`/`sheet_to_row_object_array`), HTML export (`sheet_to_html`), DOM table conversion (`table_to_sheet`, `table_to_book` — duck-typed against whatever DOM-like object is passed, no DOM library imported at runtime), formula extraction (`sheet_to_formulae`), cell metadata (`cell_set_hyperlink`, `cell_set_internal_link`, `cell_add_comment`, `sheet_set_array_formula`), sheet-visibility constants (`consts`), and text export (`sheet_to_csv`, `sheet_to_txt`)
  - `format_cell`/`cell_set_number_format` are backed by the real `ssf@0.11.2` engine (Phase 1B-2B) — `packages/xlsx`'s only runtime dependency, isolated behind a single adapter file (see `docs/xlsx-architecture.md`'s "SSF backend" decision and `THIRD_PARTY_NOTICES.md`); one genuine upstream indirection-table bug (numFmtIds 67-71) was found and corrected in that adapter, including a follow-up fix so the correction never shadows a caller's own `opts.table` override
  - Four DoS-shaped divergences from the real oracle, all empirically confirmed (not assumed) and registered as intentional safety divergences: `encode_col(Infinity)` (the oracle hangs; elixcee rejects non-finite indices) and a crafted full-grid `!ref` fed to `sheet_to_formulae`/`sheet_to_csv`/`sheet_to_txt`/`sheet_to_json`/`sheet_to_html` (the oracle takes 12s+ / doesn't return within 25s; elixcee caps iteration at 5,000,000 cells — sizing measured at 100K/1M/5M/10M cells, see `docs/limits.md`)
  - Six security fixes, all live-confirmed defects in the oracle itself, not hypothetical: two prototype-corruption fixes (`book_append_sheet`, Phase 1A, and `table_to_book`'s internal sheet-to-workbook construction, Phase 1C — both let a caller-controlled sheet name of `"__proto__"` reassign a `WorkBook.Sheets` object's own prototype), two `sheet_to_json` fixes (an explicit `opts.header` array containing `"__proto__"` could silently drop a primitive column value or reassign a row object's own prototype to a Date/object cell value, Phase 1B-3), and two `sheet_to_html` fixes (Phase 1C — `data-t`/`data-v`/`data-z`/`id` attributes built with zero escaping let a cell value or `opts.id` containing `"` inject a live event handler; `cell.l.Target` embedded into `href="..."` with no scheme check let a `javascript:` hyperlink execute code on click). A third `sheet_to_html` finding — `cell.h`'s raw-HTML passthrough — is reproduced, not fixed, since it is a documented, intentional field (see `docs/compatibility-known-defects.md`)
  - TypeScript types classified as EXACT/SAFE_EXTENSION/MISSING/INCOMPATIBLE against the real oracle's own `types/index.d.ts` (`docs/typescript-compatibility.md`) rather than described loosely; `table_to_sheet`/`table_to_book`/`sheet_add_dom` mirror the oracle's own `data: any` (not `HTMLTableElement`) and are compile-tested both with and without DOM lib present (`tsconfig.no-dom.json`, `test/smoke-dom.ts`)
  - `packages/xlsx`'s own `LICENSE`/`THIRD_PARTY_NOTICES.md` are now included in the npm tarball's own `files` (previously only the repo root had them, which never reached npm consumers); package version reset from a stale `0.1.0-phase1b1` to `0.0.0-development` pending a real publish candidate
  - Still explicitly out of scope: XLSX/ODS file reading or writing (`read`/`readFile`/`write*`), any Rust↔JS bridge, and npm publish
- **Range object variables, Union, Areas, SpecialCells, and multi-area Paste** (Milestone B7c): the VBA object-reference layer on top of the B7a/B7b foundation —
  - `Dim rng As Range` / `Set rng = Range(...)` now work with real reference semantics: `Set`-assigned variables live in a new `Vm.object_variables` namespace (`ObjectRef`), kept separate from `Vm.variables` (`Variant`s) rather than adding a `Variant::Object` variant — `Variant` is defined in the shared `elixcee-types` crate the WASM bridge also consumes. Because `ObjectRef::Range` is just B7a's `RangeRef` (coordinates, no cell values), two variables holding the same `RangeRef` already alias the same cells in `Vm.sheets` — real `Set` reference semantics with no `Rc<RefCell<_>>` needed
  - `Union(range1, range2, ...)` combines ranges into one multi-area `RangeRef`; `.Areas.Count` and `.Areas(n)` (1-based) enumerate them — `.Value`/`.Formula`/`.Areas.Count` reuse the existing generic `<var>.<field>` grammar (disambiguated at runtime by which namespace holds `var`), `.Copy`/`.Areas(n)`/`.SpecialCells(...)` are new grammar
  - `SpecialCells(xlCellTypeVisible)` consumes B7b's `sheet_visibility` directly, splitting each area into the Cartesian product of its maximal visible row-bands and column-bands
  - Multi-area Paste: the one shape both sides multi-area with matching `Areas.Count` and per-area shapes now actually executes (pairwise, in order) instead of only diagnosing — every other multi-area shape (count/shape mismatch, or either side single-area) is unchanged and stays diagnose-only via the existing `MULTI_AREA_*` root causes. `Transpose:=True` on this shape still falls through to the diagnose-only path rather than silently writing un-transposed data; merged-cell conflict checking isn't applied to it
  - `ActiveSheet` works as a dynamic sheet qualifier (`ActiveSheet.Range(...)`, `.Cells(...)`, ...); `ThisWorkbook`/`ActiveWorkbook` need no new resolution logic (elixcee only ever loads one workbook) and parse as a plain `Worksheets(...)`/`Sheets(...)` reference
  - Not supported: `Set ws = ActiveSheet` / `Set wb = ThisWorkbook` as Worksheet/Workbook object variables (only `Range` object expressions are `Set`-able); a multi-area source pasted into a single-area destination (or the reverse) actually executing; any new `--json` field (this milestone changes `MULTI_AREA_PASTE_UNSUPPORTED`'s firing condition but adds no new field or code)
- **Hidden row/column evidence** (Milestone B7b): `diagnose`/`diagnose-workbook` now report when a `.Copy`'d range overlaps hidden rows/columns —
  - New `vm::Interval`/`vm::SheetVisibility` types, threaded from XLSX's `<row hidden="1">`/`<col min=".." max=".." hidden="1">` into `Vm.sheet_visibility` the same way `merged_ranges` already is; ODS is explicitly deferred (its reader doesn't expand `table:number-rows-repeated`, so a hidden-row flag can't map to a correct absolute row number yet)
  - New `Vm::hidden_cells_observation()` computes the evidence on demand from the existing `Vm.clipboard` + `sheet_visibility` — no new stored side channel
  - A new sibling JSON field `observations` (not folded into `root_causes`, which means "why it failed" — this isn't a failure), present only when non-empty, on both success and failure: `{"code":"RANGE_CONTAINS_HIDDEN_CELLS","certainty":"observed","range":{...},"visibility":{...},"message":"..."}`
  - `diagnose-workbook` gets the same field via `FixtureResult::Passed`/`::Failed`, though honestly no additional value over a single `diagnose` call — hidden-row/column metadata is structural (workbook layout + macro text), not input-dependent
  - Copy/Paste behavior itself is unchanged — hidden cells still copy/paste exactly as before; this is observability only, laying groundwork for `SpecialCells(xlCellTypeVisible)` (B7c)
- **Multi-area Range foundation** (Milestone B7a): `Range("A1:A3,C1:C3")`-style disjoint ranges now have an underlying model —
  - New `vm::Rect`/`vm::RangeRef` types (`{ sheet, areas: Vec<Rect> }`); the existing single-rect `parse_range_addr`/`SheetRange` and their ~11 call sites are untouched — only Copy/Paste resolve through the new `parse_multi_area_addr`
  - `.Copy` now accepts a comma-separated multi-area source; `.Paste`/`.PasteSpecial` was diagnose-only for every multi-area shape in v1, even a fully-matching one — 4 new classified failures instead: `MULTI_AREA_TO_SINGLE_AREA_PASTE`, `MULTI_AREA_COUNT_MISMATCH`, `MULTI_AREA_SHAPE_MISMATCH`, and the catch-all `MULTI_AREA_PASTE_UNSUPPORTED` (the fully-matching case has since started executing — see Milestone B7c below)
  - Each area's evidence is `{"address", "rows", "columns"}`, matching the completion-condition JSON's own shape
  - `Union()`, the `Areas`/`Areas.Count`/`Areas(n)` property, and `Dim rng As Range`/`Set rng = ...` object variables were unsupported at this milestone — `Variant` gained no Range variant (still true: see Milestone B7c below for how these landed without one)
  - Foundation for B7b (hidden/filtered rows) and B7c (`SpecialCells(xlCellTypeVisible)`), sequenced ahead of shrinking (B5b) since most structural failures need this range model first
- **`diagnose-workbook` subcommand** (Milestone B6d): combines `test-workbook`'s (B5a) generated-case search with `diagnose`'s (B6a–B6c2) root-cause classification —
  - Reuses `test-workbook`'s exact fixture format, strategies, and deterministic `--seed`/`--case` replay; runs each case with `Vm::strict_resolution` on and enriches classifiable failures with the same `ResolutionFailureKind` → `RootCause` pipeline `diagnose` uses, via a new `pub(crate) diagnose::root_causes_json` entry point
  - New `--cases N` flag overrides the fixture's declared case count for one invocation (scoped to this subcommand; `test-workbook` itself is unchanged)
  - Output is `test-workbook`'s existing JSON shape plus one sibling `root_causes` field (`[]` when unclassified)
  - Most root causes are structural (merge/shape/protection kinds) and fire identically regardless of input — this command's actual value is for input-dependent kinds like `ARRAY_INDEX_OUT_OF_BOUNDS`, where a drawn value can flip an index in or out of bounds across cases
  - Shrinking (minimizing a failing case's inputs) is deliberately deferred to a later phase
- **Merged-cell-aware Paste diagnostics** (Milestone B6c2): `diagnose` now classifies Copy/Paste operations that conflict with a merged-cell layout —
  - `PASTE_INTO_NON_ANCHOR_MERGED_CELL`: the destination cell falls inside an existing merge but isn't that merge's own top-left cell (pasting into the top-left cell itself, the normal way to write to a merged cell, is unaffected)
  - `PASTE_PARTIAL_MERGED_RANGE`: a multi-cell destination partially overlaps one or more merges without fully containing them
  - `PASTE_MERGE_LAYOUT_MISMATCH`: the source's and destination's merged-cell layouts, compared by relative position (accounting for `Transpose:=True`), don't match
  - `WorkbookSheet.merged_ranges` parsed from XLSX `<mergeCell ref="...">` and ODS `table:number-columns-spanned`/`table:number-rows-spanned`, threaded into a new `Vm.merged_ranges` map
  - Unconditional hard errors in every mode that executes the macro (`run`/`diagnose`/`test-workbook`), matching real Excel's Error 1004 regardless of `On Error` state — same posture as B6b/B6c
  - Scope stays Paste-only; multi-area (`Areas`) ranges, hidden/filtered rows, and AutoFilter visible-cells-only copy remain deferred
- **Sync `XLSX.read(bytes)` MVP and the `elixcee-wasm` bridge** (Phase 2B): a real, working file-reading entry point for `@elixcee/xlsx`, backed by WebAssembly, callable with no `await init()` —
  - `src/reader.rs`'s `read_workbook` was generalized (pure extraction, no behavior change to the path-based entry point) into `read_workbook_from_archive<R: Read + Seek>` plus a new `pub fn read_workbook_from_bytes(bytes: &[u8])`; `zip::ZipArchive` was already generic, so this needed no new dependency
  - New `crates/elixcee-wasm` crate (first real use of `wasm-bindgen` in this project — deferred until this exact phase per `docs/xlsx-architecture.md`'s ADR). Node gets `wasm-pack --target nodejs` glue (`readFileSync` + synchronous `WebAssembly.Module`/`Instance` construction — genuinely synchronous, verified live by `require()`-ing it and calling the export with no `await`); the browser target inlines the compiled `.wasm` as base64 into its own glue and calls `initSync` itself rather than depending on a bundler's `.wasm` loader (verified: esbuild does not resolve `.wasm` imports by default) — both design choices come from a feasibility spike recorded in `docs/xlsx-architecture.md`'s "Phase 2B-0" section, including its one open item (an oft-cited "Safari enforces a ~4KB sync-compile ceiling" claim could not be substantiated from current MDN docs and is reported unverified, not as fact)
  - `XLSX.read(data, opts)` added to `packages/xlsx` — accepts a `Buffer`/`Uint8Array` or `opts.type === 'base64'`; differential-tested against the real oracle via real file-format round-trips (`compat/differential/xlsx-read.test.mjs`): 9 MATCH + 2 registered `UNSUPPORTED` out of 11 cases, on the scope it actually claims (`SheetNames`, per-sheet `!ref`/`!merges`, per-cell `{t,v}` — no formulas, no styles/dates, no `!rows`/`!cols` yet)
  - `zip`'s default features (`zstd-sys`, `getrandom`'s aes-crypto path) don't compile for `wasm32-unknown-unknown` in this toolchain; trimmed to `default-features = false, features = ["deflate"]` after confirming (by grep) the codebase never uses another compression method — a real fix, not a workaround
  - Node-only for now: the browser-target artifact is built and verified at the bridge level, but `packages/xlsx`'s public `read()` doesn't yet dispatch to it (no `browser` export condition wired up)
- **Oracle-neutral VBA differential corpus infrastructure** (Phase 2B): a reusable, backend-swappable differential-testing pipeline for VBA macro execution, under `compat/corpus/` —
  - Backend-agnostic scenario schema (`compat/corpus/SCHEMA.md`), 581 generated scenarios across 25 categories, an elixcee runner, a LibreOffice UNO runner, a normalizer, and a classifier (`compat/corpus/classify.mjs`, its own file — reuses `compat/differential/classify.mjs`'s verdict vocabulary and anti-laundering discipline rather than importing it, since that file's registries are keyed to the npm API surface, not VBA scenarios) with one new verdict this domain needed: `ORACLE_UNAVAILABLE`
  - Every result record carries an explicit `oracle` field (`"libreoffice"` today; `"microsoft_excel"` is defined in the schema but has never produced a record — see "Compatibility" below). LibreOffice and Excel results are never merged into one number by this pipeline
  - An Excel COM adapter **contract** (`compat/oracle-excel-com/CONTRACT.md`, I/O schema, PowerShell scaffold, Windows execution instructions) is defined but explicitly marked `UNVERIFIED` — no Windows/Excel environment exists in this project's current toolchain to run it
- **VBA foundational syntax — `Mod`/`\`/`^`/logical operators/typed `Function`/`With Range(...)`/`Set` object references** (Phase 2C): closes gaps a Phase 2B integration review found by direct execution — ordinary VBA syntax, not advanced object-model features, that previously stopped a macro at the parse stage —
  - `Mod`, `\` (integer division, real VBA round-half-to-even rounding of each operand before dividing — e.g. `5 \ 0.5` divides by `0`, a real division-by-zero, not a bug), `^` (exponentiation, left-associative, binds tighter than unary minus), and infix `And`/`Or`/`Xor`/`Not` in expressions, all at real VBA operator precedence (`^` > unary `-` > `*`/`/` > `\` > `Mod` > `+`/`-` > `&` > relational > `Not` > `And` > `Or` > `Xor` — pinned by a test asserting `2 + 3 * 2 ^ 2 == 14`)
  - `With Range("A1:B2") ... End With` (previously only `With Sheets(...)` worked); typed `Function` parameters and return types (`Function f(x As Integer) As Integer`)
  - `Set ws = ActiveSheet` / `Set wb = ThisWorkbook` are now real object variables (a new `ObjectRef::Worksheet`/`ObjectRef::Workbook`, alongside B7c's `ObjectRef::Range`) rather than a silent no-op — `ws.Range(...)`/`wb.Worksheets(...)` work through them afterward
  - Measured against the frozen 581-scenario corpus from Phase 2B: parse-error count **132 → 8**. The 8 remaining are all comma-separated multi-declarator `Dim` (`Dim a As Integer, b As Range`, only one declarator per `Dim` parses today) — not one of this phase's items, flagged as the next highest-value parser fix since it's now the entire remaining corpus parse-error surface
  - `&` was moved to its own, correctly-lower precedence tier below `+`/`-` (previously equal precedence) to match real VBA — a small behavior change beyond the 8 named items, required by the precedence table above
- **`XLSX.read()` completeness — formulas, dates, dimension, hidden rows/cols, browser wiring** (Phase 2C): closes every gap Phase 2B's `read()` MVP left open —
  - Formula text (`.f`) captured from `<f>...</f>`; `!rows`/`!cols` (hidden row/column metadata, already used by VBA's `SpecialCells` since B7b) now surfaced in `read()`'s output, gated behind `opts.cellStyles` to match the oracle's own gating
  - `.w` (formatted display string) and date-typed cells (`t:'d'`) — the largest item, requiring `styles.xml` (`<numFmts>`/`<cellXfs>`) and `<workbookPr date1904>` parsing that `reader.rs` did not do at all before this phase — landed in full, not partially: `.w` always computed, `.z` gated behind `opts.cellNF`, `t:'d'` gated behind `opts.cellDates` and a date-like resolved format
  - A real oracle inconsistency was found and deliberately reproduced, not "fixed": in a `date1904` workbook, the oracle's `.w` display string reflects the 1462-day epoch shift but its `cellDates` `.v` Date object does not, because the oracle's own read-direction date conversion doesn't accept `date1904` while its write-direction one does
  - `packages/xlsx`'s `exports["."]` gained a `"browser"` condition routing to the already-built inlined-bytes/`initSync` artifact from Phase 2B — confirmed live via a real subprocess `import.meta.resolve()` + `read()` call under `--conditions=browser`, not just "should work." The browser entry point still assumes bundled consumption (its shared code has a CJS `require('ssf')`) — not literal no-build `<script type=module>` usage
  - Both Phase 2B `UNSUPPORTED_ALLOWLIST` entries under `'read'` (empty-string cells, `<dimension>`) removed now that their underlying `reader.rs` defects are fixed — the registry is empty again

### Changed

- `read_workbook`'s XLSX-archive-consuming body is now generic over `R: Read + Seek` internally (`read_workbook_from_archive`); the public `read_workbook(path)` signature and behavior are unchanged
- Root `Cargo.toml`'s `zip` dependency narrowed from its default feature set to `deflate`-only (see above)
- Binary string-concatenation (`&`) is now its own precedence tier, below `+`/`-` and above relational operators, matching real VBA (see Phase 2C above)

### Fixed

- A silent-wrong-result bug in the new (this release) matching-shape multi-area Paste: `Transpose:=True` was being ignored, writing un-transposed data instead of either transposing correctly or erroring. Caught during the same milestone's self-review, before ever reaching a released version — fixed with a regression test (`matching_shape_multi_area_paste_with_transpose_still_errors_instead_of_silently_mis_pasting`)
- `ssf@0.11.2`'s numFmtId 67-71 indirection-table bug (see Phase 1B-2B above) — carried forward from before this release, listed here for completeness since it ships in 0.2.0's first tagged release
- Empty-string cells (`<c t="str"><v></v></c>`) were silently dropped instead of read as `{t:'s', v:''}`; `<dimension>` was never parsed, so `!ref` always came from the populated-cell bounding box even when a file legitimately declared a wider `<dimension>` (Phase 2C, both shared by `read_workbook` and `read_workbook_from_bytes`)
- `<col hidden="true">` wasn't recognized (only `hidden="1"`) — the oracle's own writer emits `"true"` for columns but `"1"` for rows, an xsd:boolean inconsistency that silently dropped hidden-column detection until caught via a live round-trip (Phase 2C)
- `worksheet_json` always emitted a colon-form `!ref` (`"B2:B2"`) even for single-cell sheets; the oracle collapses `start === end` to a colon-less ref (`"A1"`) — surfaced once Phase 2C's `.w`/date fixtures started using single-cell sheets (Phase 2C)
- An object-qualifier parsing bug (`<var>.Worksheets(`/`.Sheets(`) that could misfire without guarding on an immediate `(` — caught in the same phase's self-review before release (Phase 2C)

### Compatibility

Two independent oracle-differential efforts ran across 2B and 2C. Read both before treating any "compatibility" claim below as broader than what it says.

- **`@elixcee/xlsx` vs. the real `xlsx@0.18.5` npm package** (`compat/differential/`, Node-side, oracle always available): 512 MATCH + 14 registered intentional divergences on the `utils.*` surface, 1831/1831 MATCH on the SSF number-format engine, 34/34 export metadata matches, and `read()` now at **19/19 MATCH, 0 UNSUPPORTED, 0 BUG, 0 UNCLASSIFIED** (up from 2B's 9 MATCH + 2 UNSUPPORTED, and on a widened comparison scope that now includes `!rows`/`!cols`, `.w`, and `.z`). This axis has a real, complete oracle and these numbers are direct measurements.
- **VBA macro execution vs. LibreOffice** (`compat/corpus/`, oracle: `"libreoffice"`): of 581 generated scenarios, only **1 produced an actual `MATCH` comparison** and 2 were `NONDETERMINISTIC`; **578 are `ORACLE_UNAVAILABLE`** — LibreOffice, driven headless via `getScriptProvider().getScript().invoke()` in this project's sandboxed environment, hangs indefinitely (confirmed >90s, no CPU activity) on any `Range`/`Cells` access, which is most of what the corpus exercises. Non-object-model code runs and compares fine (proven by a dedicated smoke scenario). **This is a real, reproducibly-measured negative result, not a partial success**, and it is **unchanged by Phase 2C** — fixing the LibreOffice hang was explicitly out of scope this phase (it doesn't raise elixcee's own product quality, only this one oracle's usability). What Phase 2C *did* measure against the same 581 scenarios is elixcee's own parse-error rate in isolation (132 → 8, see "Added" above) — a real, useful signal, but not a LibreOffice-comparison signal.
- **VBA macro execution vs. Microsoft Excel**: **not attempted, at all, across either phase.** No Windows or licensed Excel environment exists in this project's toolchain. `compat/oracle-excel-com/`'s adapter is a contract only — treat every "LibreOffice" result above as informative but **not** a proxy for Excel compatibility; LibreOffice's own VBA support is its own compatibility layer, not Microsoft Excel.

### Known limitations

- `Not` still evaluates via boolean-truthy coercion, while `And`/`Or`/`Xor` (Phase 2C) do real bitwise math — so `Not 5 And 3` doesn't match real VBA's bitwise result (`2`). Phase 2C's own scope was adding the operators, not reconciling `Not`'s pre-existing evaluation semantics with them.
- Comma-separated multi-declarator `Dim` (`Dim a As Integer, b As Range`) isn't parsed — only one declarator per `Dim` statement. This is now the entire remaining parse-error surface on the 581-scenario corpus (8/581).
- Matching-shape multi-area Paste is the only multi-area Paste shape that executes; every other combination (count/shape mismatch, single↔multi either direction) remains diagnose-only.
- The LibreOffice headless `Range`/`Cells` hang described under "Compatibility" is unresolved — root-caused (headless UNO script invocation, not a scenario-specific issue) but deliberately not fixed this release (out of scope both phases), so the corpus cannot yet produce broad VBA-vs-LibreOffice compatibility signal in this environment.
- The Excel COM adapter is a contract and PowerShell scaffold only; it has never been run, against anything, by anyone, in this project's history to date. This is the single largest remaining gap toward any formal Excel-compatibility claim.
- `@elixcee/xlsx`'s browser `read()` entry point assumes bundled consumption (its shared code `require()`s `ssf`, a CJS-only dependency) — verified via a real subprocess resolving the `"browser"` export condition, not via an actual bundler build (none is installed in this project's toolchain).

## [0.1.2]

Root `elixcee` (Rust crate + Python package) only — `elixcee-types` stays
`0.1.0`/unpublished and `@elixcee/xlsx` stays `0.0.0-development`/unpublished;
neither is part of this release. No public Rust path, CLI behavior, Python
API, or `--json` output shape changed — this release is accuracy,
structure, and test coverage, not new capability.

### Fixed

- **`xml_unescape` (`src/reader.rs`)**: only decoded the 5 named XML entities
  (`&amp;`/`&lt;`/`&gt;`/`&quot;`/`&apos;`), not numeric character
  references (`&#65;`/`&#x41;`); its chained `.replace()` calls also
  double-unescaped input like the literal text `&amp;lt;` (first pass turns
  it into `&lt;`, the very next pass then corrupts that into `<`).
  Rewritten as a single forward pass that decodes numeric references too,
  with the entity-body search bounded to a small window so a run of
  unterminated `&` stays O(n), not O(n²)
- **ODS `table:number-columns-repeated`/`table:number-rows-repeated`**
  (`src/reader.rs`): never read at all — ODS's sparse-representation
  mechanism, used by real producers (LibreOffice) for any run of matching
  cells/rows, not just trailing empty ones, so a real value following a
  repeated block landed at the wrong row/column. Tracked as an arithmetic
  skip rather than a literal expansion loop, so a pathological repeat count
  costs O(1)

### Internal

- **`elixcee-types` crate** (`crates/elixcee-types/`): `ExcelError`,
  `Variant`, `CellContent`, `serial_to_display`, `parse_cell_addr`/
  `parse_range_addr`, and the date-serial helpers `serial_to_ymd`/
  `is_leap`/`days_in_month` extracted from `src/vm/mod.rs`/
  `src/formula/eval.rs` into a new std-only, zero-dependency workspace
  member — the precondition (per `docs/xlsx-architecture.md`'s ADR) for a
  future crate that depends on elixcee's value types without pulling in
  the full VBA parser/VM. Every existing `crate::vm::X` path still resolves
  via re-export; no public module path changed. Root `Cargo.toml` becomes a
  non-virtual workspace root (`fuzz/` explicitly excluded — different
  edition, its own `cargo fuzz` toolchain)
- **Mechanical `clippy` fixes** across the crate (let-chain collapses,
  `.get(0)` → `.first()`, `% 2 == 0` → `.is_multiple_of()`, `Option::
  map_or(true, _)` → `is_none_or()`, redundant closures/casts, 5
  `needless_range_loop` rewrites, 2 duplicated `#[test]` attributes that
  were silently double-registering the same test) — logic-preserving only;
  `docs`/`tasks/todo.md`'s remaining clippy backlog (8 `approx_constant`
  false positives on literal test values, 1 `too_many_arguments`) is
  unchanged, tracked separately, not a release blocker

### Tests

- **Real-producer E2E fixtures** (`tests/fixtures/e2e/`): `.xlsx`/`.ods`
  generated by real LibreOffice (not hand-crafted), read via both
  `calamine` (independent oracle) and elixcee's own reader and asserted
  equal — closes the "zero binary fixtures from a real office suite" gap
  from the `@elixcee/xlsx` Phase 0 investigation. Verified to actually
  catch the two `Fixed` bugs above (both new tests fail against the
  pre-fix reader on this real-producer input, pass after)
- **15 new unit tests** directly in `elixcee-types`, covering the extracted
  surface at its new crate boundary (previously only indirect coverage via
  the much larger `vm`/`formula` test suites)

## [0.1.1]

### Added

- **CLI binary** (`src/main.rs`): standalone `elixcee` executable — no Python required
  - Usage: `elixcee <vba_file> <MacroName> [--file xlsx] [--sheet name] [--output xlsx]`
  - `MsgBox` output printed to stdout; result cells printed as `A1\t<value>` per line
  - Pre-built binaries for Windows x64, Linux x64, macOS Apple Silicon on GitHub Releases
- **GitHub Actions release workflow** (`.github/workflows/release.yml`): builds CLI binaries on `bin-v*` tag push; attaches them to a GitHub Release via `softprops/action-gh-release`
- **`pub fn save_workbook`**: public Rust API for writing `.xlsx` / `.ods` from non-Python callers
- **`Vm::print_msgbox`** field: when `true`, `MsgBox` writes to stdout instead of being silently dropped
- **pyo3 optional feature**: `pyo3` is now an optional dependency behind the `python` feature; `cargo build --bin elixcee` compiles a Python-free binary; `maturin build` continues to use `features = ["python"]`
- **Math & Combinatorics**: `FACT`, `PERMUT`, `GCD`, `LCM`, `QUOTIENT`, `SIGN`
- **Statistical**: `CORREL`, `COVARIANCE.S`, `COVARIANCE.P`, `NORM.DIST`, `NORM.INV`, `T.DIST` — uses Stirling lgamma + Lentz incomplete-beta CF
- **Financial functions**: `FV`, `PV`, `NPER`, `RATE` (Newton-Raphson), `IPMT`, `PPMT`, `NPV`, `IRR`, `MIRR`, `XNPV`, `XIRR` — all share the `annuity_fv` / `compute_pmt` helpers
- **Database functions**: `DSUM`, `DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DMAX`, `DMIN` — all take `(database, field, criteria)` and reuse the existing `db_row_matches_criteria` / `resolve_db_field` infrastructure from `DGET`
- **GitHub Actions CI/CD**: `.github/workflows/publish.yml` — builds wheels for Linux x86_64/aarch64, Windows x86_64, macOS universal2, and an sdist; publishes to PyPI via OIDC Trusted Publisher on `v*` tag push
- **README_zh.md**: Simplified Chinese translation of README

### Added — JSON Agent Contract & Static Analysis (Milestones A, A.1, A.5, B1, B1.1, B2, B3, B4)

- **`--json` output** (`src/diagnostics.rs`): single machine-readable JSON object (result or error) instead of plain text — error classification (`ElixceeError`), a hand-rolled JSON writer/escaper (no serde in the release binary), and `Vm::msgbox_log` (`MsgBox` calls recorded into `messages` instead of printed directly, drained via `take_messages()` so a reused `Vm` never leaks a prior run's messages)
- **Source location tracking** (`SourceSpan`/`SpannedStmt`, char-offset based): parse and runtime errors report `{file, line, column}` in `--json` mode; non-JSON output is unchanged
- **`check` subcommand** (`src/check.rs`): static analysis without executing the macro — parse diagnostics, entrypoint existence, undefined Sub/Function call detection anywhere in the body (probes the real builtin-function dispatch table directly, so there's no allowlist to drift), and unsupported-construct/no-op detection (`I1002`), all with source locations
- **Multi-module projects**: pass more than one `.bas`/`.vbs` file to run a project spanning several modules; `Module.Sub`-qualified entrypoints (module name from `Attribute VB_Name`, else the filename); cross-module Sub/Function name collisions are rejected at load time
- **Deterministic black-box tests** (`tests/blackbox.rs`): declarative `.toml` fixtures (VBA source + CLI args + expected JSON) diffed byte-for-byte against the real binary's `--json` output; adding a new regression case needs no Rust
- **`snapshot` subcommand** (`src/snapshot.rs`): reads a `.xlsx`/`.xlsm`/`.ods` file directly (no VBA execution) and prints every sheet's non-empty cells as Markdown or JSON, with a `sheet_id`/`stable_id` pair for cross-sheet identity (not to be confused with VBA's real `CodeName`)

### Added — Property-Based Testing & Excel Operation Diagnostics (Milestones B5a, B6a, B6b, B6c)

- **`test-workbook` subcommand** (`src/testworkbook.rs`): reruns a macro against a starting workbook many times with generated boundary-value inputs (`boundary_numeric`/`boundary_string`), checking each independent case for panics, runtime errors, timeouts, and Excel error values; failures report `seed`/`case_index` for exact replay via `--seed`/`--case`
- **`diagnose` subcommand** (`src/diagnose.rs`): runs a macro once and classifies *why* Excel would reject an operation, with evidence, instead of a bare error string —
  - `WORKSHEET_NOT_FOUND` / `WORKBOOK_NOT_FOUND` / `ARRAY_INDEX_OUT_OF_BOUNDS`, with a hand-rolled Levenshtein "did you mean" suggestion (opt-in `Vm::strict_resolution` turns off the usual auto-vivify-on-write/silent-`Empty`-on-read behavior only for this command)
  - `Sheets(name).Range(addr)`, `Worksheets(idx)` numeric index, and a minimal `Workbooks(name).Worksheets(...)` all newly parseable, needed to even express the sheet-resolution scenarios this command diagnoses
  - `PASTE_SHAPE_MISMATCH` / `PASTE_WITHOUT_COPY`: a VM clipboard (`Vm.clipboard`) populated by `.Copy`/`.Copy Destination:=` and consumed by `.Paste`/`.PasteSpecial [Transpose:=]`/`Worksheets(sheet).Paste`, with both the Copy and Paste statement locations and a mechanically-derived resize suggestion
  - `SHEET_PROTECTED`: `Sheets(name).Protect`/`.Unprotect` (including `UserInterfaceOnly:=True`, which blocks manual edits but not macro writes, matching real Excel) blocks any cell-content mutation on that sheet — writes, clears, inserts, sorts, paste, delete — unconditionally in every mode, while reads are never blocked
  - Shape mismatches, empty-clipboard pastes, and writes to a protected sheet are unconditional hard errors in every mode that executes the macro (`run`/`diagnose`/`test-workbook`), matching real Excel's Error 1004/protection behavior regardless of `On Error` state

### Changed

- `pyproject.toml`: `features = ["pyo3/extension-module"]` → `features = ["python"]` to align with the new optional-feature approach
- **`diagnose`'s entrypoint is now a positional argument** (`elixcee diagnose <vba_file>... <MacroName> --file <path> [--json]`) instead of `--entrypoint <MacroName>` — matches `run` mode's own convention (entrypoint is always mandatory for both, unlike `check`, where it's optional and therefore needs an explicit flag to stay unambiguous). Breaking change; `--entrypoint` is removed, not kept as an alias.
- **PyPI package metadata**: `pyproject.toml` now declares a description, `readme`, `license`, `keywords`, `classifiers`, and `[project.urls]` (Homepage/Documentation/Repository/Issues/Changelog) — the published package previously had none of these

### Removed

- `FUNCTIONS_ja.md`: duplicate of `FUNCTIONS.md`; `README_ja.md` now links to the English reference

### Performance (Round 4)

- **`SUM` fast path**: single-range `SUM` iterates cell refs directly — no `Vec<Variant>` allocation
- **`range_nums_fast!` macro**: `AVERAGE`, `MIN`, `MAX` on a single range skip `Vec<Variant>` and collect `f64` directly
- **`RangeWrite` / `RangeClear` dirty-flag batching**: writes go directly to the sheet map; `cell_index_dirty` set once after the loop instead of once per cell

### Tests

503 unit tests (↑ from 329) + `tests/cli_json.rs` (14) + `tests/cli_check.rs` (15) + `tests/blackbox.rs` (1 test scanning 12 `.toml` fixtures) + `tests/cli_snapshot.rs` (7) + `tests/cli_test_workbook.rs` (7) + `tests/cli_diagnose.rs` (12) + `tests/prop_tests.rs` (17)

---

## [0.1.0] — Initial Release

### Added — VBA Parser & Interpreter

- **Sub / End Sub** with parameter passing
- **Variable assignment** and arithmetic expressions
- **Cell read/write** via `Cells(row, col).Value` and `Range("A1").Value`
- **For / Next** loops with optional `Step`
- **For Each** iteration over cell ranges
- **If / ElseIf / Else / End If** conditional branches
- **Do While / Loop** and **While / Wend** loops
- **Select Case** with value, range (`To`), and comparison (`Is`) patterns
- **Exit For**, **Exit Do**, **Exit Sub**, **Exit Function**
- **Function / End Function** with return values; **Call** statement
- **On Error Resume Next**, **On Error GoTo label**, **Resume**, **GoTo**
- **With / End With** blocks (plain and `With Sheets("name")`)
- **Const** declarations; `Option Explicit` / `Option Base` ignored
- **Dim** variable declarations; `Dim arr(n)` and `ReDim [Preserve]` arrays
- **Type ... End Type** user-defined types with typed field initialization
- **Public / Private / Friend / Static** modifiers on Sub/Function (modifier ignored)
- **Debug.Print** / **Debug.Assert** as no-ops
- **MsgBox** — configurable skip or RuntimeError
- Range operations: `ClearContents`, `Clear`, `Copy`, `Delete`, `Insert`, `Sort`, `Offset.Value`
- Sheet operations: `Sheets.Add`, `Sheets.Delete`, `Sheets("name").Cells`
- `Application.Calculation` (Manual / Automatic); `ScreenUpdating`, `EnableEvents`, `DisplayAlerts`, `StatusBar`, `Cursor`, `CutCopyMode` as no-ops
- `WorksheetFunction.*` prefix forwarding to formula engine
- `Cells(Rows.Count, col).End(xlUp).Row` and related `.End(dir).Row/Column` — indexed with `BTreeSet` for O(log n) performance

### Added — Named Ranges

- `Range("A1:B5").Name = "MyName"` registers a workbook-level named range
- All Range operations (Read/Write/Clear/Delete/Insert/Sort/Copy/ForEach) transparently resolve named range strings

### Added — Formula Engine (200+ functions)

#### Arithmetic & Statistical
`SUM`, `AVERAGE`, `AVERAGEIF`, `AVERAGEIFS`, `MIN`, `MAX`, `MINIFS`, `MAXIFS`,
`COUNT`, `COUNTA`, `COUNTIF`, `COUNTIFS`, `COUNTBLANK`,
`SUMIF`, `SUMIFS`, `SUMPRODUCT`, `PRODUCT`, `MEDIAN`, `MODE.MULT`,
`LARGE`, `SMALL`, `RANK`, `PERCENTILE` / `PERCENTILE.INC`, `PERCENTRANK` / `PERCENTRANK.INC`,
`ROUND`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `TRUNC`, `MOD`,
`RAND`, `RANDBETWEEN`, `SUBTOTAL`, `AGGREGATE`

#### Statistical
`STDEV` / `STDEV.S`, `STDEVP` / `STDEV.P`, `VAR` / `VAR.S`, `VARP` / `VAR.P`

#### Math & Trigonometry
`ABS`, `SQRT`, `POWER`, `EXP`, `LN`, `LOG`, `LOG10`, `PI`,
`SIN`, `COS`, `TAN`, `ASIN`, `ACOS`, `ATAN`, `ATAN2`, `DEGREES`, `RADIANS`,
`FLOOR` / `FLOOR.MATH`, `CEILING` / `CEILING.MATH`, `MROUND`

#### Logical
`IF`, `IFS`, `SWITCH`, `AND`, `OR`, `NOT`, `XOR`, `IFERROR`

#### Text
`LEFT`, `RIGHT`, `MID`, `LEFTB`, `RIGHTB`, `MIDB`,
`LEN`, `LENB`, `UPPER`, `LOWER`, `PROPER`, `TRIM`,
`FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`,
`CONCATENATE`, `CONCAT`, `TEXTJOIN`, `TEXT`, `VALUE`, `EXACT`,
`CHAR`, `UNICHAR`, `CODE`, `UNICODE`, `ASC`, `JIS`

#### Date & Time
`DATE`, `TODAY`, `NOW`, `YEAR`, `MONTH`, `DAY`, `WEEKDAY`, `DAYS`,
`EDATE`, `EOMONTH`, `DATEDIF`, `DATEVALUE`,
`TIME`, `TIMEVALUE`, `HOUR`, `MINUTE`, `SECOND`,
`NETWORKDAYS`, `NETWORKDAYS.INTL`, `WORKDAY.INTL`

#### Lookup & Reference
`VLOOKUP`, `HLOOKUP`, `XLOOKUP`, `LOOKUP`,
`INDEX`, `MATCH`, `XMATCH`, `CHOOSE`,
`ROW`, `COLUMN`, `INDIRECT`, `OFFSET`, `ADDRESS`

#### Information
`ISBLANK`, `ISERROR`, `ISERR`, `ISNA`, `ISNUMBER`, `ISTEXT`, `ISLOGICAL`, `ISNONTEXT`

#### Array / Spill
`FILTER`, `UNIQUE`, `SORT`, `SORTBY`, `SEQUENCE`, `TRANSPOSE`,
`TOCOL`, `TOROW`, `WRAPCOLS`, `WRAPROWS`, `RANDARRAY`

#### Lambda & Higher-Order
`LET`, `LAMBDA`, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`

### Added — Formula Engine Features

- **Topological sort** for formula recalculation (`topo_sort_formulas`): formulas are evaluated in dependency order; circular references fall back to best-effort ordering
- **Application.Calculation** mode — Manual suppresses recalc; switching to Automatic triggers full recalc
- A1-notation and R1C1-notation cell references; range references (`A1:B10`)
- DBCS byte semantics (`LENB`, `LEFTB`, `RIGHTB`, `MIDB`) matching Excel's 2-byte-per-CJK rule
- Excel 1900 leap-year bug compatibility in date serial arithmetic

### Added — Python API

| Method | Description |
|---|---|
| `Vm(on_msgbox=)` | Create a VM; `"skip"` or `"error"` on MsgBox |
| `vm.run(vba, name)` | Execute a Sub |
| `vm.set_cell(r, c, v)` / `get_cell(r, c)` | 1-based cell read/write |
| `vm.cells()` | All non-empty cells as `{(r, c): value}` |
| `vm.cells_df()` | Active sheet as pandas DataFrame (requires pandas) |
| `vm.variables()` | VBA variables as `{name: value}` |
| `vm.set_cell_formula(r, c, f)` | Set and evaluate a formula string |
| `vm.set_cell_formula_batch(d)` | Batch formula set: `{(r,c): formula}` |
| `vm.recalculate()` | Re-evaluate all formula cells |
| `vm.set_sheet(name)` / `active_sheet()` / `sheet_names()` | Sheet management |
| `vm.get_sheet(name)` | Cells of a named sheet |
| `vm.save_workbook(path)` | Save to `.xlsx` or `.ods` |
| `vm.named_ranges` | Dict of registered named ranges |
| `elixcee.run_macro(vba, name)` | One-shot macro runner |
| `elixcee.load_workbook(path)` | Load `.xlsx` / `.ods` into a `Vm` |

- `Variant::Date` → Python `datetime.date` conversion
- `Variant::Error` → Python `ExcelError` class with `.code` attribute (bidirectional)
- Type stubs `elixcee.pyi` for IDE completion

### Added — File I/O

- **Read**: `.xlsx`, `.xlsm`, `.ods` — hand-written XML parser (no calamine at runtime)
- **Write**: `.xlsx` — hand-written XML + zip; `.ods` — hand-written XML + zip
- Multi-sheet support: all sheets loaded on `load_workbook`; saved on `save_workbook`

### Performance

- `Cells.End` searches (`xlUp`, `xlDown`, `xlToLeft`, `xlToRight`) use a lazy `BTreeSet` index — O(log n) per query after O(n) rebuild on cell mutation
- Zero-copy formula parse caching via `recalculate_all` with topological ordering

### Dependencies (runtime)

| Library | Purpose |
|---|---|
| `pyo3` | Python bindings |
| `zip` | XLSX / ODS archive read-write |

`calamine` is kept as a `[dev-dependencies]` oracle for diff-testing the hand-written reader.

### Tests

299 unit tests covering parser, formula engine, VM interpreter, file round-trips, and diff tests against calamine.
