# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased]

Root `elixcee` and `elixcee-types` only — proposed `elixcee` 0.7.0 / `elixcee-types` 0.3.0,
neither yet applied to any file (`crates/elixcee-types/Cargo.toml`'s own `version` and the
root `Cargo.toml`'s `elixcee-types = { ..., version = "0.2.0" }` dependency pin both still say
0.2.0 — the latter isn't checked by `scripts/check-versions.sh`, which only compares
`Cargo.toml` against `pyproject.toml`, so it needs a manual edit at bump time). `elixcee-wasm`'s
packaged output (vendored into `@elixcee/xlsx`) was refreshed to reflect these changes even
though its own source is untouched — see the "refresh packaged artifacts" entry below. Three
independent Rust-side items: real multi-dimensional VBA arrays; call-frame-scoped `On Error`
with `Err.Source`/`Err.HelpFile`/`Err.HelpContext`/full 5-argument `Err.Raise`; and moving
undefined-procedure calls, argument-count mismatches, and undefined `GoTo` labels to a
compile-time, `On Error`-uncatchable check — plus, found as necessary groundwork for the third
item, a pre-existing parameter-parsing bug affecting any macro using
`ByVal`/`ByRef`/`Optional`/`ParamArray`. `cargo test --workspace` 952/952 (up from 872 at `HEAD`
before this round — 80 new tests: 826 lib + 82 integration + 25 `elixcee-types` + 19
`elixcee-wasm`), `cargo build --release --workspace` clean, `compat/vba-semantics` 386 cases (0
`BUG`/0 `UNCLASSIFIED`, `KNOWN_LIMITATION` 14 — down from 16), `compat/corpus` 581 scenarios (0
`UNEXPLAINED`/0 `MISMATCH`). `RuntimeErrorKind`/a typed `RuntimeError` struct (this round's own
lower-priority item, referenced by the 0.6.0 entry below as "not fixed here") is still not
started. `@elixcee/xlsx` (a separate, independently-versioned package — see its own entries
below) got `write()`/`writeFile()`/`writeFileSync()` in this round, prepared as
`0.1.0-alpha.1` but not published — `packages/xlsx/package.json`'s `version`/`private`/
`publishConfig` fields are unchanged from `0.0.0-development`/`private:true`/no
`publishConfig` in this round; only its `description` was updated to match the new scope.

### Real multi-dimensional VBA arrays

Previously, `Dim arr(3, 2)` allocated storage as if it were 1-D (dimension 1's element count
only), so `arr(1, 1)` and `arr(1, 2)` silently aliased the same storage slot and
`UBound(arr, 2)` returned dimension 1's bound regardless of which dimension was actually asked
for — disclosed as `KNOWN_LIMITATION` in `compat/vba-semantics` (`two_dimensional_array_second_index_is_silently_dropped`,
`ubound_second_dimension_argument_ignored`), now fixed and reclassified.

- **New `elixcee_types::VbaArray`/`ArrayBound`** (`crates/elixcee-types/src/lib.rs`) — flat,
  row-major storage (`idx = idx * bound.len() + (sub - bound.lower)`, first dimension varies
  slowest) with real per-dimension bounds, replacing 1-D-only storage for every `Dim`-declared
  array. New `Variant::VbaArray(VbaArray)` enum variant — additive, `Variant::Array(Vec<Variant>)`
  itself unchanged (still used for Range-value multi-cell reads, `formula::eval`'s array-formula
  results, and `DimArrayRecord`/`ArrayRecordSet` storage, none of which have per-dimension bounds
  to track). Element count is overflow-checked (`checked_mul`) and capped at 10,000,000 elements,
  surfacing real VBA's own "Out of memory" wording rather than a Rust-side allocation panic.
- **`LBound`/`UBound`** now honor the dimension argument per-dimension; `Option Base` applies
  independently to every dimension; `Erase` resets elements while preserving all dimensions'
  bounds; **`ReDim Preserve`** is correctly restricted to real VBA's own rule — only the *last*
  dimension's *upper* bound may change, every other dimension (and the last one's own lower
  bound) must stay identical or it's Error 9 (`redim_preserve` in `src/vm/mod.rs`, found and
  fixed its own bug during this round: an earlier version of the check missed that the last
  dimension's *lower* bound is equally protected, not just the non-last dimensions).
- Shape preserved through variable assignment, function-argument passing, and function-return
  values; `Array()`/`Split()` migrated to `VbaArray` while keeping their externally-observable
  0-based rank-1 shape unchanged.
- **PyO3 bindings** (`src/lib.rs`): new `vba_array_to_py` recursively reshapes flat `VbaArray`
  storage into nested Python lists matching the array's real dimensional shape — verified against
  a real `maturin`-built wheel, not just `cargo test`.
- `crates/elixcee-wasm` needed no changes — grep-confirmed it references none of `vm::`/
  `Variant::`/`VbaArray` directly, and it already compiled clean.

### Call-frame-scoped `On Error`

Previously, `On Error Resume Next`/`On Error GoTo <label>` state was a single `Vm`-wide flag —
a callee's own body could see and mistakenly try to resolve a caller's still-active `GoTo`
label, and (found and fixed as part of the same rework) a callee's remaining statements kept
running under a caller's `On Error Resume Next` even after the callee itself failed, since the
catch fired inside the callee's own `exec_stmt`, not at the call site.

- **New `Vm::call_stack: Vec<CallFrame>`**, each frame holding its own `ErrorMode` (`Disabled`/
  `ResumeNext`/`GoTo(String)`), replacing the old `on_error_resume_next: bool`/
  `on_error_goto_label: Option<String>` fields. Pushed/popped around every `call_sub_def`/
  `call_func_def` invocation, so a callee always starts with `Disabled` regardless of the
  caller's own mode — matching real VBA (error handling doesn't inherit into a callee). A
  `GoTo` handler is consumed (reset to `Disabled`) the moment it fires, matching real VBA: a
  second failure while already inside a handler propagates to the caller rather than
  re-entering the same handler.
- **Deliberate behavior change**: under the old flag, a caller's `On Error Resume Next` catching
  an error from inside a called Sub let that Sub's *remaining* statements keep running (the
  catch happened inside the callee's own body). Now the error propagates out of the callee
  entirely and is caught at the `Call` statement in the caller's own frame — the callee's
  remaining statements do not run. This is the correct real-VBA behavior, but a macro that
  depended on the old leniency will observe the difference.
- Incidental fix: `run_sub`/`run_sub_multi` never reset the old `on_error_resume_next`/
  `on_error_goto_label` fields between runs on a reused `Vm` (the Python bindings' own usage
  pattern) — `call_stack.clear()` at the start of each run closes that.

### `Err.Source`/`Err.HelpFile`/`Err.HelpContext`, full `Err.Raise`

- **`Err.Source`/`Err.HelpFile`/`Err.HelpContext`** added as readable properties (`Expr::ErrSource`/
  `ErrHelpFile`/`ErrHelpContext`), joining the existing `Err.Number`/`Err.Description`.
  **`Err.Raise`** now accepts and threads through all five real positional arguments (`Number,
  Source, Description, HelpFile, HelpContext`), correctly handling a bare comma skipping any of
  the last four at any position (`Err.Raise 513, , "text"` still means Number=513,
  Description="text", not Source="text"). **`Err.Clear`** now resets all five properties, not
  just Number/Description.
- Not done: the richer `RuntimeError`/`RuntimeErrorKind` struct this phase's own spec also
  asked for (a `span`/`kind`-classified error type, replacing string-matching-based
  `classify_vba_error_number`) — the five `Err.*` properties above are still backed by flat
  `Vm` fields (`err_number`/`err_description`/`err_source`/`err_help_file`/`err_help_context`),
  not a unified struct. Deliberately deferred, not an oversight.

### Undefined-procedure calls, argument-count mismatches, and undefined `GoTo` labels are now compile errors

Real VBA fails a macro that calls an undefined Sub/Function, passes the wrong number of
arguments to one it can see, or `GoTo`s a label that doesn't exist — *before* running a single
statement, and never lets `On Error` trap it (a whole-project compile phase, not a runtime
check). Previously all three were ordinary runtime errors here, reachable partway through
execution and (incorrectly, relative to real VBA) catchable by an active `On Error Resume
Next`/`GoTo`.

- **New `check::compile_check_errors`** (`src/check.rs`) walks the whole program (every Sub and
  Function, not just the ones the entrypoint's call chain actually reaches — matching real
  VBA's whole-project compile-then-run semantics) for exactly these three conditions, reusing
  the same `is_resolvable` logic `check`'s existing undefined-name diagnostic (E1002) already
  used. **`Vm::run_sub`/`run_sub_multi`** now run this once, before `call_sub_def`, and return
  its finding as an ordinary `Err(String)` — the "uncatchable by `On Error`" property comes for
  free from running it before any statement (including any `On Error`) has executed.
  Multi-module runs build each module's own cross-module name set the same way `elixcee check`'s
  own multi-module path already did, so a legitimate unqualified cross-module call isn't
  misreported as undefined.
- A deliberately-unimplemented `WorksheetFunction.*` call (e.g. `.TextJoin`) is still reported,
  but with the exact message its real dispatch path (`eval_wsf`) would give at runtime — a new
  `vm::builtin_call_error` helper asks the VM itself instead of guessing generic wording, so
  `check`'s pre-flight rejection always matches word-for-word what actually running it would
  have said.
- **`elixcee check` learned the same two checks** (new diagnostic codes `E1008`/
  `argument_count_mismatch`, `E1009`/`undefined_label` — see `docs/agent-contract.md`), closing
  a gap this round's own testing surfaced: without this, `elixcee check` could report a program
  clean that `run_sub`'s new pre-flight pass would then refuse to run at all. Every violation is
  reported (not just the first, unlike `run_sub`'s own short-circuit-on-first-violation pass).
- **Measured, not assumed, performance impact**: `is_known_builtin_function`/
  `builtin_call_error` each construct a throwaway `Vm` and run a real dispatch probe per
  unresolved name, and the whole pre-flight check now re-runs on every `run_sub` call —
  relevant since `test-workbook` reruns the same macro across many generated cases. Measured
  with a `test-workbook` fixture calling 10 distinct built-in/`WorksheetFunction` names
  (deliberately adversarial — a typical macro has far fewer), 3000 cases: roughly 5% slower
  than the pre-this-round build (~0.68s → ~0.72s wall-clock for the whole run, ~13µs/case). Not
  optimized (no memoization of the builtin probe) — the absolute cost is small and the fixture
  used to measure it overstates a typical macro's actual builtin-call density.
- **Deliberately not checked**: "invalid assignment target" (e.g. calling a Function's result on
  the left of `=` as if it were an array element) — `name(args) = value` parses unconditionally
  as `Stmt::ArrayWrite` regardless of whether `name` is a real array or (invalidly) a Function
  name, and telling those apart isn't syntactically decidable without type inference this
  project stays out of by design. Also not checked: a cross-module call's argument count (this
  pass only ever sees one module's own declared Sub/Function arities), and a recursive call's
  own argument count (a procedure's own name is already in its local scope, so it's never
  treated as a checkable external call).
- **Found while building this**: `Sub Foo(ByVal x As Integer)` used to silently misparse —
  `parse_params` had no special handling for the `ByVal`/`ByRef`/`Optional`/`ParamArray`
  keywords, so `consume_ident()` swallowed `byval` itself as a bogus extra parameter name,
  making `Foo` a real *2*-parameter Sub (`["byval", "x"]`). Calling `Foo(5)` bound `5` to the
  phantom `byval` parameter and left `x` unbound — `x` inside `Foo`'s body raised "Undefined
  variable: 'x'", not a type/argument error, so this was easy to miss. Pre-existing, unrelated
  to any array/call-frame work above; found because it would have made the new argument-count
  check wrong for any macro using these (very common) parameter modifiers. Fixed:
  `ByVal`/`ByRef` are now recognized and discarded (this VM already treats every parameter as
  effectively by-value, with no `ByRef` write-back modeled for *any* parameter, so discarding
  the keyword is correct, not a simplification); `Optional`/`ParamArray` are not implemented and
  now fail with a clear parse error instead of the same silent misparse — a deliberate behavior
  change, not a regression, for any macro that happened to declare one of these keywords before.

### `@elixcee/xlsx` — `write()`/`writeFile()`/`writeFileSync()`

Independent of the root `elixcee` crate: no Rust changes, no new npm dependency.
`bookType: "xlsx"` only, output `type: "buffer" | "array" | "base64"`, producing a real
OOXML ZIP via a hand-rolled ZIP/XML writer (no zip/xml-builder dependency added) —
strings/numbers/booleans/dates/formulas, multiple worksheets, merges, sheet visibility,
hidden rows/columns, basic number formats, safe XML escaping. Unsupported input (a
non-`"xlsx"` `bookType`, an unrecognized `type`, an unsupported cell shape/type, a
non-finite numeric/formula-cached value, an oversized declared `!ref`) throws an explicit
`ELIXCEE_*` error, never silently ignored or truncated.

- **`packages/xlsx/src/internal/xlsx-writer.cjs`** (new) — the OOXML XML generator:
  `[Content_Types].xml`, both `.rels` parts, `docProps/{core,app}.xml`, `xl/workbook.xml`,
  `xl/worksheets/sheetN.xml`, `xl/styles.xml`. Output is deliberately constrained to
  shapes `src/reader.rs` (elixcee's own reader) already parses, verified by reading
  `reader.rs` directly — inline strings (not shared strings), a small built-in
  numFmtId table plus custom `<numFmts>` entries (164+) for anything else.
- **`packages/xlsx/src/internal/zip-writer.cjs`** (new) — a hand-rolled ZIP archive
  writer (local file headers, central directory, end-of-central-directory record,
  table-based CRC-32) with a deterministic fixed epoch, so two `write()` calls on the
  same `WorkBook` produce byte-identical output. Platform-agnostic by design: no
  `Buffer`, every byte buffer is a plain `Uint8Array` built with `DataView`/
  `TextEncoder` — real browsers never had `Buffer` regardless of bundler, so the shared
  writer core is built to work on both platforms from the start. DEFLATE compression is
  supplied by the caller as an optional callback rather than required internally (falls
  back to STORED, a legal ZIP/OOXML method, when omitted — this is what lets the browser
  entry reuse the same writer with no `zlib` access at all).
- **`compat/differential/xlsx-write.test.mjs`** (new) — 36 MATCH + 1 disclosed
  UNSUPPORTED case (`bookType: "ods"`, registered in `classify.mjs`'s
  `UNSUPPORTED_ALLOWLIST`), covering all three round-trip directions (own write -> own
  read, own write -> oracle read, oracle write -> own read) against a fourth,
  independently-computed baseline (oracle write -> oracle read); plus standalone checks
  for OOXML ZIP/XML structural validation (CRC-32, balanced XML, `[Content_Types].xml`/
  `.rels` cross-references), 12 malformed-workbook rejection cases, output-type
  agreement (buffer/array/base64 carry identical bytes), write-determinism, a real
  filesystem round trip for `writeFile`/`writeFileSync`, and the browser entry's
  behavior (both throwing `ELIXCEE_UNSUPPORTED_IN_BROWSER` for `writeFile`/
  `writeFileSync`, and `write()` itself working with no filesystem).
- `compat/differential/metadata.test.mjs` extended: `write`/`writeFile`/`writeFileSync`
  now among the 39/39 exports checked (name/length/property-descriptor/CJS-ESM-identity
  against the oracle), plus a `writeFile === writeFileSync` aliasing check.

### `@elixcee/xlsx` — make writer bundles work in Node ESM and browsers

**Two real bundler bugs found by actually bundling and running the code, not assumed, and
both fixed at the source**:

1. An esbuild `--format=esm --platform=node` bundle can never synchronously `require()`
   anything reached through CJS-origin code — confirmed neither a lazy require,
   `require('node:zlib')`, nor `--external:zlib` changes this; the documented, correct
   pattern is marking the whole package `external` (`--packages=external`), verified
   end-to-end and pinned as a permanent regression check in `scripts/wasm-smoke.mjs`
   (step 6).
2. An esbuild `--platform=browser` bundle refused to even build at all with a
   `require('zlib')` reachable anywhere in its module graph (dead code included, since
   esbuild can't tree-shake CommonJS `module.exports` properties). Fixed by isolating the
   Node-only `zlib.deflateRawSync` wrapper into its own new file,
   **`packages/xlsx/src/internal/deflate-node.cjs`**, and stubbing that exact path (plus
   bare `zlib`) to `false` in `package.json`'s `browser` field — the same mechanism
   already used for `elixcee_wasm.node.cjs`. This works because `browser`-field
   path-remapping happens at module-resolution time, before the stubbed file's contents
   are ever parsed; moving the `require('zlib')` around *within* `index.cjs` (tried
   first) did not work, since `index.cjs` itself is wholesale-included in the browser
   bundle graph via `index.browser.mjs`'s re-export of its other, browser-safe exports.

- `scripts/wasm-smoke.mjs` extended (step 6): `bundleAndRunWrite`/`runWriteBundle` verify
  all four combinations — inlined-ESM-must-throw, inlined-CJS-must-run,
  externalized-ESM-must-run, externalized-CJS-must-run — pinned as a permanent regression
  check for bug 1 above.
- `scripts/browser-smoke.mjs` extended: the bundled entry now calls `write()` then
  `read()` and asserts the round trip, plus a build-time assertion that the bundle
  contains zero `zlib` references at all — verified against a real headless Chrome
  process, not just a passing build.
- `scripts/pack-consumer-smoke.mjs` extended: a shared `WRITE_ROUNDTRIP` snippet exercises
  `write()`/`writeFile()`/`writeFileSync()` from inside a real `npm pack` + `npm install`,
  both from CJS and ESM consumers, plus a new step for `writeFile()`/`writeFileSync()`
  against a real filesystem.
- `docs/xlsx-architecture.md` — new "Phase D: `write()`'s Node-builtin bundling posture"
  section documents both bugs, why each fix works, and why bug 2's fix (isolating the
  Node-only `zlib` access) is a different problem from bug 1's (ESM+Node package
  externalization) and needs a different solution.

## [0.6.0]

Root `elixcee` (Rust crate + Python package) only — `elixcee-types` stays `0.2.0` (no
public-surface change this round: the new array lower-bound tracking lives in a `Vm`-side
`HashMap`, not on `Variant::Array` itself, so nothing semver-relevant moved), `elixcee-wasm`
stays `0.1.0` (no source changes to `crates/elixcee-wasm/src`; its vendored build output was
regenerated to pick up the `src/reader.rs` fix below), and `@elixcee/xlsx` stays
`0.0.0-development`/unpublished/`"private": true` (none touched, though its vendored WASM
artifact was refreshed too). Four independent items, in the order they were authorized: real
VBA `Err` object semantics; `compat/vba-semantics`'s 386-case suite and `compat/corpus`'s
581-scenario suite wired into a new `compat-vba` CI job (previously local-only); four of five
disclosed array-declaration/bounds gaps fixed, the fifth (`UBound`'s dimension argument,
needing real multi-dimensional array storage) deliberately deferred with its own disclosure
corrected after its registry entry turned out to make a false claim; and `src/reader.rs`'s
`xml:space="preserve"` whitespace-trimming defect on `t="str"` cells, fixed
(`compat/differential:read` 30 MATCH + 3 disclosed → 33/33 MATCH). Full detail in the four
sections below. Plus one bug found during this round's own pre-release verification, not part
of the four authorized items: the Python `elixcee.load_workbook()` binding panicked with
`"active sheet must exist"` on any sheet named the way Excel itself defaults to naming one
(`"Sheet1"`, capital S) — a hand-rolled duplicate of the CLI's sheet-population loop had never
picked up a mixed-case-sheet-name fix the CLI path got back in July. Fixed by routing through
the same already-tested helper the CLI uses instead of maintaining a second copy. Pre-existing
since that July commit, unrelated to any 0.6.0-phase work. `cargo test --workspace` 872/872,
`compat/vba-semantics` 386 cases (0 `BUG`/0 `UNCLASSIFIED`, 16 `KNOWN_LIMITATION` — down from
19), `compat/corpus` 581 scenarios (0 `UNEXPLAINED`/0 `MISMATCH`) all green as of this bump;
every real GitHub Actions job (including the new `compat-vba` job's first real run) green on
`master` before this bump.

### `Err` object: `Err.Number` / `Err.Description` / `Err.Clear` / `Err.Raise`

First item of the 0.6.0 phase. `On Error Resume Next`/`On Error GoTo <label>` already
existed but had no way for the running macro to inspect *what* error was caught — the
single most common real-world idiom this blocked was
`On Error Resume Next : <risky op> : On Error GoTo 0 : If Err.Number <> 0 Then ...`.

- **New `Err.Number`/`Err.Description` expressions**, `Err.Clear`/`Err.Raise` statements
  (`src/parser/ast.rs`, `src/parser/mod.rs`, `src/vm/mod.rs`). Parser recognition is
  guarded on the exact member name (`Err.Number`/`Err.Description`/`Err.Clear`/`Err.Raise`
  specifically), matching the existing `ThisWorkbook`/`ActiveWorkbook` precedent — a
  genuine user variable named `err` with an unrelated field (`err.code = 1`) still parses
  as ordinary assignment/field access, unaffected (test:
  `a_bare_err_variable_is_unaffected_by_err_object_parsing`).
- **`Vm::err_number`/`err_description`** are set at both existing error-catch sites
  (`On Error Resume Next`'s per-statement catch, `On Error GoTo <label>`'s jump) via a new
  `classify_vba_error_number(msg: &str) -> (i64, String)`. Only maps a handful of
  elixcee-internal message strings that are **confirmed exact matches against Microsoft's
  own long-stable, publicly documented VBA runtime error constants** (unchanged since
  VB6 — a fact independent of this project's lack of a live Excel/VBA oracle, see
  `ROADMAP.md`'s Known gap #1): Division by zero → 11, Subscript out of range → 9, Type
  mismatch → 13, Invalid procedure call or argument → 5, Invalid use of Null → 94, Object
  variable or With block variable not set → 91. Everything else elixcee itself raises
  (undefined variable, sheet/sub/workbook not found, etc.) defaults to 1004
  ("Application-defined or object-defined error", real VBA's own generic catch-all for
  Excel-object-related failures) — a disclosed default, not independently confirmed per
  condition. Several of those (calling an undefined Sub/Function, in particular) would
  actually be a *compile*-time failure in real VBA, never reaching `On Error` at runtime
  at all — a known, disclosed divergence, not fixed here.
- **`Err.Raise Number[, Source][, Description]`** parses real VBA's full positional-slot
  grammar, including the idiomatic `Err.Raise 513, , "custom text"` form that skips
  `Source` — a naive two-positional-argument implementation would misread that
  `"custom text"` as `Source` instead of `Description`, since real VBA's slot order is
  (Number, Source, Description, HelpFile, HelpContext), not (Number, Description).
  `Source` is parsed (so this can't happen) but not modeled as a readable property —
  `Err.Source` doesn't exist here, matching this project's existing choice not to model a
  VBA project/module naming concept elsewhere. `HelpFile`/`HelpContext` aren't parsed at
  all. Raising without an explicit `Description` fills in the real VBA description text
  for the numbers above, or the 1004 catch-all text otherwise.
- `Err.Number`/`Err.Description` reset to `0`/`""` at the start of each `run_sub`/
  `run_sub_multi` call and on `Err.Clear`. Deliberately does **not** auto-clear on `On
  Error GoTo 0`/a fresh `On Error` statement — the common idiom above relies on
  `Err.Number` surviving past `On Error GoTo 0` to be inspectable at all, and the exact
  real-VBA clearing rule around `Resume`/`On Error` re-statements wasn't independently
  confirmed, so this stays conservative rather than guessing.
- 16 new tests (10 `src/vm/mod.rs`, 6 `src/parser/mod.rs` covering AST shape, the
  `Err.Raise`-skips-`Source` case at both layers, and a regression test confirming
  `pending_raised_error` — the side channel `Err.Raise` uses to preserve its own
  number/description across the generic error-classification path — can't leak into
  an unrelated later error: it's consumed synchronously by the first `On Error Resume
  Next`/`GoTo` catch on the same unwind, before any other statement can run) —
  `cargo test --workspace` 857/857
  (857 = 741 lib + 1 + 15 + 16 + 5 + 14 + 7 + 7 + 17 + 15 + 19, every test binary summed),
  no regressions in the 581-scenario corpus classifier or the 386-case `vba-semantics`
  suite (still 0 `BUG`/0 `UNCLASSIFIED`, same 19 `KNOWN_LIMITATION`).
- **Known limitation found while verifying the above, not fixed here (pre-existing,
  unrelated to `Err.Raise` specifically — reproduces with any error, e.g. `1 / 0`):**
  `On Error GoTo <label>` set in a caller does not run its handler if the error instead
  occurs inside a *called* Sub/Function that has no `On Error` of its own. Root cause:
  `on_error_goto_label` is a single `Vm`-wide field with no per-call-frame scoping, so
  the callee's own `exec_body` sees the caller's still-armed label, tries (and fails) to
  find it among the callee's own statements, and returns a synthetic "label not found"
  error instead of ever reaching the caller's handler. Needs a real design decision
  (save/restore `on_error_goto_label` — and likely `on_error_resume_next` — per call
  frame) rather than a local patch, so it's deliberately left for a dedicated fix rather
  than folded into this feature commit.

### Array declaration/bounds gaps: `Dim arr(lo To hi)`, `Option Base 1`, `Dim arr()`, `Erase`

Fixes four of the five `array_bounds` `KNOWN_LIMITATION` cases in `compat/vba-semantics/`
(19 → 16 — see that suite's own CHANGELOG-adjacent note below for the fifth, newly
disclosed rather than fixed):

- **`Dim arr(2 To 8)`** — an explicit non-zero lower bound — now parses (`ArrayDim { lower:
  Option<Expr>, upper: Expr }` replaces a bare `Expr` per dimension in `DimArray`/`ReDim`'s
  AST) and is honored: `LBound(arr)` is `2`, and `arr(2)`/`arr(8)` address the real first/
  last elements. `Option Base 1` — previously parsed and silently discarded at module level
  — now sets the default lower bound for declarators that don't give an explicit `lo To hi`
  (`Program.option_base`, read by `Dim`/`ReDim` at execution time). Storage stays a flat
  `Vec<Variant>` (`elixcee-types`' public `Variant::Array` is untouched — no semver bump):
  the lower bound is tracked separately, per array *variable name*, in a new `Vm`-side
  `array_lower_bounds` map (`LBound`/`UBound`/`ArrayWrite`/array-subscript reads all resolve
  arrays by name already, so this needed no public-surface change). An array value with no
  name to key on — `Split()`/`Array()`'s return, or any array-valued expression not bound to
  a `Dim`'d variable — defaults to lower bound 0, unchanged from before.
- **`Dim arr()`** (empty parens, a dynamic array sized later by `ReDim`) now parses — the
  declarator's dimension list is simply empty — and creates an unsized placeholder array
  ReDim can then legally resize, matching the one documented use this suite tests (`Dim
  arr()` immediately followed by `ReDim arr(5)`). elixcee doesn't model the stricter real-
  VBA behavior of raising "Subscript out of range" if `UBound`/an element is accessed
  *before* the first `ReDim` — not exercised by any case, not attempted.
- **`Erase arr`** — verified (checked the pre-change `parse_ident_stmt`, not inferred from
  the old registry entry's "IsEmpty is still False" description) to have had no `Erase`
  statement dispatch at all: `erase` wasn't a recognized keyword, so `erase arr` fell all
  the way through to the generic "bare identifier statement" fallback and became a
  `Stmt::Unsupported` no-op. Is now a real `Stmt::Erase { name }`: resets every element of a
  fixed-size array back to `Empty` in place, leaving its bounds untouched (matching real
  VBA's documented behavior for a statically-declared array). Real VBA's comma-separated
  `Erase a, b` form isn't parsed — no case needs it.
- `array_oob_error`'s `ArrayIndexOutOfBounds` diagnostic evidence used to hardcode
  `lower: 0` unconditionally; now reports the array's actual lower bound and the VBA-facing
  index that was attempted (not an internal, bound-shifted one), for the two call sites that
  now track a real bound. The two UDT-array call sites (`DimArrayRecord`/`ArrayRecordSet`/
  `ArrayRecordGet` — `Dim arr(10) As MyType`) are unaffected and still report `lower: 0`,
  matching their existing (unchanged) always-0-based behavior.
- **Found while verifying the above, and separately disclosed (not fixed — see below):**
  the fifth `array_bounds` `KNOWN_LIMITATION` case's own description ("UBound(arr, 2)
  ignores its dimension argument ... even though the array's own storage genuinely is
  two-dimensional") turned out to be **factually wrong**. elixcee's array storage is
  genuinely 1-D: `Dim arr(3, 2)` only ever allocates dimension 1's 4 elements (dimension
  2's size is parsed and discarded), and every array write/read (`Stmt::ArrayWrite`, the
  `Expr::FuncCall` array-subscript read path) indexes using only the first index
  expression — a second or later index is silently ignored on *both* sides, so `arr(2,
  0) = 111` followed by `arr(2, 1) = 222` overwrites the same element (confirmed live:
  both `arr(2,0)` and `arr(2,1)` then read back `222`). The suite's own
  `two_dimensional_array_write_and_read_round_trips` case had been cited as evidence 2-D
  addressing worked — it passed only because its single write and single read happened to
  use the *same* second index on both sides, a coincidental round-trip that never exercised
  the collision. Renamed to `two_dimensional_array_second_index_is_silently_dropped`,
  reshaped to actually discriminate (write two elements differing only in the second
  index, confirm the first wasn't clobbered — it currently is), and registered as a new,
  previously-undisclosed `KNOWN_LIMITATION`. `ubound_second_dimension_argument_ignored`'s
  own `knownLimitation` text is corrected to stop citing the now-corrected sibling case as
  proof of working 2-D storage. Fixing this for real needs shape metadata and stride
  arithmetic in both the write and read paths — comparable in scope to this project's other
  deferred Variant-surface work (see the Date/Time note elsewhere in this file) — and was
  deliberately not attempted alongside the smaller, independent lower-bound-tracking fixes
  above.
- 13 net new tests (8 `src/vm/mod.rs`; `src/parser/mod.rs` net +5 — 6 added, replacing the
  now-stale `test_option_base_ignored` with two narrower tests that assert the captured
  value instead of just "didn't break parsing") — `cargo test --workspace`
  870/870 (870 = 754 lib + 1 + 15 + 16 + 5 + 14 + 7 + 7 + 17 + 15 + 19, every test binary
  summed), `cargo clippy -p elixcee-types --all-targets -- -D warnings` clean, `cargo build
  --release --workspace` clean, `cargo check --features python --lib` clean (only the
  pre-existing, disclosed pyclass deprecation warning). `compat/corpus/` unaffected (581
  scenarios, 0 `UNEXPLAINED`/0 `MISMATCH`). `compat/vba-semantics/`: 386 cases, 0 `BUG`/
  0 `UNCLASSIFIED`, 16 `KNOWN_LIMITATION` (down from 19: four fixed, one newly disclosed
  as described above) — see `compat/vba-semantics/README.md`'s "Current state" for the
  full breakdown.

### `src/reader.rs`'s `xml:space="preserve"` whitespace defect on `t="str"` cells

`xlsx_sheet_cells`'s `<v>`-text handler called `xlsx_parse_cell(text.trim(), ...)`
unconditionally — for a `t="str"` cell (whose literal text lives directly in `<v>`, unlike
`t="s"` shared-string cells or inline `<is><t>` strings, neither of which this call site
ever trims), that silently dropped significant leading/trailing whitespace whenever the
source XML marked it with `xml:space="preserve"`. Confirmed live against
`compat/corpus/workbooks/with_text.xlsx`'s own raw `sheet1.xml`: cell A3 is `<c t="str"><v
xml:space="preserve">  padded  </v></c>`, read back as `"padded"` instead of `"  padded
  "`. Disclosed since the round that found it via `compat/differential/classify.mjs`'s
`UNSUPPORTED_ALLOWLIST` (`XML_SPACE_PRESERVE_DEFECT`, registered under both `read` and
`readFile`); reachable through `read()`/`readFile()`/`readFileSync()` alike, since the
latter two are thin wrappers over `read()`.

Fix: `xlsx_sheet_cells` now reads `<v>`'s own `xml:space` attribute (a new `v_preserve_space`
local, re-read fresh on every `<v>` open — no stale carry-over from a previous cell in the
same row) and skips the trim when it's `"preserve"`, matching plain XML `xml:space`
semantics rather than special-casing `t="str"` specifically. Real Excel/SheetJS writers
never emit this attribute on a numeric/boolean `<v>` (whitespace is never meaningful there),
so this doesn't change default behavior for any realistic file — a regression test confirms
a numeric `<v xml:space="preserve">42</v>` still parses even though `f64::parse` itself
rejects surrounding whitespace, which the fix's unconditional (not `t`-gated) skip could in
principle have broken for a pathological input.

Both `UNSUPPORTED_ALLOWLIST` entries (`with_text.xlsx:xml_space_preserve_trimmed` under
`read` and `readFile`) are removed — the allowlist is empty again — and the now-dead
`unsupportedCaseId` plumbing threading them through
`compat/differential/xlsx-read.test.mjs`'s `with_text.xlsx` fixture cases is dropped too,
matching this project's established precedent for closing a disclosed reader defect (see
this same file's `classify.mjs` comment history). `differential:read`: 33/33 MATCH, 0
disclosed (was 30 MATCH + 3 disclosed). Vendored WASM artifact
(`packages/xlsx/src/internal/wasm/`) rebuilt via `crates/elixcee-wasm/build.sh` so
`@elixcee/xlsx`'s `read()` actually carries the fix — `wasm:smoke` and
`differential:utils`/`:ssf-format`/`:metadata` all still clean, confirming no regression
from the rebuild itself.

2 new tests in `src/reader.rs` — `cargo test --workspace` 872/872 (872 = 756 lib + 1 + 15 +
16 + 5 + 14 + 7 + 7 + 17 + 15 + 19, every test binary summed), `cargo clippy -p
elixcee-types --all-targets -- -D warnings` clean, `cargo build --release --workspace`
clean. `compat/corpus/` and `compat/vba-semantics/` unaffected (verified by re-running both
after the reader.rs change — neither exercises this XML shape).

## [0.5.0]

Root `elixcee` (Rust crate + Python package) **and** `elixcee-types` (`0.1.0` → `0.2.0`,
a minor bump, not a patch — see "`elixcee-types` 0.2.0" below for why). `elixcee-wasm`
stays `0.1.0` (no source changes; its vendored build output was regenerated to pick up the
VM changes below, and its own `elixcee-types`/`elixcee` path dependencies carry no version
requirement to update), and `@elixcee/xlsx` stays
`0.0.0-development`/unpublished/`"private": true` (none touched). Built via two parallel,
disjoint-scope worktree branches — VBA structural
semantics (parser/VM: colon-statement separator, `Variant::Null` with documented
propagation rules, real object-variable unset/`Nothing` state, a runtime `With` stack) and
`@elixcee/xlsx` consumer/browser validation (a real packed-tarball install smoke, a real
headless-Chrome smoke, bundle-safe WASM loading, `readFile()`) — merged after each was
independently reverified, then integration-regression-tested together as a whole,
surfacing and fixing one real interaction bug neither branch's own tests caught (a bare
`.member` inside a single-line `If` nested in a `With` body). Does not claim Microsoft
Excel validation anywhere. Full detail in `[Unreleased]`'s two sections below (VBA
structural semantics; `@elixcee/xlsx` real-consumer and real-browser validation) and the
single-line-`If`/`With` interaction fix that follows them.

All gates green before this bump: `cargo test --workspace` (724 passing),
`cargo build --release --workspace`, `cargo check --features python --lib`,
`cargo clippy -p elixcee-types -- -D warnings`, a real `maturin build --release` wheel
installed into a fresh venv with `Null`/object-`Nothing`-alias-safety/the `With` stack
re-verified through the actual Python API post-install (not just `cargo check`); the
`compat/vba-semantics/` suite (386 cases, 0 `BUG`/0 `UNCLASSIFIED`, 19 `KNOWN_LIMITATION`,
down from 28, deterministic across 2 runs); the existing 581-scenario `compat/corpus/` suite
(0 `UNEXPLAINED`/0 `MISMATCH`, unchanged); `compat/differential/`'s utils/SSF/read+readFile/
metadata suites (all passing, read+readFile now 30 MATCH + 3 disclosed); a fresh
`wasm-pack` rebuild of both targets verified via the real packed-tarball consumer smoke and
a real headless-Chrome smoke (not just Node simulating the `browser` export condition).

### `elixcee-types` 0.2.0

- **Added `Variant::Null`** to the public `Variant` enum (see "VBA structural semantics"
  below for what it's for). **This is a public-enum-variant addition, not a purely additive
  change** — any downstream consumer doing an exhaustive `match` on `Variant` (rather than
  ending in a `_ =>` catch-all) fails to compile against this version until it adds a
  `Variant::Null` arm. Bumped `0.1.0` → `0.2.0` (a real minor bump, not left at `0.1.0`)
  specifically because of this — `elixcee` `0.5.0` depends on `elixcee-types = "0.2.0"`
  (previously `"0.1.0"`), so a `cargo build`/`cargo publish` of `elixcee` against the old
  `elixcee-types` `0.1.0` on crates.io would fail to resolve `Variant::Null` at all. No
  other public API surface changed.

### VBA structural semantics

Four language-level gaps closed, each verified against Microsoft's own VBA language
reference (fetched live, not recalled) before being encoded as an expectation.
`compat/vba-semantics/` grew **301 → 386 cases**, with `KNOWN_LIMITATION` **28 → 19**
(nine genuinely fixed, annotations removed rather than weakened — never by changing what a
case expects). `compat/corpus/`'s 581-scenario regression baseline stays at 0 `UNEXPLAINED`
/ 0 `MISMATCH`; `cargo test --workspace` passes; `report.json` is byte-identical across two
consecutive runs.

#### Added

- **The `:` multi-statement-per-line separator** — `a = 1: b = 2: c = 3`, `label1: a = 1`,
  `MsgBox "x": Exit Sub`, `For i = 1 To 3: … : Next i`, and a single-line `If`'s own
  `:`-separated Then/Else statement lists (per the If…Then…Else reference: "One or more
  statements separated by colons; executed if condition is True", example
  `If A > 10 Then A = A + 1 : B = B + A : C = C + B`). Handled in the parser via the
  tokenizer's existing `Tok::Colon`, **never** as a pre-tokenize `:`→newline rewrite —
  which would corrupt a colon inside a string literal, break the `label:` form, and mangle
  the single-line `If`'s clause boundary. All three are pinned by tests. Each
  colon-separated statement keeps its own `SourceSpan`, so `--json`'s `location` still
  points at the individual statement, not the line.
- **`Variant::Null`** — VBA's "no valid data" value, now genuinely distinct from `Empty`
  (an uninitialized Variant). Implements the documented rules: arithmetic propagates Null
  from either side (and *before* operand coercion, so `5 / Null` is Null, not a
  Division-by-zero error); `&` propagates only when *both* sides are Null (a single Null is
  a zero-length string); all six comparison operators produce Null, including
  `Null = Null`; `And`/`Or`/`Xor`/`Not` follow their three-valued truth tables, in which
  Null does *not* uniformly propagate (`False And Null` is False, `True Or Null` is True);
  `If Null Then` treats the condition as False (documented, not an error); `IsNull` and
  `IsEmpty` are now separate questions; `TypeName(Null)`/`VarType(Null)` are `"Null"`/`1`;
  and a Null reaching a genuinely numeric context raises error 94, `Invalid use of Null`.
  Adds **no new external surface** — Null serializes exactly as `Empty` already does (JSON
  `null` / Python `None` / blank cell), so `--json`, the Python bindings and the xlsx/ods
  writers are unchanged.
- **`ObjectRef::Nothing`** — a real unset/Nothing state for object variables.
  `Dim r As Range|Worksheet|Workbook` registers the name as declared-but-unset;
  `Set r = Nothing` assigns the null reference (it used to silently no-op); every
  member-access path raises real VBA's error 91 text, `Object variable or With block
  variable not set`, from one shared constant. `Set r = Nothing` clears only `r` — a
  `Set r2 = r` alias made earlier stays live and still reads and writes through to the same
  Range. `<var> Is Nothing` now parses and reflects each variable's own state (only the
  `Is Nothing` shape; a general `a Is b` is still unparsed rather than guessed at).
- **New stable error code `E1007`/`object_variable_not_set`**, documented in
  `docs/agent-contract.md` — a genuinely new error condition, not free-text reuse of an
  existing code.
- **`Array(...)`** builtin — builds a zero-based Variant array from its arguments.

#### Changed

- **`With`-target resolution is now a runtime mechanism, not a parse-time textual
  rewrite.** The target expression is captured unevaluated (`ast::WithTarget`), evaluated
  **once** on block entry, and pushed onto `Vm::with_stack`; a bare `.member` is a
  first-class statement and expression form (`Stmt::WithDot`/`Expr::WithDot`) resolved
  against the innermost entry wherever it appears in the AST. Consequences:
  `With Cells(r, c)` (any computed target) works; a bare `.member` works at any nesting
  depth inside `If`/`For`/`Do`/`Select Case` in the body; reassigning a target variable
  inside the body no longer could (and still cannot) retarget the block; nesting restores
  each outer target in turn; and the stack is popped on *every* exit path, including
  `Exit Sub`/`Exit For` and a runtime error, so a target can't leak into whatever runs
  next. The parser's `with_target`/`with_range_target` fields and the `Stmt::WithRecord`
  variant are gone.
- **`With ws` (a Worksheet-typed object variable) now qualifies `.Cells(r, c)` to that
  worksheet.** It previously wrote to whatever sheet happened to be active — a real,
  previously-undisclosed bug, surfaced by the runtime-stack work.
- **`For Each c In Range(...)` binds the loop variable as a live single-cell Range object**
  as well as a plain value, so `c.Value` reads that cell. It previously fell through to the
  UDT path and silently yielded `Empty`. Found by `compat/corpus/` reacting to the
  `Dim c As Range` change above, not by source audit.
- **A non-numeric string operand of `+`/`-`/`*`/`/`/`^` raises `Type mismatch`**, real
  VBA's documented wording ("One expression is a numeric data type and the other is a
  String | A `Type mismatch` error occurs"). Applied narrowly, via a new `arith_to_f64`
  wrapper used only by `eval_binop`'s `Add|Sub|Mul|Div|Pow` arm — the shared `to_f64`
  helper and its ~53 other call sites, each with its own correct wording for its own
  failure, are untouched. That blast radius was the exact reason this stayed disclosed
  rather than fixed when it was first found. **Not** extended to `\`/`Mod`, which go
  through `to_i64_rounded` and keep the previous wording; the rule cited above is from the
  `+` operator page, and widening it further would re-enter the blast radius this fix was
  scoped to avoid.
- `Dim x: x = 5` now parses — the declarator's trailing-syntax tolerance loop was swallowing
  the `:` separator. Found by a new suite case, not by source audit.
- **A bare `.member` branch inside a single-line `If` nested in a `With` body now runs.**
  `parse_stmt` gained a `Tok::Dot` arm for the runtime With-stack work above, but
  `parse_single_line_if_branch`'s own dispatch checked only `Tok::Ident` and was never
  updated to match — so `If .Value > 0 Then .Value = .Value + 1` inside `With Range("A1")`
  silently degraded to `Stmt::Unsupported` (no parse error, but the assignment never ran).
  Same bug *class* as the pre-existing `Range()`/`Cells()`-in-single-line-`If` fix (a
  single-line-`If` branch dispatch lagging behind block-form `parse_stmt`'s own statement
  coverage) — found during integration by manually exercising a README code sample, not by
  either subagent's own test suite, which is exactly the kind of interaction gap that can
  slip between two disjoint-scope changes that never ran against each other until merged.

### `@elixcee/xlsx` — real-consumer and real-browser validation

Closes the gap between "the differential suites pass" and "a real npm consumer, and a real
browser, actually works" — every prior check reached the package via a relative import into
`packages/xlsx/src`, or (for the `"browser"` export condition) via Node simulating that
condition, never an actual browser process.

#### Added

- **`XLSX.readFile()`/`readFileSync()`** — one function under both names (matching the real
  `xlsx` package's own identity: same `.name`, `.length`, key order), wrapping the existing
  byte-buffer `read()`. Differential-tested file-by-file against the real `xlsx@0.18.5`
  oracle, with and without `cellStyles`/`cellDates`. Throws `ELIXCEE_UNSUPPORTED_IN_BROWSER`
  from the browser entry point rather than faking a filesystem. `write*` remains
  unimplemented.
- **A packed-tarball consumer smoke test** (`packages/xlsx/scripts/pack-consumer-smoke.mjs`,
  `npm run pack:consumer`) — runs a real `npm pack`, `npm install`s the exact `.tgz` into a
  throwaway package under `os.tmpdir()`, and exercises `require()`, `import`, a TypeScript
  compile, `XLSX.read()`, CJS/ESM export-set identity, and the `"browser"` export condition
  entirely from inside that install — asserting the resolved paths land under the throwaway
  `node_modules/@elixcee/xlsx`, not a relative path back into this repo. Every earlier check
  in this project could have passed while the actual published tarball was broken; this one
  can't.
- **A real headless-browser smoke test** (`packages/xlsx/scripts/browser-smoke.mjs`,
  `npm run browser:smoke`) — launches an actual local Chrome/Chromium process (via its own
  `--dump-dom`, no browser-driver dependency added — evaluated and rejected
  playwright-core/puppeteer-core/chrome-remote-interface as unnecessary weight for "load one
  page, read one result"), serves an esbuild browser bundle over real `node:http`, and reads
  `XLSX.read()`'s result back out of the page's own DOM: sheet names, a real cell value, an
  exported-function count, zero page-observable console/uncaught errors, zero non-200
  responses for any page-referenced resource. **Distinct from, and strictly more than, the
  pre-existing `node --conditions=browser` check** (still present, in `wasm:smoke`) — that
  one is Node simulating an export condition and proves nothing about a browser; this one is
  a real browser. Neither is described as the other anywhere in code, CI step names, or this
  entry. Safari is not covered and not claimed.
- **CI**: the packed-tarball smoke joins `node-js` (both Node versions); the real-Chrome
  smoke and a CJS *and* ESM esbuild-bundle smoke (as distinct steps) join `wasm`, along with
  a diagnostic step that prints whatever browser the runner image actually provides, so a
  missing-Chrome failure is self-explanatory from the job's own log rather than a guess.
  `compat/differential/`'s own `classify.mjs`/`normalize.mjs` self-checks — existing package
  scripts that pin the exact contents of the disclosed-divergence registries — are now
  wired into CI too; they never ran there before.

#### Fixed

- **The Node/CJS WASM loader's `.wasm` lookup is no longer `__dirname`-relative.**
  `elixcee_wasm.node.cjs` (wasm-pack's own generated code, not hand-written) located its
  compiled WASM via a path relative to its own file location — bundle-*output*-relative once
  a consumer bundled it, not source-relative. ESM bundle output has no `__dirname` at all (a
  hard `ReferenceError`, not a silent failure); CJS bundle output technically had `__dirname`
  but pointed at the wrong directory, so it only worked if the consumer manually copied
  `elixcee_wasm_bg.wasm` next to their bundle. Fixed by inlining the compiled WASM as base64
  directly into the Node loader too (`crates/elixcee-wasm/build-node-inline.mjs`, mirroring
  the technique `build-browser-inline.mjs` already used for the browser build) — generated by
  `build.sh`, never hand-patched, so a fresh rebuild reproduces the committed artifact
  byte-for-byte. No `.wasm`-copy step is required for CJS *or* ESM bundling anymore, and
  browser bundling — previously broken outright (`esbuild --platform=browser` failed
  resolving `fs`) — now works too. The raw `elixcee_wasm_bg.wasm` is no longer vendored
  separately (both loaders already carry the bytes; shipping it too would double-ship
  263 KB), and `package.json` gained a `browser` field stubbing the Node loader out of
  browser bundles, so a browser consumer pays for the WASM payload once, not twice.
  Synchronous `read()` is unaffected — no `await init()` introduced anywhere.
  **Package-size impact**, measured against 0.4.0: packed tarball 339,098 → 380,005 bytes
  (+12.1%), unpacked 741,304 → 835,712 (+12.7%); the WASM payload itself is unchanged at
  263 KB (only its containers, base64-inlined, grew). No hard size gate — recorded so a
  future round can judge whether it grows *further* without basis.

#### Discovered, disclosed (not fixed — outside this round's scope)

- **`src/reader.rs` trims every cell's text unconditionally**
  (`xlsx_parse_cell(text.trim(), …)`), ignoring the `xml:space="preserve"` attribute real
  XLSX XML uses to mark significant leading/trailing whitespace — a cell whose real value is
  `"  padded  "` reads back as `"padded"`. Confirmed live against
  `compat/corpus/workbooks/with_text.xlsx` cell A3 (oracle: `"  padded  "`; elixcee:
  `"padded"`). Reachable through both `read()` and `readFile()`. Registered in
  `compat/differential/classify.mjs`'s `UNSUPPORTED_ALLOWLIST` (3 cases, one root cause) with
  a full writeup rather than silently excluded from the fixture set — the classifier's own
  self-check pins the exact entry count, so it can't go stale unnoticed. Fixing it means
  honoring `xml:space` on the `<t>` element rather than trimming at the call site, and
  re-checking the trim isn't load-bearing for the numeric/boolean paths sharing
  `xlsx_parse_cell` — `src/reader.rs` is shared surface, not `@elixcee/xlsx`-specific, so
  this is recorded for whoever next touches the reader, not fixed here.

## [0.4.0]

Root `elixcee` (Rust crate + Python package) only — `elixcee-types` and `elixcee-wasm`
both stay `0.1.0` (no source changes this release; `elixcee-wasm`'s vendored build output
was regenerated to pick up the fixes below, but its own `src/` wasn't touched) and
`@elixcee/xlsx` stays `0.0.0-development`/unpublished. Covers this round's `compat/vba-
semantics/` expansion (208 → 301 cases, 0 `BUG`/0 `UNCLASSIFIED`, 28 disclosed
`KNOWN_LIMITATION`s — see that entry below and `compat/vba-semantics/README.md` for the
full breakdown), the new `wasm` CI job, and several real behavior-changing bug fixes found
while building the suite (Boolean arithmetic, `WorksheetFunction` numeric coercion, `Empty`
equality, single-line-`If` statement recognition) — all gates green before this bump:
`cargo test --workspace` (683 passing), `cargo build --release --workspace`,
`cargo check --features python --lib`, `cargo clippy -p elixcee-types -- -D warnings`, the
existing 581-scenario `compat/corpus/` suite (0 `UNEXPLAINED`/0 `MISMATCH`, unchanged), and
the new `wasm`/existing `node-js` CI jobs both confirmed green on real GitHub Actions, not
just locally.

### Added

- **`compat/vba-semantics/`, a new VBA value-correctness suite** — a genuinely different
  question from `compat/corpus/`'s own "does elixcee run without erroring": is the VALUE
  elixcee produces the one real, documented VBA semantics says it should be. Needs no
  oracle — `reference/*.mjs` are small, independently-checkable pure-JS reference
  implementations of documented real VBA semantics (banker's rounding, `Str()`'s
  leading-space quirk, `Val()`'s prefix parsing, `And`/`Or`/`Xor`/`Not`'s logical-vs-bitwise
  split, ...), used to compute cases' expected outcomes programmatically. Six-
  verdict classification (`MATCH_DOCUMENTED_SEMANTICS`/`EXPECTED_ERROR`/`NONDETERMINISTIC`/
  `KNOWN_LIMITATION`/`BUG`/`UNCLASSIFIED`); `BUG`/`UNCLASSIFIED` both gate at 0. Started at
  208 cases across 12 categories; grew to **301 cases across 18** in the same round that
  added the `+`-vs-`&` operator-coercion, comparison-operator-coercion, `Select Case`
  matching, `With`-block-resolution, and array-bounds categories — each expected value
  sourced from Microsoft's own VBA language reference, fetched live while writing the
  cases, not recalled from memory. Current state: 253 `MATCH_DOCUMENTED_SEMANTICS` + 18
  `EXPECTED_ERROR` + 2 `NONDETERMINISTIC` + 28 `KNOWN_LIMITATION` = 301, 0 `BUG`,
  0 `UNCLASSIFIED`. All 28 `KNOWN_LIMITATION` cases are divergences found while building
  this suite and not fixed this round (several *other* divergences found the same way
  *were* fixed — see "Fixed" below); grouped by root cause with the full breakdown in
  `compat/vba-semantics/README.md`.
- **CI now runs `@elixcee/xlsx`'s own tests.** `.github/workflows/ci.yml` gained a `node-js`
  job (Node 20/22 matrix): `packages/xlsx`'s TypeScript typecheck (with and without the DOM
  lib present) and all four `compat/differential/` suites (`utils`/`ssf-format`/`read`/
  `metadata`). Previously none of this ran anywhere except a developer's own machine, despite
  every command already working — verified live before wiring each one in, not assumed from
  this file's own previously-claimed numbers.
- **CI also now builds and smoke-tests the WASM bridge from scratch.** A new `wasm` job runs
  both `wasm-pack build --target nodejs` and `--target web` fresh (the `node-js` job above
  only ever consumed the already-vendored/committed copy — a build-breaking change to
  `crates/elixcee-wasm`/`src/reader.rs` had no CI signal until now), then runs the new
  `packages/xlsx/scripts/wasm-smoke.mjs`: a Node synchronous `read()` call, the `"browser"`
  export condition resolving *and actually running* (via `node --conditions=browser`,
  self-referencing the package by name — more than a resolution check, but still Node
  simulating the condition; no real browser executes anywhere in this project's CI, and no
  Safari support is claimed anywhere), a minimal `esbuild` bundle with an in-bundle
  `XLSX.read()` call, and the current WASM binary size logged (263,204 bytes as of this
  round) — recorded, not gated against any threshold (no prior baseline exists to compare
  against). `esbuild` is `packages/xlsx`'s one new devDependency for this (pinned to `^0.28`,
  past the version with the known dev-server CORS advisory — irrelevant to this project's
  usage, which only ever calls its one-shot `build()`, never `serve()`, but avoided anyway).
  One real, previously-undisclosed consumer caveat found while writing the bundle check: the
  Node/CJS WASM loader (`elixcee_wasm.node.cjs`, wasm-pack's own generated code, not
  hand-written) locates its `.wasm` file via a `__dirname`-relative path, which becomes
  bundle-output-relative once bundled — a consumer bundling this package's Node entry needs
  to bundle to CJS (ESM output has no `__dirname` at all, a hard `ReferenceError`) and copy
  `elixcee_wasm_bg.wasm` next to their bundle output, or externalize the loader. Not fixed
  this round (would mean patching wasm-pack's own generated boilerplate); documented in
  `wasm-smoke.mjs`'s header comment and `ROADMAP.md`.
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
- `Range(...)`/`Cells(...)`/`MsgBox`/etc. weren't recognized inside a single-line `If`'s
  Then/Else branch — only identifier-led statements were, so `If x > 0 Then Exit Sub Else
  Range("A1").Value = 1` mis-parsed the Else branch as an array write to a variable
  literally named "range", failing with "Cannot convert 'A1' to number". Found by
  `compat/vba-semantics/` on exactly this shape, not by source audit. Fixed by extracting
  the full statement dispatch (previously duplicated as a narrower subset for single-line
  `If`) into one shared function used by both the block-form and single-line-`If` parsers.
  That extraction briefly regressed assignments to a variable literally named after a block
  keyword (`do = 0`, `select = 1`, ...) — caught by the existing property test before
  shipping, fixed by re-ordering the "bare `name = ...` is always assignment" check ahead
  of the block-construct keyword dispatch.
- VBA's `+`/comparison operators coerced Boolean `True` to `1.0` instead of VBA's own
  documented internal value of `-1` (`CInt(True)` is `-1` in real VBA) — `True + 5` was `6`,
  not the documented `4`. Found via `compat/vba-semantics/`'s operator-coercion matrix,
  fetched from Microsoft's own VBA language reference rather than recalled from memory.
  Fixing this then silently changed `WorksheetFunction.Sum`/`Max`/`Min`/`Average`/`SumIf`/
  `Round`/`Abs`/`Sqrt`/`Power`/`Log`/`Index` too (`WorksheetFunction.Sum(True, True)` went
  from `2` to `-2`) — wrong, since `Application.WorksheetFunction` bridges into Excel's own
  calculation engine and must keep using Excel's `TRUE = 1` coercion (matching a worksheet
  formula), not VBA's own arithmetic rule. Caught in the same round by checking every other
  caller of the shared coercion helper before considering the fix complete; `WorksheetFunction.*`
  now has its own copy of just the Boolean arm.
- The `=`/`<>` operators had no rule for comparing `Empty` against a number or string —
  `0 = Empty`/`"" = Empty` both fell through to an unconditional `False`, inconsistent with
  `<`/`>` on the exact same operand pairs (which already correctly treated `Empty` as `0`).
  Real VBA documents `Empty` as numeric-comparing as `0` and string-comparing as `""` for
  every comparison operator, not just some of them. Found via the same operator-coercion
  matrix.

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
