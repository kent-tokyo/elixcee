# XLSX architecture ADR

## Status

Accepted for planning purposes. Describes the **target** shape for the `@elixcee/xlsx`
initiative and the concrete resolutions for its two hardest structural problems. Phase 0
implements none of the Rust-side changes described here — it only creates the
documentation and the `compat/oracle` + `compat/differential` Node harness. Everything
under "Decision" below is scoped work for later phases.

**Partially superseded by what actually shipped** (confirmed against
`crates/*/Cargo.toml`, not assumed): the `elixcee-types` extraction (see "Formula/VM
circular-dependency resolution" below) did land, matching this plan. `crates/elixcee-wasm`
also now exists and ships (real Node/browser smoke tests, wired into CI) — but it depends
directly on the still-monolithic root `elixcee` crate, not on a new `elixcee-xlsx` crate as
planned in "Target workspace shape" and assumed by "Consequences" below; the
`elixcee-formula`/`elixcee-vba`/`elixcee-xlsx`/`elixcee-cli` split was never done. See
`ROADMAP.md`'s "Current state" for what's actually implemented today; this document
remains accurate as a record of the plan and the Phase 2B-0 WASM feasibility investigation,
just not as a description of the crate layout that was ultimately built.

## Context

elixcee today is a single Rust crate with two runtime dependencies (`zip`, and optional
`pyo3` behind the `python` feature) and a documented history of *removing* dependencies
rather than adding them (`rust_xlsxwriter`, `pest`/`pest_derive`, and `calamine` were all
dropped from the runtime build in favor of hand-written code — see `CHANGELOG.md`). Any
new dependency proposed for this initiative works against that established direction and
needs its own justification, not a default yes.

The target compat surface, `xlsx@0.18.5`, ships as `main: xlsx.js` (CJS) +
`module: xlsx.mjs` (ESM) + `types: types/index.d.ts`, with a buffer-first API
(`XLSX.read(data, opts)` accepting Buffer/Uint8Array/base64/array, not just file paths).
elixcee's own reader (`src/reader.rs`) is path-only. Reconciling these two facts —
"keep the hand-rolled, dependency-free reader" and "support a buffer-first API" — is one
of the two structural problems this ADR resolves. The other is a circular dependency
between `src/formula/` and `src/vm/` that would otherwise force any XLSX-facing crate to
depend on the full VBA execution engine.

## Decision

### Target workspace shape

```
elixcee/
├── Cargo.toml                 # workspace root (not yet — see "Cargo.toml" below)
├── crates/
│   ├── elixcee-types/         # NEW leaf: Variant, ExcelError, CellContent, date-serial math
│   ├── elixcee-formula/       # today's src/formula/, depends only on elixcee-types
│   ├── elixcee-vba/           # today's src/parser/ + src/vm/, depends on types + formula
│   ├── elixcee-xlsx/          # NEW: reader/writer generalized to buffer I/O, xlsx-compat surface
│   ├── elixcee-wasm/          # NEW: wasm-bindgen shim over elixcee-xlsx (built when a phase needs it)
│   └── elixcee-cli/           # today's src/main.rs
├── packages/
│   └── xlsx/                  # @elixcee/xlsx npm package (Phase 1)
├── compat/
│   ├── oracle/                # Phase 0 — machine-generated xlsx@0.18.5 API manifest
│   ├── fixtures/              # not yet created
│   ├── differential/          # Phase 0 — comparison/classification harness
│   └── dependents/            # not yet created
├── fuzz/
└── docs/
```

**This is the destination, not a Phase 0 task list.** Phase 0 creates only
`compat/oracle/` and `compat/differential/`. Nothing under `crates/`, `packages/`,
`compat/fixtures/`, or `compat/dependents/` exists yet — the single-crate `src/` layout
is unchanged.

Dependency direction, once the crate split happens:

```
@elixcee/xlsx (npm)
        │
        ▼
  elixcee-wasm
        │
        ▼
  elixcee-xlsx ──────────► elixcee-types
        ▲                       ▲
        │                       │
  elixcee-vba ──────────────────┘
        ▲                       ▲
        │                       │
Python / CLI / existing elixcee elixcee-formula
```

`elixcee-xlsx` must never depend on `elixcee-vba`. The next two sections describe why
that isn't true today and exactly what closes the gap.

### Formula/VM circular-dependency resolution

Verified directly against current file contents (`src/vm/mod.rs:1-260,2413-2434` and
`src/formula/eval.rs:1-15,2888,3970,4358`), not assumed from a first pass: `formula/eval.rs`
imports `Variant`/`ExcelError`/`CellContent` from `vm/mod.rs`, and `vm/mod.rs`'s own
`serial_to_display` (line 216) calls back into `formula::eval::serial_to_ymd_pub` — an
in-crate cycle invisible today only because both live in one crate.

**Resolution:** introduce `src/types.rs` (`pub mod types;` in `lib.rs`) and move into it,
verbatim where possible:

From `src/vm/mod.rs`:
- `ExcelError` + its `as_str`/`Display` impls (`vm/mod.rs:10-38`)
- `Variant` + its `Display` impl (`vm/mod.rs:202-238`)
- `CellContent` (`vm/mod.rs:241-244`)
- `serial_to_display` (`vm/mod.rs:214-218`)
- `parse_cell_addr`, `parse_range_addr`, and their private helper
  `col_letters_to_num_vm` (`vm/mod.rs:2413-2434`) — pure string↔coordinate parsing,
  confirmed to have zero further `crate::` dependencies of their own

From `src/formula/eval.rs`:
- `serial_to_ymd` / `serial_to_ymd_pub` (pure date-serial math, confirmed independent) —
  collapsed into one canonical `types::serial_to_ymd`, dropping the `_pub` shim

`src/vm/mod.rs` then re-exports what it used to define:
`pub use crate::types::{Variant, ExcelError, CellContent};` — every existing
`vm::Variant` / `crate::vm::ExcelError` call site across the codebase (hundreds of
references) keeps compiling unchanged.

`src/formula/eval.rs` retargets **all four** of its `crate::vm::*` references (not just
the `use` line — this is the correction from the initial pass, which undercounted the
extraction):

| Site | Today | After |
|---|---|---|
| `eval.rs:5` | `use crate::vm::{CellContent, ExcelError, Variant};` | `use crate::types::{CellContent, ExcelError, Variant};` |
| `eval.rs:2888` | `crate::vm::parse_range_addr(...)` | `crate::types::parse_range_addr(...)` |
| `eval.rs:3970` (test module) | `use crate::vm::ExcelError;` | `use crate::types::ExcelError;` |
| `eval.rs:4358` (test module) | `use crate::vm::ExcelError;` | `use crate::types::ExcelError;` |

Resulting graph: `types` (leaf) ← `formula` ← `vm`. The pre-existing `vm → formula` edge
(used for `set_cell_formula`/`recalculate_all`) is untouched; only `formula → vm` is
redirected to `formula → types`.

**This is a Phase 1+ task, described here for planning purposes, not executed in
Phase 0.** Whoever executes it must re-grep immediately before starting — `vm/mod.rs` is
a large file under active milestone-driven development, and new impls may have landed
since this ADR was written — and must confirm `cargo test` still shows all tests passing
afterward (642/642 at time of writing), since every listed change is a move or a
`use`-retarget, not a semantic change.

### reader.rs buffer-API resolution

`zip::ZipArchive<R>` is already generic over `R: Read + Seek` — elixcee simply never
instantiates it with anything but `std::fs::File` today. No new dependency is needed to
close this gap.

**Resolution (Phase 1+):** extract the body of `read_workbook` that operates on an
already-opened `ZipArchive<File>` into a private generic function:

```rust
fn read_workbook_from_archive<R: Read + Seek>(archive: ZipArchive<R>) -> Result<Vec<WorkbookSheet>, String>
```

Add a new public entry point:

```rust
pub fn read_workbook_from_bytes(bytes: &[u8]) -> Result<Vec<WorkbookSheet>, String>
```

wrapping `ZipArchive::new(Cursor::new(bytes))` and calling the same generic function.
`read_workbook(path: &str)` becomes a two-line wrapper: open the file, delegate. No
existing signature changes, no call-site breakage — this is purely additive plus one
internal extraction.

### Dependency-direction rule

Keep the hand-rolled `XmlIter`; harden it with the limits designed in
[`docs/xlsx-security-model.md`](xlsx-security-model.md) instead of adopting a DOM/XML
crate. This matches elixcee's own established posture (see Context) rather than
introducing a new dependency for a problem that's solvable by adding bounds to existing
code. No `wasm-bindgen` dependency is added until the phase that actually builds
`elixcee-wasm`. Any future dependency addition proposed for this initiative should update
this ADR with its justification, the same way the project's own CHANGELOG documents past
dependency removals.

### Cargo.toml

**No change in Phase 0.** Converting `Cargo.toml` to a workspace now would require moving
`src/` under `crates/elixcee/` (or similar) and updating `pyproject.toml`'s maturin
config, `.github/workflows/*`, and `fuzz/Cargo.toml`'s relative paths — a structural
change for zero immediate benefit, since Phase 0 adds no Rust code at all. Deferred to
whichever phase first creates `crates/elixcee-types`.

### Non-negotiable: `packages/xlsx` never depends on `xlsx` at runtime

An earlier draft of this ADR proposed a Phase 1 "thin passthrough" that re-exported the
real `xlsx` package 1:1 as a temporary implementation. **Rejected on user review** — the
product package must never load the real `xlsx@0.18.5` at runtime, because that would
make every differential test trivially `MATCH` (comparing the oracle against itself),
silently inherit the oracle's own known vulnerabilities (see
[`docs/xlsx-security-model.md`](xlsx-security-model.md)'s "the oracle itself is a
validating example" — `xlsx@0.18.5` carries two open high-severity advisories), give no
signal on real implementation progress, and directly contradict the "safer parser"
product claim if ever published in that state.

The real `xlsx` package is confined to exactly one place in this repository:
`compat/oracle/` (and, transitively, any `compat/differential/` test file that imports
it for comparison). `packages/xlsx/package.json` must never list `xlsx` under
`dependencies`, `peerDependencies`, or `optionalDependencies` — not even as a
`devDependency`, so that nothing under `packages/xlsx/node_modules` ever resolves it
either. Differential tests that need both the oracle and the elixcee implementation live
under `compat/`, importing the elixcee side via a relative path into `packages/xlsx/src`
— no npm linking/workspace machinery required for this to work.

## Phase 1A plan

Phase 1A's goal is a real `packages/xlsx` package skeleton plus a first slice of
genuinely-reimplemented, oracle-differential-tested pure utility functions — never a
passthrough. No Cargo workspace split, no `formula`/`vba` extraction, and no WASM in this
phase; elixcee's existing Rust code is untouched.

- `packages/xlsx/package.json`: name `@elixcee/xlsx` (placeholder — see the npm-scope open
  item), `private: true` (blocks accidental publish), `type: "commonjs"`, CJS entry
  (`main`) + ESM entry (`module`, and both wired through the `exports` map) + TypeScript
  declarations (`types`). No `dependencies` field referencing `xlsx` (see above).
- First API slice (pure functions, no file I/O, no ZIP/XML): `encode_col`, `decode_col`,
  `encode_row`, `decode_row`, `encode_cell`, `decode_cell`, `encode_range`, `decode_range`,
  `safe_decode_range`, `split_cell`, `book_new`, `book_append_sheet`,
  `book_set_sheet_visibility`, `aoa_to_sheet`.
- Each function is differential-tested against the real oracle (`compat/oracle`) across a
  boundary-value matrix (`A1`, `Z1`, `AA1`, `XFD1048576`, lowercase input, `$`-absolute
  references, malformed references, negative/zero/huge row-col numbers, reversed ranges,
  sheet-qualified references, sheet names with special characters, duplicate sheet names,
  hidden/very-hidden sheet visibility, `dense` option, `null`/`undefined`, sparse arrays)
  using [`compat/differential/classify.mjs`](../compat/differential/classify.mjs)'s
  registry-gated verdicts — never a hand-picked "looks close enough" sample.
- A TypeScript smoke test compiles a small consumer snippet against
  `packages/xlsx/src/index.d.ts`.
- Explicitly out of scope for Phase 1A: WASM build, any Rust↔JS bridge, any file-format
  I/O (`read`/`readFile`/`write`), the `elixcee-types`/`elixcee-xlsx` crate split. Those
  begin once this utility slice is solid and the phase that tackles `read()` starts.

## SSF backend (Phase 1B-2B)

`format_cell`/`sheet_to_csv`/`sheet_to_txt` need the oracle's number-format ("Format
Codes") engine — `SSF_format`/`eval_fmt` in the bundled `xlsx.js`, a ~900-line
format-string parser/evaluator, equivalent to the standalone `ssf` npm package (one of
the 7 Apache-2.0 SheetJS dependencies). Three options were compared before choosing:

1. **Take `ssf@0.11.2` (exact-pinned) as a real `packages/xlsx` runtime dependency.**
2. **Clean-room reimplementation** of the format engine, from reading the oracle source.
3. **An independent, Rust-portable reimplementation** (same as 2, but structured for an
   eventual `elixcee-ssf`-equivalent Rust/WASM port) — compatibility-equivalent to option
   2, differing only in internal structure.

**Decision: option 1**, overriding this document's earlier draft recommendation of
option 3. Rationale (user decision, recorded verbatim in intent): getting xlsx@0.18.5
users onto `@elixcee/xlsx` with full compatibility now outweighs the near-term value of a
from-scratch reimplementation that adds no user-visible capability over what the real
`ssf` package already provides correctly; a full clean-room port is deferred to
elixcee's eventual Rust-native phase, where `ssf-adapter.cjs` below becomes the single
swap point. This is **not** a reversal of the zero-dependency philosophy — it is a
deliberate, disclosed, transitional trade: borrow a correct implementation to reach
compatibility sooner, then replace it once the Rust core can absorb the work.

Before committing to option 1, its correctness was measured, not assumed: `ssf@0.11.2`'s
`.format()` was compared against the bundled engine (`XLSX.SSF.format`, a *different*
export surface than the standalone package but confirmed to run the same algorithm)
across 1800+ cases in
[`compat/differential/ssf-format.test.mjs`](../compat/differential/ssf-format.test.mjs)
— every `table_fmt` built-in numFmtId and its `SSF_default_map` indirection, crossed with
boundary/date-serial/text/boolean values, plus `date1904`, a custom format table,
multi-section formats, `[Red]`/conditional sections, fractions, exponential notation, and
percent/thousands. This surfaced one genuine defect: `ssf@0.11.2`'s own indirection table
has a copy-paste bug affecting numFmtIds 67-71 (confirmed by reading
`node_modules/ssf/ssf.js:93-105` — a loop meant to set entries for 69-71 instead reuses
the preceding block's loop bounds). Corrected in
[`packages/xlsx/src/internal/ssf-adapter.cjs`](../packages/xlsx/src/internal/ssf-adapter.cjs)
with a 5-entry pre-correction table, not a fork of `ssf` itself — a narrow, disclosed
patch for one precisely-diagnosed defect, not a reimplementation.

**Isolation architecture** — `ssf` is `require`d in exactly one file:

```
format_cell / sheet_to_csv / sheet_to_txt
    -> src/internal/number-format.cjs   (cell-level orchestration: caching, BErr lookup,
                                          the two-try fallthrough — ported from the oracle)
    -> src/internal/ssf-adapter.cjs     (the ONLY file that requires "ssf"; also where the
                                          numFmtId 67-71 correction lives)
    -> ssf@0.11.2
```

No other file reaches into `ssf` directly. Swapping the backend later — e.g. for a
Rust/WASM-based formatter once elixcee's Rust core grows one — is meant to be a
single-file change (`ssf-adapter.cjs`'s `format(fmt, v, opts)` contract stays fixed); the
1800+-case matrix that validated this backend becomes the acceptance suite for that swap
too.

See [`docs/licensing.md`](licensing.md) for the licensing consequence (an ordinary npm
`dependencies` declaration, not vendored/bundled source — but still `@elixcee/xlsx`'s
first real runtime dependency) and
[`THIRD_PARTY_NOTICES.md`](../THIRD_PARTY_NOTICES.md) for the resulting notice.

`sheet_to_csv`/`sheet_to_txt`/`sheet_to_formulae` also share a
[`range-guard.cjs`](../packages/xlsx/src/internal/range-guard.cjs) helper: all three walk
a worksheet's entire `!ref` rectangle regardless of sparsity, and a crafted full-grid
`!ref` was confirmed live to not return within 25s on the real oracle. See
[`docs/xlsx-security-model.md`](xlsx-security-model.md) for this as a registered
`ELIXCEE_RANGE_TOO_LARGE` safety divergence.

## Findings from generating the Phase 0 manifest

Running `compat/oracle/generate-manifest.mjs` against the real `xlsx@0.18.5` surfaced a
concrete CJS/ESM asymmetry that no documentation review caught: the ESM entrypoint
(`xlsx/xlsx.mjs`) exposes `set_cptable` and `set_fs` as top-level exports, while the CJS
entrypoint (`xlsx.js`, Node's `main`) does not. Any future `@elixcee/xlsx` package must
replicate this asymmetry rather than exposing a uniform surface across both module
systems — this is exactly the kind of fact
[`docs/xlsx-compatibility-goal.md`](xlsx-compatibility-goal.md)'s "no guessing, differential-test
only" rule exists to catch. See `compat/oracle/api-manifest.json`'s `entrypoints.cjs` vs.
`entrypoints.esm` for the full comparison.

## Phase 2B-0: sync WASM/`read()` bridge feasibility

The fork question, stated exactly as given: can
`const XLSX = require("@elixcee/xlsx"); const wb = XLSX.read(bytes);` be genuinely
synchronous — no consumer-facing `await init()` — across CJS, ESM, and browser/bundler
consumption? If yes, `elixcee-wasm` proceeds on wasm-bindgen. If no, the fallback (already
pre-authorized) is to split into a Node-native binding, browser-WASM, and JS-fallback
track instead.

This was investigated with a throwaway spike crate built **outside this repository**
(scratchpad, own `Cargo.toml`, `wasm-bindgen = "0.2"`, one trivial exported function —
`probe(bytes: &[u8]) -> usize`). Nothing here was added to the repo: no `wasm-bindgen`
dependency, no `crates/elixcee-wasm`, no `Cargo.lock` change. Every finding below was
produced by actually running the built artifact (real `wasm-pack` builds, a real `require`,
a real Chrome tab), not inferred from documentation, per this project's existing
確認済み事実-vs-推測 discipline.

**Node — CJS: sync, verified by execution.** `wasm-pack build --target nodejs` emits glue
that does `require('fs').readFileSync(...)` + `new WebAssembly.Module(bytes)` +
`new WebAssembly.Instance(module, imports)` at module-load time — synchronous by
construction, no `init()` export exists at all in this target. Confirmed by running
`require("./pkg-nodejs/wasm_poc.js")` and calling the exported function on the same line,
no `await`, no setup call.

**Node — ESM: sync, verified by execution, two independent paths.**
1. `import XLSX from "<nodejs-target-build>.js"` — Node's ESM-importing-CJS interop loads
   the same synchronous glue above; confirmed working, no warning.
2. A plain `import * as wasm from "./file.wasm"` (native ESM WASM-module import, no
   wasm-bindgen glue involved) also succeeded outright on Node v24.5.0, with only a
   printed `ExperimentalWarning: Importing WebAssembly module instances is an experimental
   feature`. Notable, but not something to depend on for a package targeting broad Node
   version support today — path 1 (ship `--target nodejs`-style glue, let ESM consumers
   hit it through interop) is the one to build on.

**Browser main-thread sync compile: clean up to 5MB in Chrome 151, untested beyond that.**
`wasm-bindgen` 0.2.127 ships a real `initSync(module)` export alongside the default async
`init()` — feeding it an already-in-hand `WebAssembly.Module`/raw bytes compiles and
instantiates synchronously. The open question was whether the *engine* enforces a
synchronous-compile size ceiling on the main thread. Tested empirically in a real Chrome
tab (v151, via `claude-in-chrome`, not headless-only): base64-inlined bytes decoded
synchronously via `atob`, then `new WebAssembly.Module(bytes)` + `new
WebAssembly.Instance(...)`, at five sizes (the ~12.7 KB trivial-crate floor, padded via a
harmless trailing custom section to 50 KB / 500 KB / 2 MB / 5 MB). **All five compiled and
instantiated synchronously with no error and no console warning**, 0.2–3.9 ms each. Caveat,
stated plainly: the padding is inert bytes in a custom section, not real code — this
measured the *loader's* size ceiling, not a realistic parser's compile time, and a real
`elixcee-xlsx`-sized module's actual synchronous compile latency (likely tens of ms, not
low-single-digit ms) was not measured here.

MDN's `WebAssembly.Module` documentation does note a caveat — quoted exactly: "Some
browsers may throw a `RangeError`, as they prohibit compilation and instantiation of Wasm
with large buffers on the UI thread" — but names no specific browser and gives no size
figure. A commonly-repeated claim that Safari/WebKit specifically enforces a ~4 KB
synchronous-compile ceiling could **not** be substantiated from current MDN pages fetched
this session (the `Module` reference page and the "Loading and running WebAssembly code"
guide were both checked directly; neither states a number or names Safari). **This is
reported as unverified, not as fact** — no Safari/WebKit environment was available in this
session to test directly. The exact harness used for the Chrome test
(`sync-test.html`) still exists in the scratchpad and can be pointed at real Safari to
settle this in one step if it matters before committing to the browser leg's design.

**Bundler WASM resolution: not automatic, and that's fine — own the byte-loading instead
of depending on it.** esbuild (installed fresh, isolated to the spike directory, default
settings) fails outright on `wasm-pack`'s `--target bundler` output: `No loader is
configured for ".wasm" files`. This is a real, immediate failure, not an async-related one
— it means "bundler WASM resolution works out of the box" is **false** as a blanket claim;
a consumer's bundler needs either explicit loader configuration or a plugin before it will
touch a raw `.wasm` import. The actionable conclusion is not to depend on generic bundler
support at all: `@elixcee/xlsx`'s own shipped glue should inline the compiled `.wasm`
bytes directly (base64 or an equivalent encoding, decoded at module-evaluation time) rather
than emitting a bare `import "*.wasm"` and hoping the consumer's bundler resolves it — the
same technique already proven to work synchronously in the Chrome test above. This sidesteps
the bundler-support question entirely rather than solving it.

**CSP / no-`wasm-unsafe-eval`: clean, catchable failure.** With a page CSP of
`script-src 'self' 'unsafe-inline'` (deliberately omitting `'wasm-unsafe-eval'`), the same
synchronous `new WebAssembly.Module(bytes)` call throws a `CompileError` whose message
names the CSP directive explicitly (`"...violates the following Content Security Policy
directive because 'unsafe-eval' is not an allowed source..."`). This is a normal,
catchable JS exception with an identifiable message — a `try { initSync(...) } catch`
around the browser entry point can detect this specific case and surface a clear,
documented error rather than an opaque crash.

**Verdict: proceed on wasm-bindgen, per the pre-agreed fork criterion, with two concrete
design commitments carried into whichever phase builds `elixcee-wasm`:**
1. Ship `--target nodejs`-style glue for the Node entry point (both `require` and ESM
   consumers reach it, per the two verified paths above) — no async story needed there at
   all.
2. For the browser entry point, inline the WASM bytes into the shipped glue and call
   `initSync` ourselves, rather than emitting a bundler-dependent `import "*.wasm"` or
   relying on the default `fetch`-based async `init()`. Wrap the synchronous compile call
   so a CSP-driven `CompileError` (and, if the unverified Safari ceiling turns out to be
   real, a `RangeError`) is caught and produces a clear, documented failure rather than an
   uncaught exception — this is the JS-fallback branch, scoped down from "a whole separate
   track" to "one catch clause," if the Safari question resolves against us.

No code from this spike is being merged; the crate lived entirely outside the repository
and is deletable. This section, plus the two design commitments above, is Phase 2B-0's
complete output — implementing `elixcee-wasm` itself is a separate, not-yet-started phase.

(A stale plan file from an earlier, already-superseded planning session — describing what
this document's own Phase 0/1A sections above cover, now fully executed — surfaced in this
session's context. It was recognized as stale and not acted on.)

## Consequences

Once the formula/VM split and the buffer-API extraction land, `elixcee-xlsx` can exist as
a crate with no dependency on `elixcee-vba` — the precondition for `elixcee-wasm` to ship
a small WASM binary that doesn't drag in the VBA interpreter. Until then, the existing
single-crate `elixcee` remains the only build target, and this initiative's Phase 0
output is documentation plus an npm-side (Node-only) oracle/differential harness that
touches no Rust code.

**Update: this precondition was bypassed, not fulfilled.** `crates/elixcee-wasm` shipped
(see "Status" above) depending directly on the monolithic `elixcee` crate — the
formula/VM split described above never happened, so its WASM binary does drag in the full
VBA interpreter rather than a slimmed-down `elixcee-xlsx`. Recorded here as a disclosed
fact, not evaluated as good or bad — no measurement of the resulting WASM binary's size
attributable to this specific choice was made as part of this correction; `ROADMAP.md`'s
npm/JS/WASM findings section has the measured package-size numbers that do exist.
