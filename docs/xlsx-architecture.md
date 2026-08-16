# XLSX architecture ADR

## Status

Accepted for planning purposes. Describes the **target** shape for the `@elixcee/xlsx`
initiative and the concrete resolutions for its two hardest structural problems. Phase 0
implements none of the Rust-side changes described here — it only creates the
documentation and the `compat/oracle` + `compat/differential` Node harness. Everything
under "Decision" below is scoped work for later phases.

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

## Consequences

Once the formula/VM split and the buffer-API extraction land, `elixcee-xlsx` can exist as
a crate with no dependency on `elixcee-vba` — the precondition for `elixcee-wasm` to ship
a small WASM binary that doesn't drag in the VBA interpreter. Until then, the existing
single-crate `elixcee` remains the only build target, and this initiative's Phase 0
output is documentation plus an npm-side (Node-only) oracle/differential harness that
touches no Rust code.
