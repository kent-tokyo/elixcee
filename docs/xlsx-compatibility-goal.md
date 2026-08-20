# XLSX compatibility goal

## Purpose

This document anchors every later phase's answer to "are we compatible yet." It defines
what compatibility means for the `@elixcee/xlsx` initiative — a planned npm package meant
to be a drop-in replacement for [`xlsx@0.18.5`](https://www.npmjs.com/package/xlsx)
(SheetJS), reachable via `"xlsx": "npm:@elixcee/xlsx@^1.0.0"` with no application code
changes.

## Definition of compatibility

Compatibility has two parts, and they are evaluated separately:

1. **Normal, well-formed input** → `@elixcee/xlsx` must produce output that is logically
   equivalent to `xlsx@0.18.5`'s output for the same input and options: same public API
   surface, same object shapes and key enumerability, same array/sheet ordering, same
   error types and messages where SheetJS itself errors on valid-but-unusual input, same
   cell/date/formula/error-value semantics.
2. **Malicious input, or input that exceeds a documented resource limit** →
   `@elixcee/xlsx` must return a safe, deterministic error. It must never replicate a
   vulnerability just because the oracle (`xlsx@0.18.5`) exhibits one. See
   [`docs/xlsx-security-model.md`](xlsx-security-model.md) for the specific limits and
   the policy this implies.

"Roughly the same" is not an acceptable compatibility judgment anywhere in this
initiative. Every observed divergence between `@elixcee/xlsx` and the oracle must be
explicitly classified — see [`compat/differential/classify.mjs`](../compat/differential/classify.mjs).

## Non-goals

- **VBA-execution compatibility** is a separate, already-existing, unaffected track.
  Nothing in this initiative changes how elixcee emulates VBA macros.
- **Performance parity** with `xlsx@0.18.5` is not a compatibility requirement.
  Compatibility comes first; if `@elixcee/xlsx` is slower, that gets recorded, not traded
  away for speed.
- **Full browser-bundle parity** (matching `dist/xlsx.full.min.js` exactly) remains
  deferred — not because a WASM build target doesn't exist (`elixcee-wasm` now does, with
  real browser smoke tests wired into CI — see `docs/xlsx-architecture.md`'s "Status"),
  but because `write`/`writeFile`/`writeFileSync` are still unimplemented (see
  `ROADMAP.md`'s "Known gaps").

## How compatibility is measured

- [`compat/oracle/`](../compat/oracle/) holds the machine-generated record of what
  `xlsx@0.18.5` actually exposes at runtime (not hand-transcribed from documentation).
- [`compat/differential/`](../compat/differential/) holds the harness that runs the same
  input through both the oracle and `@elixcee/xlsx`, normalizes the results, and
  classifies any divergence.
- [`docs/compatibility-known-defects.md`](compatibility-known-defects.md) records oracle
  behaviors that look like bugs but are deliberately reproduced for compatibility anyway.

## Status

**Stale — kept as the original Phase 0 framing, not the current state.** This document
described Phase 0 (investigation and scaffolding only, no compatibility logic
implemented) at the time it was written. Substantial `@elixcee/xlsx` compatibility work
has since shipped — see `ROADMAP.md`'s "Current state" for what's actually implemented
(33 `utils.*` exports differential-tested, working `XLSX.read()`/`readFile()`/
`readFileSync()`, CI-wired differential suites) and `docs/xlsx-architecture.md`'s "Status"
for how the crate layout actually turned out. The compatibility *definition* and
non-goals above remain the operative policy; only this specific status paragraph was
Phase-0-only and is now outdated.
