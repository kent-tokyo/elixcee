# Date/Time runtime model ADR

## Status

**Proposed — design only, not implemented.** This document compares three options and
recommends one; it does not change any code. Implementing the recommendation is future
work, scoped as an `elixcee-types` 0.2.0 / `elixcee` 0.4.0-shaped change (see
"Consequences"), not a patch release.

## Context

`Variant::Date(i64)` (`crates/elixcee-types/src/lib.rs`) represents a whole-day Excel
serial number — no sub-day (time-of-day) component. This is a real, structural
limitation, not a bug in the usual sense: `Date()` fits it perfectly (a calendar date has
no time component to lose), but `Time()`/`Now()` genuinely need a fractional day and
currently return `Variant::Float` instead, so `TypeName(Time())`/`TypeName(Now())` report
`"Double"` where real VBA reports `"Date"` (tracked in `ROADMAP.md`'s "Known gaps" #5,
found and disclosed rather than fixed in the round that produced this document).

Fixing this requires deciding how a date-with-time value should be represented, and that
decision has real reach: `Variant` is a public type in the independently-published
`elixcee-types` crate (already live on crates.io at 0.1.0), consumed by the main
`elixcee` crate, the Python bindings (`src/lib.rs`'s `variant_to_py`), the CLI's `--json`
cell serialization, the formula engine, and the WASM bridge
(`crates/elixcee-wasm`) — a change here is not locally contained.

### Facts this comparison is grounded in (verified directly, not assumed)

- `crate::types` in the main `elixcee` crate is a re-export of the external
  `elixcee_types` crate (`src/lib.rs:16`, `pub use elixcee_types as types;`) — `Variant`
  and its date-serial math live in that separately-versioned, already-published crate,
  not locally.
- `serial_to_ymd` (`crates/elixcee-types/src/lib.rs:97`) already correctly reproduces
  Excel's own serial-60 leap-year bug (serials 60 and 61 both decode to 1900-03-01,
  matching Excel's fictitious "1900-02-29" quirk) — confirmed via that function's own
  existing, passing tests. This is solved and reusable regardless of which option below
  is chosen; none of them need to re-derive this logic.
- `date1904` handling is entirely confined to `src/reader.rs` (XLSX-file-reading only,
  threaded into `BufferWorkbook`) — it has no connection today to `Variant::Date` or the
  VBA execution engine's own date math (`Date()`/`Now()`/formula `DATE()` always assume
  the standard 1900 epoch regardless of any loaded workbook's `date1904` setting). This
  is an existing, pre-existing scope boundary unaffected by any option below.
- `Variant::Date` currently appears in ~25 match sites in `src/formula/eval.rs` and ~11 in
  `src/vm/mod.rs` (direct grep count), plus the Python conversion (`src/lib.rs:68`, converts
  to `datetime.date` via `PyDate::new`) and the `--json` cell-value path (`serial_to_display`,
  a `"YYYY-MM-DD"` formatted string, not the raw serial number — confirmed empirically by
  `compat/vba-semantics/`, whose own first draft assumed the latter and had to be corrected).
- Arithmetic/comparison already routes `Variant::Date` through the same generic `to_f64`
  numeric-coercion path every other numeric type uses (`Variant::Date(s) => Ok(*s as f64)`,
  `src/vm/mod.rs:3357`) — this is centralized, not scattered, so it's low-risk and
  low-effort under any of the three options.

## Options considered

### A — Change `Variant::Date(i64)` to `Variant::DateSerial(f64)`

Replace the existing variant's payload type (and, most naturally, its name) so the same
one variant carries both whole-day and time-of-day values.

- **Rust public API**: Breaking. Every one of the ~36+ existing match sites needs
  updating; any external consumer with an exhaustive match on `Variant` fails to compile.
- **`elixcee-types` semver**: A real breaking change to a type already live on crates.io.
- **Python**: `variant_to_py`'s single `Date` branch must now decide, for *every* value —
  including all the ones that used to be whole-day-only — whether to produce `PyDate` or
  `PyDateTime`. Either choice changes observable behavior for existing Date-producing
  code (formula `DATE()`, etc.), not just newly-added code.
- **JSON**: Same problem — the existing `"YYYY-MM-DD"` cell-value format for `Date()` etc.
  either loses the (now-always-present) fractional part or the format itself changes,
  for values that used to render identically.
- **Formula engine**: ~25 internal call sites need updating; internal-only, so lower
  external risk than the Rust API point above, but real code churn for its own sake.
- **XLSX serial / date1904 / serial-60**: No fundamental incompatibility — Excel's own
  serial representation already conflates date+time into one float, so `f64` is a
  natural fit; `date1904` is unaffected (out of scope per the facts above); `serial_to_ymd`
  needs only a whole/fractional split before reuse, not new logic.
- **Arithmetic/comparison**: Marginal simplification (f64→f64 instead of i64→f64 cast) —
  the only point in A's favor.
- **WASM payload**: Negligible either way (an enum tag, not a size concern).
- **Backwards compatibility**: Worst of the three. Rewrites a real amount of working code
  and changes observable behavior for every existing Date-typed value, to fix a gap that
  only affects two functions (`Time()`, `Now()`).

### B — Keep `Variant::Date(i64)`, add `Variant::DateTime(f64)` alongside it

- **Rust public API**: Additive. Every existing `Variant::Date(i64)` match arm keeps
  working unchanged; the only required change is adding a new arm to each of the small
  number of *exhaustive* matches (`check.rs`'s two, per this project's own established
  "no wildcard, so a new variant forces a deliberate decision" discipline already applied
  to `Stmt` — see `CHANGELOG.md`'s `Stmt::DimBare`/`Stmt::DimMulti` history for the exact
  same pattern already exercised twice this round). A well-understood, bounded ripple, not
  a rewrite.
- **`elixcee-types` semver**: Still technically breaking by Rust's own rule (adding a
  variant to a public non-`#[non_exhaustive]` enum breaks an external exhaustive match) —
  so this is also a real `0.2.0`-shaped bump, not a patch, but with far less actual code
  churn than A.
- **Python**: `Variant::DateTime` gets its own new conversion branch → `PyDateTime`,
  cleanly separate from the existing `Date` → `PyDate` branch. Zero behavior change for
  any currently-Date-producing code path; only the new `Now()`/`Time()` values (if
  migrated to `DateTime`) gain a new, correctly-typed Python representation.
- **JSON**: `DateTime` gets its own serialization (an ISO-8601-with-time-shaped string),
  independent of `Date`'s existing `"YYYY-MM-DD"`-only format. No existing JSON output
  shape changes for any value that's `Variant::Date` today.
- **Formula engine**: Only the few new call sites that would construct `DateTime` (if
  `Now()`/`Time()` are migrated) need touching; every existing `DATE()`/`DATEVALUE()`/etc.
  call site keeps using `Variant::Date`, unchanged.
- **XLSX serial / date1904 / serial-60**: Same as A for the new variant's own math (natural
  `f64` fit, `date1904` unaffected) — but `serial_to_ymd`'s existing logic is reused by
  *composition* (split the new type's whole/fractional parts, call the untouched existing
  function on the whole-day part) rather than needing to be generalized in place.
- **Arithmetic/comparison**: Needs new match arms for the new variant (a few lines,
  mirroring `Date`'s existing ones) — doesn't touch how `Date` arithmetic already works.
- **WASM payload**: Negligible, same as A.
- **Backwards compatibility**: Best of the three for *observable behavior* — zero runtime
  behavior change for any existing `Date`-typed value or code path. The break is purely
  at the Rust type-system level (a new match arm required to compile against the new
  version), never a silent behavior change for a current Python/JSON/CLI consumer.

### C — A date+time representation that never becomes a public `Variant` variant at all

Considered because it promises zero blast radius to the public `elixcee-types` crate.

- **Does not actually solve the stated problem.** `Now()`/`Time()`'s return value is
  assigned to a VBA variable and stored in `Vm.variables: HashMap<String, Variant>` — it
  fundamentally must *be* a `Variant` to fit the existing execution model. Any
  "internal-only" richer representation that never surfaces as a `Variant` cannot be what
  `x = Now()` stores, and therefore cannot be what `TypeName(x)` inspects — the exact
  thing this whole exercise is about fixing. Making C viable at all would require a
  second, parallel value representation threaded through every place a VBA value can
  flow (variables, cells, comparisons, arithmetic) *alongside* `Variant`, which is a
  **larger**, not smaller, architectural change than A or B, for the same end result.
- Reported here for completeness, per the instruction that asked for a comparison of all
  three, rather than silently dropped — but it is not a viable candidate as stated, and
  every criterion above is moot for it.

## Recommendation

**Option B** — additive `Variant::DateTime(f64)`, `Variant::Date(i64)` untouched.

It solves the actual problem (a `Now()`/`Time()` value that reports `TypeName` `"Date"`,
via a variant real VBA would recognize as a date-family type) with materially less
backwards-compatibility risk than A, for the same `elixcee-types` semver cost (both are
`0.2.0`-shaped, not patches — B doesn't avoid the version bump, it avoids the *code
churn and behavior-change risk* that comes with it). It reuses the already-tested
`serial_to_ymd` leap-year logic by composition rather than needing to touch it. And it
matches this project's own repeatedly-applied decision ordering (see `ROADMAP.md`'s
"Decision Policy": safety, then reversibility, then backward compatibility, then
consistency with existing design, then small diff, ahead of novelty) — A is a strictly
worse version of B for this specific problem: the same version-bump cost, with
unnecessary blast radius and behavior-change risk for no additional benefit over B. C does
not solve the problem as stated and isn't a real alternative.

## Consequences (if B is later approved and implemented — not committed here)

- A real `elixcee-types` 0.2.0 / `elixcee` 0.4.0-shaped release, not a patch — the same
  judgment-call category as the crates.io-publish decisions already made explicitly by
  the maintainer earlier in this project's history, not something to bump silently.
- `Now()`/`Time()` would migrate from `Variant::Float` to `Variant::DateTime(f64)` — a
  real, disclosed behavior change to their own `TypeName`/`VarType` results (currently
  `"Double"`/`5`, would become `"Date"`/`7`, matching real VBA) — `compat/vba-semantics/`'s
  existing cases for these would need updating (their `knownLimitation`-registered
  divergence would become a genuine fix, not a disclosed gap, at that point).
- `crates/elixcee-types/src/lib.rs`'s `impl Display for Variant` (its only exhaustive
  match on `Variant` itself, confirmed by grep) would need a new arm — the same
  "compiler forces a deliberate decision per new variant" pattern already exercised twice
  this round for `Stmt::DimBare`/`Stmt::DimMulti` in `check.rs`'s exhaustive `Stmt` matches
  (a different type, same discipline). A full site-by-site audit of every place that would
  need a `DateTime` arm (vs. places where `Date`-only handling is deliberately correct to
  leave alone) is implementation-time work, not attempted here.
- Whether `DateTime` participates in every place `Date` currently does (JSON, CLI output,
  formula engine date functions, WASM bridge) needs its own implementation-time audit —
  this document establishes the *shape* of the change, not its full site-by-site plan.
