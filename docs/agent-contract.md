# elixcee CLI Agent Contract

This document is the fixed, machine-readable contract for the `elixcee` CLI's
`--json` mode. It exists so AI agents and CI scripts can rely on it without
re-deriving behavior from the source. Anything not listed here (exact field
set beyond what's documented, human-readable message wording) may change
without notice; anything listed here is a compatibility promise.

## Invocation

```
elixcee <vba_file>... <MacroName> [--file <path>] [--sheet <name>] [--output <path>] [--json]
```

`<MacroName>` is always the last argument; everything before it is a VBA
source file. A single-file invocation behaves exactly as before. With more
than one source file, see "Multi-module projects" below for how
`MacroName` and cross-module calls resolve.

Without `--json`, output is unchanged from elixcee's original plain-text
behavior (kept exactly, byte-for-byte, for scripts that already depend on
it) — see "Non-JSON mode" below.

## stdout / stderr contract (`--json` mode)

- **stdout**: exactly one line containing one JSON object — either the
  success shape or the error shape below. Nothing else is ever written to
  stdout in `--json` mode: `MsgBox` text is captured into the `messages`
  array instead of being printed, and errors that happen after partial work
  (e.g. a failing `--output` write after a successful macro run) still
  resolve to exactly one JSON object, never two.
- **stderr**: not part of the contract; treat anything on stderr as
  incidental/debug noise, not a signal to parse. In practice elixcee writes
  nothing to stderr in `--json` mode.
- Never rely on line count, coloring, or any output ordering beyond "one
  JSON object on stdout."

## Exit codes

| Code | Meaning |
|---|---|
| `0` | Success (`"ok": true`) |
| `1` | Any failure (`"ok": false` — see `error.code` for the specific category) |

The exit code is intentionally coarse. The JSON body's `"ok"` field and
`error.code` are the primary machine-readable signal for *why* something
failed; a wider per-category exit code taxonomy is not part of this
contract; a typed error enum may replace this interim implementation in a future
release.

## Success shape

```json
{
  "schema_version": 1,
  "ok": true,
  "entrypoint": "MacroName",
  "duration_ms": 1.82,
  "cells": [
    {"sheet": "Result", "address": "B2", "value": 125000}
  ],
  "messages": []
}
```

- `schema_version`: bumped only on a breaking change to this shape. Check
  it before assuming field meanings.
- `entrypoint`: the macro name passed on the command line.
- `duration_ms`: wall-clock time of the macro execution only (not file I/O).
- `cells`: every non-empty cell in the sheet that was active when the macro
  finished, sorted by (row, column) — deterministic across runs for the
  same input. `value` is a JSON number, boolean, string, or `null` (empty);
  VBA arrays/records currently serialize as placeholder strings
  (`"[array]"` / `"[record]"`), matching the pre-`--json` plain-text CLI.
- `messages`: every `MsgBox` the macro invoked, in the order they were
  shown (see "messages semantics" below).

## Error shape

```json
{
  "schema_version": 1,
  "ok": false,
  "error": {
    "code": "E1001",
    "kind": "undefined_variable",
    "message": "Undefined variable: 'totla'",
    "location": {"file": "Main.bas", "line": 2, "column": 5}
  },
  "messages": []
}
```

`message` is free text for a human/log — don't pattern-match on it. `code`
and `kind` are the stable, matchable fields. `messages` is present here too
(see "messages semantics" below) — a macro that shows progress via `MsgBox`
before hitting a runtime error still surfaces that text, it isn't dropped
just because the run ultimately failed. It's `[]` for failures that happen
before the macro starts running (parse/io/setup errors).

`location` is `{file, line, column}` (1-based) for parse errors and runtime
errors, or `null` for failures that happen before/outside macro execution
(io errors, `--sheet` setup errors, or a runtime failure that occurs before
any statement ever runs, like a missing entrypoint macro name). **Runtime
error locations are statement-level, not sub-expression-level** — `line`/
`column` point at the start of the statement that was executing, not the
exact sub-expression within it (e.g. for `x = totla + 1` failing on the
undefined `totla`, `location` points at the `x` that starts the statement,
not at `totla`). Parse error locations point at the specific token the
parser choked on. There's no "did you mean" suggestion field in *this*
contract — the `diagnose` subcommand below has its own, richer contract
that does.

### Error codes

| Code | Kind | Stage | Meaning |
|---|---|---|---|
| `E1001` | `undefined_variable` | runtime | Macro referenced a variable that was never assigned |
| `E1002` | `undefined_sub_or_function` | runtime | Entrypoint macro name doesn't exist, or the macro called an unknown Sub/Function |
| `E1003` | `sheet_not_found` | runtime | Reserved for a `Sheets("X")` reference failing *during* macro execution — still not reachable via `run`/`check`'s plain `--json` contract, since `Sheets("X")` auto-creates on write / reads `Empty` on miss unless `Vm::strict_resolution` is on, which only the `diagnose` subcommand sets (see below — it uses its own richer contract, not this error code) |
| `E1004` | `msgbox_blocked` | runtime | A `MsgBox` fired while the VM was configured to treat MsgBox as an error (Python API only; not reachable from the CLI today) |
| `E1007` | `object_variable_not_set` | runtime | A member access (`.Value`, `.Copy`, `.Range(...)`, …) through an object variable holding no live reference — `Dim r As Range` with no `Set`, or after an explicit `Set r = Nothing`. Also covers the *With-block* half of the same condition: a bare `.member` reference with no enclosing `With` block, or a `With` block whose target is an unset object variable. Real VBA's error 91; the message is always exactly `Object variable or With block variable not set` |
| `E1099` | `runtime_error` | runtime | Any other runtime failure not covered above — **also where an argument-count mismatch or an undefined `GoTo`/`On Error GoTo` label lands in run-mode's own `--json` output** (see below; `check` reports these under their own `E1008`/`E1009` codes, but run-mode's `classify_runtime_error` has no prefix match for either message yet, so they fall through to the generic bucket) |
| `E2001` | `parse_error` | parse | The VBA source failed to parse |
| `E3001` | `io_error` | io | Reading the VBA file, reading `--file`, or writing `--output` failed |
| `E3002` | `sheet_setup_error` | setup | Resolving which sheet to use failed *before* the macro ran — a workbook with no sheets, or `--sheet <name>` naming a sheet that doesn't exist in `--file` |

Classification is currently done by pattern-matching the existing
`Result<_, String>` error text at the CLI boundary (`src/diagnostics.rs`),
not by a typed error enum in the VM. This is a known, deliberate interim
trade-off — see the "runtime error 分類を型付きエラーへ" item in
the release roadmap for future hardening work.

**Undefined Sub/Function calls, argument-count mismatches, and undefined
`GoTo`/`On Error GoTo` labels are checked once, up front, before the
entrypoint's first statement executes** — matching real VBA, where these are
compile errors, not runtime ones. Three consequences worth knowing for a
`--json` consumer: (1) `On Error Resume Next`/`On Error GoTo` anywhere in the
project can never catch one of these — the error is unconditionally returned
from `run`, since no statement (including any `On Error`) has executed yet;
(2) the check covers the *whole* project, not just the statements the
entrypoint actually reaches — a Sub the entrypoint never calls, but that
itself contains one of these three problems, still fails the run; (3) `x = 1`
followed by one of these three errors on the next line means `x` was never
actually assigned — nothing in the macro ran at all. An "invalid assignment
target" check (e.g. calling a Function's result as if it were an array
element) was considered for this same pre-flight pass and deliberately
dropped: `name(args) = value` parses the same way whether `name` is a real
array or (invalidly) a Function name, and telling those apart isn't
decidable without type inference this project stays out of by design — that
case still surfaces as an ordinary runtime error, catchable by `On Error`
like any other.

## `messages` semantics

- Populated from every `MsgBox` statement the macro executed, **in
  execution order**, regardless of whether the run ultimately succeeded or
  failed (a `MsgBox` immediately followed by a runtime error still shows up
  in `messages`).
- If the VM is configured to treat MsgBox as blocking (`error_on_msgbox`,
  Python API only), the message is recorded *before* the resulting error is
  raised — `messages` always reflects every `MsgBox` the macro attempted to
  show, whether or not it was allowed to display.
- Scoped to a single `run_sub` call: reusing the same VM instance for a
  second macro run does not carry over the first run's messages.

## Non-JSON mode

Exactly the original behavior: non-empty cells printed as `<address>\t<value>`
TSV lines to stdout, `MsgBox` text printed to stdout inline at the point
it's shown, errors printed to stderr as `error: <message>` with exit code
`1`. This mode is not part of the versioned contract above and may keep
evolving independently — machine consumers should use `--json`.

## Multi-module projects (Milestone B2, Phase 1)

Both run-mode and `check` accept more than one `.bas` file. There is no
project manifest (`elixcee.toml`) yet, and `.cls` class modules aren't
supported — every file passed on the command line is a standard module.

- **Module names**: derived from `Attribute VB_Name = "..."` if the file
  has one (matching how VBA itself names modules), otherwise the file's
  stem — both lowercased. Two files resolving to the same module name is a
  load-time error (run-mode) or an `E1006`/`duplicate_module_name`
  diagnostic (`check`).
- **`Module.Sub` qualification**: `MacroName` (run-mode) or `--entry` (see
  below) may be a bare name or `Module.Sub`. A bare name is resolved
  project-wide by an unqualified search across every module.
- **Cross-module bare-name collisions are rejected, not resolved.** If two
  modules each declare a Sub (or, separately, a Function) with the same
  bare name, the whole run is refused — `run-mode` exits with an error
  before executing anything; `check` reports it as an `E1005`/
  `duplicate_sub_or_function` diagnostic. This is deliberate: real VBA
  resolves an unqualified call to *its own module's* definition first, and
  treats `Private` procedures as invisible outside their module — a flat
  cross-module namespace can't express either rule, so a genuine collision
  is refused rather than silently resolved by an arbitrary tie-break.
  **This means `Module.Sub` qualification cannot rescue a real collision**
  — it only disambiguates when the bare name would otherwise resolve fine
  on its own (e.g. for explicit clarity in a script).
- **`Type` name collisions across modules are not detected** — unlike Sub/
  Function, a `Type` defined identically-named in two modules is silently
  last-wins in the merged type table, with no rejection or diagnostic.
  Deferred as a Phase 2 item (cross-module UDTs are rare); tracked in
  the release roadmap.
- **`Module.Sub` qualification only makes sense for multi-file invocations.**
  A single-file run or check still resolves the entrypoint by bare Sub name
  only (unchanged from before this milestone, to keep the single-file path
  byte-for-byte backward compatible) — passing `Module.Sub` against a single
  file will not resolve, even if the name matches. Qualify only when passing
  more than one file.
- **Runtime error `location` is `null` for multi-module runs.** A runtime
  error's source position is a char offset into whichever module's source
  was executing, but `location` has no module identifier to attribute that
  offset to correctly (a single-source assumption from the source-location
  work that has not yet been revisited). Single-file
  runs are unaffected and keep their exact precise `location`. Parse errors
  are unaffected in both modes — each file parses independently, so its
  error location is always unambiguous.

## `check` subcommand (static analysis, no execution)

```
elixcee check <vba_file>... [--entry <MacroName>] [--json]
```

Inspects one or more `.bas` files **without running them** — useful as a
fast pre-flight signal before spending a real macro execution, or for
checking a macro that isn't safe/ready to run yet. Every positional argument
is a file; the entrypoint, if any, is always `--entry`, never positional —
unlike run-mode, `check`'s entrypoint is optional, so a positional macro
name would be ambiguous against a project with several files and no desired
entrypoint check (e.g. `elixcee check *.bas` to check every module in a
project). Omit `--entry` to check the file(s) on their own without
asserting a particular entrypoint exists.

This is a separate command from the run-mode above, with its own JSON shape
(a batch of findings, not a single result/error):

```json
{
  "schema_version": 1,
  "ok": false,
  "diagnostics": [
    {
      "severity": "error",
      "code": "E1002",
      "kind": "undefined_sub_or_function",
      "message": "Sub 'Bogus' not found",
      "location": null
    }
  ]
}
```

- `ok` is `true` iff no diagnostic has `"severity": "error"` — an `"info"`
  diagnostic (see `I1001` below) never fails the check on its own.
- Exit code: `0` if `ok`, `1` otherwise — same coarse-exit-code philosophy as
  run-mode.
- Non-JSON mode prints one line per diagnostic (`<severity> <code> <kind>:
  <message> (<file>:<line>:<column>)`, or bare `ok` when the list is empty)
  and uses the same exit-code rule.
- A parse error short-circuits everything else (nothing else is checkable
  once the file doesn't parse) — it's always the only diagnostic present
  when it occurs.

### Diagnostic codes

| Code | Kind | Severity | Meaning |
|---|---|---|---|
| `E3001` | `io_error` | error | The given `vba_file` couldn't be read (same code as run-mode's io failure) |
| `E2001` | `parse_error` | error | The VBA source failed to parse (same code as run-mode's parse failure) |
| `E1002` | `undefined_sub_or_function` | error | The given `MacroName`/`--entry` doesn't exist as a `Sub` in the file/project (same code as run-mode's missing-entrypoint failure) |
| `E1005` | `duplicate_sub_or_function` | error | (multi-module) Two modules declare a Sub, or separately a Function, with the same bare name — see "Multi-module projects" above |
| `E1006` | `duplicate_module_name` | error | (multi-module) Two files resolved to the same module name |
| `E1008` | `argument_count_mismatch` | error | A call to a Sub/Function declared *in the same file being checked* passes a different number of arguments than it declares — cross-module calls aren't checked (this diagnostic only ever sees one module's own `Program`), and neither is a call inside the callee's own body (recursion) |
| `E1009` | `undefined_label` | error | A `GoTo`/`On Error GoTo` target isn't a `Label` anywhere in the same Sub/Function (VBA label scope is the whole procedure, not the current block) |
| `E1010` | `blocked_external_effect` | error | The source contains an external-effecting construct such as `Shell`, COM/object creation, `WScript`, file-system object access, or `Open`/`Kill`; the default VM does not execute it and `check` fails explicitly |
| `I1001` | `interactive_call` | info | The macro contains a `MsgBox` call — not broken, just not fully headless |
| `I1002` | `unsupported_construct` | info | A line is a no-op because the construct on it isn't recognized/implemented (`Debug.Print`, an unrecognized `Range`/`Sheets` property or method, a property/field read without assignment, or calling a Sub without `Call`/parentheses) — the macro still runs to completion, this just makes an already-silent no-op visible |

Calls to undefined Sub/Function names *inside* the macro body (`Call Foo(...)`,
bare `Foo(...)`, or any nested `Bar(...)` buried in an expression) are also
detected, using `E1002`/`undefined_sub_or_function` — the same code as a
missing entrypoint, and the same location-granularity rule (statement-level:
the diagnostic points at the start of the enclosing statement, not the exact
sub-expression, since expressions don't carry their own span). A call is
only flagged if it doesn't resolve to a user `Sub`/`Function` *or* a built-in
VBA/`WorksheetFunction` name — resolution consults the VM's real dispatch
tables directly (a cheap throwaway probe call), not a hand-maintained mirror
of them, so this can't drift out of sync as built-in functions are added.

A call is also considered resolved if it matches an in-scope variable,
array, or record name (parameters, and anything assigned/declared anywhere
in the same `Sub`/`Function`) — this AST has no separate "array index"
expression, so `arr(i)` and `func(i)` are otherwise indistinguishable, and
indexed reads of a `Split()` result or a `Dim arr(10)` are ordinary VBA, not
errors. In a multi-module check, a bare call is also resolved against every
*other* module's Sub/Function names — an unqualified cross-module call
isn't misreported as undefined just because this diagnostic pass only sees
one module's own AST at a time.

Two more checks (`E1008`/`argument_count_mismatch` and `E1009`/
`undefined_label`) exist specifically so `check` agrees with what `run` will
actually do — `run` enforces both, and `check`, before this pair existed,
had no way to know about either: a program `check` reported clean could
still have `run` refuse to execute a single statement of it. Every violation
of either kind is reported (not just the first, unlike undefined-name
detection's own early-exit-per-statement style above), each with its own
located diagnostic. See the run-mode section above for exactly what "`run`
enforces this pre-flight" means in practice (uncatchable by `On Error`,
whole-project scope, and so on).

Unrecognized/unsupported constructs that silently became no-ops are also
detected (`I1002`/`unsupported_construct`, info severity — a plain `Dim x`
or a `Static x As Type` declaration inside a Sub are *not* flagged, since
those are intentional no-ops by design, not gaps).

This also covers two narrower cases: unsupported constructs at **module
level** (outside any `Sub`/`Function`), and unrecognized dotted access
nested inside a `With` block. At module level, only a `Const` declaration is
flagged (its value is never evaluated anywhere — a real gap, since a plain
`Public x`/`Dim x` with no value is a harmless no-op just like its Sub-level
counterpart) plus any genuinely unrecognized module-level line; inside a
`With` block, an unrecognized `.property`/`.Method` or a field read without
assignment is flagged the same way as the equivalent case outside a `With`.

### What this does not check yet

One narrow edge case remains unflagged: a `With <target>` header whose
target isn't `Sheets`/`Worksheets` or a plain identifier (e.g. the token
right after `With` isn't an identifier at all — malformed input). This
happens before any statement exists to attach a diagnostic to, and fixing it
would require a shape change to how `With` is represented in the AST for a
  very rare case — see the release roadmap for the
reasoning.

## `snapshot` subcommand (workbook inspection, Milestone B4)

```
elixcee snapshot <file> [--json]
```

Reads a `.xlsx`/`.xlsm`/`.ods` file **directly, without executing any VBA** —
same "inspect, don't execute" posture as `check`. Takes exactly one file (a
workbook, not a `.bas` VBA source — an unsupported extension is an
`io_error`, same as a nonexistent path). Prints every sheet's non-empty
cells as Markdown by default, or as JSON with `--json`. This is a separate
subcommand, not an extension of run-mode's `--json` output (which still only
reports the single active sheet, unchanged).

```json
{
  "schema_version": 1,
  "ok": true,
  "file": "Book1.xlsx",
  "sheets": [
    {
      "name": "Sheet1",
      "sheet_id": "1",
      "stable_id": "sheet1",
      "cells": [{"address": "A1", "value": 42}]
    }
  ]
}
```

- `sheet_id` is the raw `<sheet sheetId="...">` attribute from the file's
  `workbook.xml`, as a string — `null` for `.ods` (no equivalent attribute)
  or if the attribute was missing.
- `stable_id` is always present: `"sheet{sheet_id}"` when a real `sheet_id`
  exists, otherwise a synthetic `"sheet{1-based position}"`.
- **`sheet_id`/`stable_id` are deliberately not named `code_name`.** A field
  called `code_name` would suggest VBA's real `CodeName` property — an
  identifier assigned in the VBA IDE and stored in the binary
  `vbaProject.bin` OLE stream, which this reader doesn't parse (doing so
  would need a full OLE/Compound File Binary parser — well outside this
  feature's scope). What's actually exposed here is much weaker: a file
  format attribute (or a positional fallback), not a VBA-assigned identity.
  Naming it `sheet_id`/`stable_id` keeps `code_name`/`vba_code_name` free
  for that real property, if it's ever implemented.
- **"Stable" is honest for a genuine external `.xlsx`** — a real file's
  `sheetId` survives a tab rename. It is **not** stable for a file elixcee
  itself wrote: this repo's own xlsx writer regenerates `sheetId`
  sequentially from the current sheet order on every save, so
  re-snapshotting an elixcee-produced file after any sheet add/remove/
  reorder renumbers every `stable_id`.
- **Uniqueness across sheets in one snapshot holds only for a conformant
  file.** A real `.xlsx` (OOXML requires `sheetId` on every `<sheet>`), an
  elixcee-written `.xlsx` (always sequential), or an `.ods` (always
  synthetic, positions are unique) can't collide. A hand-edited or
  non-conformant `.xlsx` mixing a real `sheetId` with a sheet missing one
  could coincidentally produce the same `stable_id` for two sheets — this
  isn't detected or deduplicated.
- Cell content is intentionally minimal in this phase: address + computed
  value only, same per-cell shape as run-mode's `cells` array, just for
  every sheet instead of only the active one. No named ranges, no formula
  text, no cell formatting.
- Failures reuse run-mode's `error` shape/codes (`E3001`/`io_error` for a
  missing file or unsupported extension) via the same `messages: []`
  convention — `messages` can never be populated here since no macro ever
  executes.
- Non-JSON (Markdown) output is a top-level sheet index table (name /
  stable_id / cell count) followed by one address/value table per sheet —
  display-only, not meant to round-trip (table-unsafe characters in cell
  values are escaped for readability, not reversibly).

## `test-workbook` subcommand (property-based testing, Milestone B5a)

```
elixcee test-workbook <fixture.toml> [--json] [--seed <N>] [--case <N>]
```

Reruns a macro against a starting `.xlsx`/`.ods` workbook many times with
generated boundary-value inputs, checking each run for panics, runtime
errors, timeouts, and Excel error values in a result range. Every case is
fully independent: a fresh `Vm`, a fresh read of the workbook file, no
carried-over cells/variables/MsgBox log/deadline state.

### Fixture format

```toml
name = "order calculation"
workbook = "fixtures/orders.xlsx"   # resolved relative to the .toml file's own directory
vba_files = ["Main.bas"]            # one or more .bas files (same multi-module rules as run-mode)
macro = "Main.Process"              # bare or Module.Sub-qualified
cases = 100
seed = 42
timeout_secs = 10                   # optional, default 10

[[inputs]]
range = "Input!B2:B10"
strategy = "boundary_numeric"       # or "boundary_string"

[[assertions]]
range = "Result!A1:F100"
rule = "no_excel_errors"
```

Parsed by a hand-rolled, deliberately minimal TOML-subset parser (not the
`toml` crate — that's a `[dev-dependencies]`-only crate added for
`tests/blackbox.rs`, and pulling it into the release binary would reverse
this project's zero-new-runtime-dependency principle, the same one that
led Milestone B2 to reject a TOML project manifest). Only flat
`key = value` lines and `[[inputs]]`/`[[assertions]]` array-of-tables are
supported; anything outside that subset (inline tables, multi-line
strings, dotted keys, trailing junk after a value) is a hard parse error,
not a silent skip.

### Strategies (v1: two)

- `boundary_numeric`: `Empty, 0, 1, -1, 999999999, -999999999` — chosen
  over `i64::MAX`/`MIN` since these sit just past VBA's classic
  `Integer`/`Long` overflow boundaries, where realistic spreadsheet-macro
  bugs actually show up.
- `boundary_string`: `"", "test", "a"×1000`.

Each cell in an input range gets an independent draw from its strategy's
pool per case (not one value repeated across the whole range). Sampling is
with replacement — `cases` independent trials, matching how `proptest`
(already a dev-dependency in this repo) works, not exhaustive enumeration
of the (small) pool.

### Assertion rules

`no_panic`, `no_runtime_error`, and `no_timeout` are **always active** for
every case — not TOML-declared, since a panic or hang is never something a
property test should let you opt out of. `[[assertions]]` is specifically
for range-scoped rules; `no_excel_errors` (scans the range for any
`#DIV/0!`/`#VALUE!`/`#REF!`/etc. cell value) is the only one in v1. A
missing sheet in an assertion's range is a hard error (fixture/config
problem), not a silent "no errors found".

### Output

Success: `{"schema_version":1,"ok":true,"seed":42,"cases_run":100}`

Failure (fail-fast — stops at the first failing case, matching both
`proptest`'s own convention and this CLI's "exactly one JSON object per
invocation" contract):

```json
{
  "schema_version": 1,
  "ok": false,
  "seed": 42,
  "case_index": 17,
  "inputs": [{"address": "Input!B2", "value": -1}],
  "failure": {"rule": "no_excel_errors", "address": "Result!C8", "actual": "#DIV/0!"}
}
```

`failure.address`/`failure.actual` are only present for `no_excel_errors`;
`no_panic`/`no_runtime_error`/`no_timeout` use `failure.message` instead.
Exit code 0/`ok:true` if every run case passed, 1/`ok:false` on the first
failure — same coarse-exit-code convention as every other subcommand.

### Replay

`--case <N>` (0-based, matching `case_index`) reruns exactly one case
instead of the full `cases` loop; `--seed <N>` overrides the fixture's own
`seed`. The per-case seed is derived deterministically from
`(base_seed, case_index)`, and input draws are made in a pinned order
(`[[inputs]]` declaration order, cells row-major within each range), so
`elixcee test-workbook fixture.toml --seed 42 --case 17` always reproduces
the exact same drawn inputs as case 17 of a full run with `seed = 42`.

### Known limitation

`RANDARRAY`/`Rnd`'s PRNG (`src/formula/eval.rs`) is a **thread-local**, not
a `Vm` field, so a fresh `Vm` per case does not reset it — draws continue
across cases on the same thread. `--seed`/`--case` replay is only
guaranteed to reproduce identical *input generation* (which boundary value
gets written where), not VBA-visible randomness for a macro that calls
`RANDARRAY`/`Rnd`. Neither `boundary_numeric` nor `boundary_string` (the
only strategies in this phase) invoke any VBA-side randomness, so this
doesn't bite v1 — but it's a real constraint for any future strategy that
does.

### Explicit non-goals (deferred to a later B5 phase)

Shrinking (minimizing a failing input) is not implemented — the order is
deterministic generation → save failing case → single-case replay first,
shrinking later, per the roadmap. Only two strategies and one range-scoped
assertion rule exist in this phase; more of each are plausible later
additions, not redesigns.

## `diagnose` subcommand (Excel operation diagnostics, Milestones B6a/B6b/B6c/B6c2)

```
elixcee diagnose <vba_file>... <MacroName> --file <workbook> [--json]
```

Runs one macro once and classifies *why* it failed — a missing worksheet,
a missing workbook, an out-of-bounds array index, a Copy/Paste shape
mismatch or missing clipboard, a write to a protected sheet, or a
Copy/Paste that conflicts with a merged-cell layout — with concrete
evidence (the requested key and what was actually available, a "did you
mean" suggestion, the mismatched shapes, the protected sheet's name, or the
conflicting merged range), instead of only a bare runtime-error string.

Missing-sheet/workbook/array-bounds classification (Milestone B6a) has a
different posture from `run`/`check`/`test-workbook`: it turns on
`Vm::strict_resolution`, which makes elixcee's usual auto-vivify/silent-
`Empty` convenience for `Sheets("X")`/`Worksheets("X")` references into a
hard, classified failure — because a diagnostic tool whose whole purpose is
"what would Excel actually reject here" needs to *not* paper over the exact
class of mistake it exists to catch. Every other subcommand leaves
`strict_resolution` off (the default) and is completely unaffected.
Copy/Paste shape-mismatch, sheet-protection, and merged-cell-conflict
classification (Milestones B6b/B6c/B6c2, below) work differently: those
checks are unconditional hard errors in every mode that executes the macro
(`run`/`diagnose`/`test-workbook`) — `diagnose` doesn't need a toggle for
them, it just surfaces the same failure with structured evidence instead of
a bare error string.

### Strict-resolution mode

- **Missing worksheet** (`Sheets("X")`/`Worksheets("X")`, by name or by a
  new 1-based numeric index — `Worksheets(2)`): normally a write
  auto-creates the sheet and a read silently returns `Empty`; in strict
  mode, either is a `WORKSHEET_NOT_FOUND` failure with the requested name,
  every existing sheet name, and (if within a small bounded Levenshtein
  distance) a suggested closest match. elixcee has no real workbook
  tab-order tracking, so a numeric index resolves against sheet names
  sorted alphabetically, not Excel's actual left-to-right tab order — an
  honest fidelity gap, not a bug.
- **Missing workbook** (`Workbooks("X").Worksheets(...)`, a new construct
  in this milestone): elixcee only ever has one workbook loaded at a time
  (via `--file`), so this doesn't model real multi-workbook switching — it
  only compares the requested name/index against the one loaded workbook,
  raising `WORKBOOK_NOT_FOUND` on any mismatch. This check fires
  unconditionally (not gated behind strict mode), since `Workbooks(...)`
  is brand new — there's no pre-existing lenient behavior for it to
  preserve.
- **Array out of bounds** (`arr(i)` past its declared size): already a
  hard error in every mode before this milestone; now also carries
  structured `ARRAY_INDEX_OUT_OF_BOUNDS` evidence (`lower`/`upper` are
  elixcee's true 0-based bounds — `Dim arr(1 To N)`'s non-zero lower bound
  isn't tracked anywhere, so this reports elixcee's actual model, not a
  fabricated VBA-style `1 To N`).
- **`On Error Resume Next`/`On Error GoTo` are not honored** while
  `strict_resolution` is on — the first resolution failure always
  propagates and gets reported, rather than being silently swallowed or
  redirected by the macro's own error handling (which in real VBA usage is
  exactly the code most likely to be masking the bug this subcommand
  exists to surface).
- New syntax added alongside this: `Sheets(name).Range(addr)` (read and
  write — previously only `.Cells(r,c)` was supported off a sheet name);
  without it, none of the sheet-resolution scenarios above could even be
  written as a runnable macro.

### Copy/Paste shape mismatch and clipboard state (Milestone B6b)

`.Copy` now populates a clipboard (`ClipboardState`: the source address,
its row/column shape, and its cell values snapshotted at copy time — not
re-read at paste time, matching real Excel's copy-then-mutate-then-paste
semantics). `.Paste`/`.PasteSpecial` (new syntax: `Range(addr).Paste`,
`Range(addr).PasteSpecial [Transpose:=<expr>]`,
`Worksheets(sheet).Paste Destination:=Range(addr)`) consume it. Unlike
B6a's sheet resolution, these checks are **unconditional hard errors in
every mode that executes the macro** (`run`/`diagnose`/`test-workbook`
alike) — not gated behind
`strict_resolution` — because nothing in elixcee ever relied on the old
silently-wrong behavior (see below), and real Excel itself raises a hard
runtime error (1004) for both cases regardless of any error-handling
state. `On Error Resume Next`/`GoTo` still swallow these in normal `run`
mode exactly as they do for any other error; `diagnose` still bypasses
that (same mechanism as B6a) so the first failure is always reported.

- **`PASTE_SHAPE_MISMATCH`**: the destination — when given as an explicit
  range (`"E1:F10"`), not a single anchor cell — doesn't match the
  clipboard's shape (after accounting for `Transpose:=True`, which swaps
  rows/cols). Evidence carries `source_addr`/`source_rows`/`source_cols`,
  `dest_addr`/`dest_rows`/`dest_cols`, `transpose`, and a `copy_location`
  (the *Copy* statement's own location — `location` at the top level
  already points at the failing *Paste* statement, so this is the only
  root cause with two locations). Suggestions are mechanically derived:
  "resize the destination to `<top-left>:<computed bottom-right>`" and
  "or specify only the top-left cell `<anchor>`". Two cases are never
  shape-checked, matching real Excel: a single anchor destination cell
  (auto-expands to the clipboard's shape), and a single-*cell* clipboard
  pasted into an explicit destination range of any size (Excel's
  well-known "fill many cells with one copied value" behavior — a
  destination that's an exact multiple of a *multi-cell* clipboard, i.e.
  tiling, is a rarer sibling not modeled here).
- **`PASTE_WITHOUT_COPY`**: a `.Paste`/`.PasteSpecial` ran with nothing on
  the clipboard — either no prior `.Copy` at all, or
  `Application.CutCopyMode = False` cleared it since. Evidence carries
  only `dest_addr` (there is no copy to point at).
- Fixes a latent, previously-untested bug in `Range.Copy Destination:=`:
  the old execution parsed the destination via a single-cell-only parser
  and silently fell back to the source's own top-left cell for any real
  range address (a no-op) — nothing ever exercised this path with a
  multi-cell `Destination:=`. It now resolves the destination as a real
  range and shape-checks it, the same as bare `.Paste`.
- Non-goals for B6b: `.Cut` (only `.Copy` is modeled — `CutCopyMode` is
  only ever cleared, never set by a cut); `PasteSpecial`'s `Operation:=`/
  `SkipBlanks:=`/paste-type parameters; copying formulas with relative-
  reference adjustment (`.Copy` only ever copied baked values before this
  milestone too); real OS-level/cross-application clipboard.

### Sheet protection (Milestone B6c)

`Sheets(name).Protect`/`.Unprotect` (also reachable via
`Worksheets(name)...`/`Workbooks(...).Worksheets(...)...`) toggle a
per-sheet protected flag. Trailing kwargs (`Password:=`, `DrawingObjects:=`,
`Contents:=`, etc.) are accepted and discarded — elixcee has no security
model and doesn't enforce a real password, only the "is this sheet
protected" question a diagnostic tool needs. **`UserInterfaceOnly:=True`
is modeled**: real Excel blocks manual UI edits but *not* macro writes in
that mode, so `.Protect UserInterfaceOnly:=True` leaves the sheet
macro-writable in elixcee (there's no UI to block). Bare `.Protect` (or
`UserInterfaceOnly:=False`) blocks macro writes too. While protected, **any**
cell-content mutation on that sheet is a hard error, **unconditionally, in
every mode that executes the macro** (`run`/`diagnose`/`test-workbook`) —
writes (`Cells`/`Range.Value`/
`.Formula`), `Range.ClearContents`/`.Clear`/`.Delete`/`.Insert`/`.Sort`,
`.Copy`/`.Paste`/`.PasteSpecial` into it, and deleting the sheet itself —
matching real Excel, which raises a hard runtime error for all of these
regardless of `On Error` state (same "unconditional hard error" reasoning
as B6b's shape-mismatch/empty-clipboard checks: nothing pre-existing
relied on writes to a "protected" sheet succeeding, since the concept
didn't exist before). **Reads are never blocked** — protection only gates
edits, matching real Excel. `On Error Resume Next`/`GoTo` still swallow
the error in normal `run` mode; `diagnose` still bypasses that via the
existing B6a mechanism. Protecting or unprotecting a nonexistent sheet is
itself a `WORKSHEET_NOT_FOUND` failure, unconditionally (a brand-new
construct, same precedent as `WorkbookQualifiedSheet`'s mismatch check).

- **`SHEET_PROTECTED`**: evidence carries only `sheet` (the protected
  sheet's name) — the simplest evidence shape of any root cause so far.
  Suggestion: `"unprotect the sheet first: Worksheets(\"<sheet>\")
  .Unprotect"`.
- No bare `ActiveSheet.Protect` — elixcee has no `ActiveSheet` concept
  anywhere; `Sheets(name)`/`Worksheets(name)` qualification is required,
  same as every other sheet-level statement in this codebase.

### Merged-cell-aware Paste diagnostics (Milestone B6c2)

`WorkbookSheet` now carries `merged_ranges` (parsed from XLSX
`<mergeCell ref="...">` and ODS `table:number-columns-spanned`/
`table:number-rows-spanned`), threaded into the VM as
`merged_ranges: HashMap<sheet, Vec<rect>>` alongside `sheets`/
`active_sheet`/`protected_sheets`. `do_paste` checks the destination
against this state right after computing the fill dimensions — same
"unconditional hard error in every mode that executes the macro"
(`run`/`diagnose`/`test-workbook`) posture as B6b/B6c, for the same
reason: real Excel rejects these pastes outright regardless of `On Error`
state, and nothing pre-existing relied on lenient behavior since the
concept didn't exist before. Checks run in this order, first match wins:

1. **`PASTE_INTO_NON_ANCHOR_MERGED_CELL`**: the destination anchor cell
   falls inside an existing merge but isn't that merge's own top-left cell
   — applies regardless of destination shape, including a single-cell
   destination. Evidence: `dest_addr`, `dest_sheet`, `merged_range`.
   Pasting into a merge's own top-left cell is the normal way to write to
   a merged cell in real Excel and is never flagged.
2. **`PASTE_PARTIAL_MERGED_RANGE`**: the destination — only when genuinely
   multi-cell — partially overlaps one or more merges without fully
   containing them. Evidence: `dest_addr`, `dest_sheet`, `conflicts` (every
   overlapping merge's range).
3. **`PASTE_MERGE_LAYOUT_MISMATCH`**: only checked when the source isn't a
   single cell and the shape already matched (B6b); the source-side and
   destination-side merges, normalized to relative position within their
   own rect (transposed first if `Transpose:=True`), don't line up.
   Evidence: `source_addr`, `source_sheet`, `dest_addr`, `dest_sheet`,
   `conflicts` (the mismatching destination-side merges), and
   `copy_location` (same two-location convention as `PASTE_SHAPE_MISMATCH`).

Non-goals for B6c2: AutoFilter/`SpecialCells(xlCellTypeVisible)`
visible-cells-only copy, Excel Tables, hidden rows/columns, external
OS-level clipboard, formula relative-reference translation on copy, and
merge-awareness on any statement other than Paste (`RangeSort`/
`RangeInsert`/`RangeDelete`/plain `RangeWrite` are untouched — a merge only
blocks a *paste into* it, not other mutations, in this milestone).
Multi-area (`Areas`) ranges — deferred here — now has its own foundation
milestone, B7a (below).

### Output

Success: `{"schema_version":1,"ok":true,"messages":[...]}`

Failure — its own JSON contract (like `test-workbook`'s), not the flat
`ElixceeError` shape above, since ranked evidence doesn't fit `{code, kind,
message, location}`:

```json
{
  "schema_version": 1,
  "ok": false,
  "message": "Sheet '売上2025' not found",
  "location": {"file": "Main.bas", "line": 2, "column": 5},
  "root_causes": [
    {
      "code": "WORKSHEET_NOT_FOUND",
      "certainty": "definite",
      "expression": "Worksheets(\"売上2025\")",
      "requested": "売上2025",
      "available": ["input", "売上2026", "sheet1", "集計"],
      "suggested": "売上2026",
      "suggestions": ["did you mean '売上2026'?"]
    }
  ],
  "messages": []
}
```

`root_causes` is an array (currently at most one entry — the first
failure) rather than a bare object, so a later milestone's ranked-candidate
model ("3 possible reasons, ranked") can reuse this exact shape without a
breaking schema change. `ARRAY_INDEX_OUT_OF_BOUNDS` entries carry
`name`/`index`/`lower`/`upper` instead of the name-lookup evidence fields;
`PASTE_SHAPE_MISMATCH`/`PASTE_WITHOUT_COPY` entries (Milestone B6b) carry
the fields described above; `SHEET_PROTECTED` entries (Milestone B6c)
carry just `sheet`; `PASTE_INTO_NON_ANCHOR_MERGED_CELL`/
`PASTE_PARTIAL_MERGED_RANGE`/`PASTE_MERGE_LAYOUT_MISMATCH` entries
(Milestone B6c2) carry `dest_addr`/`dest_sheet` plus either
`merged_range` or a `conflicts` array (and, for the layout-mismatch case,
`source_addr`/`source_sheet`/`copy_location`), e.g.:

```json
{
  "code": "SHEET_PROTECTED",
  "certainty": "definite",
  "sheet": "sheet1",
  "suggestions": ["unprotect the sheet first: Worksheets(\"sheet1\").Unprotect"]
}
```

```json
{
  "code": "PASTE_SHAPE_MISMATCH",
  "certainty": "definite",
  "source_addr": "A1:C10", "source_rows": 10, "source_cols": 3,
  "dest_addr": "E1:F10", "dest_rows": 10, "dest_cols": 2,
  "transpose": false,
  "copy_location": {"file": "Main.bas", "line": 2, "column": 5},
  "suggestions": [
    "resize the destination to E1:G10",
    "or specify only the top-left cell E1"
  ]
}
```

```json
{
  "code": "PASTE_MERGE_LAYOUT_MISMATCH",
  "certainty": "definite",
  "source_addr": "A1:C10", "source_sheet": "sheet1",
  "dest_addr": "E1:G10", "dest_sheet": "sheet1",
  "conflicts": ["E1:G1"],
  "copy_location": {"file": "Main.bas", "line": 2, "column": 5},
  "suggestions": [
    "unmerge E1:G1 before pasting",
    "or make the source and destination merge layouts identical"
  ]
}
```

Exit code 0/`ok:true` on success, 1/`ok:false` on failure — same
convention as every other subcommand. `location` follows the same
single-module-only rule as run-mode's own `--json` contract (a
`SourceSpan` carries no module id, so a multi-module run reports
`location: null` rather than risk pointing at the wrong module's source).

### Explicit non-goals (deferred to later B6 phases)

B6a covers resolution failures (missing worksheet/workbook, array out of
bounds); B6b covers Copy/Paste shape mismatch and clipboard state (see its
own non-goals list above); B6c covers sheet protection (see its own
non-goals note above); B6c2 covers merged-cell conflicts on Paste (see its
own non-goals note above). Explicitly out of scope, planned for later:

- Hidden/filtered rows, AutoFilter visible-cells-only copy — the user's
  original roadmap bundled these with merged cells under "B6c," but
  merged-cell Paste conflicts shipped first as B6c2 once grounding showed
  each of the others needs its own new reader-format parsing (XLSX/ODS)
  and/or range-model change that merged-cell handling alone didn't need.
  Multi-area (`Areas`) ranges — the range-model change itself — shipped
  as its own foundation milestone, B7a (below); hidden/filtered rows and
  `SpecialCells(xlCellTypeVisible)` remain deferred to B7b/B7c, which
  build on B7a's `RangeRef`/`Rect` model.
- Excel Tables (`ListObjects`) — never part of the user's original
  roadmap; added as a placeholder non-goal during B6a's own docs and kept
  deferred (a full new VBA object model, comparable in scope to `Range`/
  `Sheets` itself, not "add diagnosis to an existing path").
- A real VBA `Collection` object — it doesn't exist in elixcee at all, so
  there is nothing to classify a failure for; adding one is a first-class
  feature, not "add diagnosis to an existing path."
- Real multi-workbook execution — only a name/index mismatch check against
  the single loaded workbook ships in this milestone.
- `Dim arr(1 To N)` non-zero-lower-bound tracking.

`diagnose` itself still runs a macro exactly once against one fixed
workbook state — integration with `test-workbook`'s generated-case search
now ships as its own subcommand, `diagnose-workbook` (Milestone B6d, below),
rather than changing `diagnose`'s own execution model.

## `diagnose-workbook` subcommand (generated-case root-cause diagnosis, Milestone B6d)

```
elixcee diagnose-workbook <fixture.toml> [--json] [--seed <N>] [--case <N>] [--cases <N>]
```

Reuses `test-workbook`'s (B5a) case generator — the identical TOML fixture
format, `boundary_numeric`/`boundary_string` strategies, deterministic
seed/case-index derivation, and fail-fast-on-first-failure model — but runs
each case with `Vm::strict_resolution` turned on (matching `diagnose`'s own
posture) and enriches whichever failures are classifiable with the same
root-cause machinery `diagnose` uses (B6a–B6c2's `ResolutionFailureKind` →
`RootCause` pipeline), instead of only reporting a bare rule/message.

**Honest scope**: most root causes are *structural* — the 3 merge kinds,
`PASTE_SHAPE_MISMATCH`, `PASTE_WITHOUT_COPY`, and `SHEET_PROTECTED` all
depend on the macro's own text and the workbook's fixed layout, not on which
boundary value a case draws, so they fire identically on case 0 (or never)
no matter how many cases run — a single plain `diagnose` invocation already
finds these in one shot. This command earns its keep specifically for
**input-dependent** kinds, chiefly `ARRAY_INDEX_OUT_OF_BOUNDS` (a drawn
value can flip an array index in or out of bounds across cases) and, in
principle, `WORKSHEET_NOT_FOUND`/`WORKBOOK_NOT_FOUND` if a drawn value were
used to build a sheet/workbook name. Every other runtime error (the
majority — type mismatches, division by zero, etc.) has no classification
at all, matching `diagnose`'s own permanent limitation.

- **`--cases <N>`**: overrides the fixture's declared `cases` count for this
  invocation only — scoped to `diagnose-workbook`; `test-workbook` itself
  still only honors the fixture's own `cases` field (no equivalent flag
  there). `--seed`/`--case` behave identically to `test-workbook`.
- Output is `test-workbook`'s existing JSON shape
  (`schema_version`/`ok`/`seed`/`case_index`/`inputs`/`failure` on failure,
  `schema_version`/`ok`/`seed`/`cases_run` on success) plus one sibling
  field on failure: `root_causes` — `[]` when unclassified, a one-item array
  in `diagnose`'s own field shape (same `code`/evidence-field spellings,
  e.g. `source_addr`/`dest_addr`/`conflicts` for the merge kinds) when
  classified. No `copy_location` resolution — there's no per-case source
  text/location tracking in the generated-case search, so paste-related
  kinds always report `"copy_location":null` here (unlike plain `diagnose`,
  which resolves it from the single source file it's given).
- Non-JSON mode appends a `root cause: <CODE>` line when classified — full
  evidence and suggestions are `--json`-only, matching every other
  subcommand's "plain text is a simplified view" convention.

```json
{
  "schema_version": 1,
  "ok": false,
  "seed": 42,
  "case_index": 3,
  "inputs": [{"address": "sheet1!B2", "value": 999999999}],
  "failure": {
    "rule": "no_runtime_error",
    "message": "Array 'arr': index 999999999 out of bounds (len=6)"
  },
  "root_causes": [
    {
      "code": "ARRAY_INDEX_OUT_OF_BOUNDS",
      "certainty": "definite",
      "name": "arr",
      "index": 999999999,
      "lower": 0,
      "upper": 5,
      "suggestions": ["check that 'arr' is large enough for index 999999999 (valid range is 0 To 5)"]
    }
  ]
}
```

### Explicit non-goals (this milestone)

Shrinking (minimizing a failing case's inputs to a smaller reproducer) is
the deliberate next step *after* this milestone, not part of it — cases are
reported as-drawn. Also deferred: filtered/hidden-visible-cell diagnosis
(already deferred from B6c/B6c2; multi-area ranges themselves shipped as
B7a, below), backporting `--cases` to `test-workbook` itself, and any new
`[[assertions]]` rules beyond the existing `no_excel_errors`.

## Multi-area ranges (Milestone B7a)

A single-rectangle `Range` can't model real Excel constructs like
`Range("A1:A3,C1:C3")` or a filtered `SpecialCells(xlCellTypeVisible)`
result (multiple disjoint rectangles, one sheet). B7a adds the underlying
model — `vm::Rect` (a 1-based inclusive rectangle) and `vm::RangeRef`
(`{ sheet, areas: Vec<Rect> }`) — and wires it into Copy/Paste diagnosis
only; every other statement (`RangeSort`, `RangeInsert`, plain cell
read/write, formula evaluation, etc.) is untouched and still single-rect
only. This is the foundation B7b (hidden/filtered rows) and B7c
(`SpecialCells(xlCellTypeVisible)`) build on — those milestones are what
actually *produce* a multi-area `RangeRef` from a real workbook; B7a only
adds the model and its Copy/Paste diagnostics.

**Supported syntax**: `Range("A1:A3,C1:C3")` — a single string-literal
argument with a comma inside it, exactly like real VBA's own union syntax.
No parser/grammar change was needed for this (the comma is inside the
string, not a VBA-level argument separator); only the runtime address
parser (`parse_multi_area_addr`, alongside the existing single-rect
`parse_range_addr`) splits on top-level commas. `Union(...)`, the `Areas`/
`Areas.Count`/`Areas(n)` VBA-visible property, and `Dim rng As Range` /
`Set rng = ...` Range object variables were all still-deferred at the time
this section was written — **they now exist, see "Range object variables,
Union, Areas, SpecialCells, and multi-area Paste (Milestone B7c)" below.**

**Diagnose-only for most shapes, one shape now executes (B7c)**: v1
(B7a) never actually completed (wrote cells for) a multi-area paste —
not even when source and destination areas fully corresponded in count
and shape. As of Milestone B7c, the one fully-matching shape (both sides
multi-area, same `Areas.Count`, matching per-area shapes, paired in
order) now actually pastes; see the B7c section below for exactly which
shape and what's still excluded (`Transpose:=`, merged-cell conflict
checking). Every other shape below is unchanged from B7a and still ends
in one of 4 classified failures instead:

- **`MULTI_AREA_TO_SINGLE_AREA_PASTE`**: source has more than one area,
  destination has exactly one. Evidence: `source_areas`, `destination_areas`
  (always 1 element).
- **`MULTI_AREA_COUNT_MISMATCH`**: both sides have more than one area, but
  `Areas.Count` differs. Evidence: `source_areas`, `destination_areas`.
- **`MULTI_AREA_SHAPE_MISMATCH`**: both sides have more than one area with
  matching counts, but at least one area pair (by position) differs in
  rows/columns. Evidence: `area_index` (1-based, the first mismatching
  pair), `source_area`, `destination_area`.
- **`MULTI_AREA_PASTE_UNSUPPORTED`**: the catch-all for shapes the other 3
  don't name — a single-area source into a multi-area destination, or (as
  of B7c) the reverse, a multi-area source into a single-area destination.
  (Before B7c this also covered the fully-matching multi-to-multi case;
  that shape now executes instead — see the B7c section below.) Real
  Excel would complete these too; elixcee doesn't yet, so this reports the
  limitation plainly rather than silently doing nothing or misreporting a
  mismatch that isn't there. Evidence: `source_areas`, `destination_areas`.

Every area's evidence is `{"address": "...", "rows": N, "columns": N}` —
`columns` (not `cols`), and nested per area — a different convention from
the existing flat `source_cols`/`dest_cols` fields `PASTE_SHAPE_MISMATCH`
uses. No `copy_location`/`copy_span` on any of the 4 kinds (unlike the
merge-conflict kinds) — geometry is the whole story here, there's no
merged-cell layout to cross-reference.

```json
{
  "code": "MULTI_AREA_TO_SINGLE_AREA_PASTE",
  "certainty": "definite",
  "source_areas": [
    {"address": "A1:A10", "rows": 10, "columns": 1},
    {"address": "C1:C10", "rows": 10, "columns": 1}
  ],
  "destination_areas": [
    {"address": "E1:F10", "rows": 10, "columns": 2}
  ],
  "suggestions": [
    "paste each source area separately",
    "copy a contiguous rectangular range",
    "use destination areas with matching count and shapes"
  ]
}
```

### Explicit non-goals (this milestone)

`Union()`, `Areas`/`Areas.Count`/`Areas(n)`, `Dim rng As Range`/
`Set rng = ...` Range object variables, actually completing (writing
cells for) any multi-area paste, hidden/filtered-row awareness (B7b), and
`SpecialCells(xlCellTypeVisible)` were all non-goals **at the time B7a
shipped** — all have since landed (B7b, B7c; see their own sections
below), except a multi-area paste still only executes for the one
fully-matching shape (B7c), not every shape. Shrinking (B5b) is still a
non-goal here — B7a was explicitly sequenced ahead of it because most
structural failures (multi-area, filtered rows) come from workbook
layout, not drawn cell values, so there was nothing for shrinking to
minimize until this structural model existed.

## Hidden row/column evidence (Milestone B7b)

Before `SpecialCells(xlCellTypeVisible)` (B7c) can turn a filtered range
into a multi-area `RangeRef` (B7a), elixcee needs to know which rows/
columns are hidden in the first place. B7b adds that foundation: reading
hidden-row/column metadata from a real workbook, holding it as interval
data, and computing its intersection with the range a `.Copy` last
touched — surfaced through `diagnose`/`diagnose-workbook` as a new,
**non-failure** observation. Copy/Paste behavior itself is unchanged:
hidden cells are still copied/pasted exactly as before. This milestone
only adds observability, not filtering.

**Model**: `vm::Interval { start, end }` (1-based inclusive, one type for
both rows and columns), `vm::SheetVisibility { hidden_rows, hidden_columns:
Vec<Interval> }`, threaded into `Vm.sheet_visibility: HashMap<sheet,
SheetVisibility>` — same lowercase-keyed, populated-by-`populate_from_sheets`
pattern as `merged_ranges` (Milestone B6c2).

**XLSX only** — `<row hidden="1">` and `<col min=".." max=".." hidden="1">`
are parsed; **ODS is explicitly deferred**. ODS's `<table:table-row>`
parsing today only counts rows, and critically doesn't expand
`table:number-rows-repeated` (common for blocks of blank/identical rows) —
without that expansion, a hidden-row flag can't be mapped to correct
absolute row numbers, so attempting it would silently produce wrong
results, worse than not supporting it at all.

**`Vm::hidden_cells_observation(&self) -> Option<HiddenCellsObservation>`**
computes the evidence on demand from `Vm.clipboard` (populated by the last
`.Copy`) intersected with `Vm.sheet_visibility` — not a new stored side
channel. Read-only and idempotent (unlike `take_resolution_failure`, it
doesn't drain anything), so it can be called any number of times after a
run regardless of success/failure. `None` when:
- nothing has been copied, or `Application.CutCopyMode = False` cleared it
  since (the clipboard's own "last surviving Copy" semantics — same
  limitation `PASTE_WITHOUT_COPY` already has on the failure side);
- the copy spanned more than one area (Milestone B7a's multi-area Copy) —
  combining "multi-area source" with "also touches hidden cells" is a
  compound case this milestone doesn't model;
- the copied sheet has no registered hidden rows/columns; or
- none of those hidden rows/columns actually overlap the copied range.

**`visible_cells = (rows − hidden_row_count) × (cols − hidden_col_count)`**
— the product form correctly avoids double-counting a cell hidden by both
a hidden row and a hidden column (inclusion-exclusion), assuming a sheet's
hidden row/column intervals don't overlap each other (true for any real
XLSX).

### Output: the `observations` field

A new **sibling** JSON field, not folded into `root_causes` — `root_causes`
means "why it failed"; `RANGE_CONTAINS_HIDDEN_CELLS` isn't a failure at
all, so it gets its own array, with its own `"certainty":"observed"` tier
(distinct from every `root_causes` entry's `"definite"`). **Present only
when non-empty, on both success and failure** — never an always-present
`"observations":[]`, since that would break every existing `--json` fixture
in `tests/blackbox.rs` that predates this field (all of them, since none
can involve hidden cells pre-B7b).

```json
{
  "schema_version": 1,
  "ok": true,
  "messages": [],
  "observations": [
    {
      "code": "RANGE_CONTAINS_HIDDEN_CELLS",
      "certainty": "observed",
      "range": {"sheet": "sheet1", "address": "A1:C100", "rows": 100, "columns": 3},
      "visibility": {
        "hidden_rows": ["11:14", "30:39"],
        "hidden_columns": ["B:B"],
        "total_cells": 300,
        "visible_cells": 172
      },
      "message": "The range contains hidden rows or columns. Excel operations using visible cells only may produce a multi-area range."
    }
  ]
}
```

Row/column intervals render as `"start:end"` address strings always
(`"11:14"`, `"B:B"`) — even for a single row/column (`"5:5"`, not `"5"`) —
deliberately not `rect_addr`'s single-cell-omits-the-colon convention used
elsewhere, matching this feature's own worked example.

`diagnose-workbook` gets the identical field via `FixtureResult::Passed`/
`::Failed` both gaining `hidden_cells: Option<Box<HiddenCellsObservation>>`
(mirroring Milestone B6d's `resolution_kind` precedent, but threaded
through *both* variants since this isn't failure-gated) — `test-workbook`'s
own `to_json`/`to_plain_text` ignore the field entirely, byte-identical to
before. **Honestly no additional value over a single plain `diagnose` call
here**, same as most of B6d's structural root causes: hidden-row/column
metadata depends only on the workbook file and the macro's `.Copy`
statement, never on drawn cell values, so it fires identically on case 0
(or never) no matter how many generated cases run.

### Explicit non-goals (this milestone)

AutoFilter condition evaluation, a strict manual-vs-filter-hidden
distinction, outline/group expand state, any Excel-faithful Copy/Paste
behavior change for hidden cells (still copies/pastes them exactly as
before — this milestone only adds observability), ODS hidden-row/column
support, and multi-area-source hidden-cell observations.
`SpecialCells(xlCellTypeVisible)` execution and generating
`RangeRef.areas` from hidden-row/column results were also non-goals here
— both have since landed in Milestone B7c (below), consuming this
milestone's `sheet_visibility` data directly rather than re-deriving it.

## Range object variables, Union, Areas, SpecialCells, and multi-area Paste (Milestone B7c)

B7a/B7b built the range-geometry and hidden-cell-metadata foundation; B7c
is the VBA object-reference layer on top of it: `Dim rng As Range`,
`Set rng = Range(...)`, `Union(...)`, `.Areas`, `SpecialCells
(xlCellTypeVisible)`, one multi-area Copy/Paste shape that now actually
executes, and `ThisWorkbook`/`ActiveWorkbook`/`ActiveSheet`. This
milestone is entirely inside the VBA language runtime (parser + `Vm`) —
it adds no new `--json` fields of its own. The one place it changes this
contract is `diagnose`'s `MULTI_AREA_PASTE_UNSUPPORTED` firing condition,
already corrected in the B7a section above.

**Object variables are a namespace, not a `Variant`**: `Set`-assigned
variables live in `Vm.object_variables` (`String` → `ObjectRef`), kept
separate from `Vm.variables` (`Variant`s) — mirroring VBA's own
`Set`-vs-`=` distinction, and deliberately *not* a new `Variant` variant,
since `Variant` is defined in the shared `elixcee-types` crate that the
WASM bridge also consumes. `ObjectRef::Range` is just the existing B7a
`RangeRef` (sheet + area coordinates) — no cell values are stored in it,
so two variables holding the same `RangeRef` already get real `Set`
reference semantics for free: cell values live in `Vm.sheets`, keyed by
coordinates, so a write through one variable is a write to that shared
store, immediately visible when reading through any other variable
referencing the same coordinates.

**`.Value`/`.Formula`/`.Areas.Count` ride existing grammar**: `<var>.Value
= x`, `x = <var>.Value`, and `x = <var>.Areas.Count` reuse the pre-existing
generic `<var>.<field>` statement/expression grammar (record field
get/set) rather than adding new syntax — disambiguated purely at runtime
by whether `var` is a key in `object_variables`, which only `Set` ever
populates. `<var>.Copy [Destination:=Range(addr)]` and `.Areas(n)`/
`.SpecialCells(xlCellTypeVisible)` are genuinely new grammar.

**`SpecialCells(xlCellTypeVisible)`** decomposes each area into the
Cartesian product of its maximal visible row-bands and column-bands
(computed from B7b's `Vm.sheet_visibility`) — exactly the visible-cell
set, though not necessarily grouped into the same number of `Areas` real
Excel would report when *both* axes have hidden spans (Excel's own
area-merging heuristic there is unmodeled; the cell set itself is
correct). Raises an error (no VBA exception type modeled — a plain
runtime error) if every cell in the range is hidden, matching real
Excel's "No cells were found."

**Multi-area Paste: one shape now executes.** `do_paste` still diagnoses
(never writes cells for) a single-area source into a multi-area
destination, the reverse, or a count/shape mismatch — see the B7a
section's 4 classified failures. The one shape that now actually pastes:
both sides multi-area, matching `Areas.Count`, matching per-area shapes,
copied pairwise in source-to-destination order. `Transpose:=True` on that
shape still falls through to the diagnose-only path rather than silently
writing un-transposed data, and merged-cell conflict checking
(Milestone B6c2) is not applied to this path.

**`ThisWorkbook`/`ActiveWorkbook`/`ActiveSheet`**: `ActiveSheet` resolves
dynamically (at each access) to `Vm.active_sheet` and works anywhere a
`Worksheets(...)`/`Sheets(...)` reference does (`.Range(...)`,
`.Cells(...)`, `.Delete`, ...). `ThisWorkbook`/`ActiveWorkbook` need no
new resolution logic — elixcee only ever loads one workbook, so
`ThisWorkbook.Worksheets(x)`/`ActiveWorkbook.Worksheets(x)` parse as a
plain `Worksheets(x)`/`Sheets(x)`, the qualifier simply skipped. **Not
supported**: `Set ws = ActiveSheet` / `Set wb = ThisWorkbook` as
Worksheet/Workbook *object* variables — only the `Range`-producing object
expressions above are `Set`-able; a bare `ActiveSheet`/`ThisWorkbook` on a
`Set` RHS with no further `.Worksheets(...)`/property suffix falls
through to a harmless no-op, same as any other unmodeled `Set` target.

### Explicit non-goals (this milestone)

`Set`-ing `ActiveSheet`/`ThisWorkbook`/`ActiveWorkbook` themselves as
object variables (only `Range` object expressions are `Set`-able); a
multi-area source pasted into a single-area destination (or the reverse)
actually executing; `Transpose:=` on the multi-area execution path;
merged-cell conflict checking on the multi-area execution path;
`SpecialCells` types other than `xlCellTypeVisible`; and any new
`--json` field (this milestone changes one existing root cause's firing
condition — see the B7a section — but adds no new field or code).
