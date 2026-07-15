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
source file. A single file (today's only shape until Milestone B2) parses
exactly as before. With more than one file, see "Multi-module projects"
below for how `MacroName` and cross-module calls resolve.

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
contract (see `tasks/todo.md` for why this was a deliberate choice, and
where to raise it if that changes).

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
| `E1099` | `runtime_error` | runtime | Any other runtime failure not covered above |
| `E2001` | `parse_error` | parse | The VBA source failed to parse |
| `E3001` | `io_error` | io | Reading the VBA file, reading `--file`, or writing `--output` failed |
| `E3002` | `sheet_setup_error` | setup | Resolving which sheet to use failed *before* the macro ran — a workbook with no sheets, or `--sheet <name>` naming a sheet that doesn't exist in `--file` |

Classification is currently done by pattern-matching the existing
`Result<_, String>` error text at the CLI boundary (`src/diagnostics.rs`),
not by a typed error enum in the VM. This is a known, deliberate interim
trade-off — see the "runtime error 分類を型付きエラーへ" item in
`tasks/todo.md` for the plan to harden it before adding more error kinds.

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
  `tasks/todo.md`.
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
  work in `tasks/todo.md`'s Milestone A.5, not yet revisited). Single-file
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
very rare case — see `tasks/todo.md`'s Milestone B1.1 entry for the
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

## `diagnose` subcommand (Excel operation diagnostics, Milestones B6a/B6b/B6c)

```
elixcee diagnose <vba_file>... <MacroName> --file <workbook> [--json]
```

Runs one macro once and classifies *why* it failed — a missing worksheet,
a missing workbook, or an out-of-bounds array index — with concrete
evidence (the requested key, what was actually available, a "did you mean"
suggestion), instead of only a bare runtime-error string. This is a
different posture from `run`/`check`/`test-workbook`: it turns on
`Vm::strict_resolution`, which makes elixcee's usual auto-vivify/silent-
`Empty` convenience for `Sheets("X")`/`Worksheets("X")` references into a
hard, classified failure — because a diagnostic tool whose whole purpose is
"what would Excel actually reject here" needs to *not* paper over the exact
class of mistake it exists to catch. Every other subcommand leaves
`strict_resolution` off (the default) and is completely unaffected.

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
every mode** (`run`/`check`/`diagnose` alike) — not gated behind
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
every mode** (`run`/`check`/`diagnose`) — writes (`Cells`/`Range.Value`/
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
carry just `sheet`, e.g.:

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

Exit code 0/`ok:true` on success, 1/`ok:false` on failure — same
convention as every other subcommand. `location` follows the same
single-module-only rule as run-mode's own `--json` contract (a
`SourceSpan` carries no module id, so a multi-module run reports
`location: null` rather than risk pointing at the wrong module's source).

### Explicit non-goals (deferred to later B6 phases)

B6a covers resolution failures (missing worksheet/workbook, array out of
bounds); B6b covers Copy/Paste shape mismatch and clipboard state (see its
own non-goals list above); B6c covers sheet protection (see its own
non-goals note above). Explicitly out of scope, planned for later:

- Merged cells, multi-area (`Areas`) ranges, hidden/filtered rows — the
  user's original roadmap bundled these with sheet protection under
  "B6c," but they were split into a future continuation once grounding
  showed each needs new reader-format parsing (XLSX/ODS) and/or a
  range-model change (a single rectangle → a list of areas) that
  protection alone doesn't need.
- Excel Tables (`ListObjects`) — never part of the user's original
  roadmap; added as a placeholder non-goal during B6a's own docs and kept
  deferred (a full new VBA object model, comparable in scope to `Range`/
  `Sheets` itself, not "add diagnosis to an existing path").
- Integration with `test-workbook`'s case generator for counterexample
  search (B6d) — `diagnose` runs a macro exactly once today.
- A real VBA `Collection` object — it doesn't exist in elixcee at all, so
  there is nothing to classify a failure for; adding one is a first-class
  feature, not "add diagnosis to an existing path."
- Real multi-workbook execution — only a name/index mismatch check against
  the single loaded workbook ships in this milestone.
- `Dim arr(1 To N)` non-zero-lower-bound tracking.
