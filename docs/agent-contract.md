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
parser choked on. There's no "did you mean" suggestion field.

### Error codes

| Code | Kind | Stage | Meaning |
|---|---|---|---|
| `E1001` | `undefined_variable` | runtime | Macro referenced a variable that was never assigned |
| `E1002` | `undefined_sub_or_function` | runtime | Entrypoint macro name doesn't exist, or the macro called an unknown Sub/Function |
| `E1003` | `sheet_not_found` | runtime | Reserved for a `Sheets("X")` reference failing *during* macro execution (not currently reachable — today `Sheets("X")` auto-creates missing sheets) |
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
