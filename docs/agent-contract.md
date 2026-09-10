# elixcee CLI Agent Contract

Schema version: `1`, elixcee version **1.0.9**. See
[CHANGELOG](../CHANGELOG.md) for released versus Unreleased changes.
VBA/API coverage belongs in [FUNCTIONS](../FUNCTIONS.md), not this wire-format reference.

## Invocation

```text
elixcee <vba_file>... <MacroName> [--file <workbook>] [--sheet <name>] [--output <path>] [--json]
elixcee check <vba_file>... [--entry <MacroName>] [--json]
elixcee snapshot <workbook> [--json] [--max-work-units <N>] [--timeout-ms <N>] [--cancel-file <path>]
elixcee test-workbook <fixture.toml> [--json] [--seed <N>] [--case <N>]
elixcee diagnose <vba_file>... <MacroName> --file <workbook> [--json]
elixcee diagnose-workbook <fixture.toml> [--json] [--seed <N>] [--case <N>] [--cases <N>]
```

Run-mode's last positional argument is the macro name. Source is UTF-8 text;
exported `.cls` files are class modules, alongside standard `.bas` sources
(and plain `.vbs`/`.txt`). Workbook input is `.xlsx`, `.xlsm`, or `.ods`.
Loading a workbook does not extract and execute its embedded VBA automatically.

Run-mode `--file` uses the default reader budget and cooperative signal cancellation.
Only `snapshot` exposes the explicit work-budget, timeout, and cancel-file flags above.
These controls govern reading, not a general hard limit on macro execution.
See [limits](limits.md) for resource ceilings.

## stdout, stderr, and exit codes

For a valid `--json` invocation, stdout contains one result/error JSON object.
Plain-text output and stderr are not parsing contracts.
Exit status is `0` for success and `1` for reported failure.
Argument/usage errors can exit before JSON mode is established and may print only
stderr; consumers must handle a nonzero exit with no JSON.

All result shapes below use `schema_version: 1`. Consumers should ignore unknown
additional fields and must not parse human-readable messages to identify a category.

## Run result

```json
{
  "schema_version": 1,
  "ok": true,
  "entrypoint": "Main",
  "duration_ms": 0,
  "cells": [{"address": "A1", "value": 42}],
  "messages": []
}
```

- `duration_ms` measures macro execution, not the whole read/save process.
- `cells` is the active sheet only, sorted by row then column, with A1 addresses.
- Values are JSON numbers, booleans, strings, or null. Empty and VBA Null both
  serialize as null; dates use display strings, Excel errors their error strings.
  Arrays/records use `"[array]"`/`"[record]"`; nonfinite floats use
  `"NaN"`, `"Infinity"`, or `"-Infinity"`.
- `messages` records attempted MsgBox calls in order, including a call that fails
  under blocking policy. It is cleared per run and retained on execution failure.
  It does not imply a GUI was shown; Debug.Print is not this log.

## Error shape

```json
{
  "schema_version": 1,
  "ok": false,
  "error": {
    "code": "E1001",
    "kind": "undefined_variable",
    "message": "Variable 'missing' not found",
    "location": {"file": "Main.bas", "line": 2, "column": 1}
  },
  "messages": []
}
```

The example message is illustrative; exact prose is not stable.
Location is a 1-based character line/column when resolvable, otherwise null.
Runtime locations identify a statement, not necessarily its offending token.
Multi-module runtime errors currently lack per-module locations; parse diagnostics
can locate the offending source. I/O/setup errors have null locations.

| Code | Kind | Meaning |
|---|---|---|
| E1001 | undefined_variable | Unresolved variable |
| E1002 | undefined_sub_or_function | Unresolved procedure |
| E1003 | sheet_not_found | Missing selected sheet |
| E1004 | msgbox_blocked | Interactive call blocked |
| E1007 | object_variable_not_set | Unset/Nothing object reference |
| E1011 | security_blocked_external_effect | Blocked external VBA effect |
| E1099 | runtime_error | Other runtime failure |
| E2001 | parse_error | Invalid VBA source |
| E3001 | io_error | Source/workbook read or output failure |
| E3002 | setup_error | Invalid execution setup |

Runtime classification still recognizes selected error-message patterns.
Do not assume every rejected operation has its own typed error code.
Same-module and cross-module UDT name collisions are rejected before execution
and reported by `check`. A module-qualified UDT name such as `Types.Point` is
preserved by the parser and resolves when that module's definition is available;
same-name UDTs across modules remain rejected because bare-reference scope is
not yet modeled.

## Multi-module projects

Module names come from `Attribute VB_Name`, otherwise the file stem, and are
case-insensitive. Duplicate module names or duplicate bare standard-module
Sub/Function names are rejected, even if the caller uses qualification.
Class methods live in their instance namespace.

Multi-module runs accept bare or `Module.Sub` entrypoints. For a single source,
use its bare Sub name. Cross-module UDT-name collision handling and runtime source
locations remain limitations. Argument-count and label validation can reject a
program before its body executes; On Error cannot recover from such preflight errors.

## check: static analysis

`check` does not execute code. All positional arguments are source files;
an optional entrypoint is passed with `--entry`.

```json
{
  "schema_version": 1,
  "ok": true,
  "diagnostics": []
}
```

Each diagnostic has `code`, `kind`, `severity` (`error` or `info`),
`message`, and nullable `location`. Any error makes `ok` false; information
alone does not fail the check. Parse failure prevents analysis of that module.

| Code | Kind | Severity |
|---|---|---|
| E1001 | undefined_variable | error |
| E1002 | undefined_sub_or_function | error |
| E1005 | duplicate_sub_or_function | error |
| E1006 | duplicate_module_name | error |
| E1008 | argument_count_mismatch | error |
| E1009 | undefined_label | error |
| E1010 | blocked_external_effect | error |
| E1012 | duplicate_type | error |
| E2001 | parse_error | error |
| E3001 | io_error | error |
| I1001 | interactive_call | info |
| I1002 | unsupported_construct | info |

A clean check is not proof that execution succeeds. This is not complete type
inference or whole-program verification: dynamic object/member resolution,
array-versus-function syntax, and cross-module argument checks have limitations.
Unsupported constructs may be reported as information rather than rejected.
The `check` code E1010 does not promise the same code in run-mode.

## snapshot: workbook inspection

No VBA executes. Output includes every sheet, unlike run-mode's active-sheet result.

```json
{
  "schema_version": 1,
  "ok": true,
  "file": "Book1.xlsx",
  "sheets": [{
    "name": "Sheet1",
    "sheet_id": "1",
    "stable_id": "sheet1",
    "cells": [{"address": "A1", "value": 42}]
  }]
}
```

`sheet_id` is the file-format identifier as a string, or null when unavailable
(e.g. ODS). Current XLSX validation rejects missing/invalid required sheet IDs.
`stable_id` is `sheet{sheet_id}`, falling back to `sheet{1-based position}`.
It is not a VBA CodeName, globally unique identifier, or cross-writer identity
guarantee; the formatter does not deduplicate synthesized IDs.

Cells contain address and stored value, not formula text or formatting.
This CLI shape is distinct from the richer Python `Vm.snapshot()` API. Python
callers may opt into `include_dependencies=True` to receive bounded,
syntax-level cell/range dependency edges; parse failures are omitted from that
array and are reported in `dependency_diagnostics`. The same field reports
unresolved worksheet references; `has_formula_cycle` reports a detected formula
cycle. `input_candidates` and `output_candidates` are bounded navigation
heuristics: range inputs retain their range boundary and are never expanded.
These fields are diagnostics, not an Excel calculation oracle.
Read errors use the common error shape with `messages: []`.
Default Markdown output is for display, not lossless round-trip serialization.

## test-workbook: generated cases

Each case starts with a fresh VM and workbook read. The fixture uses a deliberately
small TOML subset: flat key/value pairs and `[[inputs]]`/`[[assertions]]`.
Inline tables, dotted keys, multiline strings, and trailing value junk are rejected.

```toml
name = "order calculation"
workbook = "fixtures/orders.xlsx"
vba_files = ["Main.bas"]
macro = "Process"
cases = 100
seed = 42
timeout_secs = 10

[[inputs]]
range = "Input!B2:B10"
strategy = "boundary_numeric"

[[assertions]]
range = "Result!A1:F100"
rule = "no_excel_errors"
```

Paths are relative to the fixture. The cases and seed fields are required;
timeout_secs is optional and defaults to 10 seconds.
Macro naming follows run-mode's single/multi-module rules.

| Strategy | Independently sampled values for each input cell |
|---|---|
| boundary_numeric | Empty, 0, 1, -1, 999999999, -999999999 |
| boundary_string | empty string, "test", 1,000 repetitions of "a" |

The runner always checks `no_panic`, `no_runtime_error`, and `no_timeout`.
The supported range assertion is `no_excel_errors`; a missing referenced sheet is
an error, not an empty successful assertion. Timeout is cooperative VM checking.

Success: `{"schema_version":1,"ok":true,"seed":42,"cases_run":100}`.

Failure has `schema_version`, `ok:false`, `seed`, zero-based `case_index`,
`inputs:[{"address":"Input!B2","value":0}]`, and `failure`.
For Excel errors, `failure` contains `rule`, `address`, and `actual`;
other failures contain `rule` and `message`. Setup failures use the common
error envelope instead.

`--seed` overrides the generator seed; `--case` replays one zero-based case.
Replay controls generated inputs, not VBA Rnd/RANDARRAY or external randomness.
There is no shrinking or exhaustive search guarantee.

## diagnose: strict-resolution execution

`--file` is required. Execution reports the first uncaught failure, with structural
evidence when classified. A successful run does not prove Excel equivalence.

- Success: `schema_version`, `ok:true`, `messages`.
- Execution failure: `schema_version`, `ok:false`, `message`, nullable `location`,
  `root_causes`, `messages`.
- Pre-execution I/O/parse/setup failures use the common error envelope.
- `root_causes` contains zero or one object. Unclassified failures produce an
  empty array, not fabricated evidence. Each object has `code`,
  `certainty:"definite"`, the fields below, and `suggestions: string[]`.
  Evidence fields are directly on that object, not nested under an `evidence` key.

| Root-cause code | Additional fields |
|---|---|
| WORKSHEET_NOT_FOUND / WORKBOOK_NOT_FOUND | expression, requested, available, suggested (string or null) |
| ARRAY_INDEX_OUT_OF_BOUNDS | name, index, lower, upper |
| PASTE_SHAPE_MISMATCH | source_addr, source_rows, source_cols, dest_addr, dest_rows, dest_cols, transpose, copy_location |
| PASTE_WITHOUT_COPY | dest_addr |
| SHEET_PROTECTED | sheet |
| PASTE_INTO_NON_ANCHOR_MERGED_CELL | dest_addr, dest_sheet, merged_range, copy_location |
| PASTE_PARTIAL_MERGED_RANGE | dest_addr, dest_sheet, conflicts, copy_location |
| PASTE_MERGE_LAYOUT_MISMATCH | source_addr, source_sheet, dest_addr, dest_sheet, conflicts, copy_location |
| MULTI_AREA_TO_SINGLE_AREA_PASTE | source_areas, destination_areas |
| MULTI_AREA_COUNT_MISMATCH | source_areas, destination_areas |
| MULTI_AREA_SHAPE_MISMATCH | area_index, source_area, destination_area |
| MULTI_AREA_PASTE_UNSUPPORTED | source_areas, destination_areas |

Array bounds are the actual declared bounds, including Option Base/explicit lower
bounds. Area objects use `address`, `rows`, `columns`; `area_index` is 1-based.
`conflicts` is an array of range-address strings. `copy_location` is a source
location or null, separate from the failing paste location.

Strict resolution rejects missing worksheets/workbooks instead of inventing them.
Numeric sheet lookup follows the VM's sheet order. Only the loaded workbook is
available; this is not multi-workbook host automation. On Error can handle ordinary
runtime failures, so a caught error need not appear as the final root cause.

### Copy, protection, and multi-area limits

Copy/Paste uses a value snapshot, not a full Excel clipboard carrying all styles
and formulas. A single-cell destination can expand; a single-cell source can fill
a destination. Unsupported shapes fail instead of silently truncating.
CutCopyMode=False clears the clipboard. Transpose is supported for the documented
single-area path; multi-area paste has narrower shape/order requirements.

Protection is a VM sheet flag: UserInterfaceOnly=True permits macro writes.
Passwords are accepted but are not authentication/encryption.
Merged-cell and paste checks apply in normal execution too, not only diagnose.
For Range/Union/Areas/SpecialCells and object-member coverage, use
[FUNCTIONS](../FUNCTIONS.md); unsupported combinations are not implied by this schema.

## diagnose-workbook: generated diagnosis

Uses the same fixture, seed, replay, and failure fields as test-workbook, with
strict resolution. `--cases` additionally overrides the case count.
Failures add `root_causes` using the same fields as diagnose; paste
`copy_location` is null because this runner does not resolve per-case source spans.
It stops at the first failing generated case and does not shrink failures.

## Hidden row/column evidence

Diagnose and diagnose-workbook add `observations` only when an observation exists;
they omit the field rather than emit an empty array. Observations may accompany
success or failure and are not root causes.

Each hidden-cell observation contains:

- `code:"RANGE_CONTAINS_HIDDEN_CELLS"`, `certainty:"observed"`, and `message`.
- `range:{sheet,address,rows,columns}`.
- `visibility:{hidden_rows,hidden_columns,total_cells,visible_cells}`, where hidden
  intervals are strings such as `"11:14"` and `"B:B"`.

Evidence describes the last surviving contiguous Copy source, not every range
touched during execution. Cleared/no-hidden/multi-area sources do not produce this
observation. Metadata evidence is not a simulation of Excel's whole filter UI.
Run-mode and test-workbook do not expose this observations field.

## Contract verification

[CLI integration tests](../tests/cli_json.rs), [black-box fixtures](../tests/fixtures/blackbox/),
and [diagnostic serialization](../src/diagnose.rs) are the executable references.
The contract does not imply complete Excel compatibility or enable external effects.
