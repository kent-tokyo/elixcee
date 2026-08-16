# Compatibility-known defects

A running log of oracle (`xlsx@0.18.5`) behaviors that look like bugs but are
deliberately reproduced anyway, because compatibility with the real, currently-shipping
package takes priority over "fixing" something on its behalf. Each entry is a compat
decision on record, not something to silently normalize or reject later without updating
this file and its differential test coverage.

Contrast with [`docs/xlsx-security-model.md`](xlsx-security-model.md)'s intentional
*divergences* — those are cases where elixcee deliberately does NOT match the oracle
(because matching would mean replicating a DoS/injection vector). The entries below are
the opposite: elixcee DOES match the oracle, even though the oracle's behavior is itself
questionable, because there's no security reason not to.

---

```yaml
compatibility-known-defect:
  api: book_append_sheet
  case: colon in sheet name
  oracle_behavior: accepted
  excel_validity: invalid or application-dependent
  elixcee_behavior: reproduced for compatibility
```

`check_ws_name`'s thrown error message reads `"Sheet name cannot contain : \ / ? * [ ]"`,
listing `:` among the forbidden characters — but the actual character check
(`badchars = "][*?/\\".split("")`, `xlsx.js`) never includes `:`. A sheet named
`"Sheet:1"` is accepted without error. This is a genuine mismatch between the oracle's
error message and its real behavior, not a documentation choice. `packages/xlsx`
reproduces both: the same accept-with-colon behavior AND the same (technically
inaccurate) error message text, since real-world code may already depend on either. See
`compat/differential/xlsx-utils.test.mjs`'s `book_append_sheet` scenarios for the
differential coverage (`"Sheet:1"` is one of the tested special-character names).

**Applies to future write-path work too**: when `packages/xlsx` eventually implements
XLSX writing, do not add validation or normalization for colon-containing sheet names
beyond what the oracle itself does — differential-test whatever the oracle actually
writes to the ZIP for such a name, rather than assuming Excel's own (stricter, and
UI-context-dependent) rules apply.

---

```yaml
compatibility-known-defect:
  api: json_to_sheet / sheet_add_json
  case: "opts.dense with no existing _ws"
  oracle_behavior: silently ignored
  elixcee_behavior: reproduced for compatibility
```

`sheet_add_json`'s source (unlike `sheet_add_aoa`'s) never reads `opts.dense` at all — it
always creates `ws = _ws || ({})`, a plain object, regardless of the option. Confirmed
live: `XLSX.utils.json_to_sheet(data, {dense:true})` returns an ordinary sparse
(cell-ref-keyed) worksheet, not a dense array. `packages/xlsx` reproduces this exactly
(no `dense`-option handling in `sheetAddJson`/`jsonToSheet`) rather than "fixing" it to
honor the option the way `aoa_to_sheet` does. See
`compat/differential/xlsx-utils.test.mjs`'s `"opts.dense has no effect when _ws is null"`
fixture.

---

```yaml
compatibility-known-defect:
  api: sheet_add_json
  case: "_ws is an existing dense (array) worksheet"
  oracle_behavior: header row and object-typed values leak as stray string-keyed
    properties on the array instead of landing in the nested rows
  elixcee_behavior: reproduced for compatibility
```

When `sheet_add_json` IS given an existing dense array as `_ws`, only plain scalar
values (numbers/strings/booleans/Dates) are written correctly into the nested
`ws[row][col]` cells (via the internal `ws_get_cell_stub`, ported as `wsGetCellStub`).
The header row is written via a direct `ws[colLetter + rowNumber] = {...}` string-keyed
assignment — confirmed live: `sheet_add_json([], [{a:1,b:'x'}])` leaves `ws[0]` as
`null` while the header text is only reachable via the stray properties `ws.A1`/`ws.B1`.
Object-typed JSON values (e.g. a caller-supplied full cell object) hit the same
string-keyed-assignment path (`ws[ref] = v`) regardless of dense/sparse mode, so they
never actually reach the dense array's nested cell either — the stub created for that
slot stays `{t:'z'}`. `packages/xlsx` reproduces this exactly, including which specific
cases are affected (scalars work, headers and object values don't), rather than
"fixing" `sheet_add_json` to be dense-mode-consistent throughout. See
`compat/differential/xlsx-utils.test.mjs`'s `"dense target: scalar values land in the
nested array; header/object values leak as stray string-keyed props"` fixture.
