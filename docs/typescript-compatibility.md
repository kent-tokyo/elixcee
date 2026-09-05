# TypeScript surface compatibility

Reference target: `xlsx@0.18.5`. This is a declaration-level comparison, not
proof of runtime equivalence or npm publication.
The private package's [README](../packages/xlsx/README.md) defines runtime scope;
[src/index.d.ts](../packages/xlsx/src/index.d.ts) is the current declaration source.

## Classification

- **EXACT**: same covered call shapes, option fields, and return types as the reference.
- **SAFE_EXTENSION**: accepts additional valid shapes or declares a runtime export
  missing from the reference types; does not narrow the reference's accepted calls.
- **INCOMPATIBLE**: a known declaration mismatch. The consts spelling below is deliberate.
- **MISSING**: no corresponding declaration. Do not count missing reference-only
  declarations as absent runtime exports.

The table groups signatures; its row count is not a count of individual APIs.
Earlier utils comparisons do not establish exactness for all top-level read/write types.

## Utility declarations

| Declaration | Classification | Note |
|---|---|---|
| encode/decode col, row, cell, range | EXACT | Address helpers |
| split_cell | SAFE_EXTENSION | Runtime export absent from reference declarations |
| book_new / book_append_sheet / book_set_sheet_visibility | EXACT | Workbook utilities |
| aoa_to_sheet / sheet_add_aoa / json_to_sheet / sheet_add_json | EXACT, except dense below | Reference option shapes |
| AOA2SheetOpts.dense | SAFE_EXTENSION | Supported runtime option missing from reference types |
| format_cell / cell_set_number_format | EXACT | Formatting |
| sheet_to_formulae / sheet_to_csv / sheet_to_txt | EXACT | Export options |
| cell_set_hyperlink / cell_set_internal_link / cell_add_comment / sheet_set_array_formula | EXACT | Cell utilities |
| sheet_to_json | EXACT | Three overloads, including the generic result |
| sheet_get_cell | SAFE_EXTENSION | Runtime export absent from reference declarations |
| sheet_to_row_object_array | SAFE_EXTENSION | Runtime alias of sheet_to_json, absent from reference types |
| consts | INCOMPATIBLE | Types use runtime's SHEET_VERY_HIDDEN; reference types spell SHEET_VERYHIDDEN |
| sheet_to_html | SAFE_EXTENSION | Adds rawHtml?: boolean for explicitly trusted markup |
| sheet_add_dom / table_to_sheet / table_to_book | EXACT | data: any, not a narrower HTMLTableElement |

The reference declares `sheet_to_dif`, `sheet_to_slk`, and `sheet_to_eth` but
does not expose them in its 33 runtime utils keys. They are **MISSING declarations**,
not three missing members of that runtime set; they remain outside the target.

Top-level read/readFile/write/writeFile and synchronous aliases have declarations.
Their presence is not an EXACT classification. Streaming and writeFileAsync remain
unimplemented. See the package README for Node/browser restrictions.

## Checks

Run `npm run typecheck` and `npm run typecheck:no-dom` in `packages/xlsx`.
The first includes the DOM smoke consumer; the second ensures that importing the
declarations does not require a DOM library. Neither substitutes for runtime tests.
