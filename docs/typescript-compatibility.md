# TypeScript type-surface compatibility

`packages/xlsx/src/index.d.ts` is hand-written against `xlsx@0.18.5`'s real
`types/index.d.ts`, not independently designed. This file tracks, per declaration, how
each one relates to the oracle's own types — the classification a completion report's
"TypeScript surface" section must cite by name, not describe loosely.

## Classification

- **EXACT** — the declaration mirrors an oracle declaration: same accepted/rejected call
  shapes, same option fields, same return type shape. TypeScript code that compiles
  against the oracle's types compiles here, and vice versa (for the covered signature).
- **SAFE_EXTENSION** — elixcee's declaration accepts a superset of what the oracle's own
  types accept, or covers a real runtime export the oracle's types omit entirely. Never
  rejects anything the oracle's types would accept (the standing rule from Phase 1B-3's
  review), so it's compatible — but it is not a 1:1 mirror either, so it must not be
  counted toward "N of M signatures are EXACT."
- **MISSING** — the oracle exports something (value or type) that `packages/xlsx` has no
  declaration for at all yet.
- **INCOMPATIBLE** — elixcee's declaration rejects something the oracle's own types would
  accept, or accepts something the oracle's types reject in a way that changes which
  real-world consumer code compiles. Ordinarily a bug, since it violates the "never tighten
  types" rule Phase 1B-3 established — with one deliberate exception below (`consts`),
  where mirroring the oracle's own (buggy) types would make elixcee's types misdescribe
  elixcee's own runtime, which is worse than the gap it would close. Every other entry in
  this table is EXACT or SAFE_EXTENSION; `consts` is the only INCOMPATIBLE one, and it's
  intentional, not an oversight.

A completion report's "TypeScript surface" count (e.g. "N EXACT, M SAFE_EXTENSION, K
MISSING") must be built by walking this table, never estimated.

## Entries

| Declaration | Classification | Why |
|---|---|---|
| `encode_col`/`decode_col`/`encode_row`/`decode_row`/`encode_cell`/`decode_cell`/`encode_range`/`decode_range` | EXACT | Signatures mirror the oracle's `types/index.d.ts` directly (Phase 1A). |
| `split_cell` | SAFE_EXTENSION | Not present in the oracle's `types/index.d.ts` at all (confirmed: no `split_cell` entry anywhere in the file, the same gap class as `sheet_get_cell`) despite being a real runtime export. Pure addition — reclassified from an earlier draft of this table that had grouped it under EXACT by mistake. |
| `book_new`/`book_append_sheet`/`book_set_sheet_visibility` | EXACT | Same. |
| `aoa_to_sheet`/`sheet_add_aoa`/`json_to_sheet`/`sheet_add_json` | EXACT | Option interfaces (`AOA2SheetOpts`/`SheetAOAOpts`/`JSON2SheetOpts`/`SheetJSONOpts`) mirror the oracle's own field names/types — **except** `AOA2SheetOpts.dense`, see below. |
| `AOA2SheetOpts.dense` | SAFE_EXTENSION | The oracle's own `AOA2SheetOpts` (`types/index.d.ts`) has no `dense` field at all — confirmed by compiling an oracle-typed probe with `{dense: true}` against the real `.d.ts`, which fails with "Object literal may only specify known properties." elixcee's own `AOA2SheetOpts` includes it (added Phase 1B-1, predating this classification scheme) since `aoa_to_sheet`/`sheet_add_aoa` genuinely accept the option at runtime on both sides. Accepting more than the oracle's types isn't a violation of "never reject what the oracle accepts" — it's a strict superset — but it doesn't mirror the oracle either, so it stays SAFE_EXTENSION, not EXACT. |
| `format_cell`/`cell_set_number_format` | EXACT | |
| `sheet_to_formulae`/`sheet_to_csv`/`sheet_to_txt` | EXACT | `Sheet2CSVOpts`/`Sheet2TXTOpts` mirror the oracle's own option fields. |
| `cell_set_hyperlink`/`cell_set_internal_link`/`cell_add_comment`/`sheet_set_array_formula` | EXACT | |
| `sheet_to_json` | EXACT | `Sheet2JSONOpts` and the 3-overload set (`<T>(): T[]`, `(): any[][]`, `(): any[]`) mirror `types/index.d.ts` verbatim, including the two largely-unreachable non-generic overloads (Phase 1B-3). |
| `sheet_get_cell` | SAFE_EXTENSION | Not present in the oracle's `types/index.d.ts` at all — confirmed absent (no `get_cell` entry anywhere in the file) despite being a real runtime export (`sheet_get_cell: ws_get_cell_stub` in the oracle's own source). Pure addition: there is no oracle declaration to mirror or diverge from, so it cannot be EXACT by definition, but it also rejects nothing the oracle accepts (the oracle's types simply don't speak to this function at all). Added Phase 1B-3. |
| `sheet_to_row_object_array` | SAFE_EXTENSION | Not declared anywhere in the oracle's `types/index.d.ts` (confirmed absent) despite being a real runtime export — a literal alias for `sheet_to_json` (`U.sheet_to_row_object_array === U.sheet_to_json` is `true`). Typed identically to `sheet_to_json`'s own overload set. Added Phase 1C. |
| `consts` | INCOMPATIBLE (deliberate) | The oracle's own types declare `SHEET_VERYHIDDEN` (no underscore), but its real RUNTIME object's own key is `SHEET_VERY_HIDDEN` (underscore) — confirmed live via `Object.getOwnPropertyDescriptor`, a genuine mismatch already present in the oracle between its shipped types and its shipped runtime. `packages/xlsx` types `SHEET_VERY_HIDDEN`, matching its own (and the oracle's) actual runtime — typing `SHEET_VERYHIDDEN` instead would type-check but read `undefined` at runtime, which is worse than the gap it would close. The one deliberate INCOMPATIBLE entry in this table: chosen over EXACT specifically because mirroring the oracle's admittedly-buggy types here would make elixcee's own types describe elixcee's own runtime incorrectly. Added Phase 1C. |
| `sheet_to_html` | EXACT | `Sheet2HTMLOpts` mirrors the oracle's own field names/types (`id`/`editable`/`header`/`footer` — the `header`/`footer` fields share a name with `Sheet2JSONOpts`'s own `header` but mean something unrelated, matching the oracle's own two separate, non-overlapping interfaces). Added Phase 1C. |
| `sheet_add_dom`/`table_to_sheet`/`table_to_book` | EXACT | `Table2SheetOpts` mirrors the oracle's own field names/types field-for-field, and — the specific point Phase 1B-3's review would have flagged had it been gotten wrong — the `data` parameter mirrors the oracle's own `data: any` exactly, not narrowed to `HTMLTableElement`. Narrowing would REJECT code the oracle's own types accept (a plain duck-typed object, or a real DOM element in a non-DOM-lib TypeScript project where `any` still flows through) — exactly the kind of tightening the review prohibited. Compile-tested two ways: `tsconfig.no-dom.json` (`npm run typecheck:no-dom`) proves `src/index.d.ts` alone needs no DOM lib; `test/smoke-dom.ts` (under the default `tsconfig.json`, whose implicit lib already includes DOM — TypeScript's default lib inference pulls in DOM unless explicitly restricted, confirmed by inspecting this project's own `tsconfig.json`, which has no `lib` override) proves a real `HTMLTableElement` is still accepted. Added Phase 1C. |

## Not yet covered (MISSING as of this writing)

- `sheet_to_dif`, `sheet_to_slk`, `sheet_to_eth` — declared in the oracle's types
  (`Sheet2HTMLOpts`-typed) but **not present in the real runtime `Object.keys(XLSX.utils)`
  at all** (confirmed: `xlsx@0.18.5`'s actual build doesn't ship these three despite the
  types file declaring them — a types-ahead-of-runtime gap in the oracle itself, the
  mirror image of the `consts`/`SHEET_VERY_HIDDEN` case above). Not a target for any
  current phase: there is no real runtime function to differential-test against, so
  implementing these would mean inventing behavior with no oracle to verify against — out
  of scope until/unless a future oracle version actually ships them. These are the ONLY
  three entries in the oracle's own `Object.keys(XLSX.utils)` (33 total) that
  `packages/xlsx` does not implement, and they're excluded for the reason above, not an
  oversight — `packages/xlsx` implements all 33 real runtime keys as of Phase 1C.
- Anything under top-level `read`/`readFile`/`write*`/`stream` — see
  `docs/xlsx-architecture.md` for phase sequencing; not part of `utils` at all.
