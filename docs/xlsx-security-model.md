# XLSX security model

## Threat model

Spreadsheet files are untrusted input: email attachments, user uploads, files passed
between organizations. `@elixcee/xlsx` aims for behavioral compatibility with
`xlsx@0.18.5`, but **"same behavior as the oracle" must never mean "same vulnerabilities
as the oracle."** Where matching SheetJS would mean reproducing a resource-exhaustion or
object-injection vector, `@elixcee/xlsx` diverges deliberately and the divergence is
recorded, not hidden. See [`docs/xlsx-compatibility-goal.md`](xlsx-compatibility-goal.md)
for how this fits the overall compatibility definition.

## Existing limits (as of Phase 0)

| Limit | Value | Where |
|---|---|---|
| Per-ZIP-entry decompressed size | 64 MB | `ZIP_ENTRY_MAX_BYTES`, `src/reader.rs:214` |

That is the **only** resource limit in the current reader. Explicitly absent today:

- No cap on total decompressed size across all ZIP entries combined (many entries, each
  individually under 64 MB, is currently unbounded).
- No cap on ZIP entry count.
- No compression-ratio cap (classic zip-bomb detection).
- No cap on XML element count or attribute count per document.
- No cap on shared-string count/total length, cell count, merged-range count, sheet
  count, defined-name count, or formula-string length.
- No wall-clock/parse-time budget on the reader itself (a loop-execution deadline exists
  on `Vm`, but it only governs VBA execution *after* a workbook is already parsed).

XML nesting depth is a partial exception: `src/reader.rs`'s `XmlIter` is a flat,
non-recursive pull parser (no DOM tree, no recursive descent), so pathological nesting
depth cannot cause a Rust stack overflow the way a recursive-descent or DOM-building
parser could. It can still cost unbounded time/memory via element/attribute count, which
is why those are listed as planned limits below, not dismissed as already covered.

## `packages/xlsx` (JS) limits — distinct from the Rust reader above

The table above is specific to `src/reader.rs` (untrusted ZIP/XML file parsing).
`packages/xlsx` is a separate subsystem (in-memory JS worksheet-object manipulation, no
file I/O yet) with its own, much smaller limit set:

| Limit | Value | Where | Registered as |
|---|---|---|---|
| `!ref` rectangle cell count (`sheet_to_formulae`/`sheet_to_csv`/`sheet_to_txt`) | 5,000,000 cells | `packages/xlsx/src/internal/range-guard.cjs` | `ELIXCEE_RANGE_TOO_LARGE`, `compat/differential/classify.mjs`'s `SAFETY_DIVERGENCE_REGISTRY` |
| Non-finite column/row index (`encode_col`) | rejects `+Infinity` | `packages/xlsx/src/index.cjs` | `ELIXCEE_NON_FINITE_INDEX`, same registry |

Both were added only after empirically confirming the real oracle actually hangs/loops
on the corresponding input (a timeout-guarded subprocess run, not a speculative guard) —
per this project's standing rule against adding resource limits without measurement. See
[`docs/limits.md`](limits.md) for the time/RSS measurement behind the 5,000,000-cell
threshold specifically.

## Planned limits (not yet implemented — design targets for the phase that builds the
compat-hardened reader)

| Limit | Rationale | Where it would live |
|---|---|---|
| Compression-ratio cap (per entry and/or overall) | Classic zip-bomb: a tiny compressed entry expanding to gigabytes | ZIP entry read loop, `src/reader.rs` |
| Total decompressed-size budget across all entries | Many entries each under the per-entry cap can still sum to unbounded memory | Same read loop, tracked across the whole archive |
| ZIP entry-count cap | Bounds enumeration/metadata cost regardless of per-entry size | Archive open, before iterating entries |
| XML element-count cap per document | Bounds parse time/memory for pathologically tag-dense XML | `XmlIter` consumer (per-document counter) |
| XML attribute-count cap per element | Same class of concern, per-element | `parse_attrs`, `src/reader.rs` |
| Attribute-value / text-node length caps | Bounds a single absurdly long value from being accepted whole | `parse_attrs` / text accumulation in `xlsx_sheet_cells` etc. |
| Shared-string count and total-character budget | `sharedStrings.xml` is a common bomb vector (many/huge strings referenced repeatedly) | `xlsx_shared_strings`, `src/reader.rs` |
| Sheet / row / column / cell / merge / defined-name count caps | Bounds the size of the materialized workbook model itself | Sheet-cell parse loop, `src/reader.rs` |
| A single overall "work budget" for a read | Backstop against limit combinations that individually pass but compound | Top of `read_workbook`/its buffer-based successor |

Exact numeric values for every row above are an open item — deferred to the phase that
implements them, informed by real-world file surveys, not chosen arbitrarily here.

## Prototype-pollution-safe key handling

Excel data can legitimately contain the strings `__proto__`, `constructor`, or
`prototype` — as a header cell, a sheet name, or any other value that becomes a
JavaScript object key in the SheetJS-compatible surface. `@elixcee/xlsx` must preserve
that data (it is normal, if unusual, spreadsheet content) **without ever mutating an
Object's prototype.** The specific call sites where a spreadsheet-derived string becomes
an object key (net-new work for whichever phase implements them — none of this exists
yet):

- `utils.sheet_to_json`'s header-row-derived keys become the property names of every
  emitted row object.
- Sheet-name-keyed access — `workbook.Sheets[name]` and `utils.book_append_sheet`'s
  internal sheet-name map.

Any code that builds one of these key-indexed structures must use `Object.create(null)`
or a `Map`, never a plain `{}` object literal, so that a crafted `"__proto__"` key is
stored as ordinary data instead of reaching `Object.prototype`.

## Intentional non-compatibility policy

When matching the oracle's behavior on a given input would mean reproducing a resource-
exhaustion vector (zip bomb, XML/entity blowup, unbounded string/entry counts) or an
object-injection vector (prototype pollution), `@elixcee/xlsx` diverges on purpose. This
divergence is not a bug and not something to hide: the differential-testing harness
classifies it explicitly as `INTENTIONAL_SECURITY_DIVERGENCE` (defined in
[`compat/differential/classify.mjs`](../compat/differential/classify.mjs)) rather than
folding it into `MATCH` or silently omitting it from a compatibility report.

## The oracle itself is a validating example

`npm audit` against the installed `compat/` devDependencies (run during Phase 0) reports
that `xlsx@0.18.5` — the exact version pinned as the compatibility oracle, per the
project's own instruction to target npm's widely-used `0.18.5` rather than the latest
SheetJS release — carries two known high-severity advisories: Prototype Pollution
([GHSA-4r6h-8v6p-xvw6](https://github.com/advisories/GHSA-4r6h-8v6p-xvw6), fixed in
0.19.3) and a ReDoS
([GHSA-5pgg-2g8v-p4x9](https://github.com/advisories/GHSA-5pgg-2g8v-p4x9), fixed in
0.20.2). This is expected, not a Phase 0 problem to fix: `xlsx` here is a `devDependency`
used only to drive the oracle/differential harness, never shipped to an `@elixcee/xlsx`
consumer. It is, however, direct, concrete confirmation of why this document's
"intentional non-compatibility policy" exists — the oracle we are matching behavior
against is a real, currently-vulnerable version of a real package, not a hypothetical.

## Open items

- Exact numeric values for every row in the "planned limits" table.
- Whether some limits should be user-configurable (an options field) vs. fixed
  constants — SheetJS itself has no such options, so any configurability here is new
  surface area that needs its own compatibility reasoning, not an automatic yes.
