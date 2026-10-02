# Compatibility-known defects

A record of compatibility decisions against the pinned SheetJS `xlsx@0.20.3` oracle,
not necessarily the latest upstream release. Harmless quirks may be reproduced;
security-related differences are explicitly classified and must not be silently changed.

Contrast with [`docs/xlsx-security-model.md`](xlsx-security-model.md)'s intentional
*divergences* — cases where elixcee deliberately does NOT match the oracle.
Most entries below preserve harmless quirks; the final raw-HTML entry instead
documents a safe default with an explicit trusted-markup opt-in.

---

```yaml
compatibility-known-defect:
  api: sheet_to_json
  case: "default-inferred header text is \"__proto__\", \"constructor\", \"toString\", or \"hasOwnProperty\""
  oracle_behavior: silently renamed to "<text>_NaN"
  elixcee_behavior: reproduced for compatibility
```

`sheet_to_json`'s default header-inference loop de-dupes collisions via a plain
`header_cnt = {}` counter object, reading `header_cnt[v] || 0` and later writing
`header_cnt[v] = counter`. For `v === "__proto__"`, the READ triggers
`Object.prototype`'s inherited accessor (returning the actual — truthy — prototype
object), which the loop's duplicate-detection logic treats as "already seen," so it
appends a `"_" + counter` suffix; `counter++` on that non-numeric accessor value coerces
to `NaN`, producing the literal header text `"__proto___NaN"` (confirmed live). The same
happens for any OTHER string that is itself a truthy inherited `Object.prototype`
property read through plain bracket access — `"constructor"` (the `Object` function),
`"toString"`/`"hasOwnProperty"` (both functions) → `"constructor_NaN"`,
`"toString_NaN"`, `"hasOwnProperty_NaN"`, all confirmed live. `"prototype"` is the one
exception among the five poisoned keys tested: plain object instances (unlike function
objects) have no inherited `.prototype` property at all, so `header_cnt["prototype"]`
reads `undefined` (falsy) and never collides — the header stays the literal
`"prototype"`, also confirmed live. This is an accidental side effect of the collision
counter, not
intentional protection, and it is NOT a security hazard on its own: the corresponding
`header_cnt[v] = counter` write assigns `NaN` through the `__proto__` setter, which is a
spec no-op for non-object values, so `header_cnt`'s own prototype is never actually
touched. `packages/xlsx` reproduces the exact renamed text (using a plain `{}` for its own
`header_cnt`, not `Object.create(null)`) rather than "fixing" the collision logic to skip
poisoned keys specially. The one genuine hazard — an EXPLICIT `opts.header` array
containing the literal `"__proto__"` — is a different code path entirely and IS fixed;
see `docs/xlsx-security-model.md`'s "Prototype-pollution-safe key handling" section and
`compat/differential/xlsx-utils.test.mjs`'s `"default header text = ..."` fixtures.

---

```yaml
compatibility-known-defect:
  api: sheet_to_html
  case: "cell.h present (raw HTML rich-text rendering)"
  oracle_behavior: used verbatim, zero escaping
  elixcee_behavior: escaped by default; rawHtml:true opts into the oracle-compatible passthrough
```

The oracle's `make_html_row` uses `(cell.h || escapeHtmlText(...))` — when `cell.h` is
present, it is used **completely as-is**, with no HTML escaping at all. elixcee now escapes
`cell.h` by default and requires the explicit `rawHtml: true` option for passthrough.
Unlike `sheet_to_html`'s attribute-building
(`data-t`/`data-v`/`data-z`/`id`, both table-level and per-cell) or `cell.l.Target`'s
`href` construction — both **genuine bugs** with no intended purpose, fixed in
`packages/xlsx` (see `docs/xlsx-security-model.md`) and registered in
`compat/differential/classify.mjs`'s `SECURITY_DIVERGENCE_REGISTRY` — `cell.h` is a
**documented, intentional** field: a pre-rendered HTML representation of a cell's rich
text (e.g. `<b>bold</b>` for a bold run), meant to be inserted as-is. The opt-in preserves
that feature for callers that have independently established the markup is trusted.

`sheet_to_html` also rejects hyperlink targets with leading/trailing whitespace, ASCII control
characters, or backslashes to avoid browser URL-normalization ambiguity.

`packages/xlsx` now includes a file reader. Regardless of how a cell object was
obtained, callers must independently trust its markup before enabling `rawHtml`.
The default escaped path should be used for attacker-controlled `.h` content. See
`compat/differential/xlsx-utils.test.mjs`'s
`sheet_to_html` fixtures for the differential coverage of this exact behavior.
