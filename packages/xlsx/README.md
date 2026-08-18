# @elixcee/xlsx

A from-scratch, drop-in-compatible reimplementation of [`xlsx@0.18.5`](https://www.npmjs.com/package/xlsx)
(SheetJS) — no vendored SheetJS code, no runtime dependency on the real `xlsx` package.
Built on top of [`elixcee`](https://github.com/kent-tokyo/elixcee), a Rust VBA-macro
emulator, via a synchronous WebAssembly bridge.

**Not yet published to npm** (`0.0.0-development`). This README describes the current,
honest scope — not a finished drop-in replacement yet. See "What's not implemented" below
before depending on this for anything beyond the surface it actually covers.

## What's implemented

- **All 33 `utils.*` runtime exports** — `Object.keys(XLSX.utils)` matches the real `xlsx`
  package exactly, both content and insertion order. Differential-tested against the real
  `xlsx@0.18.5` package: 512 MATCH + 14 disclosed intentional divergences (six of them real
  security fixes for defects found in the real package itself — prototype pollution via a
  `"__proto__"` sheet/header name, XSS via unescaped `sheet_to_html` attributes and
  `javascript:` hyperlinks — see `THIRD_PARTY_NOTICES.md` and the root repo's
  `docs/compatibility-known-defects.md`).
- **`SSF` number formatting**, backed by a real `ssf@0.11.2` engine port: 1831/1831
  conformance against the real package on internal number-format tests.
- **`XLSX.read(data, opts)`** — synchronous, WebAssembly-backed, no `await init()` required.
  Works in Node (CJS and ESM) and in the browser (a `"browser"` export condition routes to
  an inlined-bytes/`initSync` WASM artifact — this assumes bundled consumption, since the
  shared code has a CJS `require('ssf')`; not literal no-build `<script type="module">`
  usage). Differential-tested: 19/19 MATCH on its declared scope — `SheetNames`, per-sheet
  `!ref`/`!merges`/`!rows`/`!cols`, and per-cell `{t, v, f, w, z}` (value, formula text,
  formatted display string, date-typed cells, resolved via real `styles.xml`/number-format
  parsing).
- Full TypeScript types, checked both with and without the DOM lib present (`table_to_sheet`/
  `table_to_book`/`sheet_add_dom` accept the oracle's own `data: any`, not `HTMLTableElement`,
  matching the real package's own type declarations rather than a stricter guess).

## What's not implemented

- **`read`/`readFile`** (the file-path/stream entry points) — only the buffer-first
  `XLSX.read(bytes)` form exists.
- **`write`/`writeFile`/`writeFileSync`** — no writer exists at all yet, for either XLSX or
  ODS output.
- Any `Rust ↔ JS` bridge beyond the read path described above.

Calling any of these today will fail (they're simply not exported), not silently
misbehave — this package never claims a capability it doesn't have.

## Compatibility scope

This package aims for **drop-in behavioral compatibility** with `xlsx@0.18.5` on the
surface it actually implements (see above) — not "inspired by" or "similar to." Where it
deliberately diverges (six security fixes, and a small number of DoS-shaped safety limits
like a capped cell-iteration count), that's disclosed, not silent — see
`docs/compatibility-known-defects.md` in the main repo for the full, itemized list.

## License

MIT, same as the parent `elixcee` project. `xlsx@0.18.5` (the compatibility target this
package is tested against) is Apache-2.0; no SheetJS code is vendored or copied into this
package, so no NOTICE obligation applies — see `THIRD_PARTY_NOTICES.md` for the full
attribution and the `ssf` dependency's own license.

## More

Architecture, the sync-WASM-bridge design, and the full compatibility-testing methodology
live in the parent repository's `docs/xlsx-architecture.md` and `compat/differential/`.
