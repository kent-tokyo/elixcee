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

- **`XLSX.readFile(path, opts)` / `XLSX.readFileSync(path, opts)`** — Node only, a thin
  wrapper over `read()`. One function under both names, matching the real package
  (`XLSX.readFile === XLSX.readFileSync`, and both report `.name === 'readFileSync'`). A
  missing file throws Node's own `ENOENT`, unwrapped, exactly as the real package does.
  Differential-tested file-by-file against every real `.xlsx` fixture in this repo, with and
  without `cellStyles`/`cellDates`. In the **browser** build both names exist but throw
  `ELIXCEE_UNSUPPORTED_IN_BROWSER` — a browser has no filesystem, and an explicit throw is
  more useful than a missing export or a faked result; read the bytes yourself
  (`FileReader`/`fetch`) and call `read(bytes)`.

- **`XLSX.write(wb, opts)` / `XLSX.writeFile(wb, path, opts)` / `XLSX.writeFileSync(...)`**
  — `bookType: "xlsx"` only (ODS/CSV/other book types are not implemented — an explicit
  `ELIXCEE_UNSUPPORTED_BOOK_TYPE` throw, never silently ignored), producing a real OOXML
  ZIP archive built by a hand-rolled ZIP/DEFLATE writer (no new npm dependency; Node's own
  `zlib` provides the DEFLATE codec). Supports strings/numbers/booleans/dates/formulas,
  multiple worksheets, merges, sheet visibility, hidden rows/columns, and basic number
  formats. Output `type: "buffer" | "array" | "base64"`. `writeFile`/`writeFileSync` are
  Node-only (one function under both names, matching the real package); the browser build
  throws `ELIXCEE_UNSUPPORTED_IN_BROWSER` for both, same rationale as `readFile`/
  `readFileSync` — call `write(wb, {type: "buffer"|"array"|"base64"})` and hand the bytes
  to a download/File-System-Access-API flow yourself. Round-trip tested both directions
  against the real `xlsx@0.18.5` package (own write → oracle read, oracle write → own
  read) and against itself (own write → own read).

- `utils.sheet_to_html` escapes caller-controlled `cell.h` markup by default. Use the
  explicit `rawHtml: true` option only when the rich-text HTML has been independently
  trusted; this is an intentional security extension beyond the oracle's default.

## What's not implemented

- **ODS output**, and any `bookType` other than `"xlsx"`.
- **`writeFileAsync`** — the real package's async file-write variant; only the synchronous
  `writeFile`/`writeFileSync` pair is exported.
- **Streaming reads** — only whole-buffer/whole-file input.
- Any `Rust ↔ JS` bridge beyond the read path described above.

Calling any of these today will fail (they're simply not exported), not silently
misbehave — this package never claims a capability it doesn't have.

## Bundling

Both shipped WASM loaders inline the compiled `.wasm` bytes as base64, so bundling this
package needs **no asset-loader config, no `.wasm` copy step, and no bundler externals** —
`esbuild --bundle` to either CJS or ESM output just works, verified end-to-end by
`scripts/wasm-smoke.mjs`. (Before this, the Node loader located its `.wasm` file
`__dirname`-relative, which is bundle-*output*-relative once bundled: CJS output required
the consumer to copy the file next to their bundle, and ESM output failed outright with
`__dirname is not defined`. Both are fixed at the source.)

Bundling for the browser (the `"browser"` export condition) additionally stubs out the Node
loader via package.json's `browser` field, so a browser bundle carries the WASM payload once
rather than twice. Verified in an actual headless Chrome process, not only under Node's
`--conditions=browser` — see `scripts/browser-smoke.mjs`. Chrome/Chromium is the only browser
covered; Safari is not tested and not claimed.

The cost of inlining, stated rather than glossed over: the packed tarball grew from 339,098
to 380,005 bytes (+12.1%; unpacked 741,304 -> 835,712, +12.7%) for the same 263,204-byte WASM
binary.

**`write`/`readFile`/`readFileSync` are a different story** — they reach a lazy
`require('zlib')`/`require('fs')` at call time (not at import time, so a caller who only
uses `read()` never pays for it). That's fine in Node directly and in an esbuild **CJS**
bundle, but an esbuild **ESM** bundle can never synchronously `require()` anything reached
through CJS-origin code, regardless of how that `require()` is phrased or whether its
target is marked `--external` — confirmed live, not assumed: it throws `Dynamic require of
"..." is not supported` the moment `write()`/`readFile()`/`readFileSync()` is actually
*called*, not merely imported. This is an inherent esbuild ESM-output limitation, not
something fixable from inside this package (see `src/internal/zip-writer.cjs`'s doc
comment for what was tried). **If you bundle a Node application with esbuild in
`--format=esm` and it calls `write`/`readFile`/`readFileSync`, mark this package
`external`** (`--external:@elixcee/xlsx` or `--packages=external`) and let Node's own
loader resolve it — verified end-to-end, both directions, by `scripts/wasm-smoke.mjs`'s
step 6. CJS-format bundles need no such treatment either way.

**`write` in the browser build is a different, narrower story, and it's already fixed
rather than merely documented as a limitation.** Two real bugs were found (not assumed) by
actually bundling the browser entry with esbuild `--platform=browser` and running the
result in a real Chrome tab, and both are resolved at the source: a `require('zlib')`
reachable anywhere in the bundle's module graph made esbuild refuse to even produce a
`platform: 'browser'` bundle at all ("Could not resolve zlib"), and a `Buffer` reference —
which bundles fine but doesn't exist in a real browser, unlike Node — threw
`ReferenceError: Buffer is not defined` at run time. The browser build now has its own
`write()`, built on a `Buffer`-free ZIP writer (`Uint8Array`/`DataView`/`TextEncoder`
throughout) that never touches `zlib` at all — every entry is written STORED
(uncompressed) instead of DEFLATEd, valid OOXML but larger than the Node build's output
for the same workbook. Verified in the same real headless-Chrome process
`scripts/browser-smoke.mjs` already uses for `read()`.

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
