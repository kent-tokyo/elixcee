'use strict';

// Node-only wrapper around zlib's DEFLATE-raw codec, split into its OWN file so
// package.json's `browser` field can stub it out entirely for a bundled browser build —
// the exact same precedent already used for `elixcee_wasm.node.cjs` (see that entry in
// package.json's `browser` map). This is not optional polish: a literal `require('zlib')`
// cannot appear ANYWHERE in index.cjs's own module body, even inside a function nobody
// calls from the browser, without breaking an esbuild `platform: 'browser'` build —
// confirmed live, not assumed (`Could not resolve "zlib"`, a build-time failure). esbuild
// cannot tree-shake individual properties out of a CommonJS `module.exports` object, so
// it conservatively resolves every `require()` call it can find anywhere in a bundled
// CJS module's source, whether or not that call ever actually executes. Isolating the
// zlib require in its own file lets the `browser` field substitute an empty module for
// THIS file specifically — at path-resolution time, before its contents are ever parsed —
// without needing to touch index.cjs's other (browser-safe) exports at all.
//
// index.browser.mjs never requires this file: it has its own write() built on the same
// platform-agnostic zip-writer.cjs, called with no `deflate` callback at all (STORED-only
// output) — see its own doc comment. This file exists purely so index.cjs's Node write()
// keeps real DEFLATE compression without dragging `zlib` into a browser bundle's module
// graph via index.browser.mjs's re-export of index.cjs's OTHER (non-write) exports.
function deflateRawSync(buf) {
  return require('zlib').deflateRawSync(buf, { level: 9 });
}

module.exports = { deflateRawSync };
