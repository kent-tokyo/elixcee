// Browser entry point — reached via the "browser" condition in package.json's `exports`
// map (see docs/xlsx-architecture.md's "Phase 2B-0: sync WASM/read() bridge feasibility").
// Every export except `read` is the exact same pure-JS implementation as index.cjs
// (re-exported verbatim below); only `read` differs by platform, because it's the one
// function backed by the WASM bridge, and the bridge itself has two separate builds:
//
// - Node (index.cjs's `read`): elixcee_wasm.node.cjs, which does
//   `require('fs').readFileSync(...)` to load the compiled .wasm file from disk at
//   require-time. That's fine in Node, and would simply fail to resolve in a browser
//   (no `fs`, no such file to read) even under a bundler.
// - Browser (this file's `read`): elixcee_wasm.browser.mjs, a self-contained ES module
//   that inlines the compiled .wasm bytes as base64 and calls wasm-bindgen's `initSync`
//   itself at import time — no `fetch`, no bundler-resolved `import "*.wasm"`, no
//   `require`. This was verified end-to-end in a real Chrome tab during Phase 2B-0
//   (synchronous compile from inlined bytes, no async init() needed).
//
// toBytes is a small, deliberate duplicate of index.cjs's own (not imported from there):
// index.cjs's version uses `Buffer`, a Node global with no standard browser equivalent;
// this one uses `atob`, which IS a standard global in both real browsers and Node
// (stable since Node 16) — so this file has no Node-specific dependency of its own.
//
// read-shape.cjs (the !rows/!cols expansion — see its own doc comment) IS shared as-is:
// it's pure object/array logic with no Buffer/fs dependency, so re-using it here doesn't
// reintroduce a Node-only path.
//
// Caveat, stated plainly rather than glossed over: everything BELOW `read` is reached by
// importing `./index.cjs`, a CommonJS module — safe in Node (including via this file, and
// via a bundler that resolves CJS requires as part of its own module graph, which is how
// the overwhelming majority of real-world "browser" JS consumption of an npm package
// actually happens) but not through a literal `require()` call in a browser tab with no
// bundler and no CJS shim. `read` itself has no such dependency (see above) and works in
// that fully bundler-less environment too — verified via `node --conditions=browser`
// dispatch (this package has no bundler installed to test an actual webpack/esbuild/vite
// build against).
import { readWorkbook } from './internal/wasm/elixcee_wasm.browser.mjs';
import { shapeWorkBook } from './internal/read-shape.cjs';
import { buildXlsxZipEntries } from './internal/xlsx-writer.cjs';
import { makeZip } from './internal/zip-writer.cjs';

const ELIXCEE_UNSUPPORTED_READ_TYPE = 'ELIXCEE_UNSUPPORTED_READ_TYPE';

function toBytes(data, opts) {
  const o = opts || {};
  if (data instanceof Uint8Array) return data;
  if (Array.isArray(data)) return Uint8Array.from(data);
  if (typeof data === 'string' && o.type === 'base64') {
    const bin = atob(data);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; ++i) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }
  const err = new Error(
    "read(): unsupported input — pass a Buffer/Uint8Array, or a base64 string with opts.type " +
      "=== 'base64'. Other xlsx@0.18.5 `type` values (binary/array/string/file) are not " +
      'implemented yet.'
  );
  err.code = ELIXCEE_UNSUPPORTED_READ_TYPE;
  throw err;
}

export function read(data, opts) {
  const bytes = toBytes(data, opts);
  return shapeWorkBook(JSON.parse(readWorkbook(bytes)), opts);
}

// readFile/readFileSync exist here purely so the browser entry's export SET stays identical
// to the Node entry's (checked by scripts/pack-consumer-smoke.mjs) — a missing export is a
// different, more confusing failure for a consumer than an explicit one. They throw
// unconditionally: there is no filesystem to read a path from in a browser, so there is no
// honest implementation, and silently returning an empty workbook or a rejected promise
// would both be worse than saying so. A consumer with bytes already in hand (a File/Blob
// read via FileReader, a fetch() response) should call read() with those bytes.
const ELIXCEE_UNSUPPORTED_IN_BROWSER = 'ELIXCEE_UNSUPPORTED_IN_BROWSER';

function readFileUnsupported() {
  const err = new Error(
    'readFile()/readFileSync() are unsupported in the browser build of @elixcee/xlsx: a ' +
      'browser has no filesystem to read a path from. Read the bytes yourself (FileReader, ' +
      'fetch, drag-and-drop) and pass them to read(bytes) instead.'
  );
  err.code = ELIXCEE_UNSUPPORTED_IN_BROWSER;
  throw err;
}

// Same two-names-one-function shape as the Node entry (and as the oracle itself), so the
// `readFile === readFileSync` identity holds in the browser build too.
export { readFileUnsupported as readFile, readFileUnsupported as readFileSync };

function writeFileUnsupported() {
  const err = new Error(
    'writeFile()/writeFileSync() are unsupported in the browser build of @elixcee/xlsx: a ' +
      'browser has no filesystem to write a path to. Call write(wb, {type: "buffer"|"array"|' +
      '"base64"}) instead and hand the bytes to a download/File-System-Access-API flow yourself.'
  );
  err.code = ELIXCEE_UNSUPPORTED_IN_BROWSER;
  throw err;
}

// Same two-names-one-function shape as writeFile/writeFileSync's Node counterpart.
export { writeFileUnsupported as writeFile, writeFileUnsupported as writeFileSync };

// `write` itself, unlike `writeFile`/`writeFileSync`, needs no filesystem — it's pure XML/
// ZIP generation returning bytes. Unlike every other non-`read` util below, though, it is
// NOT re-exported verbatim from index.cjs: index.cjs's write() uses Node's `zlib` for
// DEFLATE compression, and a real bundled-for-browser build cannot resolve that at all —
// confirmed live (not assumed) by actually bundling it with esbuild `--platform=browser`:
// "Could not resolve zlib", a build-time failure, not a runtime one, so no amount of lazy
// -requiring inside index.cjs would have helped a browser build specifically (see
// zip-writer.cjs's own top doc comment for the full explanation, including why this is a
// DIFFERENT problem from the ESM-bundle "Dynamic require" issue documented in
// docs/xlsx-architecture.md's "Phase D" section — that one is about a Node ESM bundle,
// this one is about a browser bundle, and they need different fixes).
//
// So this platform gets its OWN write(), built from the same platform-agnostic
// buildXlsxZipEntries()/makeZip() the Node entry uses, just called with no `deflate`
// callback — every entry is written STORED (uncompressed) instead of DEFLATEd. Valid
// OOXML (elixcee's own reader.rs, and every real spreadsheet application, accept STORED
// zip entries unconditionally), just larger — the same "state the cost, don't hide it"
// disclosure this package already makes for the WASM-inlining size cost (see README.md's
// "Bundling" section). Kept behaviorally identical to index.cjs's write() otherwise: same
// bookType/type validation, same error codes, same output-type switch.
const ELIXCEE_UNSUPPORTED_BOOK_TYPE = 'ELIXCEE_UNSUPPORTED_BOOK_TYPE';
const ELIXCEE_UNSUPPORTED_WRITE_TYPE = 'ELIXCEE_UNSUPPORTED_WRITE_TYPE';

// Uint8Array -> base64 via btoa, chunked (a single String.fromCharCode(...bytes) spread
// blows the call stack on a large array — this package's own written files easily exceed
// that in a multi-sheet workbook). btoa/atob are the standard browser (and Node 16+) pair
// for this; toBytes() above already uses atob for the reverse direction, same rationale:
// no Buffer, since a real browser doesn't have one and esbuild's `platform: 'browser'`
// does not polyfill it (confirmed live — see zip-writer.cjs's own doc comment).
const BASE64_CHUNK = 0x8000;
function bytesToBase64(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i += BASE64_CHUNK) {
    binary += String.fromCharCode.apply(null, bytes.subarray(i, i + BASE64_CHUNK));
  }
  return btoa(binary);
}

export function write(wb, opts) {
  const o = opts || {};
  const bookType = o.bookType || 'xlsx';
  if (bookType !== 'xlsx') {
    const err = new Error(
      `write(): bookType '${bookType}' is not supported — only 'xlsx' is implemented ` +
        '(no ODS/CSV/TXT/legacy .xls output yet).'
    );
    err.code = ELIXCEE_UNSUPPORTED_BOOK_TYPE;
    throw err;
  }
  const buf = makeZip(buildXlsxZipEntries(wb));
  switch (o.type) {
    case 'buffer':
      return buf;
    case 'array':
      return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    case 'base64':
      return bytesToBase64(buf);
    default: {
      const err = new Error(
        "write(): unsupported opts.type " +
          JSON.stringify(o.type) +
          " — pass 'buffer', 'array', or 'base64'. Other xlsx@0.18.5 `type` values " +
          "('binary'/'string'/'file') are not implemented yet."
      );
      err.code = ELIXCEE_UNSUPPORTED_WRITE_TYPE;
      throw err;
    }
  }
}

export {
  encode_col,
  encode_row,
  encode_cell,
  encode_range,
  decode_col,
  decode_row,
  split_cell,
  decode_cell,
  decode_range,
  format_cell,
  sheet_add_aoa,
  sheet_add_json,
  sheet_add_dom,
  aoa_to_sheet,
  json_to_sheet,
  table_to_sheet,
  table_to_book,
  sheet_to_csv,
  sheet_to_txt,
  sheet_to_json,
  sheet_to_html,
  sheet_to_formulae,
  sheet_to_row_object_array,
  sheet_get_cell,
  book_new,
  book_append_sheet,
  book_set_sheet_visibility,
  cell_set_number_format,
  cell_set_hyperlink,
  cell_set_internal_link,
  cell_add_comment,
  sheet_set_array_formula,
  consts,
} from './index.cjs';
