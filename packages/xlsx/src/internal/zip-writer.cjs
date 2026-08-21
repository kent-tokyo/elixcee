'use strict';

// Hand-rolled ZIP container writer for write()/writeFile()/writeFileSync() — no new
// dependency (matches this project's established posture, see docs/xlsx-architecture.md's
// "Context" section on removing rather than adding dependencies). Local file headers, the
// central directory, the end-of-central-directory record, and CRC-32 are implemented here
// directly, since none of it is exposed by any Node builtin.
//
// `zlib.crc32` exists from Node 22 onward, but package.json declares `engines: ">=18"` —
// a hand-rolled table-based CRC-32 (the standard ZIP/PNG/gzip polynomial, 0xEDB88320) keeps
// every supported Node version on one code path rather than branching on version.
//
// This module itself has NO Node-builtin dependency at all — neither `zlib` (DEFLATE
// compression is injected by the CALLER as a `deflate` callback, see makeZip below) nor
// `Buffer` (every byte buffer here is a plain `Uint8Array`, built with `DataView`/
// `TextEncoder`). Both were real bugs found by actually bundling write() for a browser
// target and running the bundle in a real Chrome tab (not assumed from reading the code):
//
// - `require('zlib')`: esbuild's `platform: 'browser'` refuses to even PRODUCE a bundle
//   that contains a `require('zlib')` call ANYWHERE in its reachable module graph —
//   "Could not resolve zlib" — regardless of whether that call is lazy/conditional, since
//   esbuild's bundler resolves every require() it can statically find in a bundled file,
//   not just the ones actually executed. Making the require lazy inside this file (an
//   earlier design) only helped Node/esbuild-CJS/esbuild-ESM-with-the-package-marked-
//   external (see docs/xlsx-architecture.md's "Phase D" section) — a browser bundle still
//   choked on it existing in this file's source at all. DEFLATE is now supplied by the
//   caller instead (index.cjs's Node build passes a lazy `zlib.deflateRawSync`;
//   index.browser.mjs's build passes nothing and gets STORED-only output).
// - `Buffer`: unlike `require('zlib')`, this one bundled and ran without error, then threw
//   `ReferenceError: Buffer is not defined` the moment a real (unshimmed) browser actually
//   executed it — `Buffer` is a Node global with no standard browser equivalent, and
//   esbuild's `platform: 'browser'` does not polyfill it (unlike `platform: 'node'`, or
//   older bundlers like webpack 4). `Uint8Array`/`DataView`/`TextEncoder` are standard in
//   both Node (11+) and every real browser, so using them here removes the need for a
//   platform split at all — this module is identical for both platforms.
//
// index.cjs's writeBuffer() wraps this module's Uint8Array output back into a real Node
// `Buffer` afterward, for API compatibility with the oracle's own type:'buffer' contract
// (`Buffer.prototype.toString('base64')`, `Buffer.isBuffer`, etc.) — see its own doc
// comment. index.browser.mjs's write() uses the Uint8Array directly.

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[n] = c >>> 0;
  }
  return table;
})();

function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) {
    crc = CRC_TABLE[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

// MS-DOS date/time encoding ZIP's own header format requires — fixed to a stable epoch
// (2020-01-01 00:00:00) rather than the real current time, so two writes of the same
// workbook produce byte-identical ZIP output (useful for tests, and avoids leaking the
// machine clock into a file's bytes for no benefit to the caller).
const DOS_TIME = 0;
const DOS_DATE = ((2020 - 1980) << 9) | (1 << 5) | 1;

function u16(n) {
  const b = new Uint8Array(2);
  new DataView(b.buffer).setUint16(0, n & 0xffff, true);
  return b;
}
function u32(n) {
  const b = new Uint8Array(4);
  new DataView(b.buffer).setUint32(0, n >>> 0, true);
  return b;
}

// Uint8Array has no Buffer.concat equivalent — this copies each part into one allocation.
function concatBytes(parts) {
  let total = 0;
  for (const p of parts) total += p.length;
  const out = new Uint8Array(total);
  let offset = 0;
  for (const p of parts) {
    out.set(p, offset);
    offset += p.length;
  }
  return out;
}

/**
 * Builds a ZIP archive (as a `Uint8Array`) from `entries: [{name: string, data:
 * Uint8Array}]`. Entry order in `entries` is preserved as the archive's own physical and
 * central-directory order — callers control ordering (OOXML doesn't require any
 * particular part order, but writing `[Content_Types].xml` first matches every real-world
 * XLSX writer, including the oracle's own).
 *
 * `deflate`, optional: a synchronous `(buf: Uint8Array) => Uint8Array` DEFLATE-raw
 * compressor (e.g. a `zlib.deflateRawSync` wrapper). When omitted, every entry is written
 * STORED (uncompressed) — a legal ZIP/OOXML method (elixcee's own reader.rs's zip
 * handling accepts it unconditionally, and so does every real spreadsheet application),
 * just larger. See this file's own top doc comment for why the browser build calls this
 * with no `deflate` at all rather than requiring zlib itself.
 */
function makeZip(entries, deflate) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const { name, data } of entries) {
    const nameBuf = new TextEncoder().encode(name);
    const crc = crc32(data);
    const compressed = deflate ? deflate(data) : null;
    // Fall back to STORED if DEFLATE somehow doesn't shrink the entry (never happens for
    // real XML text, but guards a pathological/incompressible input rather than emitting a
    // corrupt "compressed" size larger than what a naive reader might have budgeted for) —
    // or if no `deflate` was given at all.
    const useStore = !compressed || compressed.length >= data.length;
    const method = useStore ? 0 : 8;
    const payload = useStore ? data : compressed;

    const localHeader = concatBytes([
      u32(0x04034b50),
      u16(20), // version needed to extract (2.0 — DEFLATE)
      u16(0), // general purpose bit flag
      u16(method),
      u16(DOS_TIME),
      u16(DOS_DATE),
      u32(crc),
      u32(payload.length),
      u32(data.length),
      u16(nameBuf.length),
      u16(0), // extra field length
    ]);
    localParts.push(localHeader, nameBuf, payload);

    const centralHeader = concatBytes([
      u32(0x02014b50),
      u16(20), // version made by
      u16(20), // version needed to extract
      u16(0),
      u16(method),
      u16(DOS_TIME),
      u16(DOS_DATE),
      u32(crc),
      u32(payload.length),
      u32(data.length),
      u16(nameBuf.length),
      u16(0), // extra field length
      u16(0), // file comment length
      u16(0), // disk number start
      u16(0), // internal file attributes
      u32(0), // external file attributes
      u32(offset),
    ]);
    centralParts.push(centralHeader, nameBuf);

    offset += localHeader.length + nameBuf.length + payload.length;
  }

  const centralDir = concatBytes(centralParts);
  const centralStart = offset;

  const eocd = concatBytes([
    u32(0x06054b50),
    u16(0), // this disk
    u16(0), // disk with central directory start
    u16(entries.length), // entries on this disk
    u16(entries.length), // total entries
    u32(centralDir.length),
    u32(centralStart),
    u16(0), // comment length
  ]);

  return concatBytes([...localParts, centralDir, eocd]);
}

module.exports = { makeZip, crc32 };
