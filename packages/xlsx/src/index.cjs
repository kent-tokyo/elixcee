'use strict';

// @elixcee/xlsx — Phase 1A pure utility API (address encode/decode, workbook shells),
// extended in Phase 1B-1 with worksheet mutation (sheet_add_aoa/sheet_add_json) and a
// deliberately narrow number-format subset (format_cell/cell_set_number_format).
// No runtime dependency on the real `xlsx` package (see docs/xlsx-architecture.md's
// "Non-negotiable" section) — the only `require` below is the package's own internal
// safe-decode-range helper.
//
// Exact edge-case behavior (including quirks that look like bugs — e.g. decode_range
// never validates or swaps a reversed range; book_append_sheet's error message mentions
// ":" as a forbidden sheet-name character but the actual check never blocks it) was
// verified against the real xlsx@0.18.5 (SheetJS, Apache-2.0) source and confirmed
// against a live oracle run — see compat/differential/. Code below is an independent
// implementation, not copied text; see docs/licensing.md for the licensing boundary.

const { safeDecodeRange } = require('./internal/safe-decode-range.cjs');
const { checkRangeSize } = require('./internal/range-guard.cjs');
const { datenum } = require('./internal/datenum.cjs');
const { format: ssfFormat } = require('./internal/ssf-adapter.cjs');
const { formatCell, cellSetNumberFormat } = require('./internal/number-format.cjs');
// The WASM bridge (crates/elixcee-wasm, backed by elixcee's own hand-rolled reader —
// src/reader.rs's read_workbook_from_bytes) is vendored, prebuilt, under ./internal/wasm —
// npm consumers have no Rust/wasm-pack toolchain, so the compiled artifact ships committed
// (see crates/elixcee-wasm/build.sh). Still not a dependency on the real `xlsx` package —
// this is elixcee's own reader, just compiled to WASM.
//
// Required LAZILY (inside `read()`, not here at module top-level) rather than eagerly:
// this file's own `require('./internal/wasm/elixcee_wasm.node.cjs')` does
// `require('fs').readFileSync(...)` the moment IT loads — fine for every Node consumer,
// but the "browser" export condition entry point (index.browser.mjs) re-exports this
// file's non-read utility functions verbatim and must not trigger that fs read merely by
// being imported. A caller who never calls read() (the overwhelming majority of this
// package's existing utils-only surface) also never pays for resolving it.
let wasmBridge;
function getWasmBridge() {
  return wasmBridge || (wasmBridge = require('./internal/wasm/elixcee_wasm.node.cjs'));
}

// ---- column ----

// Elixcee-specific error codes, stable and short by design (see docs/xlsx-security-model.md
// and compat/differential/classify.mjs's SAFETY_DIVERGENCE_REGISTRY, which is keyed by
// these exact strings — do not rename casually).
const ELIXCEE_NON_FINITE_INDEX = 'ELIXCEE_NON_FINITE_INDEX';

function encodeCol(col) {
  // The oracle's loop (`for(++col; col; col=Math.floor((col-1)/26))`) never terminates
  // when col is +Infinity — Math.floor(Infinity) stays Infinity forever. Confirmed by
  // actually running it (process OOM-killed after ~30s), not assumed. This is the one
  // genuine hang in this API slice — encode_row/encode_cell/encode_range and
  // encode_col's own NaN/-Infinity/MAX_VALUE/MAX_SAFE_INTEGER paths were all empirically
  // timed and confirmed to terminate instantly, so they are deliberately left matching
  // the oracle unchanged (see compat/differential/xlsx-utils.test.mjs's encode_col
  // matrix for the exact evidence). Only +Infinity is intercepted, to avoid manufacturing
  // unnecessary compatibility divergences where the oracle is already safe.
  if (+col === Infinity) {
    const err = new RangeError('column index must be finite');
    err.code = ELIXCEE_NON_FINITE_INDEX;
    throw err;
  }
  if (col < 0) throw new Error('invalid column ' + col);
  let n = col + 1;
  let s = '';
  for (; n; n = Math.floor((n - 1) / 26)) {
    s = String.fromCharCode(((n - 1) % 26) + 65) + s;
  }
  return s;
}

// A single leading "$" immediately before an uppercase letter is stripped; anything else
// ($ elsewhere, lowercase, digits, punctuation) is left as-is and fed into the same raw
// charCode arithmetic below — decode_col does not validate its input is well-formed.
function unfixCol(colstr) {
  return colstr.replace(/^\$([A-Z])/, '$1');
}

function decodeCol(colstr) {
  const c = unfixCol(colstr);
  let d = 0;
  for (let i = 0; i !== c.length; ++i) d = 26 * d + c.charCodeAt(i) - 64;
  return d - 1;
}

// ---- row ----

function encodeRow(row) {
  return '' + (row + 1);
}

function unfixRow(rowstr) {
  return rowstr.replace(/\$(\d+)$/, '$1');
}

function decodeRow(rowstr) {
  return parseInt(unfixRow(rowstr), 10) - 1;
}

// ---- cell ----

function encodeCell(cell) {
  // Deliberately NOT encodeCol(cell.c) — this inline loop has no negative-column guard,
  // so encodeCell({c: -1, r: 0}) returns "1" (empty column part) instead of throwing,
  // unlike the standalone encode_col. Confirmed against the oracle, not an oversight.
  let col = cell.c + 1;
  let s = '';
  for (; col; col = ((col - 1) / 26) | 0) {
    s = String.fromCharCode(((col - 1) % 26) + 65) + s;
  }
  return s + (cell.r + 1);
}

// Scans every character of the whole string once: uppercase A-Z accumulates into the
// column (base 26), '0'-'9' accumulates into the row (base 10), everything else
// (lowercase, "$", "!", ":", "'", ...) is silently ignored. This is NOT a "parse the
// last cell reference" or "strip the sheet-name prefix" operation — inputs like
// "Sheet1!A1" or "A1:B2" produce large, arithmetic (not semantically meaningful)
// numbers because every stray letter/digit in the string still contributes. Confirmed
// against the oracle for exactly this reason: guessing "surely it strips non-cell
// characters" would have been wrong.
function decodeCell(cstr) {
  let r = 0;
  let c = 0;
  for (let i = 0; i < cstr.length; ++i) {
    const cc = cstr.charCodeAt(i);
    if (cc >= 48 && cc <= 57) r = 10 * r + (cc - 48);
    else if (cc >= 65 && cc <= 90) c = 26 * c + (cc - 64);
  }
  return { c: c - 1, r: r - 1 };
}

// ---- range ----

function encodeRange(cs, ce) {
  if (typeof ce === 'undefined' || typeof ce === 'number') {
    return encodeRange(cs.s, cs.e);
  }
  const start = typeof cs !== 'string' ? encodeCell(cs) : cs;
  const end = typeof ce !== 'string' ? encodeCell(ce) : ce;
  return start === end ? start : start + ':' + end;
}

// No validation, no swap: decode_range("B2:A1") returns {s: B2, e: A1} verbatim — the
// oracle does the same. "Reversed range" is not an error condition for this function.
function decodeRange(range) {
  const idx = range.indexOf(':');
  if (idx === -1) return { s: decodeCell(range), e: decodeCell(range) };
  return { s: decodeCell(range.slice(0, idx)), e: decodeCell(range.slice(idx + 1)) };
}

// safe_decode_range deliberately does NOT live here: it is not part of xlsx@0.18.5's
// public `utils` surface (confirmed:
// `Object.prototype.hasOwnProperty.call(XLSX.utils, "safe_decode_range") === false` at
// runtime). Publishing it under this compat namespace would itself be a compatibility
// divergence. See src/internal/safe-decode-range.cjs.

// ---- split_cell ----

// A leading "$?[A-Z]*" then "$?\d*" is matched (possibly both zero-width) and replaced
// with "$1,$2", then split on the comma. Lowercase input can produce a zero-width match
// at position 0, e.g. split_cell("a1") => ["", "a1"] — the WHOLE string lands in the
// second slot, not just its digits. Confirmed against the oracle.
function splitCell(cstr) {
  return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/, '$1,$2').split(',');
}

// ---- read ----
//
// Buffer-first XLSX.read(data, opts), backed by the WASM bridge over elixcee's own reader
// (src/reader.rs) — never the real `xlsx` package (see this file's top doc comment /
// docs/xlsx-architecture.md's "Non-negotiable" section). Accepts a Buffer/Uint8Array
// directly (the shape `fs.readFileSync(...)` already produces) or a base64 string with
// `opts.type === 'base64'` (the oracle's own convention for that value) — other oracle
// `type` values (binary/array/string/file) are not implemented yet.
//
// Cell formulas (`.f`), merged ranges (`!merges`), and formatted display text (`.w`) are
// always mapped. A worksheet's declared `<dimension>` is preferred over the
// populated-cell bounding box when present (see crates/elixcee-wasm's worksheet_json).
// Three more fields are gated behind opts, matching the oracle's own gates exactly
// (confirmed live — none of these are surfaced by default): hidden-row/col `!rows`/
// `!cols` with `opts.cellStyles`; resolved format-code `.z` with `opts.cellNF`
// (`opts.cellStyles` implies it); date-typed cells (`t:'d'`) with `opts.cellDates`. See
// ./internal/read-shape.cjs for all of the above.
const { shapeWorkBook } = require('./internal/read-shape.cjs');

const ELIXCEE_UNSUPPORTED_READ_TYPE = 'ELIXCEE_UNSUPPORTED_READ_TYPE';

function toBytes(data, opts) {
  const o = opts || {};
  if (data instanceof Uint8Array) return data;
  if (Array.isArray(data)) return Uint8Array.from(data);
  if (typeof data === 'string' && o.type === 'base64') return Uint8Array.from(Buffer.from(data, 'base64'));
  const err = new Error(
    "read(): unsupported input — pass a Buffer/Uint8Array, or a base64 string with opts.type " +
      "=== 'base64'. Other xlsx@0.18.5 `type` values (binary/array/string/file) are not " +
      'implemented yet.'
  );
  err.code = ELIXCEE_UNSUPPORTED_READ_TYPE;
  throw err;
}

function read(data, opts) {
  const bytes = toBytes(data, opts);
  return shapeWorkBook(JSON.parse(getWasmBridge().readWorkbook(bytes)), opts);
}

// ---- readFile / readFileSync ----
//
// Node-only file-path entry points, a thin wrapper over read() above. ONE function exported
// under BOTH names, because that is what the oracle does — confirmed live, not assumed:
// `XLSX.readFile === XLSX.readFileSync` is true, and the shared function's own `.name` is
// "readFileSync" (so `XLSX.readFile.name === 'readFileSync'`, which NAME_OVERRIDES below
// reproduces for both keys). Defining two separate functions would break that identity.
//
// `require('fs')` is INSIDE the function, not at module top level, for the same reason
// getWasmBridge() is lazy: index.browser.mjs re-exports this file's non-read functions, so
// a top-level fs require would make merely importing the browser entry pull `fs` into a
// browser bundle. package.json's `browser` field additionally maps "fs" to false, so a
// bundler that DOES statically follow this require (esbuild does, even inside a function
// body) stubs it out instead of failing to resolve it. The browser entry exports its own
// explicitly-throwing readFile/readFileSync rather than this one — see index.browser.mjs.
//
// No path validation, no encoding option, no existence check: a missing or unreadable file
// throws fs's own native ENOENT/EACCES error, exactly as the oracle's does (it calls
// `_fs.readFileSync` the same way). Deliberately not wrapped in a friendlier error — a
// caller matching on `err.code === 'ENOENT'` must keep working.
function readFileSyncImpl(filename, opts) {
  return read(require('fs').readFileSync(filename), opts);
}

// ---- write ----
//
// Buffer-first XLSX.write(wb, opts), the inverse of read() above — WorkBook object -> a
// real, Excel-openable .xlsx file's bytes. No WASM/Rust bridge (unlike read): OOXML
// writing is pure XML/ZIP generation, verified against elixcee's own reader (src/
// reader.rs) so "own write -> own read" is a meaningful round trip, not two independently
// -guessed formats — see internal/xlsx-writer.cjs's own top doc comment.
//
// bookType: 'xlsx' only (defaults to 'xlsx' when omitted, matching the oracle's own
// default) — any other value (the oracle also accepts 'ods'/'csv'/'txt'/legacy .xls
// variants/etc.) throws ELIXCEE_UNSUPPORTED_BOOK_TYPE rather than silently producing
// something else. type: 'buffer' | 'array' | 'base64' only — the oracle's other `type`
// values ('binary'/'string'/'file') are not implemented, matching read()'s own narrow
// `type` support; type has no default (matching the oracle: XLSX.write(wb, {}) throws
// "Unrecognized type undefined" there too), so it must be given explicitly.
const { makeZip } = require('./internal/zip-writer.cjs');
const { buildXlsxZipEntries } = require('./internal/xlsx-writer.cjs');

const ELIXCEE_UNSUPPORTED_BOOK_TYPE = 'ELIXCEE_UNSUPPORTED_BOOK_TYPE';
const ELIXCEE_UNSUPPORTED_WRITE_TYPE = 'ELIXCEE_UNSUPPORTED_WRITE_TYPE';

function writeBuffer(wb, opts) {
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
  // Lazy require of deflate-node.cjs (not zlib directly) — see that file's own top doc
  // comment for why the zlib access must live in its own file (so package.json's
  // `browser` field can stub it out of a bundled browser build) rather than inline here.
  const zipped = makeZip(buildXlsxZipEntries(wb), require('./internal/deflate-node.cjs').deflateRawSync);
  // zip-writer.cjs returns a plain Uint8Array (it has no Buffer dependency at all — see
  // its own doc comment). Wrapped back into a real Node Buffer here, zero-copy, so
  // type:'buffer' keeps matching the oracle's own contract exactly (a true Buffer, with
  // e.g. `.toString('base64')` — plain Uint8Array has no such method before Node 20).
  return Buffer.from(zipped.buffer, zipped.byteOffset, zipped.byteLength);
}

function write(wb, opts) {
  const o = opts || {};
  const buf = writeBuffer(wb, o);
  switch (o.type) {
    case 'buffer':
      return buf;
    case 'array':
      return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    case 'base64':
      return buf.toString('base64');
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

// ---- writeFile / writeFileSync ----
//
// Node-only file-path entry points, a thin wrapper over write() above — same one
// -function-under-two-names shape as readFile/readFileSync (confirmed live against the
// oracle: `XLSX.writeFile === XLSX.writeFileSync` is true, and the shared function's own
// `.name` is "writeFileSync"). `opts.type` is never needed here (the bytes always go to
// `filename` via fs.writeFileSync, never returned) — only `opts.bookType` has any effect,
// same 'xlsx'-only restriction as write() above.
//
// `require('fs')` is INSIDE the function for the same reason readFileSyncImpl's is — see
// that function's own doc comment. The browser entry exports its own explicitly-throwing
// writeFile/writeFileSync — see index.browser.mjs.
function writeFileSyncImpl(wb, filename, opts) {
  require('fs').writeFileSync(filename, writeBuffer(wb, opts));
}

// ---- workbook / sheet ----

function bookNew() {
  return { SheetNames: [], Sheets: {} };
}

// Characters actually rejected by the real oracle — note ":" is NOT in this list even
// though the thrown error message text claims it is (a confirmed oracle quirk: the
// message and the check disagree). Replicated as-is, message text included, since the
// message text itself is part of what a caller might match on or display.
const SHEET_NAME_BAD_CHARS = ['[', ']', '*', '?', '/', '\\'];

function checkSheetName(name) {
  if (name.length > 31) throw new Error('Sheet names cannot exceed 31 chars');
  for (const ch of SHEET_NAME_BAD_CHARS) {
    if (name.indexOf(ch) !== -1) {
      throw new Error('Sheet name cannot contain : \\ / ? * [ ]');
    }
  }
}

function bookAppendSheet(wb, ws, name, roll) {
  let i = 1;
  if (!name) {
    for (; i <= 0xffff; ++i, name = undefined) {
      name = 'Sheet' + i;
      if (wb.SheetNames.indexOf(name) === -1) break;
    }
  }
  if (!name || wb.SheetNames.length >= 0xffff) throw new Error('Too many worksheets');

  if (roll && wb.SheetNames.indexOf(name) >= 0) {
    const m = name.match(/(^.*?)(\d+)$/);
    i = (m && +m[2]) || 0;
    const root = (m && m[1]) || name;
    for (++i; i <= 0xffff; ++i) {
      name = root + i;
      if (wb.SheetNames.indexOf(name) === -1) break;
    }
  }

  checkSheetName(name);
  if (wb.SheetNames.indexOf(name) >= 0) {
    throw new Error('Worksheet with name |' + name + '| already exists!');
  }

  wb.SheetNames.push(name);
  // Object.defineProperty, not `wb.Sheets[name] = ws`: a sheet literally named
  // "__proto__" is legitimate data a caller may pass (a crafted or accidental sheet
  // name), and per docs/xlsx-security-model.md it must be retained as data, not
  // rejected — but plain bracket assignment on an ordinary object invokes
  // Object.prototype's inherited `__proto__` accessor instead of creating a normal own
  // property, which (a) silently reassigns wb.Sheets's own prototype to `ws` and (b)
  // makes the sheet unretrievable via `wb.Sheets[name]` or `Object.keys(wb.Sheets)`
  // (confirmed against the real oracle, which has this exact defect). defineProperty
  // always creates/overwrites a normal own data property regardless of the key's name,
  // closing the hole with zero shape divergence from the oracle (wb.Sheets stays a
  // plain object, `Object.getPrototypeOf(wb.Sheets) === Object.prototype` unchanged).
  Object.defineProperty(wb.Sheets, name, { value: ws, writable: true, enumerable: true, configurable: true });
  return name;
}

function wbSheetIdx(wb, sh) {
  if (typeof sh === 'number') {
    if (sh >= 0 && wb.SheetNames.length > sh) return sh;
    throw new Error('Cannot find sheet # ' + sh);
  }
  if (typeof sh === 'string') {
    const idx = wb.SheetNames.indexOf(sh);
    if (idx > -1) return idx;
    throw new Error('Cannot find sheet name |' + sh + '|');
  }
  throw new Error('Cannot find sheet |' + sh + '|');
}

function bookSetSheetVisibility(wb, sh, vis) {
  if (!wb.Workbook) wb.Workbook = {};
  if (!wb.Workbook.Sheets) wb.Workbook.Sheets = [];

  const idx = wbSheetIdx(wb, sh);
  if (!wb.Workbook.Sheets[idx]) wb.Workbook.Sheets[idx] = {};

  if (vis !== 0 && vis !== 1 && vis !== 2) {
    throw new Error('Bad sheet visibility setting ' + vis);
  }
  wb.Workbook.Sheets[idx].Hidden = vis;
}

// ---- number formats ----
//
// format_cell / cell_set_number_format live in ./internal/number-format.cjs, backed by
// the real SSF engine via ./internal/ssf-adapter.cjs (see docs/xlsx-architecture.md's
// "SSF backend" decision — a deliberate, disclosed transitional runtime dependency on
// `ssf@0.11.2`, confirmed byte-identical to the oracle's bundled engine across an
// 819-case matrix, replacing Phase 1B-1's deliberately-narrow 'General'/'m/d/yy'-only
// subset). `datenum` (Date -> Excel serial, unrelated to the format-string engine
// itself) lives in ./internal/datenum.cjs, shared with sheet_add_aoa/sheet_add_json
// below.
//
// sheet_add_aoa's Date branch calls `ssfFormat` directly (not through format_cell's
// safeFormatCell two-try fallback) — confirmed live: an aoa_to_sheet() call with a
// dateNF the SSF engine itself rejects throws, uncaught, unlike format_cell's
// best-effort ''+v fallback.

// ---- sheet_add_aoa / aoa_to_sheet ----
//
// Independent port of the oracle's sheet_add_aoa(_ws, data, opts): dense (array-of-
// -arrays) vs. sparse (cell-ref-keyed object) storage — inferred from `_ws` when given,
// else `opts.dense` — origin (number row, -1 "append after existing range", "A1"-style
// string, or {r,c} object), extending an existing `!ref` via safe_decode_range,
// preserving an existing cell's `.z` when overwriting, null/undefined cell skipping,
// number/boolean/string type inference, and Date cells (serial-number `v` plus a
// rendered `.w` via the narrow SSF subset above; `opts.cellDates` keeps `t:'d'`/`v` as a
// Date instead). `opts.dateNF` overrides the format string; anything other than the
// literal 'm/d/yy' throws ELIXCEE_NUMFMT_UNSUPPORTED (see the section above).
function sheetAddAoa(_ws, data, opts) {
  const o = opts || {};
  const dense = _ws ? Array.isArray(_ws) : !!o.dense;
  const ws = _ws || (dense ? [] : {});
  let _R = 0;
  let _C = 0;
  if (ws && o.origin != null) {
    if (typeof o.origin === 'number') _R = o.origin;
    else {
      const _origin = typeof o.origin === 'string' ? decodeCell(o.origin) : o.origin;
      _R = _origin.r;
      _C = _origin.c;
    }
    if (!ws['!ref']) ws['!ref'] = 'A1:A1';
  }
  const range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
  if (ws['!ref']) {
    const _range = safeDecodeRange(ws['!ref']);
    range.s.c = _range.s.c;
    range.s.r = _range.s.r;
    range.e.c = Math.max(range.e.c, _range.e.c);
    range.e.r = Math.max(range.e.r, _range.e.r);
    if (_R === -1) range.e.r = _R = _range.e.r + 1;
  }
  for (let R = 0; R !== data.length; ++R) {
    if (!data[R]) continue;
    if (!Array.isArray(data[R])) throw new Error('aoa_to_sheet expects an array of arrays');
    for (let C = 0; C !== data[R].length; ++C) {
      const raw = data[R][C];
      if (typeof raw === 'undefined') continue;
      let cell = { v: raw };
      const __R = _R + R;
      const __C = _C + C;
      if (range.s.r > __R) range.s.r = __R;
      if (range.s.c > __C) range.s.c = __C;
      if (range.e.r < __R) range.e.r = __R;
      if (range.e.c < __C) range.e.c = __C;
      if (raw && typeof raw === 'object' && !Array.isArray(raw) && !(raw instanceof Date)) {
        cell = raw; // caller-supplied full cell object, used verbatim
      } else {
        if (Array.isArray(cell.v)) {
          // [value, formula] pair shorthand.
          cell.f = raw[1];
          cell.v = raw[0];
        }
        if (cell.v === null) {
          if (cell.f) cell.t = 'n';
          else if (o.nullError) {
            cell.t = 'e';
            cell.v = 0;
          } else if (!o.sheetStubs) continue;
          else cell.t = 'z';
        } else if (typeof cell.v === 'number') cell.t = 'n';
        else if (typeof cell.v === 'boolean') cell.t = 'b';
        else if (cell.v instanceof Date) {
          cell.z = o.dateNF || 'm/d/yy';
          if (o.cellDates) {
            cell.t = 'd';
            cell.w = ssfFormat(cell.z, datenum(cell.v));
          } else {
            cell.t = 'n';
            cell.v = datenum(cell.v);
            cell.w = ssfFormat(cell.z, cell.v);
          }
        } else cell.t = 's';
      }
      if (dense) {
        if (!ws[__R]) ws[__R] = [];
        if (ws[__R][__C] && ws[__R][__C].z) cell.z = ws[__R][__C].z;
        ws[__R][__C] = cell;
      } else {
        const cellRef = encodeCell({ c: __C, r: __R });
        if (ws[cellRef] && ws[cellRef].z) cell.z = ws[cellRef].z;
        ws[cellRef] = cell;
      }
    }
  }
  if (range.s.c < 10000000) ws['!ref'] = encodeRange(range);
  return ws;
}

function aoaToSheet(data, opts) {
  return sheetAddAoa(null, data, opts);
}

// ---- sheet_add_json / json_to_sheet ----
//
// Independent port of the oracle's sheet_add_json(_ws, js, opts): header-row inference
// from `Object.keys` of each row object (or `opts.header` if given — mutated in place by
// design, matching the oracle exactly, not cloned), `opts.skipHeader`, origin (same
// number/-1/string/object forms as sheet_add_aoa, but — confirmed against a live oracle
// run — WITHOUT the "seed !ref to A1:A1" step sheet_add_aoa does; that asymmetry is a
// real oracle quirk, reproduced as-is), and Date cells (serial `v` + `z:'m/d/yy'`, but —
// also confirmed live — no `.w` is ever computed here, unlike sheet_add_aoa).
//
// Dense (`_ws` given as an array) is supported only as faithfully as the oracle supports
// it: scalar values land correctly in the nested array via ws_get_cell_stub, but object-
// typed JSON values AND the header row are both written as stray string-ref properties
// on the array (`ws.A1 = ...`) rather than into the nested rows — confirmed live
// (`sheet_add_json([], [{a:1}])` leaves `ws[0]` `null` and the header text only
// reachable via `ws.A1`). `opts.dense` itself has no effect when `_ws` is null — also
// confirmed live (json_to_sheet(data, {dense:true}) still returns a plain sparse
// object). Reproduced exactly, not "fixed", per this project's fidelity-over-tidiness
// rule for the real oracle's own quirks. See docs/compatibility-known-defects.md.
function wsGetCellStub(ws, ref) {
  if (Array.isArray(ws)) {
    const RC = decodeCell(ref);
    if (!ws[RC.r]) ws[RC.r] = [];
    return ws[RC.r][RC.c] || (ws[RC.r][RC.c] = { t: 'z' });
  }
  return ws[ref] || (ws[ref] = { t: 'z' });
}

// Public sheet_get_cell = the oracle's ws_get_cell_stub (which it exports verbatim under
// that name — confirmed by reading compat/node_modules/xlsx/xlsx.js: `sheet_get_cell:
// ws_get_cell_stub`), accepting all 3 call shapes it does: an A1 string ref (handled by
// wsGetCellStub above), a CellAddress-like object (recurse via encode_cell), or 0-based
// (row, col) numbers (recurse via encode_cell({r,c})). Not in xlsx@0.18.5's own
// types/index.d.ts at all (confirmed: no `get_cell` entry in types/index.d.ts) even
// though it's a real runtime export — src/index.d.ts adds a type for it as pure
// ADDITION, not tightening (there is no oracle declaration to narrow). Like
// wsGetCellStub, this MUTATES: a miss creates a `{t:'z'}` stub in place (and, in dense
// mode, materializes `ws[R]` if absent) — matches the oracle exactly.
function sheetGetCell(ws, R, C) {
  if (typeof R === 'string') return wsGetCellStub(ws, R);
  if (typeof R !== 'number') return sheetGetCell(ws, encodeCell(R));
  return sheetGetCell(ws, encodeCell({ r: R, c: C || 0 }));
}

function sheetAddJson(_ws, js, opts) {
  const o = opts || {};
  const offset = +!o.skipHeader;
  const ws = _ws || {};
  let _R = 0;
  let _C = 0;
  if (ws && o.origin != null) {
    if (typeof o.origin === 'number') _R = o.origin;
    else {
      const _origin = typeof o.origin === 'string' ? decodeCell(o.origin) : o.origin;
      _R = _origin.r;
      _C = _origin.c;
    }
  }
  const range = { s: { c: 0, r: 0 }, e: { c: _C, r: _R + js.length - 1 + offset } };
  if (ws['!ref']) {
    const _range = safeDecodeRange(ws['!ref']);
    range.e.c = Math.max(range.e.c, _range.e.c);
    range.e.r = Math.max(range.e.r, _range.e.r);
    if (_R === -1) {
      _R = _range.e.r + 1;
      range.e.r = _R + js.length - 1 + offset;
    }
  } else if (_R === -1) {
    _R = 0;
    range.e.r = js.length - 1 + offset;
  }
  const hdr = o.header || [];
  let C = 0;
  js.forEach((JS, R) => {
    Object.keys(JS).forEach((k) => {
      if ((C = hdr.indexOf(k)) === -1) hdr[(C = hdr.length)] = k;
      let v = JS[k];
      let t = 'z';
      let z = '';
      const ref = encodeCell({ c: _C + C, r: _R + R + offset });
      const cell = wsGetCellStub(ws, ref);
      if (v && typeof v === 'object' && !(v instanceof Date)) {
        ws[ref] = v;
      } else {
        if (typeof v === 'number') t = 'n';
        else if (typeof v === 'boolean') t = 'b';
        else if (typeof v === 'string') t = 's';
        else if (v instanceof Date) {
          t = 'd';
          if (!o.cellDates) {
            t = 'n';
            v = datenum(v);
          }
          z = o.dateNF || 'm/d/yy';
        } else if (v === null && o.nullError) {
          t = 'e';
          v = 0;
        }
        cell.t = t;
        cell.v = v;
        delete cell.w;
        delete cell.R;
        if (z) cell.z = z;
      }
    });
  });
  range.e.c = Math.max(range.e.c, _C + hdr.length - 1);
  const __R = encodeRow(_R);
  if (offset) {
    for (C = 0; C < hdr.length; ++C) {
      ws[encodeCol(C + _C) + __R] = { t: 's', v: hdr[C] };
    }
  }
  ws['!ref'] = encodeRange(range);
  return ws;
}

function jsonToSheet(js, opts) {
  return sheetAddJson(null, js, opts);
}

// ---- sheet_to_json ----
//
// Independent port of the oracle's sheet_to_json/make_json_row. header mode dispatch:
// 1 -> 0-based column-index numbers, "A" -> column letters, Array -> the array as-is
// (offset becomes 0, i.e. row 0 of the range is treated as data, not a header row), else
// -> infer from row 0's formatted text with de-dup via header_cnt (collisions get a
// "_N" suffix). `raw` defaults to true only when the `raw` key is entirely ABSENT from
// opts (`o.raw || !hasOwnProperty(o,'raw')` — confirmed live: `{raw:undefined}` behaves
// like `raw:true`, since `hasOwnProperty` is still true, but `o.raw` is falsy... the `||`
// covers that: falsy `o.raw` alone doesn't disable raw mode unless the key is present AND
// explicitly falsy). Error cells (`t:'e'`) become `null` only when `v == 0`; otherwise
// `undefined`, which then falls through to the defval/raw-null/skip logic same as any
// other null-ish cell. checkRangeSize (see ./internal/range-guard.cjs) guards the same
// !ref-rectangle walk-cost DoS already measured for sheet_to_formulae/sheet_to_csv
// (docs/limits.md) — this function walks the identical rectangle shape, so no separate
// measurement was needed for this threshold.
//
// header_cnt is a plain `{}`, matching the oracle exactly — deliberately NOT
// Object.create(null): reading `header_cnt["__proto__"]` (a literal "__proto__" header
// cell) or `header_cnt["constructor"]` returns a truthy inherited value (the real
// Object.prototype / Object constructor respectively), which accidentally triggers the
// SAME collision-suffix path as a real duplicate header, renaming the header text to
// "__proto___NaN" / "constructor_NaN" (confirmed live against the oracle) — an oracle
// quirk, not a hazard: the write `header_cnt[v] = counter` at the end assigns a NaN
// counter through the __proto__ SETTER, which per spec is a no-op for non-object values,
// so header_cnt's own prototype is never actually touched. Reproduced as-is per this
// project's fidelity-over-tidiness rule (docs/compatibility-known-defects.md). By
// contrast, "prototype" has no inherited own-object value on a plain `{}` (only
// FUNCTIONS have `.prototype`), so it never collides and passes through unchanged — also
// confirmed live and reproduced by using an ordinary `{}` rather than special-casing it.
//
// The one genuine hazard lives in make_json_row's row-key writes below, not here: the
// DEFAULT header-inference path above can never produce a literal "__proto__" (it always
// gets the "_NaN" rename), so the only reachable path is an explicit `opts.header` array
// containing "__proto__" verbatim — see setJsonRowKey.
function sheetToJson(sheet, opts) {
  if (sheet == null || sheet['!ref'] == null) return [];
  let header = 0;
  let offset = 1;
  const hdr = [];
  const o = opts || {};
  const range = o.range != null ? o.range : sheet['!ref'];
  if (o.header === 1) header = 1;
  else if (o.header === 'A') header = 2;
  else if (Array.isArray(o.header)) header = 3;
  else if (o.header == null) header = 0;
  let r;
  switch (typeof range) {
    case 'string':
      r = safeDecodeRange(range);
      break;
    case 'number':
      r = safeDecodeRange(sheet['!ref']);
      r.s.r = range;
      break;
    default:
      r = range;
  }
  checkRangeSize(r);
  if (header > 0) offset = 0;
  const rr = encodeRow(r.s.r);
  const cols = [];
  const out = [];
  let outi = 0;
  const dense = Array.isArray(sheet);
  let R = r.s.r;
  const header_cnt = {};
  if (dense && !sheet[R]) sheet[R] = [];
  const colinfo = (o.skipHidden && sheet['!cols']) || [];
  const rowinfo = (o.skipHidden && sheet['!rows']) || [];
  for (let C = r.s.c; C <= r.e.c; ++C) {
    if ((colinfo[C] || {}).hidden) continue;
    cols[C] = encodeCol(C);
    let val = dense ? sheet[R][C] : sheet[cols[C] + rr];
    switch (header) {
      case 1:
        hdr[C] = C - r.s.c;
        break;
      case 2:
        hdr[C] = cols[C];
        break;
      case 3:
        hdr[C] = o.header[C - r.s.c];
        break;
      default: {
        if (val == null) val = { w: '__EMPTY', t: 's' };
        const v = formatCell(val, null, o);
        let vv = v;
        let counter = header_cnt[v] || 0;
        if (!counter) header_cnt[v] = 1;
        else {
          do {
            vv = v + '_' + counter++;
          } while (header_cnt[vv]);
          header_cnt[v] = counter;
          header_cnt[vv] = 1;
        }
        hdr[C] = vv;
      }
    }
  }
  for (R = r.s.r + offset; R <= r.e.r; ++R) {
    if ((rowinfo[R] || {}).hidden) continue;
    const row = makeJsonRow(sheet, r, R, cols, header, hdr, dense, o);
    if (row.isempty === false || (header === 1 ? o.blankrows !== false : !!o.blankrows)) out[outi++] = row.row;
  }
  out.length = outi;
  return out;
}

// A caller-supplied `opts.header` array is the one reachable way a literal "__proto__"
// (or any other string) becomes hdr[C] verbatim (see sheetToJson's doc comment above).
// Plain `row[hdr[C]] = v` on that key invokes Object.prototype's inherited __proto__
// accessor instead of creating a normal own property — confirmed live against the oracle
// in two distinct ways: (1) a PRIMITIVE value is silently dropped (the setter no-ops for
// non-object values, so the column's data is lost with no error), and (2) an OBJECT value
// (e.g. a Date cell under cellDates:true) reassigns the ROW's own [[Prototype]] to that
// object (`row instanceof Date === true`, `Object.keys(row).length === 0` — a genuine
// prototype-corruption hazard on that specific row object, not the global
// Object.prototype, which stays clean either way). Same Object.defineProperty precedent
// as book_append_sheet above: it always creates a normal own data property regardless of
// key name — same key order, same enumerability, zero shape divergence from the oracle
// for every OTHER key, and for "__proto__" itself it retains the data as literal own data
// instead of losing it (primitive case) or corrupting the row (object case). See
// docs/xlsx-security-model.md and compat/differential/xlsx-utils.test.mjs's
// sheet_to_json dangerous-header fixtures.
function setJsonRowKey(row, key, value) {
  Object.defineProperty(row, key, { value, writable: true, enumerable: true, configurable: true });
}

function makeJsonRow(sheet, r, R, cols, header, hdr, dense, o) {
  const rr = encodeRow(R);
  const defval = o.defval;
  const raw = o.raw || !Object.prototype.hasOwnProperty.call(o, 'raw');
  let isempty = true;
  const row = header === 1 ? [] : {};
  if (header !== 1) {
    Object.defineProperty(row, '__rowNum__', { value: R, enumerable: false });
  }
  // header === 1 rows are arrays keyed by a numeric index (never "__proto__"), so plain
  // assignment is always safe there and matches the oracle exactly; every other header
  // mode can carry a caller-supplied string key and needs setJsonRowKey.
  function setRow(key, value) {
    if (header === 1) row[key] = value;
    else setJsonRowKey(row, key, value);
  }
  if (!dense || sheet[R]) {
    for (let C = r.s.c; C <= r.e.c; ++C) {
      const val = dense ? sheet[R][C] : sheet[cols[C] + rr];
      if (val === undefined || val.t === undefined) {
        if (defval === undefined) continue;
        if (hdr[C] != null) setRow(hdr[C], defval);
        continue;
      }
      let v = val.v;
      switch (val.t) {
        case 'z':
          if (v == null) break;
          continue;
        case 'e':
          v = v == 0 ? null : undefined;
          break;
        case 's':
        case 'd':
        case 'b':
        case 'n':
          break;
        default:
          throw new Error('unrecognized type ' + val.t);
      }
      if (hdr[C] != null) {
        if (v == null) {
          if (val.t === 'e' && v === null) setRow(hdr[C], null);
          else if (defval !== undefined) setRow(hdr[C], defval);
          else if (raw && v === null) setRow(hdr[C], null);
          else continue;
        } else {
          setRow(hdr[C], raw && (val.t !== 'n' || (val.t === 'n' && o.rawNumbers !== false)) ? v : formatCell(val, v, o));
        }
        if (v != null) isempty = false;
      }
    }
  }
  return { row, isempty };
}

// ---- sheet_to_formulae ----
//
// Independent port of the oracle's sheet_to_formulae(sheet): walks !ref in row-major
// order, emitting one "REF=value" string per non-empty cell. Branch order was confirmed
// against a live oracle run, not guessed — value precedence is: array formula range
// (x.F, only the top-left cell of the range, which alone carries x.f, contributes; the
// key becomes the range itself) > formula text (x.f) > stub skip (x.t=='z') > numeric
// value (x.t=='n') > boolean > CACHED DISPLAY STRING (x.w, checked BEFORE x.t=='s' — a
// string cell with a `.w` that differs from `.v` emits the `.w`, not `.v`) > skip if `.v`
// is undefined > quoted string (x.t=='s') > String(x.v) fallback (e.g. a bare Date cell
// with neither `.f` nor `.w` renders via Date.prototype.toString(), confirmed live — not
// an ISO string or a serial number). A reversed !ref (s > e after safe_decode_range,
// which never swaps) makes the loop body never run, returning [] — confirmed live, not
// a special case in this port. checkRangeSize (see ./internal/range-guard.cjs) rejects
// pathologically large ranges instead of iterating them — a fix added after this
// function's initial Phase 1B-2A implementation, once a crafted full-grid !ref was
// confirmed live to make sheet_to_csv (which walks !ref the same way) not return within
// 25s on the real oracle.
function sheetToFormulae(sheet) {
  if (sheet == null || sheet['!ref'] == null) return [];
  const r = safeDecodeRange(sheet['!ref']);
  checkRangeSize(r);
  const dense = Array.isArray(sheet);
  const cols = [];
  for (let C = r.s.c; C <= r.e.c; ++C) cols[C] = encodeCol(C);
  const cmds = [];
  for (let R = r.s.r; R <= r.e.r; ++R) {
    const rr = encodeRow(R);
    for (let C = r.s.c; C <= r.e.c; ++C) {
      let y = cols[C] + rr;
      const x = dense ? (sheet[R] || [])[C] : sheet[y];
      if (x === undefined) continue;
      let val;
      if (x.F != null) {
        y = x.F;
        if (!x.f) continue;
        val = x.f;
        if (y.indexOf(':') === -1) y = y + ':' + y;
      }
      if (x.f != null) val = x.f;
      else if (x.t === 'z') continue;
      else if (x.t === 'n' && x.v != null) val = '' + x.v;
      else if (x.t === 'b') val = x.v ? 'TRUE' : 'FALSE';
      else if (x.w !== undefined) val = "'" + x.w;
      else if (x.v === undefined) continue;
      else if (x.t === 's') val = "'" + x.v;
      else val = '' + x.v;
      cmds.push(y + '=' + val);
    }
  }
  return cmds;
}

// ---- sheet_to_csv / sheet_to_txt ----
//
// Independent port of the oracle's make_csv_row/sheet_to_csv/sheet_to_txt.
//
// make_csv_row: a cell's text comes from format_cell(val, null, o) UNLESS o.rawNumbers
// and the cell is numeric (uses the raw value's plain string form then); quoting is
// triggered by the field-separator char code, the record-separator char code, a literal
// double-quote, or o.forceQuotes. A rendered value of exactly "ID" always gets quoted (a
// SYLK-file-detection legacy in the real oracle, confirmed live, reproduced as-is). A
// formula cell without a cached/array value (`.f` set, `.F` NOT set) renders as
// "=<formula>", quoted only if it contains a literal comma (not FS) — confirmed live.
// blankrows:false skips an all-empty row entirely (no cell had a `.v`), and the row
// separator only precedes an actually-EMITTED row (a skipped row doesn't consume a
// leading separator) — confirmed live via aoa_to_sheet([[1],[],[3]]).
//
// sheet_to_csv mutates its `opts` argument: sets o.dense (Array.isArray(sheet)) for the
// duration of the call, then deletes it — confirmed live (an opts object with a
// pre-existing `dense` key comes back WITHOUT that key afterward). o.strip builds
// `new RegExp((FS=="|" ? "\\|" : FS)+"+$")` — only `|` is escaped; any other FS is used
// RAW as a regex fragment, so e.g. FS:"." strips the entire row (matches "any char,
// greedy") and FS:"(" throws a native SyntaxError (invalid regex) — both confirmed live,
// reproduced as-is with no extra escaping added.
const CSV_QUOTE_RE = /"/g;

function makeCsvRow(sheet, r, R, cols, fs, rs, FS, o) {
  let isempty = true;
  const row = [];
  const rr = encodeRow(R);
  for (let C = r.s.c; C <= r.e.c; ++C) {
    if (!cols[C]) continue;
    const val = o.dense ? (sheet[R] || [])[C] : sheet[cols[C] + rr];
    let txt;
    if (val == null) txt = '';
    else if (val.v != null) {
      isempty = false;
      txt = '' + (o.rawNumbers && val.t === 'n' ? val.v : formatCell(val, null, o));
      for (let i = 0, cc = 0; i !== txt.length; ++i) {
        cc = txt.charCodeAt(i);
        if (cc === fs || cc === rs || cc === 34 || o.forceQuotes) {
          txt = '"' + txt.replace(CSV_QUOTE_RE, '""') + '"';
          break;
        }
      }
      if (txt === 'ID') txt = '"ID"';
    } else if (val.f != null && !val.F) {
      isempty = false;
      txt = '=' + val.f;
      if (txt.indexOf(',') >= 0) txt = '"' + txt.replace(CSV_QUOTE_RE, '""') + '"';
    } else txt = '';
    row.push(txt);
  }
  if (o.blankrows === false && isempty) return null;
  return row.join(FS);
}

function sheetToCsv(sheet, opts) {
  const out = [];
  const o = opts == null ? {} : opts;
  if (sheet == null || sheet['!ref'] == null) return '';
  const r = safeDecodeRange(sheet['!ref']);
  checkRangeSize(r);
  const FS = o.FS !== undefined ? o.FS : ',';
  const fs = FS.charCodeAt(0);
  const RS = o.RS !== undefined ? o.RS : '\n';
  const rs = RS.charCodeAt(0);
  const endregex = new RegExp((FS === '|' ? '\\|' : FS) + '+$');
  const cols = [];
  o.dense = Array.isArray(sheet);
  const colinfo = (o.skipHidden && sheet['!cols']) || [];
  const rowinfo = (o.skipHidden && sheet['!rows']) || [];
  for (let C = r.s.c; C <= r.e.c; ++C) {
    if (!(colinfo[C] || {}).hidden) cols[C] = encodeCol(C);
  }
  let w = 0;
  for (let R = r.s.r; R <= r.e.r; ++R) {
    if ((rowinfo[R] || {}).hidden) continue;
    let row = makeCsvRow(sheet, r, R, cols, fs, rs, FS, o);
    if (row == null) continue;
    if (o.strip) row = row.replace(endregex, '');
    if (row || o.blankrows !== false) out.push((w++ ? RS : '') + row);
  }
  delete o.dense;
  return out.join('');
}

// UTF-16LE-encoding codepage 1200 needs no lookup table (unlike e.g. codepage 932) —
// verified byte-exact against the real oracle's codepage encoder across ASCII/BMP/
// astral/lone-surrogate cases, so this package implements it directly rather than
// taking on the separate "codepage" npm package as another dependency just for this.
function utf16leEncode(str) {
  let out = '';
  for (let i = 0; i < str.length; ++i) {
    const cu = str.charCodeAt(i);
    out += String.fromCharCode(cu & 0xff) + String.fromCharCode((cu >> 8) & 0xff);
  }
  return out;
}

// sheet_to_txt sets opts.FS='\t'/opts.RS='\n' on the CALLER'S opts object (mutates it in
// place, confirmed live — an opts:{} object comes back as {FS:'\t',RS:'\n'} after the
// call), then delegates to sheet_to_csv. Unless opts.type === 'string', the oracle
// UTF-16LE-encodes the result with a leading BOM — but ONLY when its internal
// "$cptable" codepage support happens to be loaded, which differs by how the real
// oracle package itself is reached: confirmed live that both `require('xlsx')` and a
// bare `import 'xlsx'` in Node ESM resolve to its CJS build (no `exports` map on the
// real package, so ESM falls back to `main`), which auto-loads codepage support — so
// BOM+UTF-16LE is what any normal Node consumer of the real "xlsx" package name
// actually observes by default, regardless of require/import. Only a bundler-resolved
// deep import of the oracle's separate xlsx.mjs file (not reachable through the
// "xlsx" package name) sees the opposite default. This package has one canonical
// sheet_to_txt implementation shared by its own CJS and ESM entrypoints, so it always
// matches the require/bare-import default.
function sheetToTxt(sheet, opts) {
  const o = opts || {};
  o.FS = '\t';
  o.RS = '\n';
  const s = sheetToCsv(sheet, o);
  if (o.type === 'string') return s;
  return String.fromCharCode(255) + String.fromCharCode(254) + utf16leEncode(s);
}

// ---- sheet_to_html ----
//
// Independent port of the oracle's sheet_to_html/make_html_row/make_html_preamble.
// opts.header/opts.footer here are the HTML document prefix/suffix strings (default
// HTML_BEGIN/HTML_END) — an entirely different meaning from sheet_to_json's opts.header
// (JSON key-derivation mode); the two option interfaces are NOT related despite sharing a
// field name, matching the oracle's own separate Sheet2HTMLOpts/Sheet2JSONOpts. Merged
// cells: a non-top-left cell inside a merge is skipped entirely (continue), the top-left
// gets rowspan/colspan. checkRangeSize (see ./internal/range-guard.cjs) guards the same
// !ref-rectangle walk-cost DoS already measured for sheet_to_formulae/sheet_to_csv/
// sheet_to_json — reuses that measurement basis, no separate one needed. Uses the PUBLIC
// decodeRange (matching the oracle's own call site — sheet_to_html calls decode_range,
// not safe_decode_range), so a malformed !ref throws here exactly as it does on the
// oracle, unlike sheet_to_csv/sheet_to_json which use the lenient internal parser.
// o.dense is set but deliberately never deleted (matches the oracle exactly -- confirmed
// live sheet_to_html leaks `dense` onto a caller's own opts object, unlike sheet_to_csv
// which does `delete o.dense`) -- an opts-mutation-fidelity quirk, not a bug to "fix".
//
// SECURITY: three distinct HTML-injection-shaped findings from reading + live-probing the
// oracle's source, each handled differently -- see docs/xlsx-security-model.md for the
// full writeup:
// 1. FIXED -- data-t/data-v/data-z/id (both the per-cell id and opts.id, table-level and
//    per-cell) are built via raw string concatenation with NO escaping on the oracle
//    (confirmed live: a cell value or opts.id containing a `"` breaks out of the
//    attribute and injects an arbitrary onXXX handler -- e.g. a cell.v of
//    `x" onmouseover="alert(1)" y="` produces a LIVE onmouseover handler on the real
//    oracle's output). escapeHtmlAttr (below) is used for every attribute value this
//    function builds.
// 2. FIXED -- cell.l.Target is embedded into `href="..."` with no scheme check (confirmed
//    live: a `javascript:` Target produces a clickable, code-executing link on the real
//    oracle -- quote-escaping alone does NOT fix this, since no quote is needed to make a
//    href value dangerous). isSafeHrefTarget (below) allow-lists http(s)/mailto/tel/ftp/
//    relative/fragment targets; anything else renders as plain text, no <a> wrapper.
// 3. FIXED BY DEFAULT -- cell.h is a documented raw-HTML rich-text field, but it is still
//    caller-controlled markup. Escape it by default; raw passthrough requires the explicit
//    `rawHtml: true` opt-in below. This keeps the rich-text escape hatch without rendering
//    an untrusted value as markup accidentally.
const HTML_ENTITY_MAP = { '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&apos;', '"': '&quot;' };
const HTML_DANGEROUS_CHARS_RE = /[&<>'"]/g;
const HTML_CONTROL_CHARS_RE = /[\u0000-\u001f]/g;

function hexEntity(ch) {
  return '&#x' + ('000' + ch.charCodeAt(0).toString(16)).slice(-4) + ';';
}

// Matches the oracle's own escapehtml exactly (text-content context: '\n' -> '<br/>').
function escapeHtmlText(text) {
  return String(text).replace(HTML_DANGEROUS_CHARS_RE, (ch) => HTML_ENTITY_MAP[ch]).replace(/\n/g, '<br/>').replace(HTML_CONTROL_CHARS_RE, hexEntity);
}

// Attribute-value context: deliberately does NOT substitute '\n' with '<br/>' -- a literal
// '<br/>' tag injected into a quoted attribute value is itself an escaping bug (the oracle
// never does this either, since it never escapes attribute values at all; this function is
// new, not a port of anything in xlsx.js).
function escapeHtmlAttr(text) {
  return String(text).replace(HTML_DANGEROUS_CHARS_RE, (ch) => HTML_ENTITY_MAP[ch]).replace(HTML_CONTROL_CHARS_RE, hexEntity);
}

const HTML_TAG_PRESERVE_WS_RE = /(^\s|\s$|\n)/;

// Matches the oracle's own writextag, plus escaping every attribute value (the fix) --
// unescaped attribute-value concatenation was the oracle's own bug, not a feature.
function buildHtmlTag(tag, inner, attrs) {
  let attrStr = '';
  if (attrs) {
    for (const k of Object.keys(attrs)) attrStr += ' ' + k + '="' + escapeHtmlAttr(attrs[k]) + '"';
  }
  if (inner != null) {
    return '<' + tag + attrStr + (HTML_TAG_PRESERVE_WS_RE.test(inner) ? ' xml:space="preserve"' : '') + '>' + inner + '</' + tag + '>';
  }
  return '<' + tag + attrStr + '/>';
}

const SAFE_HREF_SCHEME_RE = /^(?:https?|mailto|tel|ftp):/i;
function isSafeHrefTarget(target) {
  if (typeof target !== 'string' || target === '') return false;
  if (target.charAt(0) === '#' || target.charAt(0) === '/') return true;
  if (target.charAt(0) === '.' && (target.charAt(1) === '/' || (target.charAt(1) === '.' && target.charAt(2) === '/'))) return true;
  return SAFE_HREF_SCHEME_RE.test(target);
}

function makeHtmlRow(sheet, r, R, o) {
  const M = sheet['!merges'] || [];
  const oo = [];
  for (let C = r.s.c; C <= r.e.c; ++C) {
    let RS = 0;
    let CS = 0;
    for (let j = 0; j < M.length; ++j) {
      if (M[j].s.r > R || M[j].s.c > C) continue;
      if (M[j].e.r < R || M[j].e.c < C) continue;
      if (M[j].s.r < R || M[j].s.c < C) {
        RS = -1;
        break;
      }
      RS = M[j].e.r - M[j].s.r + 1;
      CS = M[j].e.c - M[j].s.c + 1;
      break;
    }
    if (RS < 0) continue;
    const coord = encodeCell({ r: R, c: C });
    const cell = o.dense ? (sheet[R] || [])[C] : sheet[coord];
    const cellText = cell && cell.h && o.rawHtml === true ? cell.h : (cell && cell.w || (cell && formatCell(cell), cell && cell.w) || '');
    const renderedText = cell && cell.h && o.rawHtml === true ? cellText : escapeHtmlText(cellText);
    let w = (cell && cell.v != null && renderedText) || '';
    const sp = {};
    if (RS > 1) sp.rowspan = RS;
    if (CS > 1) sp.colspan = CS;
    if (o.editable) {
      w = '<span contenteditable="true">' + w + '</span>';
    } else if (cell) {
      sp['data-t'] = cell.t || 'z';
      if (cell.v != null) sp['data-v'] = cell.v;
      if (cell.z != null) sp['data-z'] = cell.z;
      if (cell.l && (cell.l.Target || '#').charAt(0) !== '#') {
        w = isSafeHrefTarget(cell.l.Target) ? '<a href="' + escapeHtmlAttr(cell.l.Target) + '">' + w + '</a>' : w;
      }
    }
    sp.id = (o.id || 'sjs') + '-' + coord;
    oo.push(buildHtmlTag('td', w, sp));
  }
  return '<tr>' + oo.join('') + '</tr>';
}

function makeHtmlPreamble(o) {
  return '<table' + (o && o.id ? ' id="' + escapeHtmlAttr(o.id) + '"' : '') + '>';
}

const HTML_BEGIN = '<html><head><meta charset="utf-8"/><title>SheetJS Table Export</title></head><body>';
const HTML_END = '</body></html>';

function sheetToHtml(sheet, opts) {
  const o = opts || {};
  const header = o.header != null ? o.header : HTML_BEGIN;
  const footer = o.footer != null ? o.footer : HTML_END;
  const out = [header];
  const r = decodeRange(sheet['!ref']);
  checkRangeSize(r);
  o.dense = Array.isArray(sheet);
  out.push(makeHtmlPreamble(o));
  for (let R = r.s.r; R <= r.e.r; ++R) out.push(makeHtmlRow(sheet, r, R, o));
  out.push('</table>' + footer);
  return out.join('');
}

// ---- cell_set_hyperlink / cell_set_internal_link ----
//
// Independent port. A falsy target (including "", 0, false, NaN, null, undefined)
// deletes any existing .l instead of setting one — confirmed live across that whole
// matrix. No null-check on `cell` itself: a null/undefined cell throws the same native
// TypeError the oracle throws ("Cannot set properties of null (setting 'l')"),
// reproduced by deliberately NOT adding a defensive guard.
function cellSetHyperlink(cell, target, tooltip) {
  if (!target) {
    delete cell.l;
  } else {
    cell.l = { Target: target };
    if (tooltip) cell.l.Tooltip = tooltip;
  }
  return cell;
}

// "#" + range is concatenated unconditionally — range=null/undefined produce the
// (truthy) strings "#null"/"#undefined" as the target, not a no-op — confirmed live.
function cellSetInternalLink(cell, range, tooltip) {
  return cellSetHyperlink(cell, '#' + range, tooltip);
}

// ---- cell_add_comment ----
//
// Independent port. Returns undefined, NOT the cell — confirmed live (no `return`
// statement in the oracle source), unlike cell_set_hyperlink/cell_set_number_format,
// which do return their cell argument. `author` defaults to "SheetJS" only when falsy
// (matching `author||"SheetJS"`, so an explicit "" also defaults); `text` has no such
// default and is stored verbatim, including null/undefined.
function cellAddComment(cell, text, author) {
  if (!cell.c) cell.c = [];
  cell.c.push({ t: text, a: author || 'SheetJS' });
}

// ---- sheet_set_array_formula ----
//
// Independent port. `range` as a STRING is used verbatim for the stored `.F` value
// (e.g. "A1:A1" stays "A1:A1"); as an OBJECT it goes through encode_range, which
// collapses a single-cell range to its short form ("A1", no colon) — confirmed live
// these two produce different `.F` strings for the same logical range. Every cell in
// the range gets t:'n' / F / v-deleted; only the top-left cell additionally gets `.f`
// (unconditionally, even when `formula` is undefined — matching `cell.f = formula` with
// no guard; the `f` key exists with value undefined rather than being omitted,
// confirmed live) and `.D` (only when `dynamic` is truthy — omitted entirely otherwise,
// never set to false). Never touches `!ref` — confirmed live, even when the range
// extends past an existing one. `range`/`formula` are never used as object keys
// anywhere in this function, so there is no __proto__-style hazard to guard against.
function sheetSetArrayFormula(ws, range, formula, dynamic) {
  const rng = typeof range !== 'string' ? range : safeDecodeRange(range);
  const rngstr = typeof range === 'string' ? range : encodeRange(range);
  for (let R = rng.s.r; R <= rng.e.r; ++R) {
    for (let C = rng.s.c; C <= rng.e.c; ++C) {
      const cell = wsGetCellStub(ws, encodeCell({ r: R, c: C || 0 }));
      cell.t = 'n';
      cell.F = rngstr;
      delete cell.v;
      if (R === rng.s.r && C === rng.s.c) {
        cell.f = formula;
        if (dynamic) cell.D = true;
      }
    }
  }
  return ws;
}

// ---- sheet_add_dom / table_to_sheet / table_to_book ----
//
// Independent port of the oracle's sheet_add_dom/parse_dom_table/table_to_book. These are
// the "BROWSER ONLY!" functions (per the oracle's own .d.ts comment) — they consume a
// DOM-like <table> element, not spreadsheet data. `packages/xlsx` imports no DOM library
// at runtime (still zero runtime dependencies beyond ssf): this port calls only the DOM
// methods the oracle itself calls (getElementsByTagName, .children, hasAttribute,
// getAttribute, .innerHTML, .style, .ownerDocument.defaultView.getComputedStyle) on
// whatever object it's handed, so any real DOM element (browser or a devDependency like
// jsdom used only for testing) — or a hand-built duck-typed stub satisfying the same
// shape — works identically. `opts.dense`'s DENSE-global override
// (`if(DENSE != null) opts.dense = DENSE;` in the oracle) is omitted: `DENSE` is a
// module-level variable the shipped xlsx.js declares once (`var DENSE = null;`) and never
// reassigns anywhere in the bundle — confirmed by grepping the whole file — so that line
// is dead code in the real npm package, not a behavior to reproduce.
function isDomElementHidden(element) {
  let display = '';
  const getComputedStyleFn = getComputedStyleFunction(element);
  if (getComputedStyleFn) display = getComputedStyleFn(element).getPropertyValue('display');
  if (!display) display = element.style && element.style.display;
  return display === 'none';
}

function getComputedStyleFunction(element) {
  if (element.ownerDocument.defaultView && typeof element.ownerDocument.defaultView.getComputedStyle === 'function') {
    return element.ownerDocument.defaultView.getComputedStyle;
  }
  if (typeof getComputedStyle === 'function') return getComputedStyle;
  return null;
}

// Matches the oracle's own fuzzynum exactly: strips thousands separators/currency/percent
// (tracking percent's implied /100 as a weight) and a trailing/wrapping "(...)" for
// negative-in-parens accounting notation, retrying Number() after each strip.
function fuzzyNum(s) {
  let v = Number(s);
  if (!isNaN(v)) return isFinite(v) ? v : NaN;
  if (!/\d/.test(s)) return v;
  let wt = 1;
  let ss = s
    .replace(/([\d]),([\d])/g, '$1$2')
    .replace(/[$]/g, '')
    .replace(/[%]/g, () => {
      wt *= 100;
      return '';
    });
  if (!isNaN((v = Number(ss)))) return v / wt;
  ss = ss.replace(/[(](.*)[)]/, ($$, $1) => {
    wt = -wt;
    return $1;
  });
  if (!isNaN((v = Number(ss)))) return v / wt;
  return v;
}

const FUZZY_DATE_LOWER_MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];

// Matches the oracle's own fuzzydate exactly. Uses .getYear() (not .getFullYear()) for
// its year-range sanity check, deliberately reproducing that exact legacy API choice —
// .getYear() returns year-1900 (e.g. 117 for 2017), which is why the "0 < y < 8099" bound
// below is meaningless as a literal year check but is what the oracle actually tests.
function fuzzyDate(s) {
  const o = new Date(s);
  const n = new Date(NaN);
  const y = o.getYear();
  const m = o.getMonth();
  const d = o.getDate();
  if (isNaN(d)) return n;
  let lower = s.toLowerCase();
  if (lower.match(/jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec/)) {
    lower = lower.replace(/[^a-z]/g, '').replace(/([^a-z]|^)[ap]m?([^a-z]|$)/, '');
    if (lower.length > 3 && FUZZY_DATE_LOWER_MONTHS.indexOf(lower) === -1) return n;
  } else if (lower.match(/[a-z]/)) {
    return n;
  }
  if (y < 0 || y > 8099) return n;
  if ((m > 0 || d > 1) && y !== 101) return o;
  if (s.match(/[^-0-9:,/\\]/)) return n;
  return o;
}

// Matches the oracle's own parseDate for the ONLY call shape sheetAddDom ever uses
// (single-argument, so fixdate is always undefined and the timezone-shift branches never
// fire) — the oracle's own function also has non-good_pd fallback branches for engines
// where `new Date('2017-02-19T19:06:09.000Z').getFullYear() !== 2017`; confirmed live
// that condition is true in this package's target environments (any real Node.js), so
// those branches are dead code here too, not a behavior being dropped.
function fuzzyParseDate(str) {
  return new Date(str);
}

function sheetAddDom(ws, table, opts) {
  const o = opts || {};
  let orR = 0;
  let orC = 0;
  if (o.origin != null) {
    if (typeof o.origin === 'number') {
      orR = o.origin;
    } else {
      const origin = typeof o.origin === 'string' ? decodeCell(o.origin) : o.origin;
      orR = origin.r;
      orC = origin.c;
    }
  }
  const rows = table.getElementsByTagName('tr');
  const sheetRows = Math.min(o.sheetRows || 10000000, rows.length);
  const range = { s: { c: 0, r: 0 }, e: { c: orC, r: orR } };
  if (ws['!ref']) {
    const existingRange = safeDecodeRange(ws['!ref']);
    range.s.r = Math.min(range.s.r, existingRange.s.r);
    range.s.c = Math.min(range.s.c, existingRange.s.c);
    range.e.r = Math.max(range.e.r, existingRange.e.r);
    range.e.c = Math.max(range.e.c, existingRange.e.c);
    if (orR === -1) range.e.r = orR = existingRange.e.r + 1;
  }
  const merges = [];
  const rowinfo = ws['!rows'] || (ws['!rows'] = []);
  let R = 0;
  let C = 0;
  let domR = 0;
  if (!ws['!cols']) ws['!cols'] = [];
  for (; domR < rows.length && R < sheetRows; ++domR) {
    const row = rows[domR];
    if (isDomElementHidden(row)) {
      if (o.display) continue;
      rowinfo[R] = { hidden: true };
    }
    const elts = row.children;
    C = 0;
    for (let domC = 0; domC < elts.length; ++domC) {
      const elt = elts[domC];
      if (o.display && isDomElementHidden(elt)) continue;
      let v = elt.hasAttribute('data-v') ? elt.getAttribute('data-v') : elt.hasAttribute('v') ? elt.getAttribute('v') : htmlDecode(elt.innerHTML);
      const z = elt.getAttribute('data-z') || elt.getAttribute('z');
      for (let midx = 0; midx < merges.length; ++midx) {
        const m = merges[midx];
        if (m.s.c === C + orC && m.s.r < R + orR && R + orR <= m.e.r) {
          C = m.e.c + 1 - orC;
          midx = -1;
        }
      }
      const CS = +elt.getAttribute('colspan') || 1;
      const RS = +elt.getAttribute('rowspan') || 1;
      if (RS > 1 || CS > 1) {
        merges.push({ s: { r: R + orR, c: C + orC }, e: { r: R + orR + (RS || 1) - 1, c: C + orC + (CS || 1) - 1 } });
      }
      let cellObj = { t: 's', v };
      const cellT = elt.getAttribute('data-t') || elt.getAttribute('t') || '';
      if (v != null) {
        if (v.length === 0) {
          cellObj.t = cellT || 'z';
        } else if (o.raw || v.trim().length === 0 || cellT === 's') {
          // leave as string, matching the oracle's empty-block branch
        } else if (v === 'TRUE') {
          cellObj = { t: 'b', v: true };
        } else if (v === 'FALSE') {
          cellObj = { t: 'b', v: false };
        } else if (!isNaN(fuzzyNum(v))) {
          cellObj = { t: 'n', v: fuzzyNum(v) };
        } else if (!isNaN(fuzzyDate(v).getDate())) {
          cellObj = { t: 'd', v: fuzzyParseDate(v) };
          if (!o.cellDates) cellObj = { t: 'n', v: datenum(cellObj.v) };
          cellObj.z = o.dateNF || 'm/d/yy';
        }
      }
      if (cellObj.z === undefined && z != null) cellObj.z = z;
      let l = '';
      const aElts = elt.getElementsByTagName('A');
      if (aElts && aElts.length) {
        for (let ai = 0; ai < aElts.length; ++ai) {
          if (aElts[ai].hasAttribute('href')) {
            l = aElts[ai].getAttribute('href');
            if (l.charAt(0) !== '#') break;
          }
        }
      }
      if (l && l.charAt(0) !== '#') cellObj.l = { Target: l };
      if (o.dense) {
        if (!ws[R + orR]) ws[R + orR] = [];
        ws[R + orR][C + orC] = cellObj;
      } else {
        ws[encodeCell({ c: C + orC, r: R + orR })] = cellObj;
      }
      if (range.e.c < C + orC) range.e.c = C + orC;
      C += CS;
    }
    ++R;
  }
  if (merges.length) ws['!merges'] = (ws['!merges'] || []).concat(merges);
  range.e.r = Math.max(range.e.r, R - 1 + orR);
  ws['!ref'] = encodeRange(range);
  // Matches the oracle's own comma-operator line exactly (including its own comment
  // about the tradeoff): `!fullref` is set whenever R reached sheetRows at loop exit —
  // which, since sheetRows = min(opts.sheetRows||10000000, rows.length), is the NORMAL
  // exit case whenever no opts.display-hidden row was skipped (R and domR both end up
  // equal to rows.length then) — NOT only when opts.sheetRows truncated the parse. A
  // opts.display-skipped hidden row can make R fall short of domR, in which case
  // !fullref is NOT set. This is a genuinely confusing but precise boundary condition,
  // confirmed live rather than assumed from the name alone.
  if (R >= sheetRows) {
    range.e.r = rows.length - domR + R - 1 + orR;
    ws['!fullref'] = encodeRange(range);
  }
  return ws;
}

// Matches the oracle's own htmldecode exactly — strips leading/trailing whitespace,
// collapses interior whitespace runs, turns <br> into '\n', strips all other tags, then
// decodes a small fixed entity set (nbsp/middot/quot/apos/gt/lt/amp).
const HTML_DECODE_ENTITIES = [
  [/&nbsp;/gi, ' '],
  [/&middot;/gi, '·'],
  [/&quot;/gi, '"'],
  [/&apos;/gi, "'"],
  [/&gt;/gi, '>'],
  [/&lt;/gi, '<'],
  [/&amp;/gi, '&'],
];
function htmlDecode(str) {
  let o = str
    .replace(/^[\t\n\r ]+/, '')
    .replace(/[\t\n\r ]+$/, '')
    .replace(/>\s+/g, '>')
    .replace(/\s+</g, '<')
    .replace(/[\t\n\r ]+/g, ' ')
    .replace(/<\s*[bB][rR]\s*\/?>/g, '\n')
    .replace(/<[^>]*>/g, '');
  for (const [re, repl] of HTML_DECODE_ENTITIES) o = o.replace(re, repl);
  return o;
}

function parseDomTable(table, opts) {
  const o = opts || {};
  const ws = o.dense ? [] : {};
  return sheetAddDom(ws, table, opts);
}

// The oracle's own table_to_book delegates to a shared sheet_to_workbook(sheet, opts)
// helper (`var n = opts.sheet || "Sheet1"; var sheets = {}; sheets[n] = sheet; return
// {SheetNames:[n], Sheets:sheets};`) that is NOT book_new()/book_append_sheet() — it skips
// all of book_append_sheet's own validation (name length/forbidden characters/uniqueness)
// entirely, so this port deliberately does NOT route through bookAppendSheet either
// (that would add oracle-incompatible throws for e.g. a >31-char opts.sheet name).
// SECURITY: confirmed live the oracle's own `sheets[n] = sheet` is the exact same
// prototype-corruption hazard book_append_sheet had in Phase 1A (opts.sheet:'"__proto__"'
// reassigns the resulting WorkBook's own Sheets prototype instead of storing a
// retrievable entry) — reachable here since opts.sheet is directly caller-controlled.
// Fixed the same way, with Object.defineProperty, from this function's first
// implementation rather than shipping the hazard and fixing it in a later commit (unlike
// sheet_to_html's escaping fix): the fix is a single line with no wider ripple, so there
// is no benefit to a separate "vulnerable" commit here.
function sheetToWorkbookSafe(sheet, opts) {
  const n = opts && opts.sheet ? opts.sheet : 'Sheet1';
  const sheets = {};
  Object.defineProperty(sheets, n, { value: sheet, writable: true, enumerable: true, configurable: true });
  return { SheetNames: [n], Sheets: sheets };
}

function tableToBook(table, opts) {
  return sheetToWorkbookSafe(parseDomTable(table, opts), opts);
}

// Every function above is declared with a camelCase internal name (encodeCol, ...) —
// ordinary JS convention within this file — but a plain `{ encode_col: encodeCol }`
// object-literal assignment does NOT rename an already-named function's own `.name`.
// Confirmed against the live oracle: every exported function's `.name` there equals its
// exact snake_case public key (e.g. `XLSX.utils.encode_col.name === "encode_col"`),
// since the oracle source declares its functions with those names directly. Without this
// step, every elixcee export's `.name` would stay the camelCase internal one (a real,
// previously-undetected divergence — reflection-based consumer code, or anything that
// does `Object.values(utils).find(fn => fn.name === "encode_col")`, would silently
// break). `.length` needs no such correction — it's the function's own declared arity,
// unaffected by what name it's assigned under, and already matches (verified per-export
// against the oracle). `Object.defineProperty` with `configurable: true` (matching the
// oracle's own `.name` property descriptor, verified live) — not a plain `fn.name = ...`
// assignment, which silently no-ops since `.name` is non-writable by default.
// Exceptions to "every export's .name equals its exact snake_case public key" — the
// oracle itself breaks that pattern in four ways:
// - `sheet_get_cell: ws_get_cell_stub` assigns the internal helper directly without a
//   wrapper, so `XLSX.utils.sheet_get_cell.name` is genuinely "ws_get_cell_stub"
//   (confirmed live), not "sheet_get_cell".
// - `sheet_to_row_object_array` is a literal alias for the SAME function object as
//   `sheet_to_json` (confirmed live: `U.sheet_to_row_object_array === U.sheet_to_json`),
//   which the oracle declared once as `function sheet_to_json(...)` — aliasing a key to
//   an existing function never renames it, so both keys' `.name` reads "sheet_to_json".
// - `table_to_sheet: parse_dom_table` — same pattern as sheet_get_cell: the internal
//   helper assigned directly, so `.name` reads "parse_dom_table" (confirmed live), not
//   "table_to_sheet". `table_to_book` and `sheet_add_dom`, by contrast, ARE declared with
//   their own exact public names in the oracle's source, so neither needs an override.
// - `read: readSync` — the oracle's top-level `read` export (confirmed live:
//   `XLSX.read.name === "readSync"`, not "read"; `Object.keys(XLSX)` lists `read` third,
//   right after `version`/`parse_xlscfb`, not under `.utils` at all — this package's own
//   flat namespace already diverges from that nesting, a pre-existing Phase 1A decision,
//   not something this override changes).
// All four reproduced as-is per this project's fidelity-over-tidiness rule
// (compat/differential/metadata.test.mjs is what caught the first three; read's own
// differential test — compat/differential/xlsx-read.test.mjs — checks the fourth).
const NAME_OVERRIDES = {
  sheet_get_cell: 'ws_get_cell_stub',
  sheet_to_row_object_array: 'sheet_to_json',
  table_to_sheet: 'parse_dom_table',
  read: 'readSync',
  // Both keys, same target name: the oracle exports ONE function under both `readFile` and
  // `readFileSync`, and its `.name` is "readFileSync" for both (confirmed live). The two
  // entries here are the same function object, so the rename loop below simply sets the
  // same name twice.
  readFile: 'readFileSync',
  readFileSync: 'readFileSync',
  write: 'writeSync',
  // Same one-function-two-keys shape as readFile/readFileSync above, confirmed live
  // against the oracle: `XLSX.writeFile === XLSX.writeFileSync`, `.name` is
  // "writeFileSync" for both.
  writeFile: 'writeFileSync',
  writeFileSync: 'writeFileSync',
};

// nameAs is for FUNCTION exports only — `consts` is a plain data object (no `.name`
// concept; the oracle's own `consts` has no `.name` property either), so the loop below
// skips it rather than adding a property the oracle's export doesn't have.
function nameAs(fn, publicName) {
  Object.defineProperty(fn, 'name', { value: NAME_OVERRIDES[publicName] || publicName, configurable: true });
  return fn;
}

// SHEET_VERY_HIDDEN (underscore before "HIDDEN"), not SHEET_VERYHIDDEN — confirmed live
// against the real oracle's RUNTIME object via Object.getOwnPropertyDescriptor, not
// assumed from its own types/index.d.ts, which actually declares the key as
// `SHEET_VERYHIDDEN` (no underscore) — a genuine mismatch already present in the oracle
// itself between its shipped types and its shipped runtime. Reproduced here matching the
// RUNTIME name, since that's what `XLSX.utils.consts.SHEET_VERY_HIDDEN` actually is at
// call time; see docs/typescript-compatibility.md for the types-side discrepancy.
const consts = { SHEET_VISIBLE: 0, SHEET_HIDDEN: 1, SHEET_VERY_HIDDEN: 2 };

// The object literal below (not an intermediate variable reassigned to
// `module.exports`) is required as-is: Node's ESM loader synthesizes named imports from
// a CJS module via cjs-module-lexer, a static syntax scan that only recognizes this
// exact `module.exports = { ... }` literal-object pattern (or a sequence of
// `module.exports.foo = ...` assignments) — routing through a separate variable first
// breaks that static detection and `import { encode_col } from './index.cjs'` fails at
// load time with "Named export 'encode_col' not found" (confirmed by trying it). The
// nameAs() loop after this assignment is a pure runtime side effect (mutating each
// function's own `.name`), which doesn't affect what cjs-module-lexer already parsed.
//
// Key order matches the real oracle's own `Object.keys(XLSX.utils)` insertion order
// exactly (confirmed live), filtered down to the keys this package currently implements
// — not alphabetical, not grouped by this file's own section comments. A previous
// version of this literal was grouped by function (address utils, then workbook, then
// worksheet mutation, etc.) instead, which happened to never be checked against the
// oracle's actual order until this comparison was added; see
// compat/differential/metadata.test.mjs's key-order assertion.
module.exports = {
  read,
  // Immediately after `read`, matching the oracle's own top-level key order
  // (`Object.keys(XLSX)` is [..., "read", "readFile", "readFileSync", "write", ...] —
  // confirmed live). Like `read`, these are top-level oracle exports, not `utils.*`
  // members, so metadata.test.mjs compares them against `XLSX`, not `XLSX.utils`.
  readFile: readFileSyncImpl,
  readFileSync: readFileSyncImpl,
  write,
  writeFile: writeFileSyncImpl,
  writeFileSync: writeFileSyncImpl,
  encode_col: encodeCol,
  encode_row: encodeRow,
  encode_cell: encodeCell,
  encode_range: encodeRange,
  decode_col: decodeCol,
  decode_row: decodeRow,
  split_cell: splitCell,
  decode_cell: decodeCell,
  decode_range: decodeRange,
  format_cell: formatCell,
  sheet_add_aoa: sheetAddAoa,
  sheet_add_json: sheetAddJson,
  sheet_add_dom: sheetAddDom,
  aoa_to_sheet: aoaToSheet,
  json_to_sheet: jsonToSheet,
  table_to_sheet: parseDomTable,
  table_to_book: tableToBook,
  sheet_to_csv: sheetToCsv,
  sheet_to_txt: sheetToTxt,
  sheet_to_json: sheetToJson,
  sheet_to_html: sheetToHtml,
  sheet_to_formulae: sheetToFormulae,
  sheet_to_row_object_array: sheetToJson,
  sheet_get_cell: sheetGetCell,
  book_new: bookNew,
  book_append_sheet: bookAppendSheet,
  book_set_sheet_visibility: bookSetSheetVisibility,
  cell_set_number_format: cellSetNumberFormat,
  cell_set_hyperlink: cellSetHyperlink,
  cell_set_internal_link: cellSetInternalLink,
  cell_add_comment: cellAddComment,
  sheet_set_array_formula: sheetSetArrayFormula,
  consts,
};

for (const publicName of Object.keys(module.exports)) {
  if (typeof module.exports[publicName] === 'function') nameAs(module.exports[publicName], publicName);
}
