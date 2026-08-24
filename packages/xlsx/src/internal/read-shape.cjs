'use strict';

// Shapes the WASM bridge's raw parsed JSON (crates/elixcee-wasm's readWorkbook output,
// after JSON.parse) into the WorkBook shape @elixcee/xlsx's read() promises. Kept as its
// own module — not inlined into index.cjs's `read` — so both the Node and browser read()
// entry points (see index.cjs and the browser entry added for the "browser" export
// condition) share one implementation instead of drifting.
//
// This module `require`s 'ssf' (via ssf-adapter.cjs) and ./datenum.cjs, both real,
// disclosed dependencies of this package already (see docs/xlsx-architecture.md's "SSF
// backend" decision) — not new ones. That does mean index.browser.mjs's `read`, which
// shares this module for its !rows/!cols/.w/.z/date handling, is no longer reachable in a
// literal bundler-less browser tab beyond what a bundler resolves for it — see
// index.browser.mjs's own doc comment for that disclosed, accepted trade-off.
const { format: ssfFormat, resolveFormatString, isDate } = require('./ssf-adapter.cjs');
const { numdate } = require('./datenum.cjs');
//
// ---- !rows / !cols (Milestone read-item 3) ----
//
// crates/elixcee-wasm emits reader.rs's already-parsed hidden-row/col data as internal
// "!hiddenRows"/"!hiddenCols" keys: 1-based inclusive [start,end] intervals — reader.rs's
// own native shape (see src/reader.rs's WorkbookSheet.hidden_rows/hidden_columns), NOT the
// oracle's own !rows/!cols shape (a 0-based array, sparse — a real hole at every
// non-hidden index, not `undefined`/`null` filler — of {hidden:true} objects, one slot per
// row/col up to the last hidden one). expandHiddenIntervals does that expansion.
//
// opts.cellStyles gates whether !rows/!cols are surfaced at all — confirmed live against
// the real oracle (not assumed): XLSX.read(buf, {type:'buffer'}) NEVER returns !rows/!cols,
// even for a file with real hidden rows, unless the caller also passes
// {cellStyles: true} (see compat/node_modules/xlsx/xlsx.js's parse_ws_xml_cols call site
// and parse_ws_xml_data's `if(opts && opts.cellStyles)` guards around rowinfo/colinfo).
// Skipping this gate would silently diverge from the oracle's default-opts behavior on any
// fixture with hidden rows/cols — so it's threaded through here rather than always-on.
// The internal !hiddenRows/!hiddenCols keys are deleted unconditionally either way — they
// must never leak to a caller regardless of opts.
function expandHiddenIntervals(intervals) {
  let maxIdx = -1;
  for (const [, end] of intervals) {
    const idx = end - 1; // 1-based inclusive end -> 0-based index
    if (idx > maxIdx) maxIdx = idx;
  }
  const out = new Array(maxIdx + 1); // starts fully sparse — only hidden slots get set below
  for (const [start, end] of intervals) {
    for (let i = start - 1; i <= end - 1; ++i) out[i] = { hidden: true };
  }
  return out;
}

// ---- .w / .z / date-typed cells (Milestone read-item 6) ----
//
// crates/elixcee-wasm emits a per-cell "fmtId" (a raw numFmtId integer, only when
// non-zero) and workbook-level "!numFmts" (custom <numFmt> definitions) — reader.rs's raw
// parsed styles.xml data, not yet resolved to an actual format-code STRING or checked for
// date-ness. That resolution — and all of `.w`'s actual number/date/text formatting — is
// done here via the real `ssf` engine (ssf-adapter.cjs), the same one already verified
// byte-identical to the oracle's own bundled engine across 1831 cases
// (compat/differential/ssf-format.test.mjs), rather than a second, unverified
// reimplementation of SSF's format-code parsing in Rust.
//
// Three independently-confirmed-live oracle behaviors this reproduces exactly:
// 1. `.w` is unconditional (present on every cell regardless of any opts) — confirmed
//    live: even a completely unstyled cell gets a `.w` (General-formatted).
// 2. `.z` requires opts.cellNF === true (opts.cellStyles implies it, confirmed live:
//    `if(o.cellStyles) o.cellNF = true` in the oracle's own read() entry). When present,
//    `.z` is ALWAYS a resolved format STRING, even "General" literally — never the raw
//    numFmtId integer.
// 3. `t:'d'` requires opts.cellDates === true AND the resolved format is date-like
//    (isDate) AND the cell is numeric — confirmed live: XLSX.read() never returns a
//    date-typed cell without cellDates, even for an obviously date-formatted numeric
//    cell (numFmtId 14 / "m/d/yy"). The Date object itself is built via numdate(), which
//    (confirmed live against a real date1904 workbook) deliberately does NOT account for
//    date1904 — matching a genuine inconsistency in the real oracle itself, where `.w`
//    DOES shift for a date1904 file but the cellDates `.v` Date object does not. See
//    datenum.cjs's numdate doc comment for the full writeup.
//
// String cells: `.w` is always the literal `.v` (no SSF text-section formatting applied).
// Boolean cells: `.w` is always "TRUE"/"FALSE". Both are deliberate, disclosed scope
// limits — a custom format's 4th ("text") section changing a string/boolean's rendered
// text is a real but rare oracle feature this does not replicate; every fixture this
// package's own tests exercise uses the unstyled/default case, so this is an honest
// omission, not a silently wrong claim of support.
//
// Error cells (t="e"): the WASM bridge emits the same numeric BIFF code the oracle's own
// in-memory model uses for `.v` (see crates/elixcee-wasm's `ExcelError::biff_code` doc
// comment) — confirmed live against the real oracle reading a real Excel-authored `t="e"`
// cell (`XLSX.read()` never puts the display string in `.v`, only in `.w`, derived from
// this exact table). Only the 7 classic codes reader.rs itself recognizes are listed;
// there is no evidence (real fixture or otherwise) elixcee ever emits any other code.
const ERROR_CODES = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
};

function shapeCell(cell, opts, numFmts, date1904) {
  const table = numFmts || {};
  const resolved = resolveFormatString(cell.fmtId || 0, { table });
  delete cell.fmtId; // never leak the raw wire integer

  if (cell.t === 's') {
    cell.w = cell.v;
  } else if (cell.t === 'b') {
    cell.w = cell.v ? 'TRUE' : 'FALSE';
  } else if (cell.t === 'n') {
    cell.w = ssfFormat(resolved, cell.v, { date1904 });
  } else if (cell.t === 'e') {
    cell.w = ERROR_CODES[cell.v];
  }

  if (opts && (opts.cellNF || opts.cellStyles)) {
    cell.z = resolved;
  }

  if (opts && opts.cellDates && cell.t === 'n' && isDate(resolved)) {
    cell.t = 'd';
    cell.v = numdate(cell.v);
  }
}

const CELL_REF_RE = /^[A-Z]+[0-9]+$/;

function shapeSheet(ws, opts, numFmts, date1904) {
  if (ws == null) return ws;
  const hiddenRows = ws['!hiddenRows'];
  const hiddenCols = ws['!hiddenCols'];
  delete ws['!hiddenRows'];
  delete ws['!hiddenCols'];
  if (opts && opts.cellStyles) {
    if (hiddenRows) ws['!rows'] = expandHiddenIntervals(hiddenRows);
    if (hiddenCols) ws['!cols'] = expandHiddenIntervals(hiddenCols);
  }
  for (const key of Object.keys(ws)) {
    if (CELL_REF_RE.test(key)) shapeCell(ws[key], opts, numFmts, date1904);
  }
  return ws;
}

// Mutates and returns `wb` in place — the WASM bridge's JSON.parse output is a fresh
// object with no other owner, so there's no reason to clone before reshaping it.
function shapeWorkBook(wb, opts) {
  const numFmts = wb['!numFmts'];
  const date1904 = !!wb['!date1904'];
  delete wb['!numFmts'];
  delete wb['!date1904'];
  for (const name of wb.SheetNames) shapeSheet(wb.Sheets[name], opts, numFmts, date1904);
  return wb;
}

module.exports = { shapeWorkBook };
