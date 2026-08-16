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
function nameAs(fn, publicName) {
  Object.defineProperty(fn, 'name', { value: publicName, configurable: true });
  return fn;
}

// The object literal below (not an intermediate variable reassigned to
// `module.exports`) is required as-is: Node's ESM loader synthesizes named imports from
// a CJS module via cjs-module-lexer, a static syntax scan that only recognizes this
// exact `module.exports = { ... }` literal-object pattern (or a sequence of
// `module.exports.foo = ...` assignments) — routing through a separate variable first
// breaks that static detection and `import { encode_col } from './index.cjs'` fails at
// load time with "Named export 'encode_col' not found" (confirmed by trying it). The
// nameAs() loop after this assignment is a pure runtime side effect (mutating each
// function's own `.name`), which doesn't affect what cjs-module-lexer already parsed.
module.exports = {
  encode_col: encodeCol,
  decode_col: decodeCol,
  encode_row: encodeRow,
  decode_row: decodeRow,
  encode_cell: encodeCell,
  decode_cell: decodeCell,
  encode_range: encodeRange,
  decode_range: decodeRange,
  split_cell: splitCell,
  book_new: bookNew,
  book_append_sheet: bookAppendSheet,
  book_set_sheet_visibility: bookSetSheetVisibility,
  aoa_to_sheet: aoaToSheet,
  sheet_add_aoa: sheetAddAoa,
  json_to_sheet: jsonToSheet,
  sheet_add_json: sheetAddJson,
  format_cell: formatCell,
  cell_set_number_format: cellSetNumberFormat,
  sheet_to_formulae: sheetToFormulae,
  sheet_to_csv: sheetToCsv,
  sheet_to_txt: sheetToTxt,
  cell_set_hyperlink: cellSetHyperlink,
  cell_set_internal_link: cellSetInternalLink,
  cell_add_comment: cellAddComment,
  sheet_set_array_formula: sheetSetArrayFormula,
};

for (const publicName of Object.keys(module.exports)) nameAs(module.exports[publicName], publicName);
