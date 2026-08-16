'use strict';

// @elixcee/xlsx — Phase 1A pure utility API (address encode/decode, workbook shells).
// No runtime dependency on the real `xlsx` package (see docs/xlsx-architecture.md's
// "Non-negotiable" section) — this file has zero `require`s beyond itself.
//
// Exact edge-case behavior (including quirks that look like bugs — e.g. decode_range
// never validates or swaps a reversed range; book_append_sheet's error message mentions
// ":" as a forbidden sheet-name character but the actual check never blocks it) was
// verified against the real xlsx@0.18.5 (SheetJS, Apache-2.0) source and confirmed
// against a live oracle run — see compat/differential/. Code below is an independent
// implementation, not copied text; see docs/licensing.md for the licensing boundary.

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

// ---- aoa_to_sheet ----
//
// Implements the default-options subset of the oracle's sheet_add_aoa(null, data, opts):
// dense vs. sparse (object-keyed) storage, null/undefined cell skipping, sparse-array
// holes (both missing rows and missing columns within a row) skipped without shifting
// indices, number/boolean/string type inference, and "!ref" only set if at least one
// cell was actually written. NOT implemented: opts.origin, opts.nullError,
// opts.sheetStubs, Date-typed cell values (needs the SSF date-formatting engine, out of
// scope for this pure-utility slice — see compat/differential/ for how this gap is
// reported, not silently skipped).
function aoaToSheet(data, opts) {
  const o = opts || {};
  const dense = !!o.dense;
  const ws = dense ? [] : {};

  const range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };

  for (let R = 0; R !== data.length; ++R) {
    if (!data[R]) continue;
    if (!Array.isArray(data[R])) throw new Error('aoa_to_sheet expects an array of arrays');
    for (let C = 0; C !== data[R].length; ++C) {
      let raw = data[R][C];
      if (typeof raw === 'undefined') continue;
      if (raw instanceof Date) {
        // Explicit, loud gap rather than silent mishandling: real aoa_to_sheet formats
        // Date cells via the SSF date-formatting engine, which is out of scope for this
        // pure-utility slice (it's the "ssf" package — one of the 7 Apache-2.0
        // dependencies packages/xlsx deliberately does not take, see
        // docs/xlsx-architecture.md). Not in Phase 1A's boundary matrix.
        throw new Error('aoa_to_sheet: Date-typed cells are not implemented yet in Phase 1A');
      }

      let cell;
      if (raw && typeof raw === 'object' && !Array.isArray(raw)) {
        cell = raw; // caller-supplied full cell object, used verbatim
      } else {
        cell = { v: raw };
        if (Array.isArray(cell.v)) {
          // [value, formula] pair shorthand.
          cell.f = raw[1];
          cell.v = raw[0];
        }
        if (cell.v === null) continue; // Phase 1A: no nullError/sheetStubs option support yet
        else if (typeof cell.v === 'number') cell.t = 'n';
        else if (typeof cell.v === 'boolean') cell.t = 'b';
        else cell.t = 's';
      }

      if (range.s.r > R) range.s.r = R;
      if (range.s.c > C) range.s.c = C;
      if (range.e.r < R) range.e.r = R;
      if (range.e.c < C) range.e.c = C;

      if (dense) {
        if (!ws[R]) ws[R] = [];
        ws[R][C] = cell;
      } else {
        ws[encodeCell({ c: C, r: R })] = cell;
      }
    }
  }

  if (range.s.c < 10000000) ws['!ref'] = encodeRange(range);
  return ws;
}

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
};
