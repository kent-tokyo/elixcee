'use strict';

// Isolated boundary around the `ssf` npm package (Apache-2.0, SheetJS) — the ONLY file
// in this package that `require`s it. This is a deliberate, disclosed, TRANSITIONAL
// runtime dependency (see docs/xlsx-architecture.md's "SSF backend" decision and
// docs/licensing.md): xlsx@0.18.5 bundles the identical format-string engine inline
// (no `require('ssf')` in its own source), confirmed empirically against the bundled
// engine across a permanent 900+ case matrix (compat/differential/ssf-format.test.mjs)
// — every table_fmt built-in and its SSF_default_map indirection, crossed with
// boundary/date-serial/text/boolean values, plus date1904, a custom format table,
// multi-section formats, [Red]/conditional sections, fractions, exponential notation,
// and percent/thousands.
//
// Nothing outside this file reaches into `ssf` directly, so swapping the backend later
// (e.g. for a Rust/WASM-based formatter once elixcee's Rust core grows one) is meant to
// be a single-file change — everything else in this package only ever calls
// `format(fmt, v, opts)` below.
const ssf = require('ssf');

// Confirmed root cause (compat/node_modules/ssf/ssf.js lines 93-105): the numFmtId ->
// numFmtId indirection table (`default_map`, ssf@0.11.2's equivalent of the bundled
// engine's SSF_default_map) has a genuine copy-paste bug — the loop meant to set
// default_map[69..71] to [12,13,14] (per its own comment, "69 -> 12 ... 71 -> 14")
// instead reuses the PRECEDING block's loop bounds (`defi = 67; defi <= 68`), which (a)
// overwrites default_map[67]/[68] a second time with the wrong values (10/11 instead of
// 9/10) and (b) never sets default_map[69..71] at all. Confirmed live against the
// bundled engine that these are the ONLY 5 numFmtIds (of 0-100+ swept) where the two
// diverge.
//
// The oracle's own numFmtId resolution is a 3-step chain (SSF_format's "number" case):
//   1. o.table[fmt]  (if o.table was passed at all; otherwise table_fmt[fmt] — but
//      table_fmt has no entry for 67-71, they're indirection-only ids)
//   2. o.table[default_map[fmt]] || table_fmt[default_map[fmt]]
//   3. SSF_default_str[fmt] || "General"
// Step 1 doesn't touch default_map at all, so a caller's own `opts.table[67]` override
// already works correctly even through the buggy ssf@0.11.2 — confirmed live. The bug is
// isolated to step 2. A naive "remap fmt then delegate" fix (e.g. rewrite 67 to 9 and
// call ssf.format(9, v, opts)) breaks BOTH directions: it skips step 1 entirely (losing
// a literal opts.table[67] override), AND produces the wrong step-2 fallback (id 9 has
// no indirection of its own, so ssf.format(9, v, opts) falls to "General" when opts.table
// lacks a `9` key, instead of correctly reaching table_fmt[9] — confirmed live these
// differ: the real oracle still renders via table_fmt[9] in that case). This resolves
// steps 1-2 explicitly instead, delegating to ssf.format only with an already-resolved
// format STRING, so ssf never has to consult its own broken default_map for these ids.
const DEFAULT_MAP_CORRECTION = { 67: 9, 68: 10, 69: 12, 70: 13, 71: 14 };

// table_fmt entries for the 5 corrected targets — the literal built-in strings step 2
// falls back to when o.table has neither the original id nor the target id. Verified
// against the bundled engine's own table_fmt (compat/node_modules/xlsx/xlsx.js).
const TARGET_FORMAT = { 9: '0%', 10: '0.00%', 12: '# ?/?', 13: '# ??/??', 14: 'm/d/yy' };

function hasOwn(obj, key) {
  return !!obj && Object.prototype.hasOwnProperty.call(obj, key);
}

function format(fmt, v, opts) {
  if (typeof fmt === 'number' && hasOwn(DEFAULT_MAP_CORRECTION, fmt) && !hasOwn(opts && opts.table, fmt)) {
    const target = DEFAULT_MAP_CORRECTION[fmt];
    const sfmt = hasOwn(opts && opts.table, target) ? opts.table[target] : TARGET_FORMAT[target];
    // dateNF must NOT apply here: confirmed live that SSF_format(71, v, {dateNF}) does
    // NOT substitute (only the LITERAL id 14 / string 'm/d/yy' path does, per the
    // oracle's own source — the indirected step-2 path never reaches that check).
    // Stripping it prevents ssf's string-branch dateNF check (`fmt === 'm/d/yy' &&
    // o.dateNF`) from incorrectly firing once `sfmt` resolves to the literal 'm/d/yy'.
    const fallbackOpts = opts && 'dateNF' in opts ? Object.assign({}, opts, { dateNF: undefined }) : opts;
    return ssf.format(sfmt, v, fallbackOpts);
  }
  return ssf.format(fmt, v, opts);
}

// ---- read()'s .z / date-typed-cell support (Milestone read-item 6) ----
//
// resolveFormatString mirrors format()'s own 3-step numFmtId resolution (opts.table
// override -> DEFAULT_MAP_CORRECTION indirection -> ssf's own built-in table), but
// returns the resolved format-code STRING itself rather than delegating to ssf.format —
// read-shape.cjs needs the string form both to decide date-ness (isDate, below, only
// accepts a string, confirmed live: `ssf.is_date(14)` is false even though built-in id 14
// IS "m/d/yy") and to surface as `.z`, which the oracle itself always resolves to a
// string (confirmed live: even an unstyled cell's `.z` reads back as the literal string
// "General", never the number 0). A separate function, not a refactor of format() itself,
// to avoid touching the exact resolution order/dateNF-stripping behavior already verified
// byte-identical across 819 cases.
function resolveFormatString(fmt, opts) {
  if (typeof fmt !== 'number') return fmt == null ? 'General' : fmt;
  const table = (opts && opts.table) || {};
  if (hasOwn(table, fmt)) return table[fmt];
  if (hasOwn(DEFAULT_MAP_CORRECTION, fmt)) {
    const target = DEFAULT_MAP_CORRECTION[fmt];
    return hasOwn(table, target) ? table[target] : TARGET_FORMAT[target];
  }
  const builtin = ssf.get_table()[fmt];
  return builtin == null ? 'General' : builtin;
}

// ssf.is_date operates on a resolved format-code STRING (never a numFmtId) — re-exported
// verbatim under this package's own naming convention (camelCase, like format/isDate's
// sibling exports here) rather than requiring 'ssf' a second time elsewhere.
const isDate = ssf.is_date;

module.exports = { format, resolveFormatString, isDate };
