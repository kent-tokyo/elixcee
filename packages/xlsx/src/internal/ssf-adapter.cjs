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
// diverge. This corrects the 5 affected ids to their bundled-engine-correct target
// before delegating — a narrow, disclosed patch for one precisely-diagnosed upstream
// defect, not a reimplementation of any part of the format engine itself.
const DEFAULT_MAP_CORRECTION = { 67: 9, 68: 10, 69: 12, 70: 13, 71: 14 };

function format(fmt, v, opts) {
  const correctedFmt = typeof fmt === 'number' && Object.prototype.hasOwnProperty.call(DEFAULT_MAP_CORRECTION, fmt) ? DEFAULT_MAP_CORRECTION[fmt] : fmt;
  return ssf.format(correctedFmt, v, opts);
}

module.exports = { format };
