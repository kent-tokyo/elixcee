'use strict';

// Isolated boundary around the `ssf` npm package (Apache-2.0, SheetJS) — the ONLY file
// in this package that `require`s it. This is a deliberate, disclosed, TRANSITIONAL
// runtime dependency (see docs/xlsx-architecture.md's "SSF backend" decision and
// docs/licensing.md): xlsx@0.18.5 bundles the identical format-string engine inline
// (no `require('ssf')` in its own source), but confirmed empirically byte-identical to
// the standalone `ssf@0.11.2` package across an 819-case matrix (every table_fmt
// built-in and its SSF_default_map indirection, crossed with boundary/date-serial/
// text/boolean values, plus date1904, a custom format table, multi-section formats,
// [Red]/conditional sections, fractions, exponential notation, and percent/thousands —
// see compat/differential/ssf-format.test.mjs) BEFORE this dependency was added, not
// assumed from matching version strings.
//
// Nothing outside this file reaches into `ssf` directly, so swapping the backend later
// (e.g. for a Rust/WASM-based formatter once elixcee's Rust core grows one) is meant to
// be a single-file change — everything else in this package only ever calls
// `format(fmt, v, opts)` below.
const ssf = require('ssf');

function format(fmt, v, opts) {
  return ssf.format(fmt, v, opts);
}

module.exports = { format };
