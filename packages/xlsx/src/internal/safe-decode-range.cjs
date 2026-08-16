'use strict';

// INTERNAL — not exported from the package's public entrypoint (src/index.cjs / .mjs /
// .d.ts) and not part of `Object.keys` on the public namespace. `xlsx@0.18.5` does not
// export `utils.safe_decode_range` either (confirmed:
// `Object.prototype.hasOwnProperty.call(XLSX.utils, "safe_decode_range") === false` at
// runtime, despite an internal same-named function existing in the SheetJS source and
// being used internally, e.g. by sheet_add_aoa). Publishing this under the compat
// namespace would itself be a compatibility divergence, not a convenience — see
// docs/xlsx-architecture.md. Kept here as an internal helper / test helper only, reached
// by test files via this exact path, never via `require('@elixcee/xlsx')`.
//
// Behavior mirrors the internal oracle algorithm: strict left-to-right character
// scanning (uppercase-letters-then-digits, twice, separated by any single non-digit
// byte) that never throws — malformed input degrades to {s:{c:-1,r:-1},e:{c:-1,r:-1}}
// rather than an exception, unlike decode_range. There is no public oracle export to
// differential-test this against, so it is verified only by a self-check (see
// compat/differential/xlsx-utils.test.mjs), not by oracle MATCH classification.
function safeDecodeRange(range) {
  const o = { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
  const len = range.length;
  let i = 0;
  let cc = 0;
  let idx = 0;

  for (idx = 0; i < len; ++i) {
    cc = range.charCodeAt(i) - 64;
    if (cc < 1 || cc > 26) break;
    idx = 26 * idx + cc;
  }
  o.s.c = --idx;

  for (idx = 0; i < len; ++i) {
    cc = range.charCodeAt(i) - 48;
    if (cc < 0 || cc > 9) break;
    idx = 10 * idx + cc;
  }
  o.s.r = --idx;

  if (i === len || cc !== 10) {
    o.e.c = o.s.c;
    o.e.r = o.s.r;
    return o;
  }
  ++i;

  for (idx = 0; i !== len; ++i) {
    cc = range.charCodeAt(i) - 64;
    if (cc < 1 || cc > 26) break;
    idx = 26 * idx + cc;
  }
  o.e.c = --idx;

  for (idx = 0; i !== len; ++i) {
    cc = range.charCodeAt(i) - 48;
    if (cc < 0 || cc > 9) break;
    idx = 10 * idx + cc;
  }
  o.e.r = --idx;

  return o;
}

module.exports = { safeDecodeRange };
