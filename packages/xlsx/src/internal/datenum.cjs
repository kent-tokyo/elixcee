'use strict';

// Ported from the oracle's `datenum` (not `datenum_local`, which is a separate
// SSF-internal function used only when a raw Date reaches SSF_format directly — every
// caller in this package pre-converts via this one first). Pure Date -> Excel-serial
// arithmetic, independent of the SSF format-string engine — verified against a live
// oracle run in Phase 1B-1 and unaffected by the Phase 1B-2B SSF-backend change.
const DATENUM_BASEDATE = new Date(1899, 11, 30, 0, 0, 0);

function datenum(v, date1904) {
  let epoch = v.getTime();
  if (date1904) epoch -= 1462 * 24 * 60 * 60 * 1000;
  const dnthresh = DATENUM_BASEDATE.getTime() + (v.getTimezoneOffset() - DATENUM_BASEDATE.getTimezoneOffset()) * 60000;
  return (epoch - dnthresh) / (24 * 60 * 60 * 1000);
}

// Ported from the oracle's own `numdate` (the reverse direction: Excel serial -> Date),
// used by read()'s opts.cellDates support (Milestone read-item 6). Deliberately does NOT
// take a date1904 parameter, unlike datenum above — confirmed live against the real
// oracle with an actual date1904 workbook (compat/node_modules/xlsx's own `numdate`,
// called from its cellDates read path, has no date1904 branch at all): a date1904 file's
// `.w` (formatted display text) DOES shift by the 1462-day 1904 offset (via SSF_format's
// own opts.date1904, applied separately — see read-shape.cjs), but the cellDates `.v`
// Date object does NOT — the real oracle's `.w` and `.v` genuinely disagree on which
// epoch a date1904 file uses. Reproduced as-is (fidelity over tidiness — see
// docs/compatibility-known-defects.md), not "fixed" here.
//
// `refdate`/`dnthresh`/`refoffset` are computed fresh per call (the oracle computes its
// module-scoped equivalents once, at requires-time) — a deliberate simplification: the
// two only differ when a long-running process straddles a DST transition between module
// load and this call, an edge case not worth a shared-mutable-module-state footgun for.
function numdate(serial) {
  const refdate = new Date();
  const refoffset = refdate.getTimezoneOffset();
  const dnthresh = DATENUM_BASEDATE.getTime() + (refoffset - DATENUM_BASEDATE.getTimezoneOffset()) * 60000;
  const out = new Date();
  out.setTime(serial * 24 * 60 * 60 * 1000 + dnthresh);
  if (out.getTimezoneOffset() !== refoffset) {
    out.setTime(out.getTime() + (out.getTimezoneOffset() - refoffset) * 60000);
  }
  return out;
}

module.exports = { datenum, numdate };
