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

module.exports = { datenum };
