'use strict';

// SheetJS 0.20.x interprets Date values through their absolute UTC timestamp.
// Keep this helper equivalent at the arithmetic boundary so workbook serials do not
// vary with the host timezone. Excel's fictitious 1900-02-29 is retained through the
// serial < 60 adjustment, matching the oracle.
const DATENUM_THRESHOLD = Date.UTC(1899, 11, 30, 0, 0, 0);
const DAY_MS = 24 * 60 * 60 * 1000;

function datenum(v, date1904) {
  let serial = (v.getTime() - DATENUM_THRESHOLD) / DAY_MS;
  if (date1904) {
    serial -= 1462;
    return serial < -1402 ? serial - 1 : serial;
  }
  return serial < 60 ? serial - 1 : serial;
}

// Reverse conversion used by read()'s opts.cellDates support. Serial values in Excel's
// fictitious leap-day interval cannot be represented by a JavaScript Date, so the oracle
// preserves those values numerically.
function numdate(serial) {
  if (serial >= 60 && serial < 61) return serial;
  return new Date((serial > 60 ? serial : serial + 1) * DAY_MS + DATENUM_THRESHOLD);
}

module.exports = { datenum, numdate };
