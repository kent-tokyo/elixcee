'use strict';

// format_cell / cell_set_number_format, backed by the real SSF engine (see
// ./ssf-adapter.cjs). Independent port of the oracle's format_cell/safe_format_cell
// cell-level orchestration — the format-string evaluation itself is delegated to the
// transitional ssf dependency, not reimplemented (see docs/xlsx-architecture.md).
const { format: ssfFormat } = require('./ssf-adapter.cjs');
const { datenum } = require('./datenum.cjs');

// Excel BIFF error-code -> display-string lookup, used by format_cell for error cells
// (cell.v is a numeric code there, not the "#DIV/0!" string).
const B_ERR = {
  0x00: '#NULL!',
  0x07: '#DIV/0!',
  0x0f: '#VALUE!',
  0x17: '#REF!',
  0x1d: '#NAME?',
  0x24: '#NUM!',
  0x2a: '#N/A',
  0x2b: '#GETTING_DATA',
  0xff: '#WTF?',
};

function cellSetNumberFormat(cell, fmt) {
  cell.z = fmt;
  return cell;
}

// Two-try fallthrough, matching the oracle's safe_format_cell exactly: a `cell.z` that
// SSF_format rejects (confirmed live — e.g. an unterminated quoted section like `["`)
// silently falls through to the numFmtId-based fallback (always 0/General or 14/m-d-yy
// in this package, since nothing here ever sets cell.XF) rather than propagating the
// error; only a SECOND failure returns the ''+v fallback — and that final catch uses
// the ORIGINAL `v` parameter, not the datenum-converted value, also confirmed live (a
// Date `v` with both tries failing renders via Date.prototype.toString(), not a serial
// number). The success path caches its result onto `cell.w`; the final-catch path does
// NOT (confirmed live: cell.w stays unset when both tries throw).
function safeFormatCell(cell, v) {
  const q = cell.t === 'd' && v instanceof Date;
  if (cell.z != null) {
    try {
      cell.w = ssfFormat(cell.z, q ? datenum(v) : v);
      return cell.w;
    } catch (e) {
      // fall through to the numFmtId-based fallback below
    }
  }
  try {
    cell.w = ssfFormat((cell.XF || {}).numFmtId || (q ? 14 : 0), q ? datenum(v) : v);
    return cell.w;
  } catch (e) {
    return '' + v;
  }
}

function formatCell(cell, v, o) {
  if (cell == null || cell.t == null || cell.t === 'z') return '';
  if (cell.w !== undefined) return cell.w;
  if (cell.t === 'd' && !cell.z && o && o.dateNF) cell.z = o.dateNF;
  if (cell.t === 'e') return B_ERR[cell.v] || cell.v;
  return safeFormatCell(cell, v == null ? cell.v : v);
}

module.exports = { formatCell, cellSetNumberFormat };
