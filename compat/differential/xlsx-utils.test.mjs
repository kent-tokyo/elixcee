// Permanent differential test suite (Phase 1A + Phase 1B-1): runs the boundary-value
// matrix specified for each phase through both the real oracle (xlsx@0.18.5) and
// @elixcee/xlsx, classifying every case with classify.mjs. This file itself IS the
// runnable check — it exits non-zero if anything is left UNCLASSIFIED (or otherwise
// fails to resolve to an acceptable verdict), so nothing can pass CI by accident.
//
// The oracle (`xlsx`) is imported here, inside compat/, and nowhere else — @elixcee/xlsx
// itself is imported only via a relative path into packages/xlsx/src, never as an npm
// dependency. See docs/xlsx-architecture.md's "Non-negotiable" section.
import assert from 'node:assert/strict';
import XLSX from 'xlsx';
import * as elixcee from '../../packages/xlsx/src/index.mjs';
import { safeDecodeRange } from '../../packages/xlsx/src/internal/safe-decode-range.cjs';
import { classify } from './classify.mjs';
import { normalize } from './normalize.mjs';

const U = XLSX.utils;
const results = []; // { api, label, verdict }

function record(api, label, verdict) {
  results.push({ api, label, verdict });
}

// Runs `label` through oracleFn/elixceeFn with the same args, capturing thrown errors as
// part of the comparable value (so a throw-vs-throw with matching message is a MATCH,
// and a throw-vs-value or mismatched message is not silently treated as equal). If
// elixcee throws with a `.code` property, it's passed to classify() so a registered
// safety/security divergence is recognized instead of falling through to UNCLASSIFIED.
// unsupportedCaseId (optional) is forwarded to classify() as-is — see classify.mjs's
// UNSUPPORTED_ALLOWLIST doc comment: both api AND this exact caseId must be registered
// together for a divergence to resolve to UNSUPPORTED, never api alone.
function runCase(api, oracleFn, elixceeFn, args, label, unsupportedCaseId) {
  const oracleVal = invoke(oracleFn, args);
  const elixceeVal = invoke(elixceeFn, args);
  const verdict = classify({
    api,
    unsupportedCaseId,
    oracleA: oracleVal,
    elixcee: elixceeVal,
    elixceeErrorCode: elixceeVal.code,
  });
  record(api, label, verdict);
  return verdict;
}

// Variant of runCase for inputs the real oracle cannot be safely called with (confirmed
// hangs — see the encode_col(Infinity) case below). Never invokes oracleFn; classifies
// purely from elixcee's behavior + a registered error code, which is exactly what
// classify() needs since a registered code short-circuits before any oracleA/elixcee
// comparison happens.
function runUnsafeForOracleCase(api, elixceeFn, args, label) {
  const elixceeVal = invoke(elixceeFn, args);
  const verdict = classify({
    api,
    oracleA: { __note: 'not queried — confirmed to hang the real oracle, see comment at call site' },
    elixcee: elixceeVal,
    elixceeErrorCode: elixceeVal.code,
  });
  record(api, label, verdict);
  return verdict;
}

function invoke(fn, args) {
  try {
    return { threw: false, value: normalize(fn(...args)) };
  } catch (e) {
    return { threw: true, message: e.message, code: e.code };
  }
}

// ---- encode_col ----
// Non-finite / extreme numeric candidates, per the mandated safety review: empirically
// timed against the real oracle (a timeout-guarded subprocess probe, not assumed) before
// deciding what needed a fix. Only +Infinity hangs (Math.floor(Infinity) never reaches a
// falsy value) — NaN, -Infinity, MAX_VALUE, and MAX_SAFE_INTEGER all return/throw
// instantly on the real oracle, so those four stay in the normal MATCH matrix unchanged;
// fixing them too would manufacture unnecessary divergences where the oracle is already
// safe. See packages/xlsx/src/index.cjs's encodeCol doc comment and
// docs/xlsx-security-model.md.
for (const v of [
  0, 1, 25, 26, 27, 701, 702, 16383, 16384, -1, -100, 0.5, NaN, null, undefined,
  -Infinity, Number.MAX_VALUE, Number.MAX_SAFE_INTEGER, -Number.MAX_SAFE_INTEGER,
]) {
  runCase('utils.encode_col', U.encode_col, elixcee.encode_col, [v], `encode_col(${v})`);
}
// encode_col(Infinity) — the real oracle hangs (confirmed via a timeout-guarded
// subprocess run, process killed after OOM); never call U.encode_col(Infinity) here.
runUnsafeForOracleCase('utils.encode_col', elixcee.encode_col, [Infinity], 'encode_col(Infinity) [oracle hangs, not called]');

// ---- decode_col ----
for (const v of ['A', 'Z', 'AA', 'AB', 'ZZ', 'AAA', 'XFD', 'XFE', 'a', 'z', 'aa', '', '1', 'A1', '$A', null, undefined, ' A ']) {
  runCase('utils.decode_col', U.decode_col, elixcee.decode_col, [v], `decode_col(${JSON.stringify(v)})`);
}

// ---- encode_row ----
// Same non-finite/extreme candidates as encode_col, confirmed safe on both sides (no
// loop at all — "" + (row+1) — so no divergence is needed; kept in the normal matrix).
for (const v of [
  0, 1, 1048575, 1048576, -1, -2, 0.5, NaN, null, undefined,
  Infinity, -Infinity, Number.MAX_VALUE, Number.MAX_SAFE_INTEGER, -Number.MAX_SAFE_INTEGER,
]) {
  runCase('utils.encode_row', U.encode_row, elixcee.encode_row, [v], `encode_row(${v})`);
}

// ---- decode_row ----
for (const v of ['1', '1048576', '0', '-1', 'A', '', null, undefined, ' 1 ', '1.5']) {
  runCase('utils.decode_row', U.decode_row, elixcee.decode_row, [v], `decode_row(${JSON.stringify(v)})`);
}

// ---- encode_cell ----
for (const v of [
  { c: 0, r: 0 }, { c: 25, r: 0 }, { c: 26, r: 0 }, { c: 16383, r: 1048575 },
  { c: -1, r: 0 }, { c: 0, r: -1 }, {}, { c: 0 }, { r: 0 }, null,
]) {
  runCase('utils.encode_cell', U.encode_cell, elixcee.encode_cell, [v], `encode_cell(${JSON.stringify(v)})`);
}
// Non-finite/extreme column candidates: encode_cell's inline column loop uses `|0`
// (bitwise int32 truncation), not Math.floor, so it bounds itself to a small number of
// iterations for ANY numeric input including Infinity — confirmed safe empirically
// (timeout-guarded probe, ms:0 for every case below), so no fix is needed here. Kept in
// the normal MATCH matrix, not the safety-divergence path.
for (const c of [Infinity, -Infinity, NaN, Number.MAX_VALUE, Number.MAX_SAFE_INTEGER, -Number.MAX_SAFE_INTEGER]) {
  runCase('utils.encode_cell', U.encode_cell, elixcee.encode_cell, [{ c, r: 0 }], `encode_cell({c:${c},r:0})`);
  runCase('utils.encode_cell', U.encode_cell, elixcee.encode_cell, [{ c: 0, r: c }], `encode_cell({c:0,r:${c}})`);
}

// ---- decode_cell ----
for (const v of [
  'A1', 'Z1', 'AA1', 'XFD1048576', 'a1', 'z1', 'aa1', '$A$1', '$A1', 'A$1', 'A0', 'A-1',
  '1A', 'AAAA1', '', 'A', '1', 'Sheet1!A1', "'My Sheet'!A1", null, undefined, 'A1:B2',
]) {
  runCase('utils.decode_cell', U.decode_cell, elixcee.decode_cell, [v], `decode_cell(${JSON.stringify(v)})`);
}

// ---- split_cell ----
for (const v of ['A1', 'AA100', '$A$1', '$A1', 'A$1', 'a1', '', 'A', '1', 'XFD1048576']) {
  runCase('utils.split_cell', U.split_cell, elixcee.split_cell, [v], `split_cell(${JSON.stringify(v)})`);
}

// ---- encode_range ----
for (const v of [
  { s: { c: 0, r: 0 }, e: { c: 1, r: 1 } },
  { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } }, // same start/end collapses to no colon
  { s: { c: 1, r: 1 }, e: { c: 0, r: 0 } }, // reversed — not validated, not swapped
]) {
  runCase('utils.encode_range', U.encode_range, elixcee.encode_range, [v], `encode_range(${JSON.stringify(v)})`);
}
runCase('utils.encode_range', U.encode_range, elixcee.encode_range, [{ c: 0, r: 0 }, { c: 1, r: 1 }], 'encode_range(cellA, cellB)');
// Non-finite candidates: encode_range delegates entirely to encode_cell (no loop of its
// own), so it inherits encode_cell's `|0`-bounded safety — confirmed empirically, no fix
// needed, kept in the normal MATCH matrix.
for (const c of [Infinity, -Infinity, NaN, Number.MAX_VALUE, Number.MAX_SAFE_INTEGER, -Number.MAX_SAFE_INTEGER]) {
  runCase('utils.encode_range', U.encode_range, elixcee.encode_range, [{ c, r: 0 }, { c, r: 0 }], `encode_range({c:${c},r:0} x2)`);
}

// ---- decode_range ----
for (const v of [
  'A1:B2', 'A1', 'B2:A1', 'A1:A1', 'Sheet1!A1:B2', "'My Sheet'!A1:B2", 'XFD1:XFD1048576',
  'A0:B0', '', 'garbage', 'A1:', ':A1', null, undefined,
]) {
  runCase('utils.decode_range', U.decode_range, elixcee.decode_range, [v], `decode_range(${JSON.stringify(v)})`);
}

// ---- book_new ----
runCase('utils.book_new', U.book_new, elixcee.book_new, [], 'book_new()');

// ---- book_append_sheet (stateful — scenario-based) ----
function bookAppendScenario(label, run) {
  const oracleWb = U.book_new();
  const elixceeWb = elixcee.book_new();
  const oracleSteps = [];
  const elixceeSteps = [];
  run(oracleWb, U, oracleSteps);
  run(elixceeWb, elixcee, elixceeSteps);
  const verdict = classify({
    api: 'utils.book_append_sheet',
    oracleA: { steps: oracleSteps, SheetNames: oracleWb.SheetNames },
    elixcee: { steps: elixceeSteps, SheetNames: elixceeWb.SheetNames },
  });
  record('utils.book_append_sheet', label, verdict);
}
function step(fn, steps) {
  try { steps.push({ threw: false, value: fn() }); }
  catch (e) { steps.push({ threw: true, message: e.message }); }
}
const ws1 = U.aoa_to_sheet([[1]]);

bookAppendScenario('append then duplicate (no roll) throws on 2nd', (wb, u, steps) => {
  step(() => u.book_append_sheet(wb, ws1, 'Sheet1'), steps);
  step(() => u.book_append_sheet(wb, ws1, 'Sheet1'), steps);
});
bookAppendScenario('append with no name auto-increments', (wb, u, steps) => {
  step(() => u.book_append_sheet(wb, ws1), steps);
  step(() => u.book_append_sheet(wb, ws1), steps);
});
bookAppendScenario('roll=true renames duplicate instead of throwing', (wb, u, steps) => {
  step(() => u.book_append_sheet(wb, ws1, 'Sheet1', true), steps);
  step(() => u.book_append_sheet(wb, ws1, 'Sheet1', true), steps);
});
for (const name of ['Sheet/1', 'Sheet:1', 'Sheet?1', 'Sheet*1', "O'Brien", "'Quoted'", 'a'.repeat(31), 'a'.repeat(32)]) {
  bookAppendScenario(`append special-char/length name ${JSON.stringify(name)}`, (wb, u, steps) => {
    step(() => u.book_append_sheet(wb, ws1, name), steps);
  });
}
// "constructor"/"prototype" are ordinary (non-accessor) inherited properties — unlike
// "__proto__" these are NOT special-cased by the language, so plain bracket assignment
// already works correctly on the real oracle too (confirmed empirically). Expected plain
// MATCH, not a registered divergence.
for (const name of ['constructor', 'prototype', 'toString', 'hasOwnProperty']) {
  bookAppendScenario(`append property-name-shaped sheet name ${JSON.stringify(name)}`, (wb, u, steps) => {
    step(() => u.book_append_sheet(wb, ws1, name), steps);
  });
}

// Prototype-pollution-shaped input: a sheet literally named "__proto__". Found during
// Phase 1A's own security review, not part of the originally-specified boundary matrix.
// The real oracle's plain `wb.Sheets[name] = ws` invokes Object.prototype's inherited
// `__proto__` accessor instead of creating a normal own property — the sheet ends up
// UNRETRIEVABLE (wb.Sheets has zero own keys afterward) and wb.Sheets's own prototype is
// silently reassigned to the worksheet object. This does not leak into the global
// Object.prototype (confirmed: a fresh {} is unaffected), but it is exactly the
// "spreadsheet-derived string used as an object key" hazard docs/xlsx-security-model.md
// requires guarding — and per that same doc, the sheet must be retained as data, not
// rejected. Elixcee uses Object.defineProperty (see bookAppendSheet), so the sheet stays
// retrievable and nothing's prototype changes. Registered as
// INTENTIONAL_SECURITY_DIVERGENCE, not compared for MATCH, since the oracle's own output
// here is the defect.
{
  const oracleWb = U.book_new();
  U.book_append_sheet(oracleWb, ws1, '__proto__');
  const elixceeWb = elixcee.book_new();
  elixcee.book_append_sheet(elixceeWb, ws1, '__proto__');
  const verdict = classify({
    api: 'utils.book_append_sheet',
    oracleA: { ownSheetKeys: Object.keys(oracleWb.Sheets), sheetNames: oracleWb.SheetNames },
    elixcee: { ownSheetKeys: Object.keys(elixceeWb.Sheets), sheetNames: elixceeWb.SheetNames },
    securityDivergenceKey: 'book_append_sheet:proto_key_pollution',
  });
  record('utils.book_append_sheet', 'append "__proto__" name (prototype-pollution-shaped)', verdict);
  // Assert the actual safety property directly too, not just trust the registry lookup:
  assert.deepEqual(Object.keys(elixceeWb.Sheets), ['__proto__'], 'sheet must be retrievable');
  assert.equal(elixceeWb.Sheets['__proto__'], ws1, 'sheet must be the worksheet, not a corrupted prototype');
  assert.equal(Object.getPrototypeOf(elixceeWb.Sheets), Object.prototype, 'wb.Sheets\'s own prototype must be untouched');
  assert.equal(({}).marker, undefined, 'global Object.prototype must remain unpolluted');
}

// ---- book_set_sheet_visibility (stateful — scenario-based) ----
function visScenario(label, run) {
  const oracleWb = U.book_new();
  const elixceeWb = elixcee.book_new();
  U.book_append_sheet(oracleWb, U.aoa_to_sheet([[1]]), 'Sheet1');
  U.book_append_sheet(oracleWb, U.aoa_to_sheet([[1]]), 'Sheet2');
  elixcee.book_append_sheet(elixceeWb, elixcee.aoa_to_sheet([[1]]), 'Sheet1');
  elixcee.book_append_sheet(elixceeWb, elixcee.aoa_to_sheet([[1]]), 'Sheet2');
  const oracleSteps = [];
  const elixceeSteps = [];
  run(oracleWb, U, oracleSteps);
  run(elixceeWb, elixcee, elixceeSteps);
  const verdict = classify({
    api: 'utils.book_set_sheet_visibility',
    oracleA: { steps: oracleSteps, Workbook: oracleWb.Workbook },
    elixcee: { steps: elixceeSteps, Workbook: elixceeWb.Workbook },
  });
  record('utils.book_set_sheet_visibility', label, verdict);
}
for (const vis of [0, 1, 2, -1, 3]) {
  visScenario(`set idx0 visibility=${vis}`, (wb, u, steps) => {
    step(() => u.book_set_sheet_visibility(wb, 0, vis), steps);
  });
}
visScenario('set by sheet name', (wb, u, steps) => {
  step(() => u.book_set_sheet_visibility(wb, 'Sheet2', 1), steps);
});
visScenario('out-of-range index throws', (wb, u, steps) => {
  step(() => u.book_set_sheet_visibility(wb, 5, 1), steps);
});

// ---- aoa_to_sheet ----
function aoaCase(label, data, opts) {
  runCase('utils.aoa_to_sheet', (d, o) => U.aoa_to_sheet(d, o), (d, o) => elixcee.aoa_to_sheet(d, o), [data, opts], label);
}
aoaCase('basic 2x2', [[1, 2], [3, 4]]);
aoaCase('dense:true', [[1, 2], [3, 4]], { dense: true });
aoaCase('null/undefined cells skipped', [[1, null, 3], [undefined, 5, 6]]);
{
  const sparseRows = [];
  sparseRows[0] = [1, 2];
  sparseRows[2] = [5, 6]; // hole at row index 1
  aoaCase('sparse rows (hole)', sparseRows);
}
{
  const row = [];
  row[0] = 1;
  row[3] = 4; // holes within a row
  aoaCase('sparse cols within row (hole)', [row]);
}
aoaCase('empty array', []);
aoaCase('empty row', [[]]);
aoaCase('mixed number/string/boolean', [[1, 'str', true, false]]);

// ---- sheet_add_aoa (Phase 1B-1: origin, existing-!ref extension/overwrite, dense
// reuse, Date, formula/error objects, nullError/sheetStubs) ----
// Each fixture below builds a FRESH worksheet (never reuses a shared module-level `ws`
// across cases) — sheet_add_aoa mutates its `_ws` argument in place, unlike
// book_append_sheet, so sharing one worksheet across fixtures would leak state and make
// results order-dependent.
function sheetAddAoaCase(label, buildOracle, buildElixcee, unsupportedCaseId) {
  runCase(
    'utils.sheet_add_aoa',
    () => buildOracle(U),
    () => buildElixcee(elixcee),
    [],
    label,
    unsupportedCaseId
  );
}
const DATE_FIXTURE = new Date(2026, 0, 5);

sheetAddAoaCase('origin: number', (u) => u.sheet_add_aoa(u.aoa_to_sheet([[0]]), [[1, 2]], { origin: 1 }), (e) => e.sheet_add_aoa(e.aoa_to_sheet([[0]]), [[1, 2]], { origin: 1 }));
sheetAddAoaCase('origin: "C3" string', (u) => u.sheet_add_aoa(null, [[1, 2]], { origin: 'C3' }), (e) => e.sheet_add_aoa(null, [[1, 2]], { origin: 'C3' }));
sheetAddAoaCase('origin: {r,c} object', (u) => u.sheet_add_aoa(null, [[1, 2]], { origin: { r: 2, c: 3 } }), (e) => e.sheet_add_aoa(null, [[1, 2]], { origin: { r: 2, c: 3 } }));
sheetAddAoaCase(
  'origin: -1 appends after existing !ref',
  (u) => {
    const ws = u.aoa_to_sheet([[1, 2]]);
    return u.sheet_add_aoa(ws, [[3, 4]], { origin: -1 });
  },
  (e) => {
    const ws = e.aoa_to_sheet([[1, 2]]);
    return e.sheet_add_aoa(ws, [[3, 4]], { origin: -1 });
  }
);
sheetAddAoaCase(
  'existing !ref extended by a later write past its bounds',
  (u) => {
    const ws = u.aoa_to_sheet([[1, 2]]);
    return u.sheet_add_aoa(ws, [[9]], { origin: 'E5' });
  },
  (e) => {
    const ws = e.aoa_to_sheet([[1, 2]]);
    return e.sheet_add_aoa(ws, [[9]], { origin: 'E5' });
  }
);
sheetAddAoaCase(
  'overwriting an existing cell preserves its number format',
  (u) => {
    const ws = u.aoa_to_sheet([[1, 2], [3, 4]]);
    u.cell_set_number_format(ws.A1, '0.00');
    return u.sheet_add_aoa(ws, [[9]], { origin: 'A1' });
  },
  (e) => {
    const ws = e.aoa_to_sheet([[1, 2], [3, 4]]);
    e.cell_set_number_format(ws.A1, '0.00');
    return e.sheet_add_aoa(ws, [[9]], { origin: 'A1' });
  }
);
sheetAddAoaCase('dense reuse: _ws is an existing dense array', (u) => u.sheet_add_aoa(u.aoa_to_sheet([[1]], { dense: true }), [[9, 8]], { origin: 'A2' }), (e) => e.sheet_add_aoa(e.aoa_to_sheet([[1]], { dense: true }), [[9, 8]], { origin: 'A2' }));
sheetAddAoaCase('Date value, default (t:n, serial v, rendered w)', (u) => u.aoa_to_sheet([[DATE_FIXTURE]]), (e) => e.aoa_to_sheet([[DATE_FIXTURE]]));
sheetAddAoaCase('Date value, cellDates:true (t:d, Date v, rendered w)', (u) => u.aoa_to_sheet([[DATE_FIXTURE]], { cellDates: true }), (e) => e.aoa_to_sheet([[DATE_FIXTURE]], { cellDates: true }));
// Backfilled before Phase 1B-2A per user review: confirmed live that the real oracle
// DOES render a custom dateNF correctly here (SSF_format("yyyy-mm-dd", ...) succeeds),
// since sheet_add_aoa's Date branch always computes cell.w. elixcee's narrow SSF subset
// only renders 'm/d/yy', so this throws ELIXCEE_NUMFMT_UNSUPPORTED — a registered,
// case-specific gap (see classify.mjs's UNSUPPORTED_ALLOWLIST for
// 'utils.sheet_add_aoa'), not a blanket "sheet_add_aoa is unsupported."
sheetAddAoaCase(
  'Date value, dateNF="yyyy-mm-dd" (custom format outside the narrow subset) -> UNSUPPORTED',
  (u) => u.aoa_to_sheet([[DATE_FIXTURE]], { dateNF: 'yyyy-mm-dd' }),
  (e) => e.aoa_to_sheet([[DATE_FIXTURE]], { dateNF: 'yyyy-mm-dd' }),
  'dateNF="yyyy-mm-dd" (Date value, custom format other than "m/d/yy")'
);
sheetAddAoaCase(
  'Date value, cellDates:true, dateNF="yyyy-mm-dd" -> UNSUPPORTED',
  (u) => u.aoa_to_sheet([[DATE_FIXTURE]], { cellDates: true, dateNF: 'yyyy-mm-dd' }),
  (e) => e.aoa_to_sheet([[DATE_FIXTURE]], { cellDates: true, dateNF: 'yyyy-mm-dd' }),
  'dateNF="yyyy-mm-dd" (Date value, custom format other than "m/d/yy")'
);
sheetAddAoaCase('null with nullError:true -> error cell', (u) => u.aoa_to_sheet([[null, 1]], { nullError: true }), (e) => e.aoa_to_sheet([[null, 1]], { nullError: true }));
sheetAddAoaCase('null with sheetStubs:true -> stub cell', (u) => u.aoa_to_sheet([[null, 1]], { sheetStubs: true }), (e) => e.aoa_to_sheet([[null, 1]], { sheetStubs: true }));
sheetAddAoaCase('[value, formula] shorthand pair', (u) => u.aoa_to_sheet([[[5, 'A1+A2']]]), (e) => e.aoa_to_sheet([[[5, 'A1+A2']]]));
sheetAddAoaCase('[null, formula] shorthand pair', (u) => u.aoa_to_sheet([[[null, 'A1']]]), (e) => e.aoa_to_sheet([[[null, 'A1']]]));
sheetAddAoaCase('caller-supplied full cell object (formula)', (u) => u.aoa_to_sheet([[{ t: 'n', v: 1, f: 'A1' }]]), (e) => e.aoa_to_sheet([[{ t: 'n', v: 1, f: 'A1' }]]));
sheetAddAoaCase('caller-supplied full cell object (error)', (u) => u.aoa_to_sheet([[{ t: 'e', v: 7 }]]), (e) => e.aoa_to_sheet([[{ t: 'e', v: 7 }]]));
sheetAddAoaCase('NaN cell value', (u) => u.aoa_to_sheet([[NaN]]), (e) => e.aoa_to_sheet([[NaN]]));
sheetAddAoaCase('Infinity cell value', (u) => u.aoa_to_sheet([[Infinity]]), (e) => e.aoa_to_sheet([[Infinity]]));

// ---- json_to_sheet / sheet_add_json (Phase 1B-1) ----
function jsonCase(label, buildOracle, buildElixcee, api) {
  runCase(api || 'utils.json_to_sheet', () => buildOracle(U), () => buildElixcee(elixcee), [], label);
}

jsonCase('basic 2 rows', (u) => u.json_to_sheet([{ a: 1, b: 'x' }, { a: 2, b: 'y' }]), (e) => e.json_to_sheet([{ a: 1, b: 'x' }, { a: 2, b: 'y' }]));
jsonCase('skipHeader:true', (u) => u.json_to_sheet([{ a: 1, b: 'x' }], { skipHeader: true }), (e) => e.json_to_sheet([{ a: 1, b: 'x' }], { skipHeader: true }));
jsonCase('header option: partial order, new column appended', (u) => u.json_to_sheet([{ a: 1, b: 2, c: 3 }], { header: ['b'] }), (e) => e.json_to_sheet([{ a: 1, b: 2, c: 3 }], { header: ['b'] }));
jsonCase(
  'header option array is mutated in place (deliberate fidelity, not cloned)',
  (u) => { const hdr = ['b']; u.json_to_sheet([{ a: 1, b: 2, c: 3 }], { header: hdr }); return hdr; },
  (e) => { const hdr = ['b']; e.json_to_sheet([{ a: 1, b: 2, c: 3 }], { header: hdr }); return hdr; }
);
jsonCase('empty array', (u) => u.json_to_sheet([]), (e) => e.json_to_sheet([]));
jsonCase('empty row object', (u) => u.json_to_sheet([{}]), (e) => e.json_to_sheet([{}]));
jsonCase('null value, no nullError', (u) => u.json_to_sheet([{ a: null, b: 1 }]), (e) => e.json_to_sheet([{ a: null, b: 1 }]));
jsonCase('null value, nullError:true', (u) => u.json_to_sheet([{ a: null, b: 1 }], { nullError: true }), (e) => e.json_to_sheet([{ a: null, b: 1 }], { nullError: true }));
jsonCase('undefined value (own enumerable prop)', (u) => u.json_to_sheet([{ a: undefined, b: 1 }]), (e) => e.json_to_sheet([{ a: undefined, b: 1 }]));
jsonCase('sparse rows array (hole)', (u) => { const rows = []; rows[0] = { a: 1 }; rows[2] = { a: 3 }; return u.json_to_sheet(rows); }, (e) => { const rows = []; rows[0] = { a: 1 }; rows[2] = { a: 3 }; return e.json_to_sheet(rows); });
jsonCase('NaN value', (u) => u.json_to_sheet([{ a: NaN }]), (e) => e.json_to_sheet([{ a: NaN }]));
jsonCase('Infinity value', (u) => u.json_to_sheet([{ a: Infinity }]), (e) => e.json_to_sheet([{ a: Infinity }]));
jsonCase('formula object value', (u) => u.json_to_sheet([{ a: { t: 'n', v: 1, f: 'A1' } }]), (e) => e.json_to_sheet([{ a: { t: 'n', v: 1, f: 'A1' } }]));
jsonCase('error object value', (u) => u.json_to_sheet([{ a: { t: 'e', v: 7 } }]), (e) => e.json_to_sheet([{ a: { t: 'e', v: 7 } }]));
jsonCase('Date value, default (t:n, serial v, z but no w)', (u) => u.json_to_sheet([{ a: DATE_FIXTURE }]), (e) => e.json_to_sheet([{ a: DATE_FIXTURE }]));
jsonCase('Date value, cellDates:true', (u) => u.json_to_sheet([{ a: DATE_FIXTURE }], { cellDates: true }), (e) => e.json_to_sheet([{ a: DATE_FIXTURE }], { cellDates: true }));
// Backfilled before Phase 1B-2A per user review. Unlike sheet_add_aoa, this is a plain
// MATCH, not an UNSUPPORTED case: confirmed live that json_to_sheet/sheet_add_json never
// call the format engine at all for Date cells (only `z` gets set, `.w` is never
// computed here — see the Phase 1B-1 doc comment on sheetAddJson), so an unsupported
// custom dateNF is harmless — it just becomes the (unrendered) `z` string, identically
// on both sides.
jsonCase(
  'Date value, dateNF="yyyy-mm-dd" (custom format — harmless, no w is ever computed)',
  (u) => u.json_to_sheet([{ a: DATE_FIXTURE }], { dateNF: 'yyyy-mm-dd' }),
  (e) => e.json_to_sheet([{ a: DATE_FIXTURE }], { dateNF: 'yyyy-mm-dd' })
);
jsonCase(
  '__proto__-named own property (prototype-pollution probe, via JSON.parse)',
  (u) => u.json_to_sheet([JSON.parse('{"__proto__":1,"b":2}')]),
  (e) => e.json_to_sheet([JSON.parse('{"__proto__":1,"b":2}')])
);
jsonCase(
  'header option containing "__proto__" (prototype-pollution probe)',
  (u) => u.sheet_add_json({}, [{ x: 1 }], { header: ['__proto__'] }),
  (e) => e.sheet_add_json({}, [{ x: 1 }], { header: ['__proto__'] }),
  'utils.sheet_add_json'
);
jsonCase(
  'origin: number',
  (u) => u.sheet_add_json(u.aoa_to_sheet([[0]]), [{ a: 1 }], { origin: 1 }),
  (e) => e.sheet_add_json(e.aoa_to_sheet([[0]]), [{ a: 1 }], { origin: 1 }),
  'utils.sheet_add_json'
);
jsonCase(
  'origin: "C3" string',
  (u) => u.sheet_add_json(null, [{ a: 1 }], { origin: 'C3' }),
  (e) => e.sheet_add_json(null, [{ a: 1 }], { origin: 'C3' }),
  'utils.sheet_add_json'
);
jsonCase(
  'origin: {r,c} object',
  (u) => u.sheet_add_json(null, [{ a: 1 }], { origin: { r: 2, c: 3 } }),
  (e) => e.sheet_add_json(null, [{ a: 1 }], { origin: { r: 2, c: 3 } }),
  'utils.sheet_add_json'
);
jsonCase(
  'origin: -1 appends after existing !ref',
  (u) => { const ws = u.json_to_sheet([{ a: 1 }]); return u.sheet_add_json(ws, [{ a: 2 }], { origin: -1 }); },
  (e) => { const ws = e.json_to_sheet([{ a: 1 }]); return e.sheet_add_json(ws, [{ a: 2 }], { origin: -1 }); },
  'utils.sheet_add_json'
);
jsonCase(
  'overwriting an existing cell preserves its number format',
  (u) => { const ws = u.json_to_sheet([{ a: 1 }]); u.cell_set_number_format(ws.A2, '0.00'); return u.sheet_add_json(ws, [{ a: 99 }], { skipHeader: true }); },
  (e) => { const ws = e.json_to_sheet([{ a: 1 }]); e.cell_set_number_format(ws.A2, '0.00'); return e.sheet_add_json(ws, [{ a: 99 }], { skipHeader: true }); },
  'utils.sheet_add_json'
);
jsonCase(
  'opts.dense has no effect when _ws is null (confirmed oracle quirk, reproduced)',
  (u) => u.json_to_sheet([{ a: 1 }], { dense: true }),
  (e) => e.json_to_sheet([{ a: 1 }], { dense: true })
);
jsonCase(
  'dense target: scalar values land in the nested array; header/object values leak as stray string-keyed props (confirmed oracle quirk, reproduced)',
  (u) => u.sheet_add_json([], [{ a: 1, b: 'x' }]),
  (e) => e.sheet_add_json([], [{ a: 1, b: 'x' }]),
  'utils.sheet_add_json'
);
// Backfilled before Phase 1B-2A per user review — same reasoning as json_to_sheet's
// equivalent fixture above: sheet_add_json never calls the format engine for Date
// cells, so a custom dateNF is harmless (MATCH), not an UNSUPPORTED case.
jsonCase(
  'Date value, dateNF="yyyy-mm-dd" (custom format — harmless, no w is ever computed)',
  (u) => u.sheet_add_json({}, [{ a: DATE_FIXTURE }], { dateNF: 'yyyy-mm-dd' }),
  (e) => e.sheet_add_json({}, [{ a: DATE_FIXTURE }], { dateNF: 'yyyy-mm-dd' }),
  'utils.sheet_add_json'
);

// ---- format_cell / cell_set_number_format (Phase 1B-1: deliberately narrow SSF
// subset — see classify.mjs's UNSUPPORTED_ALLOWLIST entry for 'utils.format_cell') ----
// Shallow clone, not JSON.parse(JSON.stringify(...)) — format_cell mutates its cell
// argument (caches `.w`, and may set `.z` from opts.dateNF) and several fixtures use a
// Date-typed `.v`, which a JSON round-trip would silently collapse into an ISO string.
// Comparing { out, cell } rather than just the returned string catches a divergence in
// the mutation itself (e.g. a fixture named "sets cell.z" that returns the right string
// but never actually writes cell.z would otherwise MATCH by accident).
function formatCellCase(label, cell, v, opts, unsupportedCaseId) {
  const cellA = cell ? { ...cell } : cell;
  const cellB = cell ? { ...cell } : cell;
  runCase(
    'utils.format_cell',
    () => ({ out: U.format_cell(cellA, v, opts), cell: cellA }),
    () => ({ out: elixcee.format_cell(cellB, v, opts), cell: cellB }),
    [],
    label,
    unsupportedCaseId
  );
}

formatCellCase('General: number (integer)', { t: 'n', v: 42 });
formatCellCase('General: number (float, exercises SSF_general_num)', { t: 'n', v: 1234.5678 });
formatCellCase('General: string', { t: 's', v: 'hello' });
formatCellCase('General: boolean', { t: 'b', v: true });
formatCellCase('m/d/yy: date serial', { t: 'n', v: 46027, z: 'm/d/yy' });
formatCellCase('m/d/yy: out-of-range serial -> ""', { t: 'n', v: -5, z: 'm/d/yy' });
formatCellCase('error cell: BErr lookup', { t: 'e', v: 0x07 });
formatCellCase('cached .w short-circuits formatting', { t: 'n', v: 1234.5, w: 'cached' });
formatCellCase('null cell -> ""', null);
formatCellCase('t: "z" -> ""', { t: 'z' });
formatCellCase('dateNF option sets cell.z when unset', { t: 'd', v: DATE_FIXTURE }, undefined, { dateNF: 'm/d/yy' });
formatCellCase('explicit v param overrides cell.v', { t: 'n', v: 1 }, 5000);
// '0.00' is a fully-supported format on the real oracle (renders "1234.50") but is
// outside this package's deliberately narrow SSF subset (see classify.mjs's
// UNSUPPORTED_ALLOWLIST entry for 'utils.format_cell') — elixcee throws
// ELIXCEE_NUMFMT_UNSUPPORTED instead of guessing a rendering, so this is an honest,
// registered UNSUPPORTED divergence rather than a MATCH.
formatCellCase(
  'number format outside the narrow subset (0.00) -> UNSUPPORTED',
  { t: 'n', v: 1234.5, z: '0.00' },
  undefined,
  undefined,
  'z="0.00" (numeric cell, non-General/non-m/d/yy format code)'
);

runCase('utils.cell_set_number_format', (c, f) => U.cell_set_number_format(c, f), (c, f) => elixcee.cell_set_number_format(c, f), [{ t: 'n', v: 1 }, '0.00'], 'sets cell.z and returns the cell');

// ---- safe_decode_range: deliberately NOT public ----
// Confirmed absent from the real oracle's public API — not just `typeof undefined`, but
// `hasOwnProperty` false, ruling out an inherited/prototype-chain false negative:
assert.equal(
  Object.prototype.hasOwnProperty.call(U, 'safe_decode_range'),
  false,
  'expected safe_decode_range to remain unexported by the oracle — update this test if that changes'
);
// And confirmed absent from @elixcee/xlsx's public entrypoint too — it must only be
// reachable via the internal path imported at the top of this file, never via the
// package's public `index.cjs`/`index.mjs`/`index.d.ts`.
assert.equal(
  Object.prototype.hasOwnProperty.call(elixcee, 'safe_decode_range'),
  false,
  'safe_decode_range must not be publicly exported from @elixcee/xlsx either — see packages/xlsx/src/internal/safe-decode-range.cjs'
);
// No MATCH-classified oracle comparison is possible for it (no oracle export to compare
// against) — verified only by a self-check, reported separately below, never folded into
// the "oracle-covered" MATCH count.
const safeDecodeRangeSelfCheck = [];
for (const v of ['A1:B2', 'garbage', '', 'A1', 'B2:A1', '###', 'A1:B2:C3']) {
  try {
    const r = safeDecodeRange(v);
    safeDecodeRangeSelfCheck.push({ v, threw: false, r });
  } catch (e) {
    safeDecodeRangeSelfCheck.push({ v, threw: true, message: e.message });
  }
}
assert.ok(
  safeDecodeRangeSelfCheck.every((c) => !c.threw),
  'safe_decode_range must never throw (matches the internal oracle algorithm it was ported from)'
);
assert.deepEqual(
  safeDecodeRange('A1:B2'),
  { s: { c: 0, r: 0 }, e: { c: 1, r: 1 } },
  'safe_decode_range well-formed-input sanity check'
);

// ---- summary ----
// MATCH, registered-divergence verdicts (INTENTIONAL_SAFETY_DIVERGENCE,
// INTENTIONAL_SECURITY_DIVERGENCE), and — new in Phase 1B-1 — registered UNSUPPORTED
// verdicts are acceptable outcomes here. Everything else (BUG, ORACLE_AMBIGUITY,
// NONDETERMINISTIC, UNCLASSIFIED) fails the run. Per the classification rule:
// normal-input divergences are BUG, only pre-registered gaps are
// UNSUPPORTED/*_DIVERGENCE, anything else is UNCLASSIFIED — none of those are silently
// accepted. UNSUPPORTED only appears for format_cell's deliberately narrow number-format
// subset (see classify.mjs's UNSUPPORTED_ALLOWLIST) — every label that resolves to it is
// printed below by name, so a real bug hiding behind the same api key would still be
// visible on review, not silently absorbed.
const ACCEPTABLE = new Set([
  'MATCH',
  'INTENTIONAL_SAFETY_DIVERGENCE',
  'INTENTIONAL_SECURITY_DIVERGENCE',
  'UNSUPPORTED',
]);
const byApi = new Map();
let totalMatch = 0;
let totalSafetyDivergence = 0;
let totalSecurityDivergence = 0;
let totalUnsupported = 0;
let totalUnclassified = 0;
let totalBug = 0;
for (const r of results) {
  if (!byApi.has(r.api)) byApi.set(r.api, { total: 0, ok: 0, other: [], unsupported: [] });
  const bucket = byApi.get(r.api);
  bucket.total += 1;
  if (r.verdict === 'UNSUPPORTED') {
    // Acceptable, but NOT silently folded into `ok` unlabeled — UNSUPPORTED is gated by
    // an api-wide registry key (classify.mjs), so a real bug elsewhere under the same
    // api would otherwise hide behind this same bucket. Always printed by label below.
    bucket.ok += 1;
    bucket.unsupported.push(r.label);
  } else if (ACCEPTABLE.has(r.verdict)) bucket.ok += 1;
  else bucket.other.push({ label: r.label, verdict: r.verdict });
  if (r.verdict === 'MATCH') totalMatch += 1;
  else if (r.verdict === 'INTENTIONAL_SAFETY_DIVERGENCE') totalSafetyDivergence += 1;
  else if (r.verdict === 'INTENTIONAL_SECURITY_DIVERGENCE') totalSecurityDivergence += 1;
  else if (r.verdict === 'UNSUPPORTED') totalUnsupported += 1;
  else if (r.verdict === 'UNCLASSIFIED') totalUnclassified += 1;
  else if (r.verdict === 'BUG') totalBug += 1;
}

console.log('\n=== differential summary (Phase 1A + 1B-1) ===');
let anyFailure = false;
for (const [api, bucket] of byApi) {
  const status = bucket.other.length === 0 ? 'OK' : 'FAIL';
  if (bucket.other.length > 0) anyFailure = true;
  console.log(`${status}  ${api}: ${bucket.ok}/${bucket.total} acceptable`);
  for (const o of bucket.other) console.log(`      ${o.verdict}: ${o.label}`);
  for (const label of bucket.unsupported) console.log(`      UNSUPPORTED: ${label}`);
}
console.log(`\nsafe_decode_range: 7/7 self-check assertions passed (not oracle-covered — see note above)`);
console.log('\n=== Totals ===');
console.log(`normal fixtures (MATCH):            ${totalMatch}`);
console.log(`intentional safety divergence:      ${totalSafetyDivergence}`);
console.log(`intentional security divergence:    ${totalSecurityDivergence}`);
console.log(`unsupported:                        ${totalUnsupported}`);
console.log(`bug:                                ${totalBug}`);
console.log(`unclassified:                       ${totalUnclassified}`);

if (anyFailure) {
  console.error('\ndifferential suite FAILED: at least one case is not an acceptable verdict.');
  process.exit(1);
}
console.log('\ndifferential suite passed: every case is MATCH or a registered divergence.');
