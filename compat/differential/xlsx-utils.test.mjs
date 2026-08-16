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
// Backfilled before Phase 1B-2A as an UNSUPPORTED case (elixcee's then-narrow SSF
// subset only rendered 'm/d/yy'); flipped to a plain MATCH in Phase 1B-2B now that
// sheet_add_aoa's Date branch is backed by the real SSF engine (see
// packages/xlsx/src/internal/ssf-adapter.cjs) — confirmed live the oracle renders
// "2026-01-05" here and elixcee now does too.
sheetAddAoaCase(
  'Date value, dateNF="yyyy-mm-dd" (custom format, now fully supported via the real SSF engine)',
  (u) => u.aoa_to_sheet([[DATE_FIXTURE]], { dateNF: 'yyyy-mm-dd' }),
  (e) => e.aoa_to_sheet([[DATE_FIXTURE]], { dateNF: 'yyyy-mm-dd' })
);
sheetAddAoaCase(
  'Date value, cellDates:true, dateNF="yyyy-mm-dd" (now fully supported)',
  (u) => u.aoa_to_sheet([[DATE_FIXTURE]], { cellDates: true, dateNF: 'yyyy-mm-dd' }),
  (e) => e.aoa_to_sheet([[DATE_FIXTURE]], { cellDates: true, dateNF: 'yyyy-mm-dd' })
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
// Was a registered UNSUPPORTED case through Phase 1B-2A (elixcee's then-narrow SSF
// subset); flipped to a plain MATCH in Phase 1B-2B — format_cell is now backed by the
// real SSF engine and renders "1234.50" identically to the oracle.
formatCellCase('number format previously outside the narrow subset (0.00, now fully supported)', { t: 'n', v: 1234.5, z: '0.00' });

runCase('utils.cell_set_number_format', (c, f) => U.cell_set_number_format(c, f), (c, f) => elixcee.cell_set_number_format(c, f), [{ t: 'n', v: 1 }, '0.00'], 'sets cell.z and returns the cell');

// ---- sheet_to_formulae (Phase 1B-2A) ----
// Branch order confirmed against a live oracle run (see packages/xlsx/src/index.cjs's
// sheetToFormulae doc comment) rather than assumed from reading the source alone.
function formulaeCase(label, wsFactory) {
  runCase('utils.sheet_to_formulae', () => U.sheet_to_formulae(wsFactory()), () => elixcee.sheet_to_formulae(wsFactory()), [], label);
}

formulaeCase('sparse worksheet, mixed types', () => ({
  A1: { t: 'n', v: 1 },
  B1: { t: 's', v: 'hi' },
  A2: { t: 'b', v: true },
  B2: { t: 'b', v: false },
  '!ref': 'A1:B2',
}));
formulaeCase('dense worksheet', () => {
  const ws = [];
  ws[0] = [{ t: 'n', v: 1 }, { t: 's', v: 'x' }];
  ws['!ref'] = 'A1:B1';
  return ws;
});
formulaeCase('!ref absent -> []', () => ({}));
formulaeCase('sheet is null -> []', () => null);
formulaeCase('string cell', () => ({ A1: { t: 's', v: 'plain' }, '!ref': 'A1:A1' }));
formulaeCase('number cell', () => ({ A1: { t: 'n', v: 42 }, '!ref': 'A1:A1' }));
formulaeCase('number cell v=0', () => ({ A1: { t: 'n', v: 0 }, '!ref': 'A1:A1' }));
formulaeCase('boolean true', () => ({ A1: { t: 'b', v: true }, '!ref': 'A1:A1' }));
formulaeCase('boolean false', () => ({ A1: { t: 'b', v: false }, '!ref': 'A1:A1' }));
formulaeCase('Date cell, no w/f (String(v) fallback, not ISO/serial)', () => ({ A1: { t: 'd', v: new Date(2020, 0, 1) }, '!ref': 'A1:A1' }));
formulaeCase('Date cell with w (w wins over the fallback)', () => ({ A1: { t: 'd', v: new Date(2020, 0, 1), w: '1/1/20' }, '!ref': 'A1:A1' }));
formulaeCase('error cell, no w', () => ({ A1: { t: 'e', v: 7 }, '!ref': 'A1:A1' }));
formulaeCase('error cell with w', () => ({ A1: { t: 'e', v: 7, w: '#DIV/0!' }, '!ref': 'A1:A1' }));
formulaeCase('formula with cached value (formula text wins, not the cached v)', () => ({ A1: { t: 'n', v: 5, f: '1+4' }, '!ref': 'A1:A1' }));
formulaeCase('formula only, no cached v', () => ({ A1: { t: 'n', f: '1+4' }, '!ref': 'A1:A1' }));
formulaeCase('array formula, multi-cell (only top-left contributes, keyed by the range)', () => ({
  A1: { t: 'n', F: 'A1:B2', f: 'SUM(1,2)' },
  B1: { t: 'n', F: 'A1:B2' },
  A2: { t: 'n', F: 'A1:B2' },
  B2: { t: 'n', F: 'A1:B2' },
  '!ref': 'A1:B2',
}));
formulaeCase('array formula, single cell (F has no colon, gets doubled into "A1:A1")', () => ({
  A1: { t: 'n', F: 'A1', f: 'RAND()' },
  '!ref': 'A1:A1',
}));
formulaeCase('F present, f falsy on the "top-left" -> skipped entirely (shared-formula-like non-contributor)', () => ({
  A1: { t: 'n', F: 'A1:B1', f: '' },
  B1: { t: 'n', F: 'A1:B1' },
  '!ref': 'A1:B1',
}));
formulaeCase('empty-ish cell: t set, no v/f/w -> skipped', () => ({ A1: { t: 's' }, '!ref': 'A1:A1' }));
formulaeCase('stub cell t="z" -> skipped', () => ({ A1: { t: 'z' }, '!ref': 'A1:A1' }));
formulaeCase('w present, v absent (w still wins, string cell)', () => ({ A1: { t: 's', w: 'rendered' }, '!ref': 'A1:A1' }));
formulaeCase('invalid cell type with v -> String(v) fallback', () => ({ A1: { t: 'bogus', v: 'raw' }, '!ref': 'A1:A1' }));
formulaeCase('invalid cell type, no v/f/w -> skipped', () => ({ A1: { t: 'bogus' }, '!ref': 'A1:A1' }));
formulaeCase('non-ASCII string content', () => ({ A1: { t: 's', v: 'こんにちは' }, '!ref': 'A1:A1' }));
formulaeCase('reversed !ref -> loop body never runs -> []', () => ({ A1: { t: 'n', v: 1 }, B2: { t: 'n', v: 2 }, '!ref': 'B2:A1' }));
formulaeCase('sparse array hole (dense, missing row)', () => {
  const ws = [];
  ws[0] = [{ t: 'n', v: 1 }];
  ws[2] = [{ t: 'n', v: 3 }];
  ws['!ref'] = 'A1:A3';
  return ws;
});
formulaeCase('sparse object hole (missing cell key entirely)', () => ({ A1: { t: 'n', v: 1 }, '!ref': 'A1:B1' }));
formulaeCase('column boundary (single-column range)', () => ({ C1: { t: 'n', v: 3 }, '!ref': 'C1:C1' }));
// A crafted full-grid !ref (~17.18 billion cells) is confirmed to not return within 25s
// on the real oracle (timeout-guarded subprocess, not assumed) — never call it here;
// see packages/xlsx/src/internal/range-guard.cjs and classify.mjs's
// SAFETY_DIVERGENCE_REGISTRY for ELIXCEE_RANGE_TOO_LARGE.
runUnsafeForOracleCase(
  'utils.sheet_to_formulae',
  elixcee.sheet_to_formulae,
  [{ A1: { t: 'n', v: 1 }, '!ref': 'A1:XFD1048576' }],
  'full-grid !ref (A1:XFD1048576) [oracle does not return within 25s, not called]'
);

// ---- cell_set_hyperlink (Phase 1B-2A) ----
// Each fixture captures both the mutated cell and a boolean for return-value identity
// (ret === cell) — comparing the boolean, not object identity itself, since identity
// doesn't survive normalize()'s structural comparison, but a boolean does.
function hyperlinkCase(label, cellFactory, target, tooltip) {
  runCase(
    'utils.cell_set_hyperlink',
    () => { const cell = cellFactory(); const ret = U.cell_set_hyperlink(cell, target, tooltip); return { cell, identity: ret === cell }; },
    () => { const cell = cellFactory(); const ret = elixcee.cell_set_hyperlink(cell, target, tooltip); return { cell, identity: ret === cell }; },
    [],
    label
  );
}

hyperlinkCase('external URL', () => ({}), 'https://example.com/path');
hyperlinkCase('mailto:', () => ({}), 'mailto:someone@example.com');
hyperlinkCase('file URL', () => ({}), 'file:///C:/Users/test/doc.xlsx');
hyperlinkCase('relative URL', () => ({}), '../sibling/file.xlsx');
hyperlinkCase('fragment', () => ({}), '#Sheet2!A1');
hyperlinkCase('tooltip present', () => ({}), 'https://x.com', 'a tooltip');
hyperlinkCase('tooltip absent', () => ({}), 'https://x.com');
hyperlinkCase('tooltip empty string (falsy, omitted)', () => ({}), 'https://x.com', '');
hyperlinkCase('empty string target -> deletes any .l', () => ({ l: { Target: 'old' } }), '');
hyperlinkCase('null target -> deletes any .l', () => ({ l: { Target: 'old' } }), null);
hyperlinkCase('undefined target -> deletes any .l', () => ({ l: { Target: 'old' } }), undefined);
hyperlinkCase('0 target (falsy, non-nullish) -> deletes', () => ({ l: { Target: 'old' } }), 0);
hyperlinkCase('overwrite existing link', () => ({ l: { Target: '#Old!A1', Tooltip: 'old tip' } }), 'https://new.example.com', 'new tip');
hyperlinkCase('existing cell object with other fields preserved', () => ({ t: 's', v: 'label', z: '@' }), 'https://x.com');
hyperlinkCase('empty object cell', () => ({}), 'https://x.com');
hyperlinkCase('__proto__ as target (stored as a value, not a key)', () => ({}), '__proto__');
hyperlinkCase('constructor as target', () => ({}), 'constructor');
hyperlinkCase('non-ASCII URL', () => ({}), 'https://例え.jp/ページ');
hyperlinkCase('very long URL', () => ({}), 'https://example.com/' + 'a'.repeat(5000));

// ---- cell_set_internal_link (Phase 1B-2A) ----
// Delegates to cell_set_hyperlink with "#"+range — internal representation is the same
// { l: { Target, Tooltip? } } shape, just with the target string always "#"-prefixed;
// verified this matches cell_set_hyperlink's own shape exactly (no separate encoding).
function internalLinkCase(label, cellFactory, range, tooltip) {
  runCase(
    'utils.cell_set_internal_link',
    () => { const cell = cellFactory(); const ret = U.cell_set_internal_link(cell, range, tooltip); return { cell, identity: ret === cell }; },
    () => { const cell = cellFactory(); const ret = elixcee.cell_set_internal_link(cell, range, tooltip); return { cell, identity: ret === cell }; },
    [],
    label
  );
}

internalLinkCase('Sheet1!A1', () => ({}), 'Sheet1!A1');
internalLinkCase('#Sheet1!A1 (double-hash — range is used verbatim)', () => ({}), '#Sheet1!A1');
internalLinkCase('sheet name with spaces', () => ({}), 'Sheet 1!A1');
internalLinkCase("sheet name needing quotes ('Sheet 1'!A1)", () => ({}), "'Sheet 1'!A1");
internalLinkCase('sheet name with apostrophe', () => ({}), "O'Brien!A1");
internalLinkCase('named range (no sheet qualifier)', () => ({}), 'MyNamedRange');
internalLinkCase('fragment only (bare cell ref)', () => ({}), 'A1');
internalLinkCase('empty string range -> "#" target (truthy, creates a link)', () => ({}), '');
internalLinkCase('null range -> "#null" target', () => ({}), null);
internalLinkCase('undefined range -> "#undefined" target', () => ({}), undefined);
internalLinkCase('overwrite existing link', () => ({ l: { Target: '#Old!A1', Tooltip: 'old' } }), 'Sheet2!B2', 'new tip');

// ---- cell_add_comment (Phase 1B-2A) ----
// cell_add_comment returns undefined (no `return` in the oracle source) — unlike
// cell_set_hyperlink/cell_set_internal_link, which return the cell. Captured as a
// boolean (retIsUndefined) alongside the mutated cell, same reasoning as the identity
// checks above.
function commentCase(label, cellFactory, text, author) {
  runCase(
    'utils.cell_add_comment',
    () => { const cell = cellFactory(); const ret = U.cell_add_comment(cell, text, author); return { cell, retIsUndefined: ret === undefined }; },
    () => { const cell = cellFactory(); const ret = elixcee.cell_add_comment(cell, text, author); return { cell, retIsUndefined: ret === undefined }; },
    [],
    label
  );
}

commentCase('text only, author omitted -> defaults to "SheetJS"', () => ({}), 'hello');
commentCase('author present', () => ({}), 'hello', 'Alice');
commentCase('author empty string (falsy, defaults to "SheetJS")', () => ({}), 'hello', '');
commentCase('non-ASCII text', () => ({}), 'こんにちは、世界');
commentCase('text with newlines', () => ({}), 'line1\nline2\r\nline3');
commentCase('empty text (not defaulted, unlike author)', () => ({}), '');
commentCase('null text (stored verbatim, not defaulted)', () => ({}), null);
commentCase('undefined text (stored verbatim, not defaulted)', () => ({}), undefined);
commentCase('long text', () => ({}), 'x'.repeat(5000));
commentCase('__proto__ as author (stored as a value, not a key)', () => ({}), 'hi', '__proto__');
commentCase('__proto__ as text', () => ({}), '__proto__');
commentCase('empty object cell', () => ({}), 'hi');
// Multiple comments: two independent calls on the same cell append, not replace.
runCase(
  'utils.cell_add_comment',
  () => { const cell = {}; U.cell_add_comment(cell, 'first', 'A'); U.cell_add_comment(cell, 'second', 'B'); return cell; },
  () => { const cell = {}; elixcee.cell_add_comment(cell, 'first', 'A'); elixcee.cell_add_comment(cell, 'second', 'B'); return cell; },
  [],
  'multiple comments append to the same .c array'
);
// Existing comment array (pre-populated .c, e.g. from a parsed file) is appended to.
runCase(
  'utils.cell_add_comment',
  () => { const cell = { c: [{ t: 'existing', a: 'Prior' }] }; U.cell_add_comment(cell, 'new', 'New'); return cell; },
  () => { const cell = { c: [{ t: 'existing', a: 'Prior' }] }; elixcee.cell_add_comment(cell, 'new', 'New'); return cell; },
  [],
  'pre-existing .c comment array is appended to, not replaced'
);
// Comment object key order — not captured by classify()'s order-insensitive deepEqual,
// so checked with a direct assertion (same pattern as Phase 1A's safe_decode_range
// self-check).
{
  const oracleCell = {};
  U.cell_add_comment(oracleCell, 'txt', 'auth');
  const elixceeCell = {};
  elixcee.cell_add_comment(elixceeCell, 'txt', 'auth');
  assert.deepEqual(
    Object.keys(oracleCell.c[0]),
    Object.keys(elixceeCell.c[0]),
    'cell_add_comment: pushed comment object key order must match the oracle (t, then a)'
  );
  assert.deepEqual(Object.keys(oracleCell.c[0]), ['t', 'a'], 'sanity: the oracle itself pushes {t, a} in that order');
}

// ---- sheet_set_array_formula (Phase 1B-2A) ----
// Each fixture captures both the resulting worksheet and a boolean for return-value
// identity (ret === ws).
function arrayFormulaCase(label, wsFactory, range, formula, dynamic) {
  runCase(
    'utils.sheet_set_array_formula',
    () => { const ws = wsFactory(); const ret = U.sheet_set_array_formula(ws, range, formula, dynamic); return { ws, identity: ret === ws }; },
    () => { const ws = wsFactory(); const ret = elixcee.sheet_set_array_formula(ws, range, formula, dynamic); return { ws, identity: ret === ws }; },
    [],
    label
  );
}

arrayFormulaCase('"A1:B2" string range', () => ({}), 'A1:B2', 'SUM(A1:B2)');
arrayFormulaCase('range object', () => ({}), { s: { r: 0, c: 0 }, e: { r: 1, c: 1 } }, 'SUM(A1:B2)');
arrayFormulaCase(
  'single cell: string "A1:A1" kept verbatim vs object range collapsed to "A1" by encode_range',
  () => ({}),
  'A1:A1',
  'X'
);
arrayFormulaCase('single cell: object range (collapses .F to "A1", no colon)', () => ({}), { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }, 'X');
arrayFormulaCase('multiple cells (3x1)', () => ({}), 'A1:C1', 'X');
arrayFormulaCase('reversed range (s>e after safe_decode_range) -> loop never runs, ws unchanged', () => ({}), 'B2:A1', 'X');
arrayFormulaCase(
  'existing cell: other fields (e.g. .z) survive the mutation',
  () => ({ A1: { t: 's', v: 'old', z: '@' } }),
  'A1:A1',
  'NEW'
);
arrayFormulaCase(
  'existing formula: overwritten by the new one',
  () => ({ A1: { t: 'n', v: 1, f: 'OLD()' } }),
  'A1:A1',
  'NEW()'
);
arrayFormulaCase('dense worksheet target', () => { const ws = []; ws[0] = [{ t: 's', v: 'x' }]; return ws; }, 'A1:B2', 'D');
arrayFormulaCase('sparse worksheet target', () => ({}), 'A1:B2', 'S');
arrayFormulaCase('!ref absent -> stays absent (never set by this function)', () => ({}), 'A1:A1', 'X');
arrayFormulaCase(
  '!ref present -> stays exactly as-is, NOT extended even when the range goes past it',
  () => { const ws = { A1: { t: 'n', v: 1 }, '!ref': 'A1:A1' }; return ws; },
  'D5:E6',
  'X'
);
arrayFormulaCase('dynamic:true -> .D=true on the top-left cell only', () => ({}), 'A1:B1', 'X', true);
arrayFormulaCase('dynamic:false -> no .D key at all (not set to false)', () => ({}), 'A1:A1', 'X', false);
arrayFormulaCase('dynamic omitted -> no .D key', () => ({}), 'A1:A1', 'X');
arrayFormulaCase('formula with leading =', () => ({}), 'A1:A1', '=SUM(1,2)');
arrayFormulaCase('formula without leading =', () => ({}), 'A1:A1', 'SUM(1,2)');
arrayFormulaCase('empty formula string', () => ({}), 'A1:A1', '');
arrayFormulaCase('null formula (.f set to null, not omitted)', () => ({}), 'A1:A1', null);
arrayFormulaCase('undefined formula (.f key present with value undefined, not omitted)', () => ({}), 'A1:A1', undefined);
arrayFormulaCase('invalid range string (no colon, garbage) -> safe_decode_range degrades, does not throw', () => ({}), 'garbage', 'X');
runCase(
  'utils.sheet_set_array_formula',
  () => { try { U.sheet_set_array_formula({}, null, 'X'); return { threw: false }; } catch (e) { return { threw: true, ctor: e.constructor.name }; } },
  () => { try { elixcee.sheet_set_array_formula({}, null, 'X'); return { threw: false }; } catch (e) { return { threw: true, ctor: e.constructor.name }; } },
  [],
  'null range -> native TypeError on both sides (no defensive guard added)'
);
runCase(
  'utils.sheet_set_array_formula',
  () => { try { U.sheet_set_array_formula({}, undefined, 'X'); return { threw: false }; } catch (e) { return { threw: true, ctor: e.constructor.name }; } },
  () => { try { elixcee.sheet_set_array_formula({}, undefined, 'X'); return { threw: false }; } catch (e) { return { threw: true, ctor: e.constructor.name }; } },
  [],
  'undefined range -> native TypeError on both sides (no defensive guard added)'
);

// ---- Prototype pollution / prototype corruption probes across Phase 1B-2A APIs ----
// None of the four cell-metadata functions use a caller-supplied string as an object
// key (confirmed by reading the ported algorithms: hyperlink targets/tooltips and
// comment text/author are always VALUES; sheet_set_array_formula's cell keys always
// come from encode_cell's well-formed output) — these probes exist to have that
// confirmed on record empirically, not just by code reading, matching this project's
// established practice (see book_append_sheet's Phase 1A fix for the one case where
// this assumption was actually wrong).
for (const poisoned of ['__proto__', 'constructor', 'prototype', 'toString', 'hasOwnProperty']) {
  hyperlinkCase(`hyperlink target="${poisoned}"`, () => ({}), poisoned);
  commentCase(`comment author="${poisoned}"`, () => ({}), 'txt', poisoned);
  commentCase(`comment text="${poisoned}"`, () => ({}), poisoned);
}
{
  const before = Object.getPrototypeOf({});
  assert.equal(before, Object.prototype, 'sanity: Object.prototype unpolluted before the probes above');
  // Re-check after every poisoned-key fixture ran above — must still be Object.prototype,
  // and a fresh {} must still behave like a normal empty object.
  assert.equal(Object.getPrototypeOf({}), Object.prototype, 'Phase 1B-2A probes must not have polluted the global Object.prototype');
  assert.equal(({}).polluted, undefined, 'Phase 1B-2A probes must not have added a "polluted" property reachable from a fresh {}');
}

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
// accepted. classify.mjs's UNSUPPORTED_ALLOWLIST is empty as of Phase 1B-2B (the SSF
// backend closed the two cases that used to be registered here) — kept in ACCEPTABLE
// and reported below regardless, since any future gap is expected to be registered by
// exact (api, caseId), never by api alone, and printed below by label so a real bug
// hiding behind the same api key would still be visible on review, not silently
// absorbed.
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
