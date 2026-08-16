// Permanent Phase 1A differential test suite: runs the boundary-value matrix specified
// for Phase 1A through both the real oracle (xlsx@0.18.5) and @elixcee/xlsx, classifying
// every case with classify.mjs. This file itself IS the runnable check — it exits
// non-zero if anything is left UNCLASSIFIED (or otherwise fails to MATCH), so nothing
// can pass CI by accident.
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
function runCase(api, oracleFn, elixceeFn, args, label) {
  const oracleVal = invoke(oracleFn, args);
  const elixceeVal = invoke(elixceeFn, args);
  const verdict = classify({
    api,
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
// Only MATCH and registered-divergence verdicts (INTENTIONAL_SAFETY_DIVERGENCE,
// INTENTIONAL_SECURITY_DIVERGENCE) are acceptable outcomes here — everything else
// (UNSUPPORTED, BUG, ORACLE_AMBIGUITY, NONDETERMINISTIC, UNCLASSIFIED) fails the run.
// Per the classification rule: normal-input divergences are BUG, only pre-registered
// gaps are UNSUPPORTED/*_DIVERGENCE, anything else is UNCLASSIFIED — none of those are
// silently accepted.
const ACCEPTABLE = new Set(['MATCH', 'INTENTIONAL_SAFETY_DIVERGENCE', 'INTENTIONAL_SECURITY_DIVERGENCE']);
const byApi = new Map();
let totalMatch = 0;
let totalSafetyDivergence = 0;
let totalSecurityDivergence = 0;
let totalUnsupported = 0;
let totalUnclassified = 0;
let totalBug = 0;
for (const r of results) {
  if (!byApi.has(r.api)) byApi.set(r.api, { total: 0, ok: 0, other: [] });
  const bucket = byApi.get(r.api);
  bucket.total += 1;
  if (ACCEPTABLE.has(r.verdict)) bucket.ok += 1;
  else bucket.other.push({ label: r.label, verdict: r.verdict });
  if (r.verdict === 'MATCH') totalMatch += 1;
  else if (r.verdict === 'INTENTIONAL_SAFETY_DIVERGENCE') totalSafetyDivergence += 1;
  else if (r.verdict === 'INTENTIONAL_SECURITY_DIVERGENCE') totalSecurityDivergence += 1;
  else if (r.verdict === 'UNSUPPORTED') totalUnsupported += 1;
  else if (r.verdict === 'UNCLASSIFIED') totalUnclassified += 1;
  else if (r.verdict === 'BUG') totalBug += 1;
}

console.log('\n=== Phase 1A differential summary ===');
let anyFailure = false;
for (const [api, bucket] of byApi) {
  const status = bucket.other.length === 0 ? 'OK' : 'FAIL';
  if (bucket.other.length > 0) anyFailure = true;
  console.log(`${status}  ${api}: ${bucket.ok}/${bucket.total} acceptable`);
  for (const o of bucket.other) console.log(`      ${o.verdict}: ${o.label}`);
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
  console.error('\nPhase 1A differential suite FAILED: at least one case is not an acceptable verdict.');
  process.exit(1);
}
console.log('\nPhase 1A differential suite passed: every case is MATCH or a registered divergence.');
