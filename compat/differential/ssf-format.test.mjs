// Permanent acceptance suite for the SSF backend decision (Phase 1B-2B): compares the
// real oracle's bundled format engine (XLSX.SSF.format, xlsx.js's inline SSF_format —
// NOT the standalone `ssf` package, which is a different export surface even though
// confirmed behaviorally identical) against packages/xlsx/src/internal/ssf-adapter.cjs
// (the thin wrapper around the `ssf@0.11.2` runtime dependency this package actually
// ships). This is the same 819+ case comparison run once, ad hoc, before deciding to
// take `ssf` as a dependency — kept here permanently as the regression guard for that
// decision and for any future backend swap (see docs/xlsx-architecture.md).
//
// Tests the internal adapter directly, not through format_cell — format_cell's own
// cell-level orchestration (caching, BErr lookup, the two-try fallthrough) is covered by
// compat/differential/xlsx-utils.test.mjs's format_cell fixtures. This file is purely
// about "does the chosen SSF backend evaluate format strings/numFmtIds identically to
// the oracle's own engine."
import assert from 'node:assert/strict';
import XLSX from 'xlsx';
import { format as elixceeSsfFormat } from '../../packages/xlsx/src/internal/ssf-adapter.cjs';
import { classify } from './classify.mjs';
import { normalize } from './normalize.mjs';

const oracleSsfFormat = XLSX.SSF.format;

const results = [];
function record(label, verdict) {
  results.push({ label, verdict });
}

function invoke(fn, args) {
  try {
    return { threw: false, value: normalize(fn(...args)) };
  } catch (e) {
    return { threw: true, message: e.message, code: e.code };
  }
}

function ssfCase(label, fmt, v, opts) {
  const args = opts === undefined ? [fmt, v] : [fmt, v, opts];
  const oracleVal = invoke(oracleSsfFormat, args);
  const elixceeVal = invoke(elixceeSsfFormat, args);
  const verdict = classify({ api: 'internal.ssf_format', oracleA: oracleVal, elixcee: elixceeVal, elixceeErrorCode: elixceeVal.code });
  record(label, verdict);
  return verdict;
}

// ---- table_fmt built-ins (numFmtId -> literal format string, no indirection) ----
const TABLE_FMT_IDS = [0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 37, 38, 39, 40, 45, 46, 47, 48, 49, 56];

// ---- SSF_default_map indirection ids (numFmtId -> another numFmtId's format) ----
const DEFAULT_MAP_IDS = [
  5, 6, 7, 8, 23, 24, 25, 26, 27, 28, 29, 30, 31, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82,
];

// ids with no table_fmt entry and no SSF_default_map entry -> SSF_default_str/"General" fallback
const UNDEFINED_IDS = [100, 200, 999, 65535];

const VALUES = [0, 1, -1, 0.5, -0.5, 42, -42, 1234.5678, -1234.5678, 1e10, 1e-10, 46027, 60, 2958465, 0.0001, -0.0001, 'text', true, false];

for (const id of [...TABLE_FMT_IDS, ...DEFAULT_MAP_IDS, ...UNDEFINED_IDS]) {
  for (const v of VALUES) {
    ssfCase(`numFmtId ${id}, v=${JSON.stringify(v)}`, id, v);
  }
}

// ---- custom format-code strings (not numFmtIds) ----
const CUSTOM_FORMATS = [
  'General',
  '0.00',
  '#,##0.00',
  '@',
  '0%',
  '0.00%',
  '0.00;[Red]-0.00',
  '0.00;(0.00)',
  '[>100]0.00;[<0]-0.00;0.00',
  '# ?/?',
  '# ??/??',
  '0.00E+00',
  'yyyy-mm-dd',
  'm/d/yy h:mm',
  '[h]:mm:ss',
  '#,##0',
  '"$"#,##0.00',
  '0.00;[Red]-0.00;0.00;"TEXT:"@', // 4-section: pos;neg;zero;text
  '[>100]"big";[<0]"neg";"small"', // conditional literal-string sections
];
for (const fmt of CUSTOM_FORMATS) {
  for (const v of VALUES) {
    ssfCase(`format ${JSON.stringify(fmt)}, v=${JSON.stringify(v)}`, fmt, v);
  }
}

// ---- date1904 / custom table option passthrough ----
ssfCase('date1904 true', 14, 46027, { date1904: true });
ssfCase('date1904 false explicit', 14, 46027, { date1904: false });
ssfCase('custom table option', 200, 42, { table: { 200: '0.0"custom"' } });

// ---- date-serial boundaries ----
for (const v of [-1, 0, 60, 2958465, 2958466]) {
  ssfCase(`date boundary v=${v}`, 14, v);
}

// ---- null / undefined values ----
ssfCase('null value', 'General', null);
ssfCase('undefined value', 'General', undefined);
ssfCase('NaN value', 'General', NaN);
ssfCase('Infinity value', 'General', Infinity);
ssfCase('-Infinity value', 'General', -Infinity);

// ---- security probes: verified live (not assumed) before being added here as MATCH
// fixtures. Each was run through the real oracle directly to confirm it neither hangs
// nor throws unexpectedly before being trusted as a normal differential case. ----
ssfCase('very long format code (2000 "0" chars)', '0'.repeat(2000), 42);
ssfCase('many sections (50 semicolon-separated)', new Array(50).fill('0.00').join(';'), 42);
ssfCase('many quoted/escaped literals', '"a"0"b"0"c"0"d"0"e"0.00"f"', 42);
ssfCase('abnormally long numeric string', Number('1'.repeat(300)), 'General');
ssfCase('crafted deeply-nested conditional format', '[>1][>2][>3]0;[<1]0;0', 5);

// ---- summary ----
console.log('\n=== ssf-format.test.mjs summary ===');
const ACCEPTABLE = new Set(['MATCH', 'INTENTIONAL_SAFETY_DIVERGENCE', 'INTENTIONAL_SECURITY_DIVERGENCE', 'UNSUPPORTED']);
let totalMatch = 0;
let totalOther = 0;
const other = [];
for (const r of results) {
  if (r.verdict === 'MATCH') totalMatch += 1;
  else if (ACCEPTABLE.has(r.verdict)) totalOther += 1; // acceptable but non-MATCH (none expected here)
  else other.push(r);
}
console.log(`total cases: ${results.length}`);
console.log(`MATCH: ${totalMatch}`);
console.log(`other acceptable (safety/security/unsupported): ${totalOther}`);
console.log(`UNCLASSIFIED/BUG: ${other.length}`);
for (const o of other) console.log(`      ${o.verdict}: ${o.label}`);

if (other.length > 0) {
  console.error('\nssf-format.test.mjs FAILED: at least one case is not MATCH or a registered divergence.');
  process.exit(1);
}
assert.ok(results.length >= 819, `expected at least 819 comparisons (the original ad hoc measurement), got ${results.length}`);
console.log('\nssf-format.test.mjs passed: every case is MATCH or a registered divergence.');
