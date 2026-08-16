// Phase 0 "prove the plumbing works" demo — not real coverage. Runs one input through
// the real oracle twice (proving oracle-vs-oracle comparison, after normalization,
// yields MATCH) and once through a placeholder elixcee stub that is NOT registered in
// classify.mjs's UNSUPPORTED_ALLOWLIST (proving an un-registered divergence comes back
// UNCLASSIFIED, not a free pass — see classify.mjs for why auto-UNSUPPORTED is not a
// thing). Exits non-zero if either assertion fails; this script itself IS the runnable
// check for this harness, per Ponytail's "non-trivial logic needs one runnable check"
// rule.
import XLSX from 'xlsx';
import { classify } from './classify.mjs';

// Strips fields that legitimately vary run-to-run (e.g. XLSX.write's docProps/core.xml
// timestamp) before comparing two "should be equivalent" results. See classify.mjs's
// module doc comment for why raw-byte/raw-timestamp comparison is the wrong bar.
function normalize(workbook) {
  const clone = JSON.parse(JSON.stringify(workbook));
  if (clone.Props) {
    delete clone.Props.CreatedDate;
    delete clone.Props.ModifiedDate;
  }
  return clone;
}

// Stands in for a future real `@elixcee/xlsx` call. Phase 0 has no Rust/JS-backed XLSX
// compat logic yet, so this always returns a placeholder that does NOT match the oracle.
// It is deliberately NOT registered in classify.mjs's UNSUPPORTED_ALLOWLIST, to prove an
// un-registered divergence never auto-passes as "not a bug."
function elixceeStub(_input) {
  return { __elixceeStub: true, note: 'not implemented yet (Phase 0 placeholder)' };
}

const wb = XLSX.utils.book_new();
const ws = XLSX.utils.aoa_to_sheet([
  [1, 2],
  ['a', 'b'],
]);
XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

// Case A: oracle vs. itself (same buffer, parsed twice) — must MATCH after normalization.
const oracleA = normalize(XLSX.read(buf, { type: 'buffer' }));
const oracleB = normalize(XLSX.read(buf, { type: 'buffer' }));
const verdictA = classify({ api: 'read', oracleA, oracleB });

// Case B: oracle vs. an un-registered elixcee placeholder — must classify as
// UNCLASSIFIED (and callers must treat that as a failure), never a silent pass.
const oracleSheetJson = XLSX.utils.sheet_to_json(ws, { header: 1 });
const elixceeResult = elixceeStub(buf);
const verdictB = classify({ api: 'utils.sheet_to_json', oracleA: oracleSheetJson, elixcee: elixceeResult });

const results = [
  { case: 'oracle-vs-oracle (read same buffer twice, normalized)', verdict: verdictA },
  { case: 'oracle-vs-unregistered-elixcee-stub (sheet_to_json)', verdict: verdictB },
];
console.log(JSON.stringify(results, null, 2));

const ok = verdictA === 'MATCH' && verdictB === 'UNCLASSIFIED';
if (!ok) {
  console.error('run-demo.mjs: expected [MATCH, UNCLASSIFIED], got', [verdictA, verdictB]);
  process.exit(1);
}
console.log('run-demo.mjs: plumbing check passed (unregistered divergence correctly failed closed)');
