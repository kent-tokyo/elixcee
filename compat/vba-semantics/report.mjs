// Joins cases.json + expected-results.json (documented real-VBA ground truth) against
// results/elixcee-results.json (what elixcee actually produced) and classifies every
// case. This is the suite's whole point: a function that runs without error and returns
// a plausible-but-wrong VALUE is invisible to compat/corpus/'s own PASS/FAIL-shaped
// classifiers — this one compares the actual value, not just whether something errored.
//
// Verdicts:
//   MATCH_DOCUMENTED_SEMANTICS - actual value matches the documented-real-VBA expected
//                                value exactly.
//   EXPECTED_ERROR             - expected an error, and got exactly that error message.
//   NONDETERMINISTIC           - case is registered as having no fixed expected value
//                                (e.g. Now()'s sub-second component); only checked for
//                                running without erroring.
//   KNOWN_LIMITATION           - actual differs from documented-real-VBA semantics, but
//                                this exact case is registered (via `knownLimitation` in
//                                expected-results.json) as an already-disclosed gap —
//                                never silently inferred, always requires a reason string
//                                written by a human who looked at the actual divergence.
//   BUG                        - actual differs from documented-real-VBA semantics and
//                                is NOT registered as a known limitation. This is what
//                                the suite exists to catch; must be 0 for the gate to pass.
//   UNCLASSIFIED                - something structurally wrong (no result recorded, no
//                                cell found at the expected address, an expected.kind
//                                this script doesn't recognize). Must also be 0 — an
//                                UNCLASSIFIED case is a bug in the suite itself, not a
//                                verdict on elixcee, and should never be "explained away".
//
// CI gate: BUG and UNCLASSIFIED must both be 0.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as numericRef from './reference/numeric.mjs';

const DIR = path.dirname(fileURLToPath(import.meta.url));

export const VERDICTS = /** @type {const} */ ([
  'MATCH_DOCUMENTED_SEMANTICS',
  'EXPECTED_ERROR',
  'NONDETERMINISTIC',
  'KNOWN_LIMITATION',
  'BUG',
  'UNCLASSIFIED',
]);

function loadJson(name) {
  return JSON.parse(fs.readFileSync(path.join(DIR, name), 'utf8'));
}

function findCellValue(result, address) {
  const cell = (result.cells || []).find(c => c.address === address);
  return cell ? cell.value : undefined;
}

/**
 * @param {Array<{id: string, category: string}>} cases
 * @param {Record<string, any>} expectedResults
 * @param {Array<{id: string, ok: boolean, cells?: any[], error?: {kind: string, message: string, code?: string}}>} elixceeResults
 * @param {Record<string, Function>} computedFns - name -> zero-arg function returning the fresh expected value
 */
export function classify(cases, expectedResults, elixceeResults, computedFns) {
  const resultsById = new Map(elixceeResults.map(r => [r.id, r]));
  const records = [];

  for (const c of cases) {
    const expected = expectedResults[c.id];
    const result = resultsById.get(c.id);

    if (!expected) {
      records.push({ id: c.id, category: c.category, verdict: 'UNCLASSIFIED', reason: 'no expected-results.json entry for this case id' });
      continue;
    }
    if (!result) {
      records.push({ id: c.id, category: c.category, verdict: 'UNCLASSIFIED', reason: 'no result recorded — results/elixcee-results.json is stale or incomplete; rerun run-elixcee.mjs' });
      continue;
    }

    if (expected.kind === 'nondeterministic') {
      records.push(
        result.ok
          ? { id: c.id, category: c.category, verdict: 'NONDETERMINISTIC', reason: expected.reason }
          : { id: c.id, category: c.category, verdict: 'BUG', reason: `expected to run without erroring (value itself not asserted, per "${expected.reason}"), but failed: ${result.error?.message}` },
      );
      continue;
    }

    if (expected.kind === 'error') {
      const actualMessage = result.ok ? null : result.error?.message;
      const matches = !result.ok && actualMessage === expected.errorMessage;
      if (matches) {
        records.push({ id: c.id, category: c.category, verdict: 'EXPECTED_ERROR', reason: expected.reason });
      } else if (expected.knownLimitation) {
        records.push({
          id: c.id, category: c.category, verdict: 'KNOWN_LIMITATION',
          reason: expected.knownLimitation,
          documented: `real VBA: error "${expected.errorMessage}"`,
          actual: result.ok ? 'ran without error' : `error "${actualMessage}"`,
        });
      } else {
        records.push({
          id: c.id, category: c.category, verdict: 'BUG',
          reason: `expected error "${expected.errorMessage}", got ${result.ok ? 'no error (ran successfully)' : `error "${actualMessage}"`}`,
        });
      }
      continue;
    }

    if (expected.kind === 'no_cells') {
      // The scenario must run successfully (ok:true) and write zero cells — used for
      // "this line must never execute" assertions (a guarded Range write behind an Exit/
      // GoTo/loop-bound that must not be reached). --json's `cells` array only ever lists
      // non-empty cells, so a guarded-off write simply never appears when this holds.
      if (!result.ok) {
        const failMsg = `expected to run successfully with no cells written, but the scenario errored: ${result.error?.message}`;
        records.push(
          expected.knownLimitation
            ? { id: c.id, category: c.category, verdict: 'KNOWN_LIMITATION', reason: expected.knownLimitation, actual: failMsg }
            : { id: c.id, category: c.category, verdict: 'BUG', reason: failMsg },
        );
        continue;
      }
      const cellCount = (result.cells || []).length;
      if (cellCount === 0) {
        records.push({ id: c.id, category: c.category, verdict: 'MATCH_DOCUMENTED_SEMANTICS', reason: expected.reason });
      } else if (expected.knownLimitation) {
        records.push({
          id: c.id, category: c.category, verdict: 'KNOWN_LIMITATION',
          reason: expected.knownLimitation,
          documented: 'real VBA: no cells written',
          actual: `${cellCount} cell(s) written: ${JSON.stringify(result.cells)}`,
        });
      } else {
        records.push({
          id: c.id, category: c.category, verdict: 'BUG',
          reason: `expected no cells written (${expected.reason}), but got ${cellCount}: ${JSON.stringify(result.cells)}`,
        });
      }
      continue;
    }

    if (expected.kind === 'value' || expected.kind === 'computed') {
      if (!result.ok) {
        const failMsg = `expected value ${JSON.stringify(expected.value)}, but the scenario errored: ${result.error?.message}`;
        records.push(
          expected.knownLimitation
            ? { id: c.id, category: c.category, verdict: 'KNOWN_LIMITATION', reason: expected.knownLimitation, documented: `real VBA: ${JSON.stringify(expected.value)}`, actual: failMsg }
            : { id: c.id, category: c.category, verdict: 'BUG', reason: failMsg },
        );
        continue;
      }

      let expectedValue = expected.value;
      if (expected.kind === 'computed') {
        const fn = computedFns[expected.computedBy];
        if (!fn) {
          records.push({ id: c.id, category: c.category, verdict: 'UNCLASSIFIED', reason: `no computed-value function registered for "${expected.computedBy}"` });
          continue;
        }
        expectedValue = fn();
      }

      const actualValue = findCellValue(result, expected.address);
      if (actualValue === undefined) {
        records.push({ id: c.id, category: c.category, verdict: 'UNCLASSIFIED', reason: `no cell found at ${expected.address} in the scenario's output` });
        continue;
      }

      const matches = actualValue === expectedValue
        || (typeof actualValue === 'number' && typeof expectedValue === 'number' && Math.abs(actualValue - expectedValue) < 1e-9);
      if (matches) {
        records.push({ id: c.id, category: c.category, verdict: 'MATCH_DOCUMENTED_SEMANTICS', reason: expected.reason });
      } else if (expected.knownLimitation) {
        records.push({
          id: c.id, category: c.category, verdict: 'KNOWN_LIMITATION',
          reason: expected.knownLimitation,
          documented: `real VBA: ${JSON.stringify(expectedValue)}`,
          actual: JSON.stringify(actualValue),
        });
      } else {
        records.push({
          id: c.id, category: c.category, verdict: 'BUG',
          reason: `expected ${JSON.stringify(expectedValue)} (${expected.reason}), got ${JSON.stringify(actualValue)}`,
        });
      }
      continue;
    }

    records.push({ id: c.id, category: c.category, verdict: 'UNCLASSIFIED', reason: `unrecognized expected.kind: "${expected.kind}"` });
  }

  return records;
}

function main() {
  const cases = loadJson('cases.json');
  const expectedResults = loadJson('expected-results.json');
  const elixceeResults = loadJson('results/elixcee-results.json');
  const computedFns = { todaySerial: numericRef.todaySerial, todayDateString: numericRef.todayDateString };

  const records = classify(cases, expectedResults, elixceeResults, computedFns);

  const counts = Object.fromEntries(VERDICTS.map(v => [v, 0]));
  for (const r of records) counts[r.verdict] = (counts[r.verdict] ?? 0) + 1;

  const outDir = path.join(DIR, 'results');
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(
    path.join(outDir, 'report.json'),
    JSON.stringify({ total: records.length, counts, records }, null, 2) + '\n',
  );

  console.log(`classified ${records.length} cases:`);
  for (const v of VERDICTS) {
    if (counts[v] > 0) console.log(`  ${v}: ${counts[v]}`);
  }

  const knownLimitations = records.filter(r => r.verdict === 'KNOWN_LIMITATION');
  if (knownLimitations.length > 0) {
    console.log(`\n${knownLimitations.length} KNOWN_LIMITATION cases (disclosed, not gating):`);
    for (const r of knownLimitations) console.log(`  ${r.id}: documented=${r.documented} actual=${r.actual}`);
  }

  const gateFailures = records.filter(r => r.verdict === 'BUG' || r.verdict === 'UNCLASSIFIED');
  console.log('\nwrote results/report.json');
  if (gateFailures.length > 0) {
    console.log(`\n${gateFailures.length} BUG/UNCLASSIFIED cases (gate failure):`);
    for (const r of gateFailures) console.log(`  [${r.verdict}] ${r.id}: ${r.reason}`);
    process.exitCode = 1;
    return;
  }

  console.log(`\nGate passed: 0 BUG, 0 UNCLASSIFIED across ${records.length} cases.`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main();
}
