// Classifies elixcee's own pass/fail outcome across every corpus scenario — a different,
// oracle-independent axis from classify.mjs (which compares elixcee against LibreOffice/
// Excel and needs a working oracle to produce anything). This script never needs an
// oracle: it explains "did elixcee do what elixcee itself should do", turning "570/581
// scenarios pass" into "581/581 scenarios explained" by classifying every failure against
// expected-outcomes.json, the committed registry of scenario IDs whose correct outcome is
// something other than plain success.
//
// Run: `node classify-elixcee-outcomes.mjs` from compat/corpus/, after
// `node run-elixcee.mjs` has written results/elixcee-results.json.
//
// Verdicts:
//   PASS                    - the scenario ran successfully (ok: true).
//   EXPECTED_RUNTIME_ERROR  - registered in expected-outcomes.json as a runtime error that
//                             correctly matches real VBA behavior (e.g. division by zero).
//   EXPECTED_UNSUPPORTED   - registered as hitting a deliberately-unimplemented feature.
//   NONDETERMINISTIC       - registered as depending on real-world state elixcee doesn't
//                             implement (e.g. Timer()).
//   MISMATCH                - registered in expected-outcomes.json, but the actual
//                             failure's error kind and/or message doesn't exactly match
//                             what was registered — something about *how* or *why* it
//                             fails changed. The old classification no longer applies and
//                             must be re-verified by a human, not silently kept.
//   UNEXPLAINED             - failed, and isn't registered at all. Every one of these is
//                             either a genuine bug or a scenario that needs triaging into
//                             expected-outcomes.json with a real reason.
//
// Matching is scenario-ID-keyed and checks the exact `error.kind` and exact `error.message`
// string against what's registered — never a category-wide rule and never a fuzzy/partial
// message match, so an unrelated new failure under the same category, or an existing
// failure that starts failing for a *different* reason, can't hide behind an
// already-registered entry.
//
// CI gate: MISMATCH and UNEXPLAINED must both be 0 (see main()'s exit code). A "harness"
// kind result (crash, timeout/hang, or unparseable stdout — see run-elixcee.mjs's
// catch-all) can never be registered as expected in expected-outcomes.json and always
// counts as UNEXPLAINED — there is no such thing as an "expected hang" or "expected panic".
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const DIR = path.dirname(fileURLToPath(import.meta.url));

export const VERDICTS = /** @type {const} */ ([
  'PASS',
  'EXPECTED_RUNTIME_ERROR',
  'EXPECTED_UNSUPPORTED',
  'NONDETERMINISTIC',
  'MISMATCH',
  'UNEXPLAINED',
]);

function loadJson(name) {
  return JSON.parse(fs.readFileSync(path.join(DIR, name), 'utf8'));
}

/**
 * @param {Array<{id: string, category: string}>} scenarios
 * @param {Array<{id: string, ok: boolean, error?: {kind: string, message: string, code?: string}}>} elixceeResults
 * @param {Record<string, {classification: string, expectedErrorKind: string, expectedErrorMessage: string, reason: string}>} expectedOutcomes
 */
export function classify(scenarios, elixceeResults, expectedOutcomes) {
  const resultsById = new Map(elixceeResults.map(r => [r.id, r]));
  const records = [];

  for (const scenario of scenarios) {
    const result = resultsById.get(scenario.id);
    const expected = expectedOutcomes[scenario.id];

    if (!result) {
      records.push({
        id: scenario.id,
        category: scenario.category,
        verdict: 'UNEXPLAINED',
        reason: 'no result recorded for this scenario id — results/elixcee-results.json is stale or incomplete; rerun run-elixcee.mjs',
      });
      continue;
    }

    if (result.ok) {
      if (expected) {
        // Registered as an expected failure but it now passes. An improvement, not a
        // gate failure — but the registry entry is stale and should be removed.
        records.push({
          id: scenario.id,
          category: scenario.category,
          verdict: 'PASS',
          reason: `now passes; expected-outcomes.json still has a stale "${expected.classification}" entry for this id — remove it`,
          staleRegistryEntry: true,
        });
      } else {
        records.push({ id: scenario.id, category: scenario.category, verdict: 'PASS' });
      }
      continue;
    }

    // Failed. A harness-level failure (crash/timeout/unparseable output) is never a valid
    // "expected" outcome, regardless of what's registered for this id.
    if (result.error?.kind === 'harness') {
      records.push({
        id: scenario.id,
        category: scenario.category,
        verdict: 'UNEXPLAINED',
        reason: `harness-level failure (${result.error.code}): ${result.error.message} — a crash, timeout/hang, or unparseable output is never an expected outcome`,
      });
      continue;
    }

    if (!expected) {
      records.push({
        id: scenario.id,
        category: scenario.category,
        verdict: 'UNEXPLAINED',
        reason: `failed (${result.error?.kind ?? 'unknown'}: ${result.error?.message ?? 'no message'}) and is not registered in expected-outcomes.json`,
      });
      continue;
    }

    const kindMatches = result.error?.kind === expected.expectedErrorKind;
    const messageMatches = result.error?.message === expected.expectedErrorMessage;
    if (!kindMatches || !messageMatches) {
      records.push({
        id: scenario.id,
        category: scenario.category,
        verdict: 'MISMATCH',
        reason: `registered as "${expected.classification}" expecting kind="${expected.expectedErrorKind}" message="${expected.expectedErrorMessage}", but actually got kind="${result.error?.kind}" message="${result.error?.message}"`,
      });
      continue;
    }

    records.push({
      id: scenario.id,
      category: scenario.category,
      verdict: expected.classification,
      reason: expected.reason,
    });
  }

  return records;
}

function main() {
  const scenariosRaw = loadJson('scenarios.json');
  const scenarios = Array.isArray(scenariosRaw) ? scenariosRaw : scenariosRaw.scenarios;
  const elixceeResults = loadJson('results/elixcee-results.json');
  const { _readme, ...expectedOutcomes } = loadJson('expected-outcomes.json');

  const records = classify(scenarios, elixceeResults, expectedOutcomes);

  const counts = Object.fromEntries(VERDICTS.map(v => [v, 0]));
  for (const r of records) counts[r.verdict] = (counts[r.verdict] ?? 0) + 1;

  const outDir = path.join(DIR, 'results');
  fs.mkdirSync(outDir, { recursive: true });
  fs.writeFileSync(
    path.join(outDir, 'elixcee-outcomes-classified.json'),
    JSON.stringify({ total: records.length, counts, records }, null, 2) + '\n',
  );

  console.log(`classified ${records.length} scenarios:`);
  for (const v of VERDICTS) {
    if (counts[v] > 0) console.log(`  ${v}: ${counts[v]}`);
  }

  const stale = records.filter(r => r.staleRegistryEntry);
  if (stale.length > 0) {
    console.log(`\n${stale.length} stale expected-outcomes.json entries (now passing, safe to remove):`);
    for (const r of stale) console.log(`  ${r.id}`);
  }

  const problems = records.filter(r => r.verdict === 'UNEXPLAINED' || r.verdict === 'MISMATCH');
  console.log('\nwrote results/elixcee-outcomes-classified.json');
  if (problems.length > 0) {
    console.log(`\n${problems.length} UNEXPLAINED/MISMATCH scenarios (gate failure):`);
    for (const r of problems) console.log(`  ${r.id}: ${r.reason}`);
    process.exitCode = 1;
    return;
  }

  console.log(`\nAll ${records.length} scenarios explained: 0 UNEXPLAINED, 0 MISMATCH.`);
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main();
}
