// Joins results/elixcee-results.json with results/libreoffice-results*.json (one or more
// shard files — see run-libreoffice.mjs's parallel-shard note) on scenario id, classifies
// every scenario, and writes results/classify-results.json plus a printed summary table.
// Run: `node run-classify.mjs` from compat/corpus/, after both runners have produced
// their results files.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { classifyScenario, populateUnsupportedFromElixceeErrors, summarizeByCategoryAndVerdict, summarizeOverall, VERDICTS } from './classify.mjs';

const DIR = path.dirname(fileURLToPath(import.meta.url));
const resultsDir = path.join(DIR, 'results');

const elixceeResults = JSON.parse(fs.readFileSync(path.join(resultsDir, 'elixcee-results.json'), 'utf8'));
const elixceeById = new Map(elixceeResults.map((r) => [r.id, r]));

const oracleById = new Map();
for (const file of fs.readdirSync(resultsDir)) {
  if (!/^libreoffice-results.*\.json$/.test(file)) continue;
  const shard = JSON.parse(fs.readFileSync(path.join(resultsDir, file), 'utf8'));
  for (const r of shard) oracleById.set(r.id, r);
}

populateUnsupportedFromElixceeErrors(elixceeResults);

const scenarios = JSON.parse(fs.readFileSync(path.join(DIR, 'scenarios.json'), 'utf8'));
const classifications = [];
for (const scenario of scenarios) {
  const elixcee = elixceeById.get(scenario.id);
  const oracleResult = oracleById.get(scenario.id);
  if (!elixcee) continue; // not run against elixcee (shouldn't happen once run-elixcee.mjs covers the full corpus)
  classifications.push(
    classifyScenario({ id: scenario.id, category: scenario.category, oracle: 'libreoffice', elixcee, oracleResult })
  );
}

fs.writeFileSync(path.join(resultsDir, 'classify-results.json'), JSON.stringify(classifications, null, 2) + '\n');

const overall = summarizeOverall(classifications);
console.log(`classified ${classifications.length} scenarios (oracle: libreoffice)`);
for (const v of VERDICTS) {
  if (overall.has(v)) console.log(`  ${v}: ${overall.get(v)}`);
}

const byCategory = summarizeByCategoryAndVerdict(classifications);
console.log('\nby category:');
for (const [cat, verdicts] of byCategory) {
  const parts = VERDICTS.filter((v) => verdicts.has(v)).map((v) => `${v}=${verdicts.get(v)}`);
  console.log(`  ${cat}: ${parts.join(' ')}`);
}

console.log(`\nwrote results/classify-results.json (${classifications.length} records, oracle: "libreoffice")`);
