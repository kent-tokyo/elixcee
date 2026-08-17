// Drives scenarios.json against the real elixcee CLI binary (built from this repo's own
// Rust source, not modified — see docs/agent-contract.md for the --json contract this
// relies on). Writes results/elixcee-results.json: one record per scenario, in the exact
// `cells` shape elixcee's own --json mode emits, so the normalizer has minimal work to do
// on this side.
//
// Build first: `cargo build --release --bin elixcee` from the repo root.
// Run: `node run-elixcee.mjs [path-to-elixcee-binary]` from compat/corpus/.
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { execFileSync } from 'node:child_process';

const DIR = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(DIR, '..', '..');
const ELIXCEE_BIN = process.argv[2] || path.join(REPO_ROOT, 'target', 'release', 'elixcee');

if (!fs.existsSync(ELIXCEE_BIN)) {
  console.error(`elixcee binary not found at ${ELIXCEE_BIN}. Build it first:`);
  console.error('  cargo build --release --bin elixcee');
  process.exit(1);
}

const scenarios = JSON.parse(fs.readFileSync(path.join(DIR, 'scenarios.json'), 'utf8'));
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'elixcee-corpus-'));

const results = [];
let ok = 0;
let failed = 0;

for (const scenario of scenarios) {
  const vbaPath = path.join(tmpDir, `${scenario.id}.vba`);
  fs.writeFileSync(vbaPath, scenario.vbaSource);

  const args = [vbaPath, scenario.entrypoint, '--json'];
  if (scenario.workbook) {
    args.push('--file', path.join(DIR, 'workbooks', `${scenario.workbook}.xlsx`));
  }

  let record;
  try {
    const stdout = execFileSync(ELIXCEE_BIN, args, { encoding: 'utf8', timeout: 10_000 });
    const parsed = JSON.parse(stdout.trim());
    record = { id: scenario.id, category: scenario.category, ...parsed };
    if (parsed.ok) ok++;
    else failed++;
  } catch (err) {
    // execFileSync throws on non-zero exit (elixcee's documented "any failure" exit 1) —
    // its stdout still carries the JSON error shape per the CLI contract.
    const stdout = err.stdout ? err.stdout.toString().trim() : '';
    try {
      const parsed = JSON.parse(stdout);
      record = { id: scenario.id, category: scenario.category, ...parsed };
      failed++;
    } catch {
      record = {
        id: scenario.id,
        category: scenario.category,
        ok: false,
        error: { code: 'HARNESS_SPAWN_FAILURE', kind: 'harness', message: String(err.message || err) },
      };
      failed++;
    }
  }
  results.push(record);
}

fs.rmSync(tmpDir, { recursive: true, force: true });

const outDir = path.join(DIR, 'results');
fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, 'elixcee-results.json'), JSON.stringify(results, null, 2) + '\n');

console.log(`ran ${results.length} scenarios against elixcee: ${ok} ok, ${failed} failed`);
console.log(`wrote results/elixcee-results.json`);
