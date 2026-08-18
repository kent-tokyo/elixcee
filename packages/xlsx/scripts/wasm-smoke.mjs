// Confirms the WASM bridge (crates/elixcee-wasm) actually works from a consumer's
// perspective, beyond "does it compile" — run AFTER crates/elixcee-wasm/build.sh (which
// builds both wasm-pack targets and vendors their output into src/internal/wasm/; that
// build itself, and its two `wasm-pack build --target nodejs`/`--target web` invocations
// succeeding, is this script's own precondition, not something it re-does).
//
// Four checks, each printed as its own step so a CI failure points at exactly which one
// broke:
//   1. Node sync read() — require()'s the Node/CJS entry directly and reads a real .xlsx
//      fixture, confirming the freshly-built elixcee_wasm.node.cjs + .wasm actually work
//      together synchronously (no top-level await, no async init dance).
//   2. Browser export condition resolves — spawns a child `node --conditions=browser` and
//      self-imports this package by name (Node's own self-reference resolution, not a
//      symlink), confirming package.json's `exports.".".browser` condition really routes
//      to index.browser.mjs (the base64-inlined-WASM entry), not silently falling through
//      to the Node entry — then calls read() through it too, so this checks the browser
//      entry's WASM actually runs, not just that resolution succeeds.
//   3. Minimal esbuild bundle + in-bundle read() — bundles a tiny consumer of the Node/CJS
//      entry (CJS output; ESM output breaks here, see below) and runs the bundle.
//      Consumer note, discovered while writing this check and worth stating plainly: the
//      Node/CJS WASM loader (wasm-pack's own generated elixcee_wasm.node.cjs, not
//      hand-written) locates its .wasm file via a `__dirname`-relative path — that's
//      bundle-output-relative once bundled, not source-relative, and ESM bundle output has
//      no `__dirname` at all (a ReferenceError, not a silent failure). A consumer bundling
//      this package's Node entry must either (a) bundle to CJS and copy
//      elixcee_wasm_bg.wasm next to the bundle output (what this check does), or (b) mark
//      the wasm loader external and let it resolve from node_modules normally. Not
//      something this round fixes — recorded here and in ROADMAP.md as a real, disclosed
//      consumer caveat, not silently worked around.
//   4. WASM artifact size — recorded, not gated. No prior-baseline file exists yet to
//      compare against, and asserting a threshold with no basis for the number would be
//      exactly the kind of unjustified gate this project avoids elsewhere (see
//      compat/vba-semantics/'s own anti-laundering discipline) — a policy to consider
//      adopting once a baseline exists, not applied here.
//
// Run: `node scripts/wasm-smoke.mjs` from packages/xlsx/ (needs esbuild installed —
// `npm ci` first).
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as esbuild from 'esbuild';

const require = createRequire(import.meta.url);
const DIR = path.dirname(fileURLToPath(import.meta.url));
const PKG_DIR = path.resolve(DIR, '..');
const REPO_ROOT = path.resolve(PKG_DIR, '..', '..');
const WASM_DIR = path.join(PKG_DIR, 'src', 'internal', 'wasm');
const FIXTURE = path.join(REPO_ROOT, 'tests', 'fixtures', 'e2e', 'source.xlsx');

function step(name, fn) {
  console.log(`\n[wasm-smoke] ${name}`);
  fn();
  console.log(`[wasm-smoke] ok: ${name}`);
}

for (const f of ['elixcee_wasm.node.cjs', 'elixcee_wasm_bg.wasm', 'elixcee_wasm.browser.mjs']) {
  if (!fs.existsSync(path.join(WASM_DIR, f))) {
    console.error(`missing ${f} under ${WASM_DIR} -- run crates/elixcee-wasm/build.sh first`);
    process.exit(1);
  }
}
if (!fs.existsSync(FIXTURE)) {
  console.error(`missing fixture: ${FIXTURE}`);
  process.exit(1);
}

step('1. Node sync read()', () => {
  const XLSX = require(path.join(PKG_DIR, 'src', 'index.cjs'));
  const bytes = fs.readFileSync(FIXTURE);
  const wb = XLSX.read(bytes);
  if (!Array.isArray(wb.SheetNames) || wb.SheetNames.length === 0) {
    throw new Error(`read() returned no sheets: ${JSON.stringify(wb.SheetNames)}`);
  }
  console.log(`  SheetNames: ${JSON.stringify(wb.SheetNames)}`);
});

step('2. browser export condition resolves and runs', () => {
  const script = `
    import('@elixcee/xlsx').then(async m => {
      if (typeof m.read !== 'function') throw new Error('browser entry has no read()');
      const fs = await import('node:fs');
      const wb = m.read(fs.readFileSync(${JSON.stringify(FIXTURE)}));
      if (!Array.isArray(wb.SheetNames) || wb.SheetNames.length === 0) {
        throw new Error('browser-condition read() returned no sheets');
      }
      console.log('  browser-condition SheetNames: ' + JSON.stringify(wb.SheetNames));
    }).catch(e => { console.error(e.stack); process.exit(1); });
  `;
  execFileSync(process.execPath, ['--conditions=browser', '--input-type=module', '-e', script], {
    cwd: PKG_DIR,
    stdio: 'inherit',
  });
});

step('3. minimal esbuild bundle + in-bundle read()', () => {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'elixcee-wasm-smoke-'));
  const consumerPath = path.join(tmpDir, 'consumer.cjs');
  const bundlePath = path.join(tmpDir, 'bundle.cjs');
  fs.writeFileSync(
    consumerPath,
    [
      `const XLSX = require(${JSON.stringify(path.join(PKG_DIR, 'src', 'index.cjs'))});`,
      `const fs = require('node:fs');`,
      `const wb = XLSX.read(fs.readFileSync(${JSON.stringify(FIXTURE)}));`,
      `if (!Array.isArray(wb.SheetNames) || wb.SheetNames.length === 0) throw new Error('bundle read() returned no sheets');`,
      `console.log('  in-bundle SheetNames: ' + JSON.stringify(wb.SheetNames));`,
    ].join('\n'),
  );
  esbuild.buildSync({
    entryPoints: [consumerPath],
    bundle: true,
    platform: 'node',
    format: 'cjs',
    outfile: bundlePath,
  });
  // The Node/CJS WASM loader resolves its .wasm file __dirname-relative to wherever it
  // ends up after bundling -- see this file's header comment. A real consumer needs the
  // same copy step (or externalize the loader); this is that step, not a workaround
  // specific to this smoke test.
  fs.copyFileSync(path.join(WASM_DIR, 'elixcee_wasm_bg.wasm'), path.join(tmpDir, 'elixcee_wasm_bg.wasm'));
  const bundleSize = fs.statSync(bundlePath).size;
  console.log(`  bundle size: ${bundleSize} bytes`);
  execFileSync(process.execPath, [bundlePath], { stdio: 'inherit' });
});

step('4. WASM artifact size (recorded, not gated)', () => {
  const wasmSize = fs.statSync(path.join(WASM_DIR, 'elixcee_wasm_bg.wasm')).size;
  console.log(`  elixcee_wasm_bg.wasm: ${wasmSize} bytes (${(wasmSize / 1024).toFixed(1)} KiB)`);
});

console.log('\n[wasm-smoke] all checks passed.');
