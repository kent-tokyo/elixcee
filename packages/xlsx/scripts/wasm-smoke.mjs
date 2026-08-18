// Confirms the WASM bridge (crates/elixcee-wasm) actually works from a consumer's
// perspective, beyond "does it compile" — run AFTER crates/elixcee-wasm/build.sh (which
// builds both wasm-pack targets and vendors their output into src/internal/wasm/; that
// build itself, and its two `wasm-pack build --target nodejs`/`--target web` invocations
// succeeding, is this script's own precondition, not something it re-does).
//
// Five checks, each printed as its own step so a CI failure points at exactly which one
// broke:
//   1. Node sync read() — require()'s the Node/CJS entry directly and reads a real .xlsx
//      fixture, confirming the freshly-built elixcee_wasm.node.cjs actually works
//      synchronously (no top-level await, no async init dance).
//   2. Browser export condition resolves — spawns a child `node --conditions=browser` and
//      self-imports this package by name (Node's own self-reference resolution, not a
//      symlink), confirming package.json's `exports.".".browser` condition really routes
//      to index.browser.mjs (the base64-inlined-WASM entry), not silently falling through
//      to the Node entry — then calls read() through it too, so this checks the browser
//      entry's WASM actually runs, not just that resolution succeeds. NOTE: this is Node
//      simulating the browser CONDITION, not a browser. Actual-browser coverage is a
//      separate script (scripts/browser-smoke.mjs), which launches a real Chrome process.
//   3./4. Minimal esbuild bundle + in-bundle read(), CJS output AND ESM output, run as two
//      distinct steps. Both used to be impossible: the Node/CJS WASM loader (wasm-pack's
//      own generated output) located its .wasm file via a `__dirname`-relative path, which
//      is bundle-OUTPUT-relative once bundled, so CJS output only worked if the consumer
//      manually copied elixcee_wasm_bg.wasm next to their bundle, and ESM output crashed
//      outright (`__dirname is not defined`). Both are fixed at the source: build.sh now
//      runs crates/elixcee-wasm/build-node-inline.mjs, which inlines the .wasm bytes as
//      base64 into the loader exactly as build-browser-inline.mjs already did for the
//      browser build. Neither step below copies any file next to the bundle — that step
//      being GONE is what these two checks now assert.
//   5. WASM artifact size — recorded, not gated. No prior-baseline file exists yet to
//      compare against, and asserting a threshold with no basis for the number would be
//      exactly the kind of unjustified gate this project avoids elsewhere (see
//      compat/vba-semantics/'s own anti-laundering discipline) — a policy to consider
//      adopting once a baseline exists, not applied here. Measured by decoding the
//      base64 payload out of the vendored loader, since the raw .wasm file is no longer
//      vendored (it would double-ship the same bytes; see build.sh).
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

for (const f of ['elixcee_wasm.node.cjs', 'elixcee_wasm.browser.mjs']) {
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

// Bundles a tiny consumer of the Node/CJS entry and RUNS the resulting bundle from a
// throwaway directory containing nothing but that one file — deliberately no .wasm copied
// next to it, and (since esbuild inlines the whole dependency graph) no node_modules to
// fall back on. If the loader ever reverts to a __dirname-relative filesystem lookup, both
// callers of this helper fail immediately with an ENOENT rather than passing silently.
function bundleAndRun(format, ext) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), `elixcee-wasm-smoke-${format}-`));
  const consumerPath = path.join(tmpDir, `consumer.${ext}`);
  const bundlePath = path.join(tmpDir, `bundle.${ext}`);
  const entry = JSON.stringify(path.join(PKG_DIR, 'src', 'index.cjs'));
  const fixture = JSON.stringify(FIXTURE);
  const body = [
    `const wb = XLSX.read(fs.readFileSync(${fixture}));`,
    `if (!Array.isArray(wb.SheetNames) || wb.SheetNames.length === 0) throw new Error('bundle read() returned no sheets');`,
    `console.log('  in-bundle SheetNames: ' + JSON.stringify(wb.SheetNames));`,
  ];
  fs.writeFileSync(
    consumerPath,
    (format === 'cjs'
      ? [`const XLSX = require(${entry});`, `const fs = require('node:fs');`]
      : [`import XLSX from ${entry};`, `import fs from 'node:fs';`]
    )
      .concat(body)
      .join('\n'),
  );
  esbuild.buildSync({ entryPoints: [consumerPath], bundle: true, platform: 'node', format, outfile: bundlePath });
  const bundleSrc = fs.readFileSync(bundlePath, 'utf8');
  // The regression this whole fix exists to prevent, asserted on the bundle TEXT rather
  // than only on "did it run": an ESM bundle carrying `__dirname` is broken even if some
  // other code path happens to make the process exit 0.
  if (format === 'esm' && /\b__dirname\b/.test(bundleSrc)) {
    throw new Error('ESM bundle contains __dirname — the WASM loader is doing a filesystem lookup again');
  }
  if (!bundleSrc.includes('ELIXCEE_WASM_BASE64')) {
    throw new Error('bundle does not contain the inlined WASM payload — build-node-inline.mjs did not run');
  }
  console.log(`  bundle: ${bundlePath} (${fs.statSync(bundlePath).size} bytes, format=${format})`);
  console.log(`  files next to the bundle: ${JSON.stringify(fs.readdirSync(tmpDir))} (no .wasm copied)`);
  execFileSync(process.execPath, [bundlePath], { stdio: 'inherit' });
}

step('3. esbuild CJS bundle + in-bundle read(), no .wasm copied next to it', () => bundleAndRun('cjs', 'cjs'));

step('4. esbuild ESM bundle + in-bundle read(), no .wasm copied next to it', () => bundleAndRun('esm', 'mjs'));

step('5. WASM artifact size (recorded, not gated)', () => {
  // Decoded from the vendored loader's own base64 constant: the raw .wasm file is no
  // longer vendored (build.sh stopped copying it once both loaders inlined their bytes),
  // so this measures the exact same payload from where it now actually lives.
  for (const f of ['elixcee_wasm.node.cjs', 'elixcee_wasm.browser.mjs']) {
    const src = fs.readFileSync(path.join(WASM_DIR, f), 'utf8');
    const m = src.match(/ELIXCEE_WASM_BASE64 = '([A-Za-z0-9+/=]+)'/);
    if (!m) throw new Error(`${f} has no inlined ELIXCEE_WASM_BASE64 payload`);
    const wasmSize = Buffer.from(m[1], 'base64').length;
    const fileSize = fs.statSync(path.join(WASM_DIR, f)).size;
    console.log(
      `  ${f}: ${fileSize} bytes on disk, carrying ${wasmSize} bytes of WASM ` +
        `(${(wasmSize / 1024).toFixed(1)} KiB) as ${m[1].length} base64 chars`,
    );
  }
});

console.log('\n[wasm-smoke] all checks passed.');
