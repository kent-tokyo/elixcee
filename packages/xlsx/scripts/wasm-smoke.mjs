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
//   5. WASM artifact size — recorded and diffed against crates/elixcee-wasm/
//      wasm-size-baseline.json, but never gated: asserting a pass/fail threshold with no
//      basis for the number would be exactly the kind of unjustified gate this project
//      avoids elsewhere (see compat/vba-semantics/'s own anti-laundering discipline). The
//      baseline is updated by hand, deliberately, when a size change is intentional — not
//      auto-written by this script. Measured by decoding the base64 payload out of the
//      vendored loader, since the raw .wasm file is no longer vendored (it would
//      double-ship the same bytes; see build.sh). Also written to $GITHUB_STEP_SUMMARY
//      (when set) so the number is visible on the CI run without opening logs.
//   6. write()'s Node-builtin bundling posture — a DIFFERENT concern from 1-5 above.
//      write()/readFile()/readFileSync() reach a lazy `require('zlib')`/`require('fs')`
//      at call time (see src/internal/zip-writer.cjs's doc comment), which is fine for
//      Node and for an esbuild CJS bundle, but an esbuild ESM bundle can never
//      synchronously require() anything reached through CJS-origin code — confirmed here
//      both ways: inlining this package into an ESM bundle and then calling write() must
//      still throw (pinning the known esbuild limitation so a future esbuild/toolchain
//      change that silently "fixes" or re-breaks it doesn't go unnoticed), while marking
//      the package `external` (the documented, correct consumer pattern — see README.md's
//      "Bundling" section) must let write() run to completion in both CJS and ESM output.
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

step('5. WASM artifact size (recorded + diffed against baseline, not gated)', () => {
  // Decoded from the vendored loader's own base64 constant: the raw .wasm file is no
  // longer vendored (build.sh stopped copying it once both loaders inlined their bytes),
  // so this measures the exact same payload from where it now actually lives.
  const BASELINE_PATH = path.join(REPO_ROOT, 'crates', 'elixcee-wasm', 'wasm-size-baseline.json');
  const baseline = fs.existsSync(BASELINE_PATH) ? JSON.parse(fs.readFileSync(BASELINE_PATH, 'utf8')).wasmBytes : null;
  const summaryLines = ['| file | on-disk bytes | wasm bytes | vs. baseline |', '| --- | --- | --- | --- |'];
  let wasmSize;
  for (const f of ['elixcee_wasm.node.cjs', 'elixcee_wasm.browser.mjs']) {
    const src = fs.readFileSync(path.join(WASM_DIR, f), 'utf8');
    const m = src.match(/ELIXCEE_WASM_BASE64 = '([A-Za-z0-9+/=]+)'/);
    if (!m) throw new Error(`${f} has no inlined ELIXCEE_WASM_BASE64 payload`);
    wasmSize = Buffer.from(m[1], 'base64').length;
    const fileSize = fs.statSync(path.join(WASM_DIR, f)).size;
    let diffStr = 'no baseline';
    if (baseline != null) {
      const diff = wasmSize - baseline;
      const pct = ((diff / baseline) * 100).toFixed(2);
      diffStr = `${diff >= 0 ? '+' : ''}${diff} bytes (${diff >= 0 ? '+' : ''}${pct}%)`;
    }
    console.log(
      `  ${f}: ${fileSize} bytes on disk, carrying ${wasmSize} bytes of WASM ` +
        `(${(wasmSize / 1024).toFixed(1)} KiB) as ${m[1].length} base64 chars -- ${diffStr}`,
    );
    summaryLines.push(`| ${f} | ${fileSize} | ${wasmSize} | ${diffStr} |`);
  }
  if (process.env.GITHUB_STEP_SUMMARY) {
    fs.appendFileSync(process.env.GITHUB_STEP_SUMMARY, `\n### WASM artifact size\n\n${summaryLines.join('\n')}\n`);
  }
});

// Bundles a tiny consumer that actually CALLS write() (not just imports it), either with
// the package inlined into the bundle (an absolute-path import, same convention as
// bundleAndRun above) or marked `external` (a bare `@elixcee/xlsx` specifier, left for
// Node's own loader to resolve at run time) — and reports whether it ran.
function bundleAndRunWrite(format, ext, external) {
  // The externalized case needs Node's self-referencing-package resolution (see
  // https://nodejs.org/api/packages.html#self-referencing-a-package-using-its-name),
  // which only kicks in for a file inside the package's OWN directory tree — so that
  // bundle is written under PKG_DIR itself, not os.tmpdir(), and cleaned up after.
  const tmpDir = external
    ? fs.mkdtempSync(path.join(PKG_DIR, `.write-smoke-${format}-`))
    : fs.mkdtempSync(path.join(os.tmpdir(), `elixcee-write-smoke-${format}-`));
  const consumerPath = path.join(tmpDir, `consumer.${ext}`);
  const bundlePath = path.join(tmpDir, `bundle.${ext}`);
  const entry = JSON.stringify(path.join(PKG_DIR, 'src', 'index.cjs'));
  const body = [
    `const wb = XLSX.book_new();`,
    `XLSX.book_append_sheet(wb, XLSX.aoa_to_sheet([[1, 'x']]), 'S1');`,
    `const buf = XLSX.write(wb, { type: 'buffer' });`,
    `console.log('  write() produced ' + buf.length + ' bytes');`,
  ];
  const importLine = external
    ? format === 'cjs'
      ? `const XLSX = require('@elixcee/xlsx');`
      : `import * as XLSX from '@elixcee/xlsx';`
    : format === 'cjs'
      ? `const XLSX = require(${entry});`
      : `import XLSX from ${entry};`;
  fs.writeFileSync(consumerPath, [importLine].concat(body).join('\n'));
  try {
    esbuild.buildSync({
      entryPoints: [consumerPath],
      bundle: true,
      platform: 'node',
      format,
      outfile: bundlePath,
      ...(external ? { packages: 'external' } : {}),
    });
    runWriteBundle(bundlePath, format, external);
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}

function runWriteBundle(bundlePath, format, external) {
  if (external) {
    // The documented, correct consumer pattern — must always run to completion.
    execFileSync(process.execPath, [bundlePath], { cwd: PKG_DIR, stdio: 'inherit' });
    return;
  }
  // Inlined: CJS must run fine (real `require` exists at runtime); ESM must still throw
  // the known "Dynamic require" error (see this script's step-6 doc comment above) — a
  // pinned regression check, not a wish, so a toolchain change that alters this behavior
  // gets noticed rather than silently drifting.
  let threw = false;
  try {
    execFileSync(process.execPath, [bundlePath], { stdio: 'pipe' });
  } catch (e) {
    threw = true;
    if (!/Dynamic require/.test(String(e.stderr))) {
      throw new Error(`inlined ${format} bundle failed for an unexpected reason:\n${e.stderr}`);
    }
  }
  if (format === 'esm' && !threw) {
    throw new Error(
      'inlined ESM bundle calling write() unexpectedly SUCCEEDED — the known esbuild ' +
        '"Dynamic require" limitation may have been fixed upstream, or this package\'s ' +
        'require()-of-Node-builtins pattern changed; re-check README.md\'s Bundling ' +
        'section and this script\'s doc comment for step 6.',
    );
  }
  if (format === 'cjs' && threw) {
    throw new Error('inlined CJS bundle calling write() unexpectedly failed');
  }
  console.log(`  ${format} inlined bundle behaved as expected (${threw ? 'threw' : 'ran'})`);
}

step('6a. inlined ESM bundle + write() — must still throw (known esbuild limitation)', () =>
  bundleAndRunWrite('esm', 'mjs', false),
);
step('6b. inlined CJS bundle + write() — must run (CJS `require` works normally)', () =>
  bundleAndRunWrite('cjs', 'cjs', false),
);
step('6c. externalized ESM bundle + write() — must run (the documented consumer pattern)', () =>
  bundleAndRunWrite('esm', 'mjs', true),
);
step('6d. externalized CJS bundle + write() — must run', () => bundleAndRunWrite('cjs', 'cjs', true));

console.log('\n[wasm-smoke] all checks passed.');
