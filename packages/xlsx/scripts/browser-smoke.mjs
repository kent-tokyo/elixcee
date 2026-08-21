// Runs @elixcee/xlsx in an ACTUAL browser process. This is a different, stronger claim
// than the one scripts/wasm-smoke.mjs makes, and the distinction is deliberate:
//
//   - wasm-smoke.mjs step 2 spawns `node --conditions=browser`. That proves the "browser"
//     export CONDITION resolves to index.browser.mjs and that the base64-inlined WASM in it
//     initializes and reads a workbook. It is Node, simulating one aspect of a browser.
//     No browser is involved and none of a browser's actual constraints apply.
//   - This script bundles that same browser entry with esbuild, serves it over real HTTP
//     from Node's own `http` module, and launches a real Chrome/Chromium process that
//     navigates to it and executes it. Everything asserted below comes back out of that
//     browser process.
//
// Driving the browser with zero new dependencies: Chrome's own `--headless=new --dump-dom`
// prints the serialized DOM after load, which is all this check needs — the page writes its
// result into the DOM and this script reads it back. Deliberately chosen over the
// alternatives the spec listed (playwright-core, puppeteer-core, chrome-launcher +
// chrome-remote-interface): each is a real dependency (megabytes, plus its own Chrome-
// version compatibility surface) bought to do something a documented Chrome CLI flag
// already does for a page whose entire workload is synchronous. If this check ever needs
// interaction, multiple navigations, or CDP-level introspection, a driver becomes worth its
// weight — for "load one page, read one result", it is not.
//
// What is and is not asserted about errors, stated precisely: the page installs
// window.onerror / unhandledrejection / console.error hooks BEFORE the bundle script runs,
// so it captures page-observable errors; and this script's own static server records every
// request it served, so a 404 the page swallowed still fails the check. Errors reported
// only to Chrome's internal log (e.g. some network-stack or policy warnings) are outside
// both mechanisms and are not claimed.
//
// Chrome discovery: $CHROME_PATH first, then a per-platform candidate list. On failure
// every path tried is printed. If no browser is found this exits non-zero rather than
// skipping — a silent skip is exactly how "we have browser coverage" becomes untrue.
//
// Safari is not tested and is not claimed to be supported anywhere in this project.
//
// Run: `node scripts/browser-smoke.mjs` from packages/xlsx/ (needs `npm ci` first, for
// esbuild).
import { execFileSync, spawn } from 'node:child_process';
import fs from 'node:fs';
import http from 'node:http';
import { createRequire } from 'node:module';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as esbuild from 'esbuild';

const DIR = path.dirname(fileURLToPath(import.meta.url));
const PKG_DIR = path.resolve(DIR, '..');
const REPO_ROOT = path.resolve(PKG_DIR, '..', '..');
const FIXTURE = path.join(REPO_ROOT, 'tests', 'fixtures', 'e2e', 'source.xlsx');

const CHROME_CANDIDATES = {
  darwin: [
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
  ],
  linux: [
    '/usr/bin/google-chrome',
    '/usr/bin/google-chrome-stable',
    '/usr/bin/chromium-browser',
    '/usr/bin/chromium',
    '/opt/google/chrome/chrome',
  ],
  win32: [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  ],
};

function step(name) {
  console.log(`\n[browser-smoke] ${name}`);
}

function findChrome() {
  const tried = [];
  if (process.env.CHROME_PATH) {
    tried.push(`$CHROME_PATH=${process.env.CHROME_PATH}`);
    if (fs.existsSync(process.env.CHROME_PATH)) return process.env.CHROME_PATH;
  }
  for (const c of CHROME_CANDIDATES[process.platform] || []) {
    tried.push(c);
    if (fs.existsSync(c)) return c;
  }
  console.error(
    `[browser-smoke] FAIL: no Chrome/Chromium executable found on platform "${process.platform}".\n` +
      `paths tried:\n${tried.map((t) => `  ${t}`).join('\n')}\n` +
      `Set $CHROME_PATH to a Chrome/Chromium binary. Not skipping: a skipped browser check ` +
      `would leave this project claiming browser coverage it does not have.`
  );
  process.exit(1);
}

// The page's whole workload is synchronous at module-evaluation time (the inlined WASM
// compiles via initSync, then one read()), so by the load event that --dump-dom waits for,
// __RESULT__ is already populated. Nothing here polls or races.
function buildPage(bundleJs) {
  return `<!doctype html>
<html><head><meta charset="utf-8"><title>elixcee browser smoke</title></head>
<body>
<pre id="result">PENDING</pre>
<script>
  // Installed BEFORE the module bundle below, so an error thrown while the bundle
  // evaluates is captured rather than merely crashing the page silently.
  window.__ERRORS__ = [];
  window.addEventListener('error', function (e) {
    window.__ERRORS__.push('error: ' + (e.message || String(e.error)));
  });
  window.addEventListener('unhandledrejection', function (e) {
    window.__ERRORS__.push('unhandledrejection: ' + String(e.reason));
  });
  var realConsoleError = console.error.bind(console);
  console.error = function () {
    window.__ERRORS__.push('console.error: ' + Array.prototype.join.call(arguments, ' '));
    realConsoleError.apply(null, arguments);
  };
  window.__REPORT__ = function (obj) {
    obj.errors = window.__ERRORS__;
    // base64, not raw JSON: --dump-dom serializes text content with HTML escaping, and a
    // base64 payload has no character this script then has to un-escape correctly.
    document.getElementById('result').textContent = btoa(JSON.stringify(obj));
  };
</script>
<script type="module" src="${bundleJs}"></script>
</body></html>
`;
}

const chrome = findChrome();
step('1. locate a real browser executable');
console.log(`  executable: ${chrome}`);
console.log(`  version: ${execFileSync(chrome, ['--version'], { encoding: 'utf8' }).trim()}`);

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'elixcee-browser-smoke-'));
const siteDir = path.join(tmpDir, 'site');
fs.mkdirSync(siteDir);

step('2. bundle the browser entry with esbuild');
{
  // A node_modules symlink so esbuild resolves the bare specifier "@elixcee/xlsx" exactly
  // as a consumer's bundler would — through package.json's `exports` map with the
  // "browser" condition — rather than through a path that bypasses condition resolution.
  const nm = path.join(tmpDir, 'node_modules', '@elixcee');
  fs.mkdirSync(nm, { recursive: true });
  fs.symlinkSync(PKG_DIR, path.join(nm, 'xlsx'), 'dir');

  const fixtureB64 = fs.readFileSync(FIXTURE).toString('base64');
  const entry = path.join(tmpDir, 'entry.mjs');
  fs.writeFileSync(
    entry,
    [
      `import * as XLSX from '@elixcee/xlsx';`,
      `const wb = XLSX.read(${JSON.stringify(fixtureB64)}, { type: 'base64' });`,
      `const ws = wb.Sheets[wb.SheetNames[0]];`,
      // write() is pure XML/ZIP generation, no filesystem needed — this exercises it in a
      // REAL browser, not just a bundle build (esbuild's platform:'browser' refusing to
      // resolve 'zlib' was a real bug this same script's build step already guards
      // against below; this proves the resulting bundle also WORKS, not just builds).
      `const wbOut = XLSX.book_new();`,
      `XLSX.book_append_sheet(wbOut, XLSX.aoa_to_sheet([[1, 'two', true]]), 'Written');`,
      `const written = XLSX.write(wbOut, { type: 'buffer' });`,
      `const readBack = XLSX.read(written);`,
      `window.__REPORT__({`,
      `  ok: true,`,
      `  sheetNames: wb.SheetNames,`,
      `  ref: ws['!ref'],`,
      `  a1: ws.A1 && ws.A1.v,`,
      `  b2: ws.B2 && ws.B2.v,`,
      `  csvHead: XLSX.sheet_to_csv(ws).split('\\n')[0],`,
      `  exportCount: Object.keys(XLSX).filter(k => k !== 'default').length,`,
      `  writeRoundTripSheetNames: readBack.SheetNames,`,
      `  writeRoundTripBytes: written.length,`,
      `});`,
    ].join('\n')
  );

  esbuild.buildSync({
    entryPoints: [entry],
    bundle: true,
    platform: 'browser',
    format: 'esm',
    conditions: ['browser'],
    outfile: path.join(siteDir, 'bundle.js'),
  });

  const bundleSrc = fs.readFileSync(path.join(siteDir, 'bundle.js'), 'utf8');
  // Without these assertions this check could "pass" while having bundled the NODE entry —
  // the failure mode where the browser condition silently falls through and the test proves
  // nothing about the browser build at all.
  if (!bundleSrc.includes('initSync')) {
    throw new Error('bundle has no initSync — the browser WASM entry was not the one bundled');
  }
  // The Node WASM loader's own signature line. Deliberately NOT a bare search for
  // "readFileSync": index.cjs's readFile()/readFileSync() (Node-only, stubbed out in a
  // browser build via package.json's `browser` field mapping "fs" to false) legitimately
  // put that identifier in the bundle as unreachable code, so a bare search reports a leak
  // that isn't one. This matches the loader and nothing else.
  if (bundleSrc.includes('Buffer.from(ELIXCEE_WASM_BASE64')) {
    throw new Error('bundle contains the Node WASM loader — the browser condition fell through');
  }
  // ...and the payload must appear exactly ONCE. Both loaders inline the same 263KB of
  // base64, so a browser bundle carrying two copies means the Node loader came along for
  // the ride even if its init line was tree-shaken.
  const payloadCount = (bundleSrc.match(/ELIXCEE_WASM_BASE64 = /g) || []).length;
  if (payloadCount !== 1) {
    throw new Error(`bundle carries the WASM payload ${payloadCount} times, expected exactly 1`);
  }
  if (/require\(["']fs["']\)|from\s*["']fs["']/.test(bundleSrc)) {
    throw new Error('bundle still requires "fs" — a browser build must never reach the filesystem');
  }
  // The real bug this guards against: esbuild's `platform: 'browser'` used to refuse to
  // even PRODUCE a bundle containing write()'s Node-only `require('zlib')` call, since it
  // was reachable (dead code, but textually present) via index.browser.mjs's re-export of
  // index.cjs's other utils — confirmed live, fixed by giving the browser build its own
  // zlib-free write() and stubbing deflate-node.cjs via package.json's `browser` field
  // (see index.browser.mjs's and deflate-node.cjs's own doc comments). If this bundle
  // succeeded at all with `require("zlib")` still present in it, something upstream
  // (esbuild's own behavior, or the browser-field stub) changed.
  if (/require\(["']zlib["']\)|from\s*["']zlib["']/.test(bundleSrc)) {
    throw new Error('bundle contains a "zlib" reference — write() must not reach zlib in a browser build');
  }
  console.log(`  bundle: ${fs.statSync(path.join(siteDir, 'bundle.js')).size} bytes (browser condition, esm)`);
}

fs.writeFileSync(path.join(siteDir, 'index.html'), buildPage('./bundle.js'));

step('3. serve it over real HTTP from node:http');
const served = [];
const MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8' };
const server = http.createServer((req, res) => {
  const urlPath = req.url === '/' ? '/index.html' : req.url.split('?')[0];
  const file = path.join(siteDir, path.normalize(urlPath).replace(/^(\.\.[/\\])+/, ''));
  let status = 200;
  if (!file.startsWith(siteDir) || !fs.existsSync(file)) {
    status = 404;
    res.writeHead(404).end('not found');
  } else {
    res.writeHead(200, { 'content-type': MIME[path.extname(file)] || 'application/octet-stream' });
    res.end(fs.readFileSync(file));
  }
  served.push({ urlPath, status });
});
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const url = `http://127.0.0.1:${server.address().port}/`;
console.log(`  serving ${siteDir} at ${url}`);

step('4. launch the real browser and read the result back out of it');
const dom = await new Promise((resolve, reject) => {
  const args = [
    '--headless=new',
    '--disable-gpu',
    '--no-first-run',
    '--no-default-browser-check',
    // Without these, Chrome's own updater/background children inherit the stdio pipes and
    // this process never sees EOF even after the DOM has been dumped (observed on macOS).
    '--disable-component-update',
    '--disable-background-networking',
    `--user-data-dir=${path.join(tmpDir, 'chrome-profile')}`,
    '--dump-dom',
    url,
  ];
  if (process.platform === 'linux') args.splice(1, 0, '--no-sandbox', '--disable-dev-shm-usage');
  console.log(`  ${chrome} ${args.filter((a) => a !== url).join(' ')} <url>`);

  const child = spawn(chrome, args, { stdio: ['ignore', 'pipe', 'ignore'] });
  let out = '';
  let done = false;
  // Resolve on the dump being COMPLETE ("</html>"), not on process exit: Chrome's own
  // background children (updater, crashpad) inherit the stdout pipe and can keep it open
  // long after the DOM has been written, so waiting for EOF hangs until the timeout even
  // on a fully successful run (observed on macOS). Chrome's exit code is unusable for the
  // same reason, so success is judged purely by whether a complete DOM came back.
  const settle = (ok, value) => {
    if (done) return;
    done = true;
    clearTimeout(timer);
    child.kill('SIGKILL');
    ok ? resolve(value) : reject(value);
  };
  const timer = setTimeout(
    () => settle(false, new Error(`browser did not produce a complete DOM within 60s; captured:\n${out.slice(0, 2000)}`)),
    60_000
  );
  child.stdout.on('data', (d) => {
    out += d;
    if (out.includes('</html>')) settle(true, out);
  });
  child.on('error', (e) => settle(false, e));
  child.on('close', () =>
    settle(false, new Error(`browser exited without dumping a complete DOM; stdout was:\n${out.slice(0, 2000)}`))
  );
});
server.close();

step('5. verify what the browser actually produced');
const m = dom.match(/<pre id="result">([^<]*)<\/pre>/);
if (!m) throw new Error('no #result element in the dumped DOM');
if (m[1] === 'PENDING') throw new Error('#result still PENDING — the bundle never ran to completion in the browser');
const result = JSON.parse(Buffer.from(m[1], 'base64').toString('utf8'));
console.log(`  ${JSON.stringify(result)}`);

const problems = [];
if (result.ok !== true) problems.push('page did not report ok');
if (JSON.stringify(result.sheetNames) !== JSON.stringify(['source'])) {
  problems.push(`unexpected SheetNames: ${JSON.stringify(result.sheetNames)}`);
}
if (result.ref !== 'A1:D9') problems.push(`unexpected !ref: ${result.ref}`);
if (result.a1 !== 'Name') problems.push(`unexpected A1 value: ${JSON.stringify(result.a1)}`);
if (result.b2 !== 42) problems.push(`unexpected B2 value: ${JSON.stringify(result.b2)}`);
if (result.csvHead !== 'Name,Amount,Active,Note') problems.push(`unexpected CSV header: ${result.csvHead}`);
if (JSON.stringify(result.writeRoundTripSheetNames) !== JSON.stringify(['Written'])) {
  problems.push(`write()->read() round trip failed in the browser: ${JSON.stringify(result.writeRoundTripSheetNames)}`);
}
if (!(result.writeRoundTripBytes > 0)) problems.push(`write() produced no bytes in the browser: ${result.writeRoundTripBytes}`);
// Compared against the Node entry's own live export count rather than a hard-coded number,
// so adding an export doesn't require editing this file — but a browser entry that FORGETS
// one still fails here.
const nodeExportCount = Object.keys(createRequire(import.meta.url)(path.join(PKG_DIR, 'src', 'index.cjs'))).length;
if (result.exportCount !== nodeExportCount) {
  problems.push(`browser entry exported ${result.exportCount} names, Node entry exports ${nodeExportCount}`);
}
if (result.errors.length !== 0) problems.push(`page-observable errors: ${JSON.stringify(result.errors)}`);

console.log(`  requests served: ${JSON.stringify(served)}`);
// /favicon.ico is requested by the browser itself, not by anything the page references, and
// this fixture site has none — its 404 is browser housekeeping, not a missing page resource.
// Every OTHER request must have been served 200: that is what catches a bundle or asset the
// page asked for and silently didn't get.
const bad = served.filter((r) => r.status !== 200 && r.urlPath !== '/favicon.ico');
if (bad.length) problems.push(`non-200 responses: ${JSON.stringify(bad)}`);
if (!served.some((r) => r.urlPath === '/bundle.js')) problems.push('the browser never requested /bundle.js');

if (problems.length) {
  console.error(`\n[browser-smoke] FAIL:\n${problems.map((p) => `  - ${p}`).join('\n')}`);
  process.exit(1);
}

console.log(
  `\n[browser-smoke] all checks passed in a real browser process (${execFileSync(chrome, ['--version'], {
    encoding: 'utf8',
  }).trim()}).`
);
