import assert from 'node:assert/strict';
import { execFileSync, spawn } from 'node:child_process';
import fs from 'node:fs';
import http from 'node:http';
import os from 'node:os';
import path from 'node:path';

const root = path.resolve(import.meta.dirname, '..', 'dist');
const candidates = process.env.CHROME_PATH
  ? [process.env.CHROME_PATH]
  : process.platform === 'darwin'
    ? ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/Applications/Chromium.app/Contents/MacOS/Chromium']
    : ['/usr/bin/google-chrome', '/usr/bin/chromium', '/usr/bin/chromium-browser'];
const chrome = candidates.find((candidate) => fs.existsSync(candidate));
if (!chrome) throw new Error(`no Chrome/Chromium executable found; tried ${candidates.join(', ')}`);

const mime = { '.css': 'text/css; charset=utf-8', '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8' };
const server = http.createServer((request, response) => {
  const requestPath = request.url.split('?')[0];
  const relative = requestPath === '/' ? '/index.html' : requestPath;
  const file = path.resolve(root, `.${relative}`);
  if (!file.startsWith(`${root}${path.sep}`) || !fs.existsSync(file)) { response.writeHead(404).end(); return; }
  response.writeHead(200, { 'content-type': mime[path.extname(file)] || 'application/octet-stream' });
  response.end(fs.readFileSync(file));
});
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const profile = process.env.PLAYGROUND_PROFILE || '';
const mobile = process.env.PLAYGROUND_VIEWPORT === 'mobile';
const url = `http://127.0.0.1:${server.address().port}/${profile ? `?profile=${encodeURIComponent(profile)}` : ''}`;
const output = await new Promise((resolve, reject) => {
  const child = spawn(chrome, [
    '--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
    '--disable-component-update', '--disable-background-networking',
    ...(mobile ? ['--window-size=500,844'] : []),
    `--user-data-dir=${fs.mkdtempSync(path.join(os.tmpdir(), 'elixcee-a11y-chrome-'))}`,
    '--dump-dom', url,
  ], { stdio: ['ignore', 'pipe', 'ignore'] });
  let text = '';
  const timer = setTimeout(() => { child.kill('SIGKILL'); reject(new Error('browser did not dump DOM within 30s')); }, 30_000);
  child.stdout.on('data', (chunk) => { text += chunk; if (text.includes('</html>')) { clearTimeout(timer); child.kill('SIGKILL'); resolve(text); } });
  child.on('error', (error) => { clearTimeout(timer); reject(error); });
});
server.close();

assert.match(output, /<table[^>]*role="grid"/);
assert.match(output, /http-equiv="Content-Security-Policy"[^>]*wasm-unsafe-eval/);
assert.doesNotMatch(output, /Content-Security-Policy[^>]*unsafe-inline/);
const ids = new Set([...output.matchAll(/\sid="([^"]+)"/g)].map((match) => match[1]));
const references = [...output.matchAll(/\sheaders="([^"]+)"/g)].flatMap((match) => match[1].split(/\s+/));
assert.ok(references.length > 0, 'grid cells should reference row and column headers');
assert.ok(references.every((reference) => ids.has(reference)), 'every cell header reference must resolve to a DOM id');
assert.equal([...output.matchAll(/class="[^"]*cell-input[^"]*"[^>]*tabindex="0"/g)].length, 1, 'exactly one cell should be roving-tabindex focusable');
assert.match(output, /aria-rowcount="\d+"[^>]*aria-colcount="\d+"/);
if (profile === 'large-sparse') {
  assert.match(output, /data-playground-profile="large-sparse"/);
  const rowCounts = [...output.matchAll(/aria-rowcount="(\d+)"/g)].map((match) => Number(match[1]));
  const rowCount = Math.max(...rowCounts);
  assert.ok(rowCount >= 200000, `large sparse profile should expose its full logical row count: ${rowCount}`);
  const renderedCells = [...output.matchAll(/class="[^"]*cell-input[^"]*"/g)].length;
  assert.ok(renderedCells <= 240, `virtual viewport rendered too many cells: ${renderedCells}`);
  const renderMs = Number(output.match(/data-initial-render-ms="(\d+)"/)?.[1]);
  assert.ok(Number.isFinite(renderMs), 'large sparse profile should record initial render time');
  console.log(`large sparse logical rows=${rowCount}, rendered cells=${renderedCells}, initial render=${renderMs}ms`);
}
if (profile === 'scroll-soak') {
  assert.match(output, /data-playground-profile="scroll-soak"/);
  const steps = Number(output.match(/data-virtual-scroll-soak-steps="(\d+)"/)?.[1]);
  const soakMs = Number(output.match(/data-virtual-scroll-soak-ms="(\d+)"/)?.[1]);
  const renderedCells = [...output.matchAll(/class="[^\"]*cell-input[^\"]*"/g)].length;
  assert.equal(steps, 100, 'virtual scroll soak should execute all steps');
  assert.ok(Number.isFinite(soakMs), 'virtual scroll soak should record duration');
  assert.ok(soakMs < 15_000, `virtual scroll soak exceeded 15s: ${soakMs}ms`);
  assert.ok(renderedCells <= 240, `virtual scroll soak rendered too many cells: ${renderedCells}`);
  console.log(`virtual scroll soak steps=${steps}, duration=${soakMs}ms, rendered cells=${renderedCells}`);
}
if (mobile) {
  const viewportWidth = Number(output.match(/data-viewport-width="(\d+)"/)?.[1]);
  assert.ok(Number.isFinite(viewportWidth) && viewportWidth <= 500, `narrow viewport was not applied: ${viewportWidth}`);
  assert.ok([...output.matchAll(/class="[^\"]*cell-input[^\"]*"/g)].length <= 240, 'mobile profile rendered too many cells');
  console.log(`mobile viewport width=${viewportWidth}, rendered cells=${[...output.matchAll(/class="[^\"]*cell-input[^\"]*"/g)].length}`);
}
console.log(`playground accessibility smoke passed (${execFileSync(chrome, ['--version'], { encoding: 'utf8' }).trim()})`);
