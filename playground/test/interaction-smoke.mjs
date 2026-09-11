import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
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

const driver = `
(() => {
  const selectedCount = () => document.querySelectorAll('td[role="gridcell"].selected').length;
  const cell = (ref) => document.querySelector(\`.cell-input[data-ref="\${ref}"]\`);
  const mouse = (target, type, init = {}) => target.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, button: 0, ...init }));
  const pointer = (target, type, init = {}) => target.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, pointerId: 7, pointerType: 'touch', ...init }));
  const run = async () => {
    try {
    const a1 = cell('A1'); const c3 = cell('C3');
    if (!a1 || !c3) throw new Error('sample cells were not rendered');
    mouse(a1, 'mousedown'); mouse(c3, 'mouseover'); mouse(c3, 'mouseup');
    const mouseRange = selectedCount();
    a1.focus(); a1.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', shiftKey: true, bubbles: true, cancelable: true }));
    const keyboardRange = selectedCount(); const keyboardNameBox = document.querySelector('#name-box')?.value;
    const pageCell = cell('A1'); pageCell.focus(); pageCell.dispatchEvent(new KeyboardEvent('keydown', { key: 'PageDown', bubbles: true, cancelable: true }));
    const pageDownNameBox = document.querySelector('#name-box')?.value;
    const resetScrollCell = cell('A1'); resetScrollCell.click();
    const wrap = document.querySelector('.sheet-wrap'); Object.defineProperty(wrap, 'clientHeight', { configurable: true, value: 480 }); Object.defineProperty(wrap, 'scrollHeight', { configurable: true, value: 960 }); Object.defineProperty(wrap, 'scrollTop', { configurable: true, writable: true, value: 960 }); wrap.dispatchEvent(new Event('scroll', { bubbles: true }));
    const wheelBoundaryNameBox = document.querySelector('#name-box')?.value;
    const touchA1 = cell('A1'); const touchC3 = cell('C3');
    pointer(touchA1, 'pointerdown'); pointer(touchC3, 'pointerover'); pointer(touchC3, 'pointerup');
    const touchRange = selectedCount();
    const tabTarget = cell('A1'); tabTarget.click(); tabTarget.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }));
    const tabNameBox = document.querySelector('#name-box')?.value;
    const enterTarget = cell('B1'); enterTarget.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));
    const enterNameBox = document.querySelector('#name-box')?.value;
    const homeTarget = cell('B2'); homeTarget.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', bubbles: true, cancelable: true }));
    const homeNameBox = document.querySelector('#name-box')?.value;
    const endTarget = cell('A2'); endTarget.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true, cancelable: true }));
    const endNameBox = document.querySelector('#name-box')?.value;
    const ctrlEndTarget = cell('A1'); ctrlEndTarget.click(); ctrlEndTarget.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', ctrlKey: true, bubbles: true, cancelable: true }));
    const ctrlEndNameBox = document.querySelector('#name-box')?.value;
    const ctrlHomeTarget = cell('C5'); ctrlHomeTarget.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', ctrlKey: true, bubbles: true, cancelable: true }));
    const ctrlHomeNameBox = document.querySelector('#name-box')?.value;
    const shiftTabTarget = cell('B1'); shiftTabTarget.click(); shiftTabTarget.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, cancelable: true }));
    const shiftTabNameBox = document.querySelector('#name-box')?.value;
    const shiftEnterTarget = cell('B2'); shiftEnterTarget.click(); shiftEnterTarget.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', shiftKey: true, bubbles: true, cancelable: true }));
    const shiftEnterNameBox = document.querySelector('#name-box')?.value;
    const copyTarget = cell('A1'); copyTarget.click(); document.querySelector('#copy')?.click(); await new Promise((resolve) => setTimeout(resolve, 20));
    const pasteTarget = cell('C1'); pasteTarget.click(); document.querySelector('#paste')?.click(); await new Promise((resolve) => setTimeout(resolve, 20));
    const pastedValue = cell('C1')?.value;
    const pivotCreate = document.querySelector('#create-pivot'); const pivotRefresh = document.querySelector('#refresh-pivot');
    const ribbonPivot = Boolean(pivotCreate?.classList.contains('ribbon-button') && pivotRefresh?.classList.contains('ribbon-button'));
    const pivotIcons = Boolean(pivotCreate?.querySelector('svg path[d*="M4 5h16v14"]') && pivotRefresh?.querySelector('svg path[d*="M4 5h16v14"]'));
    const result = { mouseRange, keyboardRange, touchRange, keyboardNameBox, pageDownNameBox, wheelBoundaryNameBox, tabNameBox, enterNameBox, homeNameBox, endNameBox, ctrlEndNameBox, ctrlHomeNameBox, shiftTabNameBox, shiftEnterNameBox, pastedValue, nameBox: document.querySelector('#name-box')?.value, ribbonPivot, pivotIcons };
      const node = document.createElement('pre'); node.id = 'interaction-result'; node.textContent = JSON.stringify(result); document.body.append(node);
    } catch (error) {
      const node = document.createElement('pre'); node.id = 'interaction-result'; node.textContent = JSON.stringify({ error: String(error) }); document.body.append(node);
    }
  };
  setTimeout(run, 1000);
})();
`;

const index = fs.readFileSync(path.join(root, 'index.html'), 'utf8').replace('</body>', '<script src="./interaction-driver.js"></script></body>');
const mime = { '.css': 'text/css; charset=utf-8', '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8' };
const server = http.createServer((request, response) => {
  const requestPath = request.url.split('?')[0];
  if (requestPath === '/interaction-driver.js') { response.writeHead(200, { 'content-type': mime['.js'] }); response.end(driver); return; }
  const relative = requestPath === '/' ? '/index.html' : requestPath;
  const file = path.resolve(root, `.${relative}`);
  if (!file.startsWith(`${root}${path.sep}`) || !fs.existsSync(file)) { response.writeHead(404).end(); return; }
  response.writeHead(200, { 'content-type': mime[path.extname(file)] || 'application/octet-stream' });
  response.end(relative === '/index.html' ? index : fs.readFileSync(file));
});
await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
const url = `http://127.0.0.1:${server.address().port}/`;
const output = await new Promise((resolve, reject) => {
  const child = spawn(chrome, [
    '--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--disable-background-networking',
    `--user-data-dir=${fs.mkdtempSync(path.join(os.tmpdir(), 'elixcee-interaction-chrome-'))}`, '--virtual-time-budget=3000', '--dump-dom', url,
  ], { stdio: ['ignore', 'pipe', 'ignore'] });
  let text = '';
  const timer = setTimeout(() => { child.kill('SIGKILL'); reject(new Error(`browser did not complete interaction smoke within 30s; output=${text.slice(-2000)}`)); }, 30_000);
  child.stdout.on('data', (chunk) => { text += chunk; if (text.includes('</html>') && text.includes('interaction-result')) { clearTimeout(timer); child.kill('SIGKILL'); resolve(text); } });
  child.on('error', (error) => { clearTimeout(timer); reject(error); });
});
server.close();

const encoded = output.match(/<pre id="interaction-result">([^<]+)<\/pre>/)?.[1];
assert.ok(encoded, 'interaction result was not written');
const result = JSON.parse(encoded);
assert.equal(result.mouseRange, 9, `mouse drag should select A1:C3: ${JSON.stringify(result)}`);
assert.equal(result.keyboardRange, 2, `Shift+Arrow should select A1:B1: ${JSON.stringify(result)}`);
assert.equal(result.touchRange, 9, `touch drag should select A1:C3: ${JSON.stringify(result)}`);
assert.equal(result.keyboardNameBox, 'A1:B1');
assert.equal(result.pageDownNameBox, 'A5', `PageDown should move by the virtual viewport height: ${JSON.stringify(result)}`);
assert.equal(result.wheelBoundaryNameBox, 'A5', `scrolling to the viewport boundary should advance the virtual range: ${JSON.stringify(result)}`);
assert.equal(result.tabNameBox, 'B1', `Tab should move right: ${JSON.stringify(result)}`);
assert.equal(result.enterNameBox, 'B2', `Enter should move down: ${JSON.stringify(result)}`);
assert.equal(result.homeNameBox, 'A2', `Home should move to the first column: ${JSON.stringify(result)}`);
assert.equal(result.endNameBox, 'C2', `End should move to the last used column: ${JSON.stringify(result)}`);
assert.equal(result.ctrlEndNameBox, 'C5', `Ctrl+End should move to the used range end: ${JSON.stringify(result)}`);
assert.equal(result.ctrlHomeNameBox, 'A1', `Ctrl+Home should move to the used range start: ${JSON.stringify(result)}`);
assert.equal(result.shiftTabNameBox, 'A1:B1', `Shift+Tab should extend backwards: ${JSON.stringify(result)}`);
assert.equal(result.shiftEnterNameBox, 'B1:B2', `Shift+Enter should extend upwards: ${JSON.stringify(result)}`);
assert.equal(result.pastedValue, 'Item', `internal clipboard paste should copy the selected cell: ${JSON.stringify(result)}`);
assert.equal(result.ribbonPivot, true, `Pivot actions should use ribbon-button decoration: ${JSON.stringify(result)}`);
assert.equal(result.pivotIcons, true, `Pivot actions should use the dedicated SVG icon: ${JSON.stringify(result)}`);
console.log(`playground interaction smoke passed (${result.mouseRange}/${result.keyboardRange}/${result.touchRange} selected cells, PageDown=${result.pageDownNameBox}, wheel=${result.wheelBoundaryNameBox}, pivot-icons=yes)`);
