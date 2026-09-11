import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const source = fs.readFileSync(fileURLToPath(new URL('../src/vba-worker.js', import.meta.url)), 'utf8');
async function run(data) {
  let message;
  const context = { self: { postMessage(value) { message = value; } } };
  vm.runInNewContext(source, context, { filename: 'vba-worker.js' });
  context.self.onmessage({ data });
  return message;
}

const ok = await run({
  source: 'Sub Demo()\n Cells(2, 2).Value = 4 + 5\nEnd Sub',
  macroName: 'demo',
  cells: {},
});
assert.deepEqual(JSON.parse(JSON.stringify(ok)), { ok: true, changes: [{ ref: 'B2', value: 9 }], statements: 1 });

for (const sourceText of [
  'Sub Demo()\n Cells(2, 2).Value = 9',
  'Sub Demo()\n Cells(2, 2).Value = 9\nEnd Sub\nCells(3, 3).Value = 1',
  'Sub Demo()\nEnd Sub\nEnd Sub',
]) {
  const result = await run({ source: sourceText, macroName: 'Demo', cells: {} });
  assert.equal(result.ok, false, `invalid Sub boundary should be rejected: ${sourceText}`);
}

const tooMany = await run({
  source: `Sub Demo()\n${Array.from({ length: 101 }, () => 'x = 1').join('\n')}\nEnd Sub`,
  macroName: 'Demo',
  cells: {},
});
assert.equal(tooMany.ok, false);
assert.match(tooMany.error, /statement limit/);

const tooManyCells = await run({
  source: 'Sub Demo()\nEnd Sub',
  macroName: 'Demo',
  cells: Object.fromEntries(Array.from({ length: 10001 }, (_, index) => [`A${index + 1}`, index])),
});
assert.equal(tooManyCells.ok, false);
assert.match(tooManyCells.error, /input cell limit/);

const invalidCell = await run({ source: 'Sub Demo()\nEnd Sub', macroName: 'Demo', cells: { 'not-a-cell': 'x' } });
assert.equal(invalidCell.ok, false);
assert.match(invalidCell.error, /invalid macro input cell/);
console.log('vba worker boundary tests passed');
