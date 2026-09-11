import assert from 'node:assert/strict';
import * as XLSX from '../src/index.mjs';
import { calculateWorkbook } from '../src/runtime.mjs';

const workbook = {
  SheetNames: ['Sheet1'],
  Sheets: {
    Sheet1: {
      '!ref': 'A1',
      A1: { f: 'SEQUENCE(1,3)', t: 'n' },
    },
  },
};

const calculated = JSON.parse(
  calculateWorkbook(XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })),
);
const sheet = calculated.Sheets.Sheet1;

assert.deepEqual(
  { A1: sheet.A1.v, B1: sheet.B1.v, C1: sheet.C1.v },
  { A1: 1, B1: 2, C1: 3 },
);
assert.equal(sheet['!ref'], 'A1:C1');
assert.equal(sheet.B1.f, undefined);
assert.equal(sheet.C1.f, undefined);

const rectangular = {
  SheetNames: ['Sheet1'],
  Sheets: { Sheet1: { '!ref': 'A1', A1: { f: 'SEQUENCE(2,2)', t: 'n' } } },
};
const rectangularSheet = JSON.parse(
  calculateWorkbook(XLSX.write(rectangular, { type: 'buffer', bookType: 'xlsx' })),
).Sheets.Sheet1;
assert.deepEqual(
  { A1: rectangularSheet.A1.v, B1: rectangularSheet.B1.v, A2: rectangularSheet.A2.v, B2: rectangularSheet.B2.v },
  { A1: 1, B1: 2, A2: 3, B2: 4 },
);
assert.equal(rectangularSheet['!ref'], 'A1:B2');

for (const [formula, expectedRef, expected] of [
  ['WRAPROWS(SEQUENCE(5),2)', 'A1:B3', { A1: 1, B1: 2, A2: 3, B2: 4, A3: 5, B3: '' }],
  ['WRAPCOLS(SEQUENCE(5),2)', 'A1:C2', { A1: 1, B1: 3, C1: 5, A2: 2, B2: 4, C2: '' }],
]) {
  const wrapped = { SheetNames: ['Sheet1'], Sheets: { Sheet1: { '!ref': 'A1', A1: { f: formula, t: 'n' } } } };
  const wrappedSheet = JSON.parse(
    calculateWorkbook(XLSX.write(wrapped, { type: 'buffer', bookType: 'xlsx' })),
  ).Sheets.Sheet1;
  assert.equal(wrappedSheet['!ref'], expectedRef);
  assert.deepEqual(Object.fromEntries(Object.keys(expected).map((ref) => [ref, wrappedSheet[ref]?.v])), expected);
}

const collisionWorkbook = {
  SheetNames: ['Sheet1'],
  Sheets: {
    Sheet1: {
      '!ref': 'A1:B1',
      A1: { f: 'SEQUENCE(1,3)', t: 'n' },
      B1: { v: 'keep', t: 's' },
    },
  },
};
const collision = JSON.parse(
  calculateWorkbook(XLSX.write(collisionWorkbook, { type: 'buffer', bookType: 'xlsx' })),
).Sheets.Sheet1;
assert.equal(collision.A1.v, '#SPILL!');
assert.equal(collision.B1.v, 'keep');
assert.equal(collision.C1, undefined);

console.log('formula spill projection passed');
