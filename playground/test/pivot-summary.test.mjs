import assert from 'node:assert/strict';
import { adjustPivotSourceRef, buildPivotRows } from '../src/pivot-summary.mjs';

const sheet = {
  A1: { v: 'Category' }, B1: { v: 'Amount' },
  A2: { v: 'A' }, B2: { v: 2 },
  A3: { v: 'A' }, B3: { v: 3 },
  A4: { v: 'B' }, B4: { v: 7 },
};
const range = { s: { r: 0, c: 0 }, e: { r: 3, c: 1 } };
assert.deepEqual(buildPivotRows(sheet, range, 'sum', ''), [
  ['Category', 'Sum of Amount'], ['A', 5], ['B', 7],
]);
assert.deepEqual(buildPivotRows(sheet, range, 'average', 'a'), [
  ['Category', 'Average of Amount'], ['A', 2.5],
]);
assert.equal(adjustPivotSourceRef('A1:B4', 'r', { r: 1, c: 0 }, 1), 'A1:B5');
assert.equal(adjustPivotSourceRef('A1:B4', 'c', { r: 0, c: 0 }, 1), 'B1:C4');
assert.equal(adjustPivotSourceRef('A1:B4', 'r', { r: 0, c: 0 }, -1), undefined, 'deleting the header is rejected');
console.log('pivot summary passed');
