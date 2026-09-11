import assert from 'node:assert/strict';
import * as XLSX from '../src/index.mjs';
import { buildXlsxZipEntries } from '../src/internal/xlsx-writer.cjs';

const workbook = {
  SheetNames: ['Data', 'Pivot'],
  Sheets: {
    Data: {
      '!ref': 'A1:B4',
      A1: { v: 'Category' }, B1: { v: 'Amount' },
      A2: { v: 'A' }, B2: { v: 2 },
      A3: { v: 'A' }, B3: { v: 3 },
      A4: { v: 'B' }, B4: { v: 7 },
    },
    Pivot: {
      '!ref': 'A1:B3',
      A1: { v: 'Category' }, B1: { v: 'Sum of Amount' },
      A2: { v: 'A' }, B2: { v: 5 },
      A3: { v: 'B' }, B3: { v: 7 },
      // Playground-only metadata must not turn this limited summary into a
      // misleading native OOXML PivotTable claim.
      '!pivot': { sourceSheet: 'Data', sourceRef: 'A1:B4', aggregate: 'sum', filter: '' },
    },
  },
};

const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
const entries = buildXlsxZipEntries(workbook);
const names = entries.map(({ name }) => name);
assert.ok(names.includes('xl/worksheets/sheet2.xml'));
assert.ok(!names.some((name) => name.startsWith('xl/pivotCache/')));
assert.ok(!names.some((name) => name.startsWith('xl/pivotTables/')));
const sheetXml = new TextDecoder().decode(entries.find(({ name }) => name === 'xl/worksheets/sheet2.xml').data);
assert.match(sheetXml, />Sum of Amount<\/t>/);
assert.match(sheetXml, /<v>5<\/v>/);

const roundTrip = XLSX.read(bytes, { type: 'buffer' });
assert.deepEqual(
  ['A1', 'B1', 'A2', 'B2', 'A3', 'B3'].map((ref) => roundTrip.Sheets.Pivot[ref]?.v),
  ['Category', 'Sum of Amount', 'A', 5, 'B', 7],
);
console.log('pivot round-trip passed (limited summary worksheet; no native PivotTable parts)');
