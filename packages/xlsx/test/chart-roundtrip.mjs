import assert from 'node:assert/strict';
import * as XLSX from '../src/index.mjs';
import { buildXlsxZipEntries } from '../src/internal/xlsx-writer.cjs';

const workbook = {
  SheetNames: ['Sheet1'],
  Sheets: {
    Sheet1: {
      '!ref': 'A1:C4',
      A1: { v: 'Category' }, B1: { v: 'One' }, C1: { v: 'Two' },
      A2: { v: 'A' }, B2: { v: 1 }, C2: { v: 2 },
      A3: { v: 'B' }, B3: { v: 3 }, C3: { v: 4 },
      '!charts': [
        { ref: 'A1:C3', type: 'bar', title: 'First', xAxisTitle: 'Month', yAxisTitle: 'Amount', colors: ['#12abef', '#ef7d12'] },
        { ref: 'A1:C3', type: 'line', title: 'Second', xAxisTitle: 'Period' },
        { ref: 'A1:C3', type: 'bar', title: 'No legend', legend: false },
      ],
    },
  },
};

const bytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
const chartEntry = buildXlsxZipEntries(workbook).find(({ name }) => name === 'xl/charts/chart1.xml');
assert.ok(chartEntry);
const chartXml = new TextDecoder().decode(chartEntry.data);
assert.match(chartXml, /<c:(?:strRef|numRef)>/);
assert.match(chartXml, /<c:f>'Sheet1'!\$A\$2:\$A\$3<\/c:f>/);
assert.doesNotMatch(chartXml, /<c:(?:strCache|numCache)>/);
const roundTrip = XLSX.read(bytes, { type: 'buffer' });
const charts = roundTrip.Sheets.Sheet1['!charts'];
assert.equal(charts.length, 3);
assert.deepEqual(
  charts.map(({ ref, type, title }) => ({ ref, type, title })),
  [
    { ref: 'A1:C3', type: 'bar', title: 'First' },
    { ref: 'A1:C3', type: 'line', title: 'Second' },
    { ref: 'A1:C3', type: 'bar', title: 'No legend' },
  ],
);
assert.ok(charts.every(({ widthCols, heightRows }) => widthCols > 0 && heightRows > 0));
assert.equal(charts[2].legend, false);
assert.equal(charts[0].xAxisTitle, 'Month');
assert.equal(charts[0].yAxisTitle, 'Amount');
assert.equal(charts[1].xAxisTitle, 'Period');
assert.equal(charts[1].yAxisTitle, undefined);
assert.deepEqual(charts[0].colors, ['12ABEF', 'EF7D12']);
console.log('chart round-trip passed');
