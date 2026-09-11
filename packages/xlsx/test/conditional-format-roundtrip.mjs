import assert from 'node:assert/strict';
import * as XLSX from '../src/index.mjs';

const workbook = {
  SheetNames: ['Sheet1'],
  Sheets: {
    Sheet1: {
      A1: { t: 'n', v: 15 },
      '!ref': 'A1:A1',
      '!conditionalFormats': [
        { type: 'cellIs', operator: 'greaterThan', priority: 1, stopIfTrue: true, sqref: ['A1'], formula: '10', dxf: { font: { bold: true, italic: true, underline: true, color: { rgb: 'FF9C0006' } }, fill: { fgColor: { rgb: 'FFFFF2CC' } } } },
        { type: 'cellIs', operator: 'lessThan', priority: 2, sqref: ['A1'], formula: '20', dxf: { font: { color: { rgb: 'FF0000FF' } } } },
      ],
    },
  },
};
const roundTrip = XLSX.read(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), { cellStyles: true });
const rules = roundTrip.Sheets.Sheet1['!conditionalFormats'];
assert.equal(rules.length, 2);
assert.equal(rules[0].priority, 1);
assert.equal(rules[0].stopIfTrue, true);
assert.equal(rules[0].dxf.font.bold, true);
assert.equal(rules[0].dxf.font.italic, true);
assert.equal(rules[0].dxf.font.underline, true);
assert.equal(rules[0].dxf.fill.fgColor.rgb, 'FFFFF2CC');
assert.equal(rules[1].priority, 2);
console.log('conditional-format priority/stopIfTrue round-trip smoke passed');
