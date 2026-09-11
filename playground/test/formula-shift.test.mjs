import assert from 'node:assert/strict';
import {
  shiftFormulaReferences,
  shiftQualifiedFormulaReferences,
  shiftWorkbookFormulaReferences,
} from '../src/formula-shift.mjs';

assert.equal(shiftFormulaReferences('=SUM(A1:$B$3)', 'r', 1, 1), '=SUM(A1:$B$4)');
assert.equal(shiftFormulaReferences('="A1"&A1', 'r', 0, 1), '="A1"&A2');
assert.equal(shiftFormulaReferences('=A2+B2', 'c', 1, -1), '=A2+#REF!');
assert.equal(shiftQualifiedFormulaReferences("='Sheet 1'!$A2+Other!A2", 'Sheet 1', 'r', 1, 1), "='Sheet 1'!$A3+Other!A2");

const workbook = {
  Sheets: {
    Summary: { A1: { f: "'Data'!A2+1" } },
    Data: { B3: { f: '=A2+"A2"' } },
  },
};
shiftWorkbookFormulaReferences(workbook, 'Data', 'r', 1, 1);
assert.equal(workbook.Sheets.Summary.A1.f, "'Data'!A3+1");
assert.equal(workbook.Sheets.Data.B3.f, '=A3+"A2"');

console.log('formula-shift tests passed');
