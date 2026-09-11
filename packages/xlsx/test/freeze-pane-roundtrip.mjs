import assert from 'node:assert/strict';
import { inflateRawSync } from 'node:zlib';
import * as XLSX from '../src/index.mjs';

function worksheetXml(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;
  while (offset + 30 <= bytes.length && view.getUint32(offset, true) === 0x04034b50) {
    const method = view.getUint16(offset + 8, true);
    const compressedSize = view.getUint32(offset + 18, true);
    const nameSize = view.getUint16(offset + 26, true);
    const extraSize = view.getUint16(offset + 28, true);
    const name = new TextDecoder().decode(bytes.subarray(offset + 30, offset + 30 + nameSize));
    const payloadStart = offset + 30 + nameSize + extraSize;
    const payload = bytes.subarray(payloadStart, payloadStart + compressedSize);
    if (name === 'xl/worksheets/sheet1.xml') return new TextDecoder().decode(method === 8 ? inflateRawSync(payload) : payload);
    offset = payloadStart + compressedSize;
  }
  throw new Error('worksheet XML not found');
}

const cases = [
  { rows: 1, cols: 0, topLeft: 'A2', pane: 'bottomLeft' },
  { rows: 0, cols: 1, topLeft: 'B1', pane: 'topRight' },
  { rows: 1, cols: 1, topLeft: 'B2', pane: 'bottomRight' },
];
for (const freeze of cases) {
  const bytes = XLSX.write({ SheetNames: ['Sheet1'], Sheets: { Sheet1: { '!ref': 'A1:B2', '!freezePane': freeze, A1: { v: 'A' } } } }, { type: 'buffer', bookType: 'xlsx' });
  const xml = worksheetXml(bytes);
  if (freeze.cols) assert.ok(xml.includes(`xSplit="${freeze.cols}"`));
  if (freeze.rows) assert.ok(xml.includes(`ySplit="${freeze.rows}"`));
  assert.ok(xml.includes(`topLeftCell="${freeze.topLeft}"`));
  assert.ok(xml.includes(`activePane="${freeze.pane}"`));
}
console.log('freeze pane round-trip passed');
