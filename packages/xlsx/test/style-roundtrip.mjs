import assert from 'node:assert/strict';
import fs from 'node:fs';
import * as XLSX from '../src/index.mjs';

const fixture = new URL('../../../compat/oracle-excel-com/fixtures/pristine/fixture4_hyperlink_comment_name.xlsm', import.meta.url);
const workbook = XLSX.read(fs.readFileSync(fixture), { cellStyles: true });
const sheet = workbook.Sheets[workbook.SheetNames[0]];
assert.equal(sheet.D6?.s?.font?.underline, true);
assert.equal(sheet.D6?.s?.alignment?.vertical, 'center');
const output = XLSX.read(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), { cellStyles: true });
const roundTrip = output.Sheets[output.SheetNames[0]].D6?.s;
assert.equal(roundTrip?.font?.underline, true);
assert.equal(roundTrip?.alignment?.vertical, 'center');
console.log('cell style read/write round-trip smoke passed');
