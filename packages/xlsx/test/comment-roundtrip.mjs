import assert from 'node:assert/strict';
import fs from 'node:fs';
import * as XLSX from '../src/index.mjs';

const fixture = new URL('../../../compat/oracle-excel-com/fixtures/pristine/fixture4_hyperlink_comment_name.xlsm', import.meta.url);
const workbook = XLSX.read(fs.readFileSync(fixture), { cellStyles: true });
const sheet = workbook.Sheets[workbook.SheetNames[0]];
assert.equal(sheet.C4?.c?.[0]?.a, 'tc={7FCCF7C4-7381-A54C-8681-54A93A005F37}');
assert.match(sheet.C4?.c?.[0]?.t || '', /aaaaa/);
const output = XLSX.read(XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), { cellStyles: true });
assert.match(output.Sheets[output.SheetNames[0]].C4?.c?.[0]?.t || '', /aaaaa/);
console.log('legacy comment read/write round-trip smoke passed');
