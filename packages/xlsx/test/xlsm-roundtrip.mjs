import assert from 'node:assert/strict';
import * as XLSX from '../src/index.mjs';
import { buildXlsxZipEntries } from '../src/internal/xlsx-writer.cjs';
import { extractZipEntry } from '../../../playground/src/zip-budget.js';

const vbaProject = Uint8Array.from([0x00, 0x01, 0x02, 0x4d, 0x5a]);
const workbook = XLSX.book_new();
XLSX.book_append_sheet(workbook, XLSX.aoa_to_sheet([['macro-safe', 42]]), 'Sheet1');
workbook['!vbaProject'] = vbaProject;

const bytes = XLSX.write(workbook, { bookType: 'xlsm', type: 'buffer' });
const entries = buildXlsxZipEntries(workbook);
const entryText = (name) => new TextDecoder().decode(entries.find((entry) => entry.name === name).data);
assert.ok(entries.some((entry) => entry.name === 'xl/vbaProject.bin'));
assert.match(entryText('[Content_Types].xml'), /macroEnabled\.main\+xml/);
assert.match(entryText('[Content_Types].xml'), /vbaProject/);
assert.match(entryText('xl/_rels/workbook.xml.rels'), /vbaProject/);
assert.throws(() => XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' }), (error) => error.code === 'ELIXCEE_UNSUPPORTED_BOOK_TYPE');
assert.throws(() => XLSX.write(XLSX.book_new(), { bookType: 'xlsm', type: 'buffer' }), (error) => error.code === 'ELIXCEE_UNSUPPORTED_BOOK_TYPE');
const roundTripped = XLSX.read(bytes);
assert.deepEqual(roundTripped.SheetNames, ['Sheet1']);
assert.equal(roundTripped.Sheets.Sheet1.A1.v, 'macro-safe');
assert.deepEqual(await extractZipEntry(bytes, 'xl/vbaProject.bin'), vbaProject);
console.log('XLSM opaque VBA preservation smoke passed');
