import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import * as XLSX from '../src/index.mjs';
import { extractZipEntry } from '../../../playground/src/zip-budget.js';

const fixture = path.resolve(new URL('../../../compat/oracle-excel-com/fixtures/pristine/fixture2_vba_macro.xlsm', import.meta.url).pathname);
const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'elixcee-xlsm-reopen-'));
const outputDir = path.join(temp, 'converted');
fs.mkdirSync(outputDir);
try {
  const source = new Uint8Array(fs.readFileSync(fixture));
  const workbook = XLSX.read(source, { cellStyles: true });
  const vbaProject = await extractZipEntry(source, 'xl/vbaProject.bin');
  assert.ok(vbaProject?.length, 'fixture must contain xl/vbaProject.bin');
  workbook['!vbaProject'] = vbaProject;
  const generated = Buffer.from(XLSX.write(workbook, { bookType: 'xlsm', type: 'buffer' }));
  const generatedPath = path.join(temp, 'roundtrip.xlsm');
  fs.writeFileSync(generatedPath, generated);
  execFileSync('soffice', ['--headless', '--convert-to', 'xlsx', '--outdir', outputDir, generatedPath], { stdio: 'pipe' });
  const converted = path.join(outputDir, 'roundtrip.xlsx');
  assert.ok(fs.existsSync(converted), 'LibreOffice did not produce a reopened workbook');
  const reopened = XLSX.read(fs.readFileSync(converted));
  assert.deepEqual(reopened.SheetNames, workbook.SheetNames);
  console.log(`XLSM real-fixture reopen smoke passed (${workbook.SheetNames.length} sheet(s), ${generated.length} bytes)`);
} finally {
  fs.rmSync(temp, { recursive: true, force: true });
}
