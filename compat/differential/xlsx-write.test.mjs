// Differential test suite for `write()`/`writeFile()`/`writeFileSync()` (Phase D). Unlike
// xlsx-read.test.mjs (which only reads bytes the oracle produced), this file exercises
// BOTH directions of a real file-format round trip:
//
//   - own write -> own read       (does elixcee's writer + elixcee's reader agree with
//                                   themselves, on the same ground truth as the oracle?)
//   - own write -> oracle read    (can a genuinely independent parser open a file this
//                                   package produced and recover the same data — the
//                                   closest proxy available here to "does real Excel open
//                                   it", since no real Excel install exists in CI)
//   - oracle write -> own read    (does elixcee's reader correctly parse a file an
//                                   independent, spec-compliant writer produced — narrower
//                                   overlap with xlsx-read.test.mjs, included here too so
//                                   this file's own summary covers all three legs in one
//                                   place, per the write()-specific verification checklist)
//
// All three are compared against a fourth, independently computed baseline — oracle write
// -> oracle read — rather than against each other, so a bug shared by two of the three
// legs can't cancel out and hide.
//
// The oracle (`xlsx`) is imported here, inside compat/, and nowhere else — @elixcee/xlsx
// itself is imported only via a relative path into packages/xlsx/src, never as an npm
// dependency. See docs/xlsx-architecture.md's "Non-negotiable" section.
//
// ---- projectWorkBook is intentionally duplicated from xlsx-read.test.mjs ----
//
// Same rationale as that file's own duplicated CRC32/zip-writer helpers: xlsx-read.test.mjs
// is itself a runnable, self-executing script (ends by calling process.exit on failure),
// so importing it as a module would re-run its entire case matrix as a side effect. The
// projection logic — strip both sides down to exactly the fields @elixcee/xlsx's read()
// advertises support for (SheetNames, and per sheet: !ref/!merges/!rows/!cols, and each
// cell's {t,v,f,w,z}) — must stay identical between the two files, so any edit here should
// be mirrored there and vice versa.
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import zlib from 'node:zlib';
import XLSX from 'xlsx';
import * as elixcee from '../../packages/xlsx/src/index.mjs';
import { classify, summarizeByApiAndVerdict, formatApiVerdictSummary, VERDICTS } from './classify.mjs';
import { normalize } from './normalize.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const U = XLSX.utils;

const results = []; // { api, label, verdict }
function record(api, label, verdict) {
  results.push({ api, label, verdict });
}

const CELL_REF_RE = /^[A-Z]+[0-9]+$/;

function projectRowsOrCols(arr) {
  if (arr == null) return undefined;
  const out = new Array(arr.length);
  let anyHidden = false;
  for (let i = 0; i < arr.length; i++) {
    if (i in arr) {
      out[i] = { hidden: !!arr[i].hidden };
      if (out[i].hidden) anyHidden = true;
    }
  }
  return anyHidden ? out : undefined;
}

function projectSheet(ws) {
  if (ws == null) return null;
  const out = {};
  for (const key of Object.keys(ws)) {
    if (key === '!ref') {
      out['!ref'] = ws['!ref'];
    } else if (key === '!merges') {
      out['!merges'] = ws['!merges'].map((m) => ({ s: { r: m.s.r, c: m.s.c }, e: { r: m.e.r, c: m.e.c } }));
    } else if (key === '!rows' || key === '!cols') {
      const projected = projectRowsOrCols(ws[key]);
      if (projected !== undefined) out[key] = projected;
    } else if (CELL_REF_RE.test(key)) {
      out[key] = { t: ws[key].t, v: ws[key].v, f: ws[key].f, w: ws[key].w, z: ws[key].z };
    }
  }
  return out;
}

function projectWorkBook(wb) {
  const sheets = {};
  for (const name of wb.SheetNames) sheets[name] = projectSheet(wb.Sheets[name]);
  return { SheetNames: wb.SheetNames, Sheets: sheets };
}

// ---- write-case runner ----
//
// Builds a WorkBook once, writes it with BOTH writers, reads all the resulting byte
// buffers back with BOTH readers, and compares each of the three non-baseline
// combinations against the fourth (oracle write -> oracle read), the same {threw, value}
// -or-{threw, message, code} wrapping xlsx-read.test.mjs's invokeRead uses.
const DEFAULT_READ_OPTS = { cellStyles: true, cellDates: true, cellNF: true };

function invoke(fn) {
  try {
    return { threw: false, value: normalize(projectWorkBook(fn())) };
  } catch (e) {
    return { threw: true, message: e.message, code: e.code };
  }
}

function runWriteCase(label, wb, writeOpts, readOpts) {
  const wo = Object.assign({ type: 'buffer', bookType: 'xlsx' }, writeOpts);
  const ro = Object.assign({}, DEFAULT_READ_OPTS, readOpts);

  const oracleBytes = XLSX.write(wb, wo);
  const elixceeBytes = elixcee.write(wb, wo);

  const groundTruth = invoke(() => XLSX.read(oracleBytes, Object.assign({ type: 'buffer' }, ro)));
  assert.equal(groundTruth.threw, false, `baseline oracle write -> oracle read must not throw for "${label}": ${groundTruth.message}`);

  function compare(apiSuffix, val) {
    const verdict = classify({
      api: `write:${apiSuffix}`,
      oracleA: groundTruth,
      elixcee: val,
      elixceeErrorCode: val.code,
    });
    record(`write:${apiSuffix}`, label, verdict);
    return verdict;
  }

  compare('own-round-trip', invoke(() => elixcee.read(elixceeBytes, ro)));
  compare('oracle-reads-own-output', invoke(() => XLSX.read(elixceeBytes, Object.assign({ type: 'buffer' }, ro))));
  compare('own-reads-oracle-output', invoke(() => elixcee.read(oracleBytes, ro)));
}

// ---- cases ----

runWriteCase('mixed types (string/number/float/negative/zero/boolean)', (() => {
  const wb = U.book_new();
  U.book_append_sheet(
    wb,
    U.aoa_to_sheet([
      ['Name', 'Amount', 'Active'],
      ['Alice', 42, true],
      ['Bob', 3.5, false],
      ['Carol', -7, true],
      ['Zero', 0, false],
    ]),
    'Sheet1'
  );
  return wb;
})());

runWriteCase('unicode and XML-special characters in string cells', (() => {
  const wb = U.book_new();
  U.book_append_sheet(
    wb,
    U.aoa_to_sheet([
      ['quote " amp & lt < gt >', "apostrophe ' end"],
      ['unicode: café ★ 日本語 🎉', 'newline:\nembedded'],
    ]),
    'Sheet1'
  );
  return wb;
})());

runWriteCase('multiple worksheets with distinct content', (() => {
  const wb = U.book_new();
  U.book_append_sheet(wb, U.aoa_to_sheet([['a', 1], ['b', 2]]), 'First');
  U.book_append_sheet(wb, U.aoa_to_sheet([['x', 'y'], ['z', 'w']]), 'Second');
  U.book_append_sheet(wb, U.aoa_to_sheet([[true, false]]), 'Third Sheet');
  return wb;
})());

runWriteCase('merged cells (top-left anchor + covered cells)', (() => {
  const wb = U.book_new();
  const ws = U.aoa_to_sheet([['Title'], ['a', 'b', 'c']]);
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }];
  U.book_append_sheet(wb, ws, 'S1');
  return wb;
})());

runWriteCase('multiple non-overlapping merges on one sheet', (() => {
  const wb = U.book_new();
  const ws = U.aoa_to_sheet([['A', , 'B'], ['x', 'y', 'z', 'w']]);
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } },
    { s: { r: 0, c: 2 }, e: { r: 0, c: 3 } },
  ];
  U.book_append_sheet(wb, ws, 'S1');
  return wb;
})());

runWriteCase('a worksheet with no cells at all', (() => {
  const wb = U.book_new();
  U.book_append_sheet(wb, U.aoa_to_sheet([]), 'Empty');
  return wb;
})());

runWriteCase('a sparse worksheet with a gap between populated cells', (() => {
  const wb = U.book_new();
  U.book_append_sheet(wb, U.aoa_to_sheet([['A1', , , 'D1'], [, , , ,], [, 'B3']]), 'S1');
  return wb;
})());

runWriteCase('formula cells (.f roundtrip, cached value present)', (() => {
  const wb = U.book_new();
  U.book_append_sheet(
    wb,
    U.aoa_to_sheet([
      [1, 2, { t: 'n', v: 3, f: 'SUM(A1:B1)' }],
      [4, 5, { t: 'n', v: 9, f: 'SUM(A2:B2)' }],
    ]),
    'S1'
  );
  return wb;
})());

runWriteCase('hidden rows and columns', (() => {
  const wb = U.book_new();
  const ws = U.aoa_to_sheet([[1, 2], [3, 4], [5, 6]]);
  ws['!rows'] = [];
  ws['!rows'][1] = { hidden: true };
  ws['!cols'] = [];
  ws['!cols'][0] = { hidden: true };
  U.book_append_sheet(wb, ws, 'S1');
  return wb;
})());

runWriteCase('basic number formats (built-in and custom)', (() => {
  const wb = U.book_new();
  const ws = U.aoa_to_sheet([[42, 3.14159, 45444]]);
  U.cell_set_number_format(ws.A1, '0.00'); // built-in numFmtId
  U.cell_set_number_format(ws.C1, '0.00"kg"'); // custom numFmtId (164+)
  U.book_append_sheet(wb, ws, 'S1');
  return wb;
})());

runWriteCase('date-typed cells', (() => {
  const wb = U.book_new();
  U.book_append_sheet(
    wb,
    U.aoa_to_sheet([[new Date(2024, 0, 15, 13, 45, 30)], [new Date(2020, 11, 31)]]),
    'S1'
  );
  return wb;
})());

runWriteCase('multi-cell sheet mixing date/custom-numeric/general-numeric/string/boolean styles', (() => {
  const wb = U.book_new();
  const ws = U.aoa_to_sheet([[new Date(2024, 0, 15), 3.14159, 42, 'hello', true]]);
  U.cell_set_number_format(ws.B1, '0.00"kg"');
  U.book_append_sheet(wb, ws, 'S1');
  return wb;
})());

// ---- hyperlinks in XLSX output (0.17.0) ----
{
  const wb = U.book_new();
  const ws = U.aoa_to_sheet([['external', 'internal', 'mailto']]);
  U.cell_set_hyperlink(ws.A1, 'https://example.com/path', 'Example site');
  U.cell_set_internal_link(ws.B1, 'Sheet2!A1', 'Jump to Sheet2');
  U.cell_set_hyperlink(ws.C1, 'mailto:user@example.com');
  U.book_append_sheet(wb, ws, 'Sheet1');
  U.book_append_sheet(wb, U.aoa_to_sheet([['destination']]), 'Sheet2');

  const read = XLSX.read(elixcee.write(wb, { type: 'buffer', bookType: 'xlsx' }), {
    type: 'buffer',
    cellStyles: true,
  });
  assert.equal(read.Sheets.Sheet1.A1.l.Target, 'https://example.com/path');
  assert.equal(read.Sheets.Sheet1.A1.l.Tooltip, 'Example site');
  assert.equal(read.Sheets.Sheet1.B1.l.Target, '#Sheet2!A1');
  assert.equal(read.Sheets.Sheet1.B1.l.Tooltip, 'Jump to Sheet2');
  assert.equal(read.Sheets.Sheet1.C1.l.Target, 'mailto:user@example.com');
  console.log('OK  write: external, internal, mailto hyperlinks and tooltips survive an elixcee XLSX write');
}

// ---- sheet visibility ----
//
// read() (packages/xlsx/src/internal/read-shape.cjs) never parses xl/workbook.xml's own
// per-sheet visibility at all — confirmed by reading that file, not assumed — so
// projectWorkBook's comparison scope structurally cannot see it, and running this through
// runWriteCase would silently test nothing. Checked directly instead: elixcee's own
// write() output, read by the ORACLE (which does expose it, confirmed live below), must
// report the same Hidden values that were set.
{
  const wb = U.book_new();
  U.book_append_sheet(wb, U.aoa_to_sheet([[1, 2]]), 'Visible');
  U.book_append_sheet(wb, U.aoa_to_sheet([[3, 4]]), 'Hidden');
  U.book_append_sheet(wb, U.aoa_to_sheet([[5, 6]]), 'VeryHidden');
  U.book_set_sheet_visibility(wb, 1, U.consts.SHEET_HIDDEN);
  U.book_set_sheet_visibility(wb, 2, U.consts.SHEET_VERY_HIDDEN);

  const elixceeBytes = elixcee.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const read = XLSX.read(elixceeBytes, { type: 'buffer' });
  const hiddenFlags = read.Workbook.Sheets.map((s) => s.Hidden);
  assert.deepEqual(hiddenFlags, [0, 1, 2], `oracle-read Hidden flags of elixcee's own write() output: got ${JSON.stringify(hiddenFlags)}`);
  console.log('OK  write: sheet visibility (hidden/veryHidden) survives elixcee write -> oracle read');
}

// ---- writeFile / writeFileSync: real filesystem round trip ----

{
  const wb = U.book_new();
  U.book_append_sheet(wb, U.aoa_to_sheet([[1, 'x'], [2, 'y']]), 'S1');
  const tmpFile = join(fs.mkdtempSync(join(os.tmpdir(), 'elixcee-write-fs-')), 'out.xlsx');

  elixcee.writeFile(wb, tmpFile);
  assert.ok(fs.existsSync(tmpFile), 'writeFile() must create the file');
  const viaOracle = normalize(projectWorkBook(XLSX.read(fs.readFileSync(tmpFile), Object.assign({ type: 'buffer' }, DEFAULT_READ_OPTS))));
  const viaOwn = normalize(projectWorkBook(elixcee.readFileSync(tmpFile, DEFAULT_READ_OPTS)));
  assert.deepEqual(viaOwn, viaOracle, 'a file written by writeFile() must read back identically via both readers');

  fs.rmSync(tmpFile);
  elixcee.writeFileSync(wb, tmpFile);
  assert.ok(fs.existsSync(tmpFile), 'writeFileSync() must create the file');
  assert.equal(elixcee.writeFile, elixcee.writeFileSync, 'writeFile and writeFileSync must be the same function, matching the oracle');
  fs.rmSync(tmpFile);
  console.log('OK  write: writeFile()/writeFileSync() round trip through a real file, and are the same function');
}

// ---- determinism ----
//
// zip-writer.cjs's own doc comment claims a fixed epoch makes two writes of the same
// workbook byte-identical — checked directly rather than assumed.
{
  const wb = U.book_new();
  U.book_append_sheet(wb, U.aoa_to_sheet([[1, 'x', new Date(2024, 0, 1)]]), 'S1');
  const a = elixcee.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const b = elixcee.write(wb, { type: 'buffer', bookType: 'xlsx' });
  assert.ok(Buffer.from(a).equals(Buffer.from(b)), 'two write() calls on the same WorkBook must produce byte-identical output');
  console.log('OK  write: output is byte-deterministic across repeated calls on the same WorkBook');
}

// ---- output type variations: buffer / array / base64 must agree on the same bytes ----

{
  const wb = U.book_new();
  U.book_append_sheet(wb, U.aoa_to_sheet([[1, 'x']]), 'S1');
  const buf = elixcee.write(wb, { type: 'buffer' });
  const arr = elixcee.write(wb, { type: 'array' });
  const b64 = elixcee.write(wb, { type: 'base64' });
  assert.ok(buf instanceof Uint8Array, 'type:"buffer" must return a Buffer/Uint8Array');
  assert.ok(arr instanceof ArrayBuffer, 'type:"array" must return an ArrayBuffer');
  assert.equal(typeof b64, 'string', 'type:"base64" must return a string');
  assert.ok(Buffer.from(arr).equals(Buffer.from(buf)), 'type:"array" must carry the same bytes as type:"buffer"');
  assert.equal(Buffer.from(b64, 'base64').toString('hex'), Buffer.from(buf).toString('hex'), 'type:"base64" must decode to the same bytes as type:"buffer"');
  console.log('OK  write: buffer/array/base64 output types all carry identical bytes');
}

// ---- OOXML ZIP/XML structural validation ----
//
// Not a differential comparison — a direct structural check of elixcee's own write()
// output: a hand-rolled ZIP READER (mirroring the project's existing precedent of a
// hand-rolled zip WRITER in packages/xlsx/src/internal/zip-writer.cjs, and a hand-rolled
// zip writer in xlsx-read.test.mjs — no new npm dependency) that parses the local file
// headers, verifies each entry's CRC-32 against its declared value after DEFLATE
// decompression, cross-checks the central directory and end-of-central-directory record,
// and confirms every OOXML part [Content_Types].xml declares/requires is actually present.
function readZipEntries(bytes) {
  const buf = Buffer.from(bytes);
  const eocdSig = 0x06054b50;
  let eocdOff = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === eocdSig) {
      eocdOff = i;
      break;
    }
  }
  assert.ok(eocdOff >= 0, 'no end-of-central-directory record found');
  const totalEntries = buf.readUInt16LE(eocdOff + 10);
  const centralDirSize = buf.readUInt32LE(eocdOff + 12);
  const centralDirOffset = buf.readUInt32LE(eocdOff + 16);
  assert.equal(centralDirOffset + centralDirSize, eocdOff, 'central directory must end exactly where EOCD begins');

  const entries = [];
  let p = centralDirOffset;
  for (let i = 0; i < totalEntries; i++) {
    assert.equal(buf.readUInt32LE(p), 0x02014b50, `central directory entry ${i} has a bad signature`);
    const method = buf.readUInt16LE(p + 10);
    const crc = buf.readUInt32LE(p + 16);
    const compSize = buf.readUInt32LE(p + 20);
    const uncompSize = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localOffset = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen);
    entries.push({ name, method, crc, compSize, uncompSize, localOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }
  assert.equal(p, centralDirOffset + centralDirSize, 'central directory entries must exactly fill the declared size');

  for (const e of entries) {
    assert.equal(buf.readUInt32LE(e.localOffset), 0x04034b50, `local header for ${e.name} has a bad signature`);
    const localNameLen = buf.readUInt16LE(e.localOffset + 26);
    const localExtraLen = buf.readUInt16LE(e.localOffset + 28);
    const dataStart = e.localOffset + 30 + localNameLen + localExtraLen;
    const compressed = buf.subarray(dataStart, dataStart + e.compSize);
    const data = e.method === 0 ? compressed : zlib.inflateRawSync(compressed);
    assert.equal(data.length, e.uncompSize, `${e.name}: decompressed size must match the declared uncompressed size`);
    assert.equal(crc32(data), e.crc, `${e.name}: CRC-32 of decompressed data must match the declared CRC`);
    e.data = data;
  }
  return entries;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();
function crc32(buf) {
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = CRC_TABLE[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

{
  const wb = U.book_new();
  U.book_append_sheet(wb, U.aoa_to_sheet([[1, 'x']]), 'Sheet1');
  U.book_append_sheet(wb, U.aoa_to_sheet([[2, 'y']]), 'Sheet2');
  const bytes = elixcee.write(wb, { type: 'buffer' });
  const entries = readZipEntries(bytes);
  const names = entries.map((e) => e.name).sort();
  const expected = [
    '[Content_Types].xml',
    '_rels/.rels',
    'docProps/app.xml',
    'docProps/core.xml',
    'xl/_rels/workbook.xml.rels',
    'xl/styles.xml',
    'xl/workbook.xml',
    'xl/worksheets/sheet1.xml',
    'xl/worksheets/sheet2.xml',
  ].sort();
  assert.deepEqual(names, expected, `write() must produce exactly the minimal required OOXML part set, got: ${JSON.stringify(names)}`);

  // Every XML part must be well-formed enough to at least be parseable by the oracle's
  // own writer-independent reader (already proven above for the whole archive via
  // runWriteCase — this adds a per-part balanced-tag sanity check, catching a corrupt
  // individual part that a whole-archive read might tolerate via a fallback path).
  for (const e of entries.filter((e) => e.name.endsWith('.xml') || e.name.endsWith('.rels'))) {
    const text = e.data.toString('utf8');
    assert.ok(text.startsWith('<?xml version="1.0"'), `${e.name} must start with an XML declaration`);
    // Classify every tag by its own shape (not by excluding "/" from attribute values,
    // which breaks on any xmlns URL) — self-closing first (</tag> can't also match "/>"
    // since it has no content between < and >), then open vs. close by the leading "/".
    let opens = 0;
    let closes = 0;
    let selfClosing = 0;
    for (const m of text.matchAll(/<\/?[a-zA-Z][^>]*>/g)) {
      const tag = m[0];
      if (tag.endsWith('/>')) selfClosing++;
      else if (tag.startsWith('</')) closes++;
      else opens++;
    }
    // Every opening tag not self-closing must have a matching close — a crude but
    // effective well-formedness check that catches a truncated/mis-escaped part.
    assert.equal(opens, closes, `${e.name}: unbalanced XML tags (${opens} opens vs ${closes} closes, ${selfClosing} self-closing)`);
  }

  // [Content_Types].xml must declare an Override for every worksheet part actually
  // present, and workbook.xml.rels must reference every worksheet by the same target the
  // workbook.xml <sheet r:id> points at — cross-checked directly rather than trusting
  // read() to have caught a mismatch (read() tolerates a good deal more than a strict
  // OOXML consumer would).
  const contentTypes = entries.find((e) => e.name === '[Content_Types].xml').data.toString('utf8');
  for (let i = 1; i <= 2; i++) {
    assert.ok(
      contentTypes.includes(`/xl/worksheets/sheet${i}.xml`),
      `[Content_Types].xml must declare sheet${i}.xml`
    );
  }
  const wbXml = entries.find((e) => e.name === 'xl/workbook.xml').data.toString('utf8');
  const wbRels = entries.find((e) => e.name === 'xl/_rels/workbook.xml.rels').data.toString('utf8');
  const rIds = [...wbXml.matchAll(/r:id="(rId\d+)"/g)].map((m) => m[1]);
  assert.equal(rIds.length, 2, 'workbook.xml must declare one r:id per sheet');
  for (const rid of rIds) {
    assert.ok(wbRels.includes(`Id="${rid}"`), `workbook.xml.rels must define a relationship for ${rid} referenced by workbook.xml`);
  }
  console.log('OK  write: OOXML ZIP/XML structural validation (parts, CRC-32, balanced XML, Content_Types/rels cross-references)');
}

// ---- malformed-workbook rejection: explicit errors, never silent misbehavior ----
//
// Each case must throw the documented ELIXCEE_* code — never return a truncated/garbage
// buffer and never silently drop the offending part.
function assertWriteThrows(label, wb, opts, expectedCode) {
  let threw = false;
  let code;
  try {
    elixcee.write(wb, Object.assign({ type: 'buffer' }, opts));
  } catch (e) {
    threw = true;
    code = e.code;
  }
  assert.equal(threw, true, `"${label}" must throw, not silently produce output`);
  assert.equal(code, expectedCode, `"${label}" must throw ${expectedCode}, got ${code}`);
  console.log(`OK  write: malformed input rejected — ${label} -> ${code}`);
}

assertWriteThrows('null workbook', null, {}, 'ELIXCEE_UNSUPPORTED_SHEET_SHAPE');
assertWriteThrows('workbook missing SheetNames', { Sheets: {} }, {}, 'ELIXCEE_UNSUPPORTED_SHEET_SHAPE');
assertWriteThrows('workbook with zero sheets', { SheetNames: [], Sheets: {} }, {}, 'ELIXCEE_UNSUPPORTED_SHEET_SHAPE');
assertWriteThrows(
  'a sheet listed in SheetNames but missing from Sheets',
  { SheetNames: ['S1'], Sheets: {} },
  {},
  'ELIXCEE_UNSUPPORTED_SHEET_SHAPE'
);
assertWriteThrows(
  'a worksheet that is neither an object nor an array',
  { SheetNames: ['S1'], Sheets: { S1: 'not a sheet' } },
  {},
  'ELIXCEE_UNSUPPORTED_SHEET_SHAPE'
);
assertWriteThrows(
  'a dense worksheet row that is not an array',
  { SheetNames: ['S1'], Sheets: { S1: [['ok'], 'not a row'] } },
  {},
  'ELIXCEE_UNSUPPORTED_SHEET_SHAPE'
);
assertWriteThrows(
  'a cell with an unsupported type tag',
  { SheetNames: ['S1'], Sheets: { S1: { A1: { t: 'e', v: '#N/A' } } } },
  {},
  'ELIXCEE_UNSUPPORTED_CELL_TYPE'
);
assertWriteThrows(
  'a numeric cell with a non-finite value',
  { SheetNames: ['S1'], Sheets: { S1: { A1: { t: 'n', v: Infinity } } } },
  {},
  'ELIXCEE_UNSUPPORTED_CELL_TYPE'
);
assertWriteThrows(
  'a formula cell whose cached value is not a string/boolean/finite number',
  { SheetNames: ['S1'], Sheets: { S1: { A1: { t: 'n', v: { nope: true }, f: 'A2' } } } },
  {},
  'ELIXCEE_UNSUPPORTED_CELL_TYPE'
);
assertWriteThrows(
  'an Invalid Date cell',
  { SheetNames: ['S1'], Sheets: { S1: { A1: { t: 'd', v: new Date(NaN) } } } },
  {},
  'ELIXCEE_UNSUPPORTED_CELL_TYPE'
);
assertWriteThrows(
  'a declared !ref exceeding the range-size safety limit',
  { SheetNames: ['S1'], Sheets: { S1: { '!ref': 'A1:XFD1048576' } } },
  {},
  'ELIXCEE_RANGE_TOO_LARGE'
);

// bookType/type: compared against the oracle via classify() (the oracle DOES support a
// wider surface here — 'ods' output, and 'binary'/'string'/'file' write types — so these
// are genuine, disclosed capability gaps, not bugs; registered in classify.mjs's
// UNSUPPORTED_ALLOWLIST under api 'write', matching the anti-laundering discipline every
// other UNSUPPORTED verdict in this repo follows).
{
  const wb = U.book_new();
  U.book_append_sheet(wb, U.aoa_to_sheet([[1]]), 'S1');

  function invokeWrite(fn) {
    try {
      return { threw: false, value: fn() };
    } catch (e) {
      return { threw: true, message: e.message, code: e.code };
    }
  }

  {
    const oracleVal = invokeWrite(() => XLSX.write(wb, { type: 'buffer', bookType: 'ods' }));
    const elixceeVal = invokeWrite(() => elixcee.write(wb, { type: 'buffer', bookType: 'ods' }));
    const verdict = classify({
      api: 'write',
      unsupportedCaseId: "bookType='ods' (ODS output not implemented)",
      oracleA: oracleVal,
      elixcee: elixceeVal,
      elixceeErrorCode: elixceeVal.code,
    });
    record('write', "bookType='ods' must throw ELIXCEE_UNSUPPORTED_BOOK_TYPE, not silently write something else", verdict);
  }

  assertWriteThrows('opts.type omitted entirely', wb, { type: undefined }, 'ELIXCEE_UNSUPPORTED_WRITE_TYPE');
  assertWriteThrows("opts.type: 'binary' (not implemented)", wb, { type: 'binary' }, 'ELIXCEE_UNSUPPORTED_WRITE_TYPE');
}

// ---- browser build: write() works (pure JS/ZIP, no filesystem), writeFile/writeFileSync
//      throw ELIXCEE_UNSUPPORTED_IN_BROWSER ----
//
// Same `node --conditions=browser` dispatch xlsx-read.test.mjs's own read-item-5/readFile
// blocks use — a real subprocess resolving '@elixcee/xlsx' through the "browser" export
// condition via Node's self-referencing package resolution. Note this does NOT prove
// Buffer-freeness (Node still has a global Buffer even under --conditions=browser) — that
// specific claim is verified separately, in a REAL unshimmed browser, by
// scripts/browser-smoke.mjs (see docs/xlsx-architecture.md's "Phase D" section for the
// real bug that check caught: Buffer is not defined in an actual Chrome tab).
{
  const script =
    "import * as X from '@elixcee/xlsx';\n" +
    "const wb = X.book_new();\n" +
    "X.book_append_sheet(wb, X.aoa_to_sheet([[1, 'x']]), 'S1');\n" +
    "const bytes = X.write(wb, { type: 'buffer' });\n" +
    "const b64 = X.write(wb, { type: 'base64' });\n" +
    "const readBackFromBytes = X.read(bytes);\n" +
    "const readBackFromB64 = X.read(b64, { type: 'base64' });\n" +
    "let writeFileOut = { threw: false };\n" +
    "try { X.writeFile(wb, '/tmp/whatever.xlsx'); } catch (e) { writeFileOut = { threw: true, code: e.code }; }\n" +
    "process.stdout.write(JSON.stringify({ byteLength: bytes.length, b64Length: b64.length, readBackFromBytesSheetNames: readBackFromBytes.SheetNames, readBackFromB64SheetNames: readBackFromB64.SheetNames, writeFileOut, aliased: X.writeFile === X.writeFileSync }));\n";
  const r = JSON.parse(
    execFileSync(process.execPath, ['--conditions=browser', '--input-type=module', '-e', script], {
      cwd: join(here, '../../packages/xlsx'),
      encoding: 'utf8',
    })
  );
  assert.ok(r.byteLength > 0, 'browser-entry write() must produce real bytes (no filesystem dependency)');
  assert.ok(r.b64Length > 0, 'browser-entry write({type:"base64"}) must produce a real string');
  assert.deepEqual(r.readBackFromBytesSheetNames, ['S1'], 'browser-entry write({type:"buffer"})->read() round trip must preserve sheet names');
  assert.deepEqual(r.readBackFromB64SheetNames, ['S1'], 'browser-entry write({type:"base64"})->read() round trip must preserve sheet names');
  assert.equal(r.writeFileOut.threw, true, 'browser-entry writeFile() must throw, not silently pretend to have a filesystem');
  assert.equal(r.writeFileOut.code, 'ELIXCEE_UNSUPPORTED_IN_BROWSER', `unexpected error code: ${r.writeFileOut.code}`);
  assert.equal(r.aliased, true, 'browser entry must keep writeFile === writeFileSync, matching the Node entry and the oracle');
  console.log('OK  write: browser entry — write() works with no filesystem (buffer + base64), writeFile()/writeFileSync() throw ELIXCEE_UNSUPPORTED_IN_BROWSER');
}

// ---- summary / exit code (matches xlsx-read.test.mjs's convention) ----

const ACCEPTABLE = new Set(['MATCH', 'INTENTIONAL_SAFETY_DIVERGENCE', 'INTENTIONAL_SECURITY_DIVERGENCE', 'UNSUPPORTED']);

const byApi = new Map();
const totals = new Map();
for (const r of results) {
  if (!byApi.has(r.api)) byApi.set(r.api, { other: [] });
  if (!ACCEPTABLE.has(r.verdict)) byApi.get(r.api).other.push({ label: r.label, verdict: r.verdict });
  totals.set(r.verdict, (totals.get(r.verdict) || 0) + 1);
}

const byApiVerdict = summarizeByApiAndVerdict(results);
console.log('\n=== write() differential summary (compat/differential/xlsx-write.test.mjs) ===');
console.log('Comparison scope: SheetNames, per-sheet !ref/!merges/!rows/!cols, per-cell {t,v,f,w,z} —');
console.log('same projection xlsx-read.test.mjs uses, applied to all four write/read combinations.');
let anyFailure = false;
for (const [api, bucket] of byApi) {
  const status = bucket.other.length === 0 ? 'OK' : 'FAIL';
  if (bucket.other.length > 0) anyFailure = true;
  console.log(`${status}  ${formatApiVerdictSummary(new Map([[api, byApiVerdict.get(api)]]))}`);
  for (const o of bucket.other) console.log(`      ${o.verdict}: ${o.label}`);
}
console.log('\n=== Totals ===');
for (const v of VERDICTS) {
  if (totals.has(v)) console.log(`${v}:`.padEnd(38) + totals.get(v));
}

if (anyFailure) {
  console.error('\nwrite() differential suite FAILED: at least one case is not an acceptable verdict.');
  process.exit(1);
}
console.log('\nwrite() differential suite passed: every case matches on its declared supported-field scope.');
