// Differential test suite for `read()` (Phase 2B): builds real .xlsx byte buffers with
// the real oracle (xlsx@0.18.5), then reads those exact bytes through BOTH the oracle's
// own `XLSX.read()` and @elixcee/xlsx's new `read()` (backed by crates/elixcee-wasm, a
// WASM bridge over elixcee's own hand-rolled reader — src/reader.rs's
// read_workbook_from_bytes). This is a real file-format round-trip, not a synthetic
// object comparison — the bytes elixcee parses are produced by writing a real workbook
// through the oracle's own writer.
//
// The oracle (`xlsx`) is imported here, inside compat/, and nowhere else — @elixcee/xlsx
// itself is imported only via a relative path into packages/xlsx/src, never as an npm
// dependency. See docs/xlsx-architecture.md's "Non-negotiable" section.
//
// ---- why this file projects both WorkBooks before calling classify() ----
//
// The oracle's read() returns a WorkBook carrying an order of magnitude more than
// {SheetNames, Sheets}: Directory (CFB/zip part listing), Workbook (WBProps/CalcPr/
// Views/...), Props/Custprops/Deps, Strings, Styles (NumberFmt/Fonts/Fills/Borders/
// CellXf), Themes, SSF — none of which `read_workbook_from_bytes` (src/reader.rs) parses
// or could parse without a much larger scope than this MVP (see
// crates/elixcee-wasm/src/lib.rs's `read_workbook` doc comment for the exact list: no
// cell formulas, no formatted `.w`/`.h` text, no date-typed cells, no `!rows`/`!cols`).
// Confirmed empirically (not assumed) by writing a plain 3x3 aoa_to_sheet workbook through
// the oracle and reading it back — every cell came back with `.w` (and strings also got
// `.h`) even with no formatting ever applied, and the WorkBook itself carried Styles/SSF/
// Themes objects derived purely from the oracle's own default styles.xml.
//
// classify.mjs's own doc comment says callers must normalize before calling classify() —
// normalize.mjs does that for type-tagging (NaN/undefined/-0/Date/...). This file adds one
// more normalization step specific to `read()`: projectWorkBook() strips both sides down
// to exactly the fields @elixcee/xlsx's read() advertises support for (SheetNames, and per
// sheet: !ref, !merges, and each cell's {t, v} only). This is a single, up-front,
// documented scope boundary — not a per-case escape hatch like UNSUPPORTED_ALLOWLIST — so
// if a genuinely supported field (t, v, !ref, !merges, SheetNames) ever diverges, it still
// surfaces as UNCLASSIFIED/BUG exactly as classify() intends; only fields this MVP never
// claimed to produce are excluded, and excluded identically from both sides so the
// comparison stays apples-to-apples rather than favoring elixcee's narrower shape.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
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

// Strips a WorkBook down to the fields @elixcee/xlsx's read() advertises support for — see
// this file's header comment for why. Applied to BOTH the oracle's and elixcee's output.
function projectWorkBook(wb) {
  const sheets = {};
  for (const name of wb.SheetNames) {
    sheets[name] = projectSheet(wb.Sheets[name]);
  }
  return { SheetNames: wb.SheetNames, Sheets: sheets };
}

function projectSheet(ws) {
  if (ws == null) return null;
  const out = {};
  for (const key of Object.keys(ws)) {
    if (key === '!ref') {
      out['!ref'] = ws['!ref'];
    } else if (key === '!merges') {
      out['!merges'] = ws['!merges'].map((m) => ({ s: { r: m.s.r, c: m.s.c }, e: { r: m.e.r, c: m.e.c } }));
    } else if (CELL_REF_RE.test(key)) {
      out[key] = { t: ws[key].t, v: ws[key].v };
    }
    // Every other key (!cols, !rows, !fullref, !type, ...) is out of this MVP's scope —
    // silently dropped from BOTH sides, not just elixcee's, so its absence can never look
    // like a false MATCH for a field neither side is being compared on.
  }
  return out;
}

// Builds real .xlsx bytes via the oracle's own writer — never elixcee's (there is no
// elixcee-side writer yet; this MVP is read-only). `sheets` is [[name, aoa, merges?], ...].
function buildXlsxBytes(sheets) {
  const wb = U.book_new();
  for (const [name, aoa, merges] of sheets) {
    const ws = U.aoa_to_sheet(aoa);
    if (merges) ws['!merges'] = merges;
    U.book_append_sheet(wb, ws, name);
  }
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// Both sides go through the same {threw, value}-or-{threw, message, code} wrapping (as
// xlsx-utils.test.mjs's invoke() does) so a throw-vs-value asymmetry is itself a visible
// divergence rather than an apples-to-oranges shape mismatch feeding classify().
function invokeRead(fn, bytes) {
  try {
    return { threw: false, value: normalize(projectWorkBook(fn(bytes))) };
  } catch (e) {
    return { threw: true, message: e.message, code: e.code };
  }
}

// The common comparison step, independent of where `bytes` came from (the oracle's own
// writer via buildXlsxBytes, a real fixture file read from disk, or a hand-built zip —
// see the two case groups below).
function runReadCaseBytes(label, bytes, unsupportedCaseId) {
  const oracleVal = invokeRead((b) => XLSX.read(b, { type: 'buffer' }), bytes);
  const elixceeVal = invokeRead((b) => elixcee.read(b), bytes);
  const verdict = classify({
    api: 'read',
    unsupportedCaseId,
    oracleA: oracleVal,
    elixcee: elixceeVal,
    elixceeErrorCode: elixceeVal.code,
  });
  record('read', label, verdict);
  return verdict;
}

function runReadCase(label, sheets, unsupportedCaseId) {
  return runReadCaseBytes(label, buildXlsxBytes(sheets), unsupportedCaseId);
}

// caseIds for the two registered gaps below — kept as named constants so the call sites
// and the classify.mjs registry entries can't drift out of sync with each other.
const EMPTY_STRING_CELL_CASE_ID =
  'empty-string cell value (<v></v> with zero characters — no Text event for the ' +
  "pull-XML parser to record, see reader.rs's xlsx_sheet_cells)";
const DIMENSION_WIDER_THAN_DATA_CASE_ID =
  "declared <dimension> wider than the populated cell range — reader.rs never parses " +
  "<dimension> at all, it always computes !ref from the populated-cell bounding box";

// ---- a minimal hand-built .xlsx (STORED-only zip, no compression library needed) ----
//
// Used by the "!ref" case below, which needs a sheet XML with a <dimension> tag that
// deliberately disagrees with the populated cell range — something none of the oracle's
// own utils.* writer functions (aoa_to_sheet et al.) can produce, since they always write
// <dimension> equal to the populated bounding box. A real Excel/LibreOffice file routinely
// writes a wider one (formatting-only cells past the last populated one), so this isn't a
// contrived-only-in-theory input. Hand-rolling a tiny STORED (uncompressed) zip writer
// here — rather than adding a zip-library devDependency for one test case, or shelling out
// to a system `zip`/`unzip` binary (not portable to every CI environment) — is the
// smallest self-contained way to get bytes elixcee's own zip reader (Stored-entry support
// is unconditional, no feature needed) and the oracle can both actually parse.
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

function buildStoredZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const { name, content } of entries) {
    const nameBuf = Buffer.from(name, 'utf8');
    const dataBuf = Buffer.from(content, 'utf8');
    const crc = crc32(dataBuf);

    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0); // local file header signature
    localHeader.writeUInt16LE(20, 4); // version needed to extract
    localHeader.writeUInt16LE(0, 6); // flags
    localHeader.writeUInt16LE(0, 8); // method 0 = stored
    localHeader.writeUInt16LE(0, 10); // mod time
    localHeader.writeUInt16LE(0, 12); // mod date
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(dataBuf.length, 18); // compressed size == uncompressed (stored)
    localHeader.writeUInt32LE(dataBuf.length, 22);
    localHeader.writeUInt16LE(nameBuf.length, 26);
    localHeader.writeUInt16LE(0, 28); // extra field length
    const localEntry = Buffer.concat([localHeader, nameBuf, dataBuf]);
    localParts.push(localEntry);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0); // central directory file header signature
    centralHeader.writeUInt16LE(20, 4); // version made by
    centralHeader.writeUInt16LE(20, 6); // version needed
    centralHeader.writeUInt16LE(0, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(dataBuf.length, 20);
    centralHeader.writeUInt32LE(dataBuf.length, 24);
    centralHeader.writeUInt16LE(nameBuf.length, 28);
    centralHeader.writeUInt16LE(0, 30); // extra field length
    centralHeader.writeUInt16LE(0, 32); // comment length
    centralHeader.writeUInt16LE(0, 34); // disk number start
    centralHeader.writeUInt16LE(0, 36); // internal file attributes
    centralHeader.writeUInt32LE(0, 38); // external file attributes
    centralHeader.writeUInt32LE(offset, 42); // relative offset of local header
    centralParts.push(Buffer.concat([centralHeader, nameBuf]));

    offset += localEntry.length;
  }
  const centralDir = Buffer.concat(centralParts);
  const centralOffset = offset;
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); // end of central directory signature
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(centralDir.length, 12);
  eocd.writeUInt32LE(centralOffset, 16);
  eocd.writeUInt16LE(0, 20);
  return Buffer.concat([...localParts, centralDir, eocd]);
}

// A minimal but complete single-sheet .xlsx — [Content_Types].xml, both .rels parts,
// xl/workbook.xml, and one worksheet with a <dimension> deliberately wider than the two
// populated cells it actually contains.
function buildDimensionWiderThanDataXlsxBytes() {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;
  const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
  const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
  // Only A1:B2 are populated; <dimension> claims A1:E10 (e.g. formatting-only cells past
  // the last populated one, or simply a stale/hand-edited dimension — both real-world).
  const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:E10"/>
<sheetData>
<row r="1"><c r="A1" t="str"><v>a</v></c><c r="B1" t="str"><v>b</v></c></row>
<row r="2"><c r="A2" t="str"><v>c</v></c><c r="B2" t="str"><v>d</v></c></row>
</sheetData>
</worksheet>`;
  return buildStoredZip([
    { name: '[Content_Types].xml', content: contentTypes },
    { name: '_rels/.rels', content: rootRels },
    { name: 'xl/workbook.xml', content: workbookXml },
    { name: 'xl/_rels/workbook.xml.rels', content: workbookRels },
    { name: 'xl/worksheets/sheet1.xml', content: sheetXml },
  ]);
}

// ---- cases ----

runReadCase('mixed types (string/number/float/negative/zero/boolean)', [
  ['Sheet1', [
    ['Name', 'Amount', 'Active'],
    ['Alice', 42, true],
    ['Bob', 3.5, false],
    ['Carol', -7, true],
    ['Zero', 0, false],
  ]],
]);

runReadCase('unicode and XML-special characters in string cells', [
  ['Sheet1', [
    ['quote " amp & lt < gt >', "apostrophe ' end"],
    ['unicode: café ★ 日本語 🎉', 'newline:\nembedded'],
  ]],
]);

runReadCase('multiple sheets with distinct content', [
  ['First', [['a', 1], ['b', 2]]],
  ['Second', [['x', 'y'], ['z', 'w']]],
  ['Third Sheet', [[true, false]]],
]);

// The covered (non-anchor) cells under a merge are left as array HOLES, not empty
// strings — see the dedicated "empty-string cell" case and its UNSUPPORTED_ALLOWLIST
// registration below for why '' specifically is a known, separately-tracked gap; a hole
// (aoa_to_sheet skips `undefined` entries outright) sidesteps that gap so this case tests
// only what it says it tests — merge-range fidelity, not blank-cell representation.
runReadCase('merged cells (top-left anchor + covered cells)', [
  ['S1', [
    ['Title'],
    ['a', 'b', 'c'],
  ], [{ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } }]],
]);

runReadCase('multiple non-overlapping merges on one sheet', [
  ['S1', [
    ['A', , 'B'],
    ['x', 'y', 'z', 'w'],
  ], [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } },
    { s: { r: 0, c: 2 }, e: { r: 0, c: 3 } },
  ]],
]);

runReadCase('a sheet with no cells at all', [
  ['Empty', []],
]);

runReadCase('a sparse sheet with a gap between populated cells', [
  ['S1', [
    ['A1', , , 'D1'],
    [, , , ,],
    [, 'B3'],
  ]],
]);

runReadCase('boundary numeric values (large integer, small float, negative float)', [
  ['S1', [
    [9007199254740991, 0.1, -123.456],
    [1048576, 16384, -0],
  ]],
]);

// ---- discovered gap: empty-string cell values ----
//
// Found by this test file, not assumed: the oracle's writer emits an empty-string aoa
// cell as a real `<c r="B1" t="str"><v></v></c>` (confirmed by inspecting the actual
// written sheet1.xml) — a self-closing-content `<v>` element with ZERO characters between
// its open and close tags. `reader.rs`'s hand-rolled pull-XML parser (`xlsx_sheet_cells`)
// only records a cell's value on an `Ev::Text` event; an empty element never produces one
// (there's no text to emit), so `in_v` is set on `<v>` and cleared on `</v>` with no
// `cells.insert(...)` ever happening in between — the cell is silently absent from
// elixcee's output, while the oracle reports `{t:"s", v:""}`.
//
// This is a real, narrow parser gap in `reader.rs` itself (shared by BOTH the path-based
// `read_workbook` and the new `read_workbook_from_bytes` — not something introduced by
// this WASM bridge or read() MVP), left unfixed here deliberately: fixing it means
// changing `xlsx_sheet_cells`'s shared Text-event-driven cell-recording logic, which
// affects the CLI/VM path too and was explicitly out of scope for this phase's "pure
// extraction, no behavior change" requirement (docs/xlsx-architecture.md's "reader.rs
// buffer-API resolution"). Registered below (see classify.mjs's UNSUPPORTED_ALLOWLIST)
// rather than silently worked around, so it stays visible for a future fix.
runReadCase(
  'empty-string cell value (known reader.rs gap — registered UNSUPPORTED)',
  [['S1', [['before', '', 'after']]]],
  EMPTY_STRING_CELL_CASE_ID
);

// A real .xlsx produced by a real writer (see tests/fixtures/e2e/source.xlsx's own commit,
// "add real-producer E2E fixtures, cross-checked with calamine") rather than the oracle's
// own utils.* writer — this is the one case in this file that exercises reader.rs's
// SHARED-STRINGS path (xl/sharedStrings.xml + the "s" branch of xlsx_parse_cell): every
// case above built via buildXlsxBytes uses the oracle's own aoa_to_sheet + XLSX.write,
// which (confirmed by inspecting the written sheet XML — see this file's header comment
// history) emits inline `t="str"` cells, never `t="s"` shared-string references, so without
// this case the shared-strings code path would be completely untested here.
runReadCaseBytes(
  'a real .xlsx fixture from an independent writer (exercises shared strings)',
  readFileSync(join(here, '../../tests/fixtures/e2e/source.xlsx'))
);

// Discovered gap #2: reader.rs never parses the worksheet's declared <dimension> tag at
// all — !ref always comes from the populated-cell bounding box instead. The oracle trusts
// <dimension> verbatim. These normally agree (every oracle-written file has them equal,
// which is why none of the cases above caught this), but a real Excel/LibreOffice file can
// legitimately declare a WIDER dimension (e.g. formatting-only cells past the last
// populated one) — confirmed here with a hand-built file that does exactly that (see
// buildDimensionWiderThanDataXlsxBytes above): oracle reports "!ref":"A1:E10" (the declared
// dimension), elixcee reports "!ref":"A1:B2" (the populated bounding box). Same disclosure
// rule as the empty-string-cell gap: a real, pre-existing reader.rs limitation (shared by
// read_workbook and read_workbook_from_bytes), out of scope to fix under this phase's
// "pure extraction, no behavior change" requirement, registered rather than silently
// worked around.
runReadCaseBytes(
  'declared <dimension> wider than populated cells (known reader.rs gap — registered UNSUPPORTED)',
  buildDimensionWiderThanDataXlsxBytes(),
  DIMENSION_WIDER_THAN_DATA_CASE_ID
);

// ---- summary / exit code (matches xlsx-utils.test.mjs's convention) ----
//
// UNSUPPORTED/*_DIVERGENCE are acceptable outcomes here in principle (classify.mjs's
// general contract), though this file's own case matrix never registers any — every case
// above is expected to MATCH on the projected (supported-fields-only) shape. Anything else
// (BUG, ORACLE_AMBIGUITY, NONDETERMINISTIC, UNCLASSIFIED) fails the run.
const ACCEPTABLE = new Set(['MATCH', 'INTENTIONAL_SAFETY_DIVERGENCE', 'INTENTIONAL_SECURITY_DIVERGENCE', 'UNSUPPORTED']);

const byApi = new Map();
const totals = new Map();
for (const r of results) {
  if (!byApi.has(r.api)) byApi.set(r.api, { other: [] });
  if (!ACCEPTABLE.has(r.verdict)) byApi.get(r.api).other.push({ label: r.label, verdict: r.verdict });
  totals.set(r.verdict, (totals.get(r.verdict) || 0) + 1);
}

const byApiVerdict = summarizeByApiAndVerdict(results);
console.log('\n=== read() differential summary (compat/differential/xlsx-read.test.mjs) ===');
console.log('Comparison scope: SheetNames, per-sheet !ref/!merges, per-cell {t,v} only — see');
console.log('this file\'s header comment for why the oracle\'s much richer WorkBook shape is');
console.log('projected down before comparing.');
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
  console.error('\nread() differential suite FAILED: at least one case is not an acceptable verdict.');
  process.exit(1);
}
console.log('\nread() differential suite passed: every case matches on its declared supported-field scope.');

// Self-check that projectWorkBook's field-selection logic itself is correct, independent
// of any oracle call — run via `node compat/differential/xlsx-read.test.mjs` alongside the
// cases above (this file IS the runnable check, like every other file in this directory).
{
  const projected = projectSheet({
    A1: { t: 's', v: 'hi', w: 'hi', h: 'hi' },
    '!ref': 'A1:A1',
    '!merges': [{ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }],
    '!cols': [{ hidden: true }],
  });
  assert.deepEqual(projected, {
    A1: { t: 's', v: 'hi' },
    '!ref': 'A1:A1',
    '!merges': [{ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }],
  });
  assert.equal('!cols' in projected, false);
}
