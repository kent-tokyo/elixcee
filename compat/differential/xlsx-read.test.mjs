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
// or could parse without a much larger scope than this package's read() (see
// crates/elixcee-wasm/src/lib.rs's `read_workbook` doc comment for the current exact
// list — as of this file's own read-item 1-4 work: no formatted `.w`/`.h` text, no
// date-typed cells (`t:'d'`) — both still need `styles.xml` number-format parsing, see
// this file's read-item-6 section). Confirmed empirically (not assumed) by writing a
// plain 3x3 aoa_to_sheet workbook through the oracle and reading it back — every cell
// came back with `.w` (and strings also got `.h`) even with no formatting ever applied,
// and the WorkBook itself carried Styles/SSF/Themes objects derived purely from the
// oracle's own default styles.xml.
//
// classify.mjs's own doc comment says callers must normalize before calling classify() —
// normalize.mjs does that for type-tagging (NaN/undefined/-0/Date/...). This file adds one
// more normalization step specific to `read()`: projectWorkBook() strips both sides down
// to exactly the fields @elixcee/xlsx's read() advertises support for (SheetNames, and per
// sheet: !ref, !merges, !rows, !cols, and each cell's {t, v, f}). This is a single,
// up-front, documented scope boundary — not a per-case escape hatch like
// UNSUPPORTED_ALLOWLIST — so if a genuinely supported field ever diverges, it still
// surfaces as UNCLASSIFIED/BUG exactly as classify() intends; only fields this package
// never claimed to produce (.w, .h, t:'d', ...) are excluded, and excluded identically
// from both sides so the comparison stays apples-to-apples rather than favoring elixcee's
// narrower shape.
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
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

// !rows/!cols: projected down to just the `hidden` flag on each present (non-hole) slot,
// on BOTH sides — the oracle's own !cols entries also carry a `width` (and !rows an `hpx`/
// `hpt`) computed from column-width/row-height metadata reader.rs doesn't parse and
// read()'s doc comment never promised (confirmed live: `{hidden:true,width:null}` vs this
// package's `{hidden:true}` — a field neither side is being compared on, not a false
// MATCH). Real array holes (a non-hidden row/col slot) are preserved as holes, not filled
// with `undefined`/`null` — normalize.mjs distinguishes a hole from an explicit value, and
// so does the real oracle's own output (confirmed live via `in`).
function projectRowsOrCols(arr) {
  if (arr == null) return undefined;
  const out = new Array(arr.length);
  for (let i = 0; i < arr.length; i++) {
    if (i in arr) out[i] = { hidden: !!arr[i].hidden };
  }
  return out;
}

function projectSheet(ws) {
  if (ws == null) return null;
  const out = {};
  for (const key of Object.keys(ws)) {
    if (key === '!ref') {
      out['!ref'] = ws['!ref'];
    } else if (key === '!merges') {
      out['!merges'] = ws['!merges'].map((m) => ({ s: { r: m.s.r, c: m.s.c }, e: { r: m.e.r, c: m.e.c } }));
    } else if (key === '!rows') {
      out['!rows'] = projectRowsOrCols(ws['!rows']);
    } else if (key === '!cols') {
      out['!cols'] = projectRowsOrCols(ws['!cols']);
    } else if (CELL_REF_RE.test(key)) {
      // .f/.w/.z (Milestone read-item 4/6) are safe to compare unconditionally: every
      // case that doesn't exercise them has the field `undefined` on BOTH sides (a plain
      // property access on a key neither side's cell object has for that case), so this
      // can never manufacture a false MATCH/mismatch for a case that isn't testing them.
      // .w in particular was verified live across every cell type already exercised
      // below (string/number/boolean, including the boundary-numeric-values case) before
      // being widened in here — see this file's read-item-6 section.
      out[key] = { t: ws[key].t, v: ws[key].v, f: ws[key].f, w: ws[key].w, z: ws[key].z };
    }
    // Every other key (!fullref, !type, ...) is out of this MVP's scope — silently dropped
    // from BOTH sides, not just elixcee's, so its absence can never look like a false
    // MATCH for a field neither side is being compared on.
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
// see the two case groups below). `readOpts`, when given, is passed to BOTH sides'
// read() (e.g. `{cellStyles: true}` for the !rows/!cols cases — see this file's own
// read-item-3 section for why that option matters).
function runReadCaseBytes(label, bytes, unsupportedCaseId, readOpts) {
  const oracleVal = invokeRead((b) => XLSX.read(b, Object.assign({ type: 'buffer' }, readOpts)), bytes);
  const elixceeVal = invokeRead((b) => elixcee.read(b, readOpts), bytes);
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

function runReadCase(label, sheets, unsupportedCaseId, readOpts) {
  return runReadCaseBytes(label, buildXlsxBytes(sheets), unsupportedCaseId, readOpts);
}

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

// ---- FIXED gap: empty-string cell values ----
//
// Found by this test file (originally registered UNSUPPORTED, see git history), now
// fixed: the oracle's writer emits an empty-string aoa cell as a real
// `<c r="B1" t="str"><v></v></c>` (confirmed by inspecting the actual written sheet1.xml)
// — a self-closing-content `<v>` element with ZERO characters between its open and close
// tags. `reader.rs`'s xlsx_sheet_cells now routes the empty string through the same
// xlsx_parse_cell used for the non-empty path on `</v>` when no Ev::Text ever fired for
// it, so `{t:"s", v:""}` is recorded like the oracle instead of the cell being silently
// absent. Shared by read_workbook and read_workbook_from_bytes. No longer registered in
// classify.mjs's UNSUPPORTED_ALLOWLIST — this is a plain MATCH case now.
runReadCase('empty-string cell value (formerly a reader.rs gap — now fixed)', [
  ['S1', [['before', '', 'after']]],
]);

// A formula (.f) roundtrip case (Milestone read-item 4) — reader.rs now captures
// per-cell <f>...</f> text; the oracle's own aoa writer accepts a plain cell object
// ({t,v,f}) verbatim in an aoa slot (sheet_add_aoa's "caller-supplied full cell object"
// branch — confirmed live it writes an independent <f> per cell, never a shared formula,
// so a literal capture-the-inline-text approach is enough for every fixture this suite
// builds).
runReadCase('formula cells (.f roundtrip)', [
  ['S1', [
    [1, 2, { t: 'n', v: 3, f: 'SUM(A1:B1)' }],
    [4, 5, { t: 'n', v: 9, f: 'SUM(A2:B2)' }],
  ]],
]);

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

// ---- FIXED gap #2: <dimension> parsing (Milestone read-item 2) ----
//
// reader.rs previously never parsed the worksheet's declared <dimension> tag at all —
// !ref always came from the populated-cell bounding box instead. It now does, replicating
// the oracle's own quirks exactly (a colon-less single-cell ref like "A1" is NOT trusted,
// matching the oracle's own dimregex; a reversed range is rejected too — see
// parse_dimension_ref's doc comment). A real Excel/LibreOffice file can legitimately
// declare a WIDER dimension than its populated cells (e.g. formatting-only cells past the
// last populated one) — exercised here with a hand-built file that does exactly that (see
// buildDimensionWiderThanDataXlsxBytes above). No longer registered in classify.mjs's
// UNSUPPORTED_ALLOWLIST — this is a plain MATCH case now.
runReadCaseBytes(
  'declared <dimension> wider than populated cells (formerly a reader.rs gap — now fixed)',
  buildDimensionWiderThanDataXlsxBytes()
);

// ---- read-item 3: !rows/!cols (requires opts.cellStyles, matching the oracle) ----
//
// Confirmed live against the real oracle (not assumed): XLSX.read() never returns
// !rows/!cols AT ALL — even for a file with genuinely hidden rows/columns — unless the
// caller also passes {cellStyles: true} (see compat/node_modules/xlsx/xlsx.js's
// parse_ws_xml_cols/parse_ws_xml_data cellStyles guards). packages/xlsx's read() now
// mirrors that gate exactly (see ./internal/read-shape.cjs) rather than always surfacing
// reader.rs's already-parsed hidden-row/col data, which would diverge from the oracle's
// own default-opts behavior. Two cases: with the option (both sides project down to just
// the `hidden` flag per row/col slot — see projectRowsOrCols's doc comment for why width/
// height metadata is excluded), and without it (regression guard that the gate actually
// suppresses !rows/!cols by default, not just that read() runs).
function buildHiddenRowsColsXlsxBytes() {
  const ws = U.aoa_to_sheet([[1, 2], [3, 4], [5, 6]]);
  ws['!rows'] = [];
  ws['!rows'][1] = { hidden: true };
  ws['!cols'] = [];
  ws['!cols'][0] = { hidden: true };
  const wb = U.book_new();
  U.book_append_sheet(wb, ws, 'S1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

runReadCaseBytes(
  'hidden rows and columns with opts.cellStyles: true (!rows/!cols)',
  buildHiddenRowsColsXlsxBytes(),
  undefined,
  { cellStyles: true }
);

runReadCaseBytes(
  'hidden rows and columns WITHOUT opts.cellStyles (must not surface !rows/!cols by default)',
  buildHiddenRowsColsXlsxBytes()
);

// ---- read-item 6: .w / .z / date-typed cells ────────────────────────────
//
// reader.rs now parses xl/styles.xml (numFmts + cellXfs) and the workbook's date1904
// flag; the JS layer (read-shape.cjs) resolves that into .w (always computed — confirmed
// live the oracle emits it unconditionally for every cell), .z (gated behind
// opts.cellNF/opts.cellStyles, always a resolved format STRING even "General"), and
// date-typed cells (gated behind opts.cellDates AND a date-like resolved format — the
// underlying serial value is unaffected either way) via the real `ssf` engine already
// verified byte-identical to the oracle's own across 1831 cases
// (compat/differential/ssf-format.test.mjs). See read-shape.cjs's own doc comment for the
// exact contract, confirmed live against the oracle, including a genuine oracle
// inconsistency this reproduces on purpose: a date1904 workbook's .w shifts by the
// 1462-day offset but its cellDates .v Date object does not (datenum.cjs's numdate).
//
// projectSheet was widened above to compare .w/.z unconditionally (verified first,
// per-field, across every cell type the 14 cases above already exercise — string,
// number, boolean, including the boundary-numeric-values case — before being widened,
// per this file's read-item-6 section in its own header comment).
function buildNumberFormatXlsxBytes() {
  const ws = U.aoa_to_sheet([[42, 3.14159, 45444]]);
  U.cell_set_number_format(ws.A1, '0.00'); // built-in numFmtId (2)
  U.cell_set_number_format(ws.C1, '0.00"kg"'); // custom numFmtId (164+, via <numFmts>)
  const wb = U.book_new();
  U.book_append_sheet(wb, ws, 'S1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

runReadCaseBytes(
  'number-format resolution (.w/.z) with opts.cellNF: true — built-in and custom formats',
  buildNumberFormatXlsxBytes(),
  undefined,
  { cellNF: true }
);

runReadCase(
  'date-typed cells with opts.cellDates: true (date and date-time values)',
  [['S1', [[new Date(2024, 0, 15, 13, 45, 30)], [new Date(2020, 11, 31)]]]],
  undefined,
  { cellDates: true }
);

runReadCase(
  'the same date-formatted cells WITHOUT opts.cellDates stay numeric (t:"n"), not "d"',
  [['S1', [[new Date(2024, 0, 15)]]]]
);

// A real date1904 workbook — the genuine oracle inconsistency (.w shifts, cellDates .v
// does not) called out in this section's header comment, exercised end-to-end rather
// than just asserted about.
function buildDate1904XlsxBytes() {
  const ws = U.aoa_to_sheet([[new Date(2024, 0, 15)]]);
  const wb = { SheetNames: ['S1'], Sheets: { S1: ws }, Workbook: { WBProps: { date1904: true } } };
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

runReadCaseBytes(
  'a date1904 workbook with opts.cellDates + opts.cellNF: true',
  buildDate1904XlsxBytes(),
  undefined,
  { cellDates: true, cellNF: true }
);

// ---- read-item 5: browser export condition dispatch ----
//
// Confirms LIVE — not just "should work" — that package.json's "browser" condition under
// exports["."] actually routes to the inlined-bytes/initSync WASM artifact
// (packages/xlsx/src/internal/wasm/elixcee_wasm.browser.mjs), not the Node glue. "browser"
// isn't a condition Node activates by default (only bundlers opt into it), so the only
// faithful way to exercise it without installing a bundler this repo doesn't otherwise
// need is `node --conditions=browser` — spawned here as a real subprocess reading real
// oracle-written bytes through '@elixcee/xlsx' via Node's self-referencing package
// resolution (a package resolving its own name back through its own `exports` map when
// run from inside that package's directory). `import.meta.resolve` inside the subprocess
// reports the ACTUAL resolved module URL directly — the unambiguous way to prove which
// file the "read" being called came from, rather than inferring it from behavior (which
// would still "pass" even if this silently fell through to the Node entry point, since
// both produce identical output on valid input).
{
  const bytes = buildXlsxBytes([
    ['S1', [[1, 2, { t: 'n', v: 3, f: 'SUM(A1:B1)' }], ['x', '']]],
  ]);
  const b64 = Buffer.from(bytes).toString('base64');
  const script =
    "const resolved = import.meta.resolve('@elixcee/xlsx');\n" +
    "import { read } from '@elixcee/xlsx';\n" +
    `const wb = read(${JSON.stringify(b64)}, { type: 'base64' });\n` +
    "process.stdout.write(JSON.stringify({ wb, resolved }));\n";
  const browserOut = JSON.parse(
    execFileSync(process.execPath, ['--conditions=browser', '--input-type=module', '-e', script], {
      cwd: join(here, '../../packages/xlsx'),
      encoding: 'utf8',
    })
  );
  const nodeOut = elixcee.read(bytes);
  assert.ok(
    browserOut.resolved.endsWith('/src/index.browser.mjs'),
    `expected --conditions=browser to resolve '@elixcee/xlsx' to index.browser.mjs, got: ${browserOut.resolved}`
  );
  assert.deepEqual(browserOut.wb, nodeOut, 'browser-condition read() must match the Node read() on the same bytes');
  console.log(
    'OK  read: "browser" export condition resolves to index.browser.mjs (the inlined-bytes WASM artifact) and matches the Node path'
  );
}

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
console.log('Comparison scope: SheetNames, per-sheet !ref/!merges/!rows/!cols, per-cell {t,v,f} — see');
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
    A1: { t: 's', v: 'hi', w: 'hi', h: '<b>hi</b>' },
    B1: { t: 'n', v: 3, f: 'SUM(A1:A1)', w: '3', z: 'General' },
    '!ref': 'A1:B1',
    '!merges': [{ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }],
    '!cols': [{ hidden: true, width: 12 }],
  });
  assert.deepEqual(projected, {
    A1: { t: 's', v: 'hi', f: undefined, w: 'hi', z: undefined },
    B1: { t: 'n', v: 3, f: 'SUM(A1:A1)', w: '3', z: 'General' },
    '!ref': 'A1:B1',
    '!merges': [{ s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }],
    '!cols': [{ hidden: true }],
  });
  // .h (rich-HTML text — never claimed as supported, out of this package's scope
  // entirely, unlike .w/.z which item 6 now computes) stays excluded from the per-cell
  // projection even though the field-by-field assert.deepEqual above wouldn't itself
  // catch a stray extra key (deepEqual on A1 only checks the keys it lists) — assert
  // that explicitly.
  assert.equal('h' in projected.A1, false);

  // A hole (non-hidden slot) in !rows/!cols must stay a real hole after projection, not
  // become an explicit `undefined`/`null` entry — normalize.mjs (and the real oracle's
  // own output) treats those as distinct.
  const withHole = projectSheet({ '!rows': [{ hidden: true }, , { hidden: false }] });
  assert.equal(0 in withHole['!rows'], true);
  assert.equal(1 in withHole['!rows'], false);
  assert.deepEqual(withHole['!rows'][2], { hidden: false });
}
