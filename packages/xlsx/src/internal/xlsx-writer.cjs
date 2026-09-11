'use strict';

// write()/writeFile()/writeFileSync() — WorkBook object -> a real, Excel-openable
// .xlsx file's ZIP entries. No Rust/WASM bridge (unlike read()): OOXML *writing* is pure
// XML/ZIP generation, nothing this package's own hand-rolled Rust reader is needed for —
// see docs/xlsx-architecture.md, which never planned a write-side bridge either.
//
// Output is deliberately constrained to shapes `src/reader.rs` (elixcee's own reader,
// vendored into this package as the WASM bridge `read()` uses) already parses — verified
// by reading reader.rs directly, not assumed: inline strings (`t="inlineStr"`, not shared
// strings — simpler, equally valid OOXML, and reader.rs confirmed to support both forms),
// numeric cells with no `t` attribute, `t="b"` booleans, `<f>`/`<v>` formula pairs,
// `<mergeCell>`, `<row hidden="1">`/`<col hidden="1">`, and `<dimension>`. This is what
// makes "own write -> own read" a meaningful round-trip test rather than two independently
// -guessed formats that happen to both claim OOXML compliance.
//
// bookType: 'xlsx' or 'xlsm' with an opaque !vbaProject (see write() in index.cjs for the explicit-unsupported-error
// contract for anything else) — no ODS, no legacy .xls, no CSV/HTML output formats.

const { checkRangeSize } = require('./range-guard.cjs');
const { safeDecodeRange } = require('./safe-decode-range.cjs');
const { datenum } = require('./datenum.cjs');

// Same algorithm as index.cjs's own encode_cell/decode_cell (0-based {r, c} <-> "A1"),
// duplicated rather than imported: index.cjs will itself require this module for write()
// (see below), so importing the other direction would be circular. Kept in exact lockstep
// with index.cjs's own copy — both are pure, stable, and already differential-tested
// against the oracle (compat/differential/xlsx-utils.test.mjs) via index.cjs's exports.
function encodeCell(cell) {
  let col = cell.c + 1;
  let s = '';
  for (; col; col = ((col - 1) / 26) | 0) {
    s = String.fromCharCode(((col - 1) % 26) + 65) + s;
  }
  return s + (cell.r + 1);
}

function decodeCell(cstr) {
  let r = 0;
  let c = 0;
  for (let i = 0; i < cstr.length; ++i) {
    const cc = cstr.charCodeAt(i);
    if (cc >= 48 && cc <= 57) r = 10 * r + (cc - 48);
    else if (cc >= 65 && cc <= 90) c = 26 * c + (cc - 64);
  }
  return { c: c - 1, r: r - 1 };
}

const ELIXCEE_UNSUPPORTED_CELL_TYPE = 'ELIXCEE_UNSUPPORTED_CELL_TYPE';
const ELIXCEE_UNSUPPORTED_SHEET_SHAPE = 'ELIXCEE_UNSUPPORTED_SHEET_SHAPE';

function unsupported(code, message) {
  const err = new Error(message);
  err.code = code;
  return err;
}

// ---- XML text/attribute escaping ----
//
// & < > for text; & < > " additionally for attribute values (a `'` inside a double-quoted
// attribute needs no escape, matching every real XML writer). XML 1.0 forbids most C0
// control bytes outright (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F) — tab/LF/CR (0x09/0x0A/0x0D)
// are the only ones a well-formed document may contain unescaped. Real Excel/SheetJS
// writers re-encode a forbidden byte as a literal `_xHHHH_` token a reader then has to
// specially unescape; reader.rs implements no such unescaping (grep-confirmed), so
// replicating that convention here would produce a value THIS package's own read() could
// never recover correctly. Stripped instead — a disclosed, narrow scope limit (control
// characters in cell text are vanishingly rare in real spreadsheets), not silent data
// corruption of anything a real-world caller is likely to pass.
const FORBIDDEN_XML_CHARS = /[\x00-\x08\x0b\x0c\x0e-\x1f]/g;

function xmlText(s) {
  return String(s)
    .replace(FORBIDDEN_XML_CHARS, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function xmlAttr(s) {
  return xmlText(s).replace(/"/g, '&quot;');
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

// ---- [Content_Types].xml / _rels ----

function buildContentTypes(sheetCount, commentSheets, chartSheets, tableSheets, hasVba = false) {
  const comments = new Set(commentSheets || []);
  const charts = new Set(chartSheets || []);
  const tables = new Set(tableSheets || []);
  const overrides = [
    `<Override PartName="/xl/workbook.xml" ContentType="${hasVba ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'}"/>`,
  ];
  for (let i = 1; i <= sheetCount; i++) {
    overrides.push(
      `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    );
  }
  if (comments.size) {
    overrides.push('<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>');
    for (const i of [...comments].sort((a, b) => a - b)) {
      overrides.push(
        `<Override PartName="/xl/comments${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`
      );
    }
  }
  if (charts.size) {
    const chartParts = [...charts].flatMap((entry) => typeof entry === 'number' ? [{ drawing: String(entry), charts: [String(entry)] }] : [entry]);
    for (const { drawing, charts: parts } of chartParts.sort((a, b) => Number(a.drawing) - Number(b.drawing))) {
      overrides.push(
        `<Override PartName="/xl/drawings/drawing${drawing}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`,
        ...parts.map((part) => `<Override PartName="/xl/charts/chart${part}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`)
      );
    }
  }
  if (tables.size) for (const i of [...tables].sort()) overrides.push(`<Override PartName="/xl/tables/table${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
  if (hasVba) overrides.push('<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>');
  overrides.push(
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
  );
  return (
    XML_DECL +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    overrides.join('') +
    '</Types>'
  );
}

function buildRootRels() {
  return (
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
    '</Relationships>'
  );
}

function buildWorkbookRels(sheetCount, hasVba = false) {
  const rels = [];
  for (let i = 1; i <= sheetCount; i++) {
    rels.push(
      `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`
    );
  }
  rels.push(
    `<Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
  );
  if (hasVba) rels.push(`<Relationship Id="rId${sheetCount + 2}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`);
  return (
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    rels.join('') +
    '</Relationships>'
  );
}

// ---- docProps ----

function buildCoreXml() {
  return (
    XML_DECL +
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' +
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' +
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:creator>@elixcee/xlsx</dc:creator>' +
    '</cp:coreProperties>'
  );
}

function buildAppXml(sheetNames) {
  const titles = sheetNames.map((n) => `<vt:lpstr>${xmlText(n)}</vt:lpstr>`).join('');
  return (
    XML_DECL +
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' +
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
    '<Application>@elixcee/xlsx</Application>' +
    `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${sheetNames.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
    `<TitlesOfParts><vt:vector size="${sheetNames.length}" baseType="lpstr">${titles}</vt:vector></TitlesOfParts>` +
    '</Properties>'
  );
}

// ---- xl/workbook.xml ----
//
// Sheet visibility mirrors the real oracle's own `Workbook.Sheets[i].Hidden` convention
// (0/1/2 — see index.cjs's book_set_sheet_visibility and consts.SHEET_VISIBLE/HIDDEN/
// VERY_HIDDEN) onto OOXML's own `state="hidden"|"veryHidden"` sheet attribute (omitted
// entirely for visible, matching real Excel/SheetJS writer output — there is no
// `state="visible"` in the wild).
function sheetStateAttr(hidden) {
  if (hidden === 1) return ' state="hidden"';
  if (hidden === 2) return ' state="veryHidden"';
  return '';
}

function buildDefinedNamesXml(workbookMeta) {
  const names = workbookMeta && Array.isArray(workbookMeta.Names) ? workbookMeta.Names : [];
  const entries = names.map((entry) => {
    if (!entry || typeof entry !== 'object') return '';
    const name = String(entry.Name ?? entry.name ?? '');
    const ref = String(entry.Ref ?? entry.ref ?? '');
    if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(name) || !ref || /[\[\]]/.test(ref) || (!ref.includes('!') && !Number.isInteger(entry.Sheet))) return '';
    const attrs = [`name="${xmlAttr(name)}"`];
    if (Number.isInteger(entry.Sheet) && entry.Sheet >= 0) attrs.push(`localSheetId="${entry.Sheet}"`);
    if (entry.Hidden) attrs.push('hidden="1"');
    return `<definedName ${attrs.join(' ')}>${xmlText(ref)}</definedName>`;
  }).filter(Boolean);
  return entries.length ? `<definedNames>${entries.join('')}</definedNames>` : '';
}

function buildWorkbookXml(sheetNames, workbookMeta) {
  const wbSheets = (workbookMeta && workbookMeta.Sheets) || [];
  const sheetTags = sheetNames
    .map((name, i) => {
      const hidden = wbSheets[i] && wbSheets[i].Hidden;
      return `<sheet name="${xmlAttr(name)}" sheetId="${i + 1}" r:id="rId${i + 1}"${sheetStateAttr(hidden)}/>`;
    })
    .join('');
  const date1904 = workbookMeta && workbookMeta.WBProps && workbookMeta.WBProps.date1904;
  const workbookPr = date1904 ? '<workbookPr date1904="1"/>' : '';
  return (
    XML_DECL +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    workbookPr +
    `<sheets>${sheetTags}</sheets>` +
    buildDefinedNamesXml(workbookMeta) +
    '</workbook>'
  );
}

// ---- number formats / styles.xml ----
//
// A deliberately small built-in numFmtId table — real Excel's full built-in range is
// 0-163, this covers the common "basic number formats" this phase's own scope names
// (General, integer, 2-decimal, percent, and the date/time formats aoa_to_sheet's own
// Date handling already produces via ssf-adapter.cjs's default 'm/d/yy'). Anything else
// gets a custom numFmtId (164+, incrementing), written into <numFmts> verbatim — exactly
// how real Excel represents a format string with no built-in id, and how reader.rs's own
// `BufferWorkbook::number_formats` expects to find one.
const BUILTIN_NUMFMT_IDS = new Map([
  ['General', 0],
  ['0', 1],
  ['0.00', 2],
  ['#,##0', 3],
  ['#,##0.00', 4],
  ['0%', 9],
  ['0.00%', 10],
  ['m/d/yy', 14],
  ['d-mmm-yy', 15],
  ['d-mmm', 16],
  ['mmm-yy', 17],
  ['h:mm AM/PM', 18],
  ['h:mm:ss AM/PM', 19],
  ['h:mm', 20],
  ['h:mm:ss', 21],
  ['m/d/yy h:mm', 22],
]);

// Tracks the format-code -> cellXf-index mapping for one write() call. Index 0 is always
// the default General/no-style xf (real Excel's own convention, and what an un-styled
// `<c>` with no `s` attribute implicitly means, so it never needs to be referenced
// explicitly either — allocate() is only ever called for a cell that actually has a `.z`).
function createStyleTable() {
  const numFmtIdByCode = new Map(); // format code -> numFmtId
  const customNumFmts = []; // [{id, code}] for <numFmts>
  const fonts = [{ bold: false, italic: false, underline: false, name: 'Calibri', size: 11 }];
  const fills = [{ pattern: 'none' }];
  const borders = [{ sides: {} }];
  const dxfs = [];
  const cellXfs = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: null }];
  const fontIds = new Map([['default', 0]]);
  const fillIds = new Map([['none', 0]]);
  const borderIds = new Map([['none', 0]]);
  const cellXfIds = new Map([['0|0|0|', 0]]);
  let nextCustomId = 164;

  function numFmtIdFor(code) {
    if (numFmtIdByCode.has(code)) return numFmtIdByCode.get(code);
    let id = BUILTIN_NUMFMT_IDS.get(code);
    if (id === undefined) {
      id = nextCustomId++;
      customNumFmts.push({ id, code });
    }
    numFmtIdByCode.set(code, id);
    return id;
  }

  function normalizeColor(color) {
    if (!color) return '';
    const raw = typeof color === 'string' ? color : color.rgb || color.RGB || '';
    const hex = String(raw).replace(/^#/, '').toUpperCase();
    return /^[0-9A-F]{6}$/.test(hex) ? `FF${hex}` : /^[0-9A-F]{8}$/.test(hex) ? hex : '';
  }

  function fontIdFor(style) {
    const font = style?.font || {};
    const key = JSON.stringify({ bold: !!font.bold, italic: !!font.italic, underline: !!font.underline, name: font.name || 'Calibri', size: Number(font.sz || font.size || 11), color: normalizeColor(font.color) });
    if (fontIds.has(key)) return fontIds.get(key);
    const id = fonts.length;
    fonts.push({ bold: !!font.bold, italic: !!font.italic, underline: !!font.underline, name: font.name || 'Calibri', size: Number(font.sz || font.size || 11), color: normalizeColor(font.color) });
    fontIds.set(key, id);
    return id;
  }

  function fillIdFor(style) {
    const fill = style?.fill || {};
    const color = normalizeColor(fill.fgColor || fill.fg || fill.color);
    if (!color) return 0;
    const key = `solid|${color}`;
    if (fillIds.has(key)) return fillIds.get(key);
    const id = fills.length;
    fills.push({ pattern: 'solid', color });
    fillIds.set(key, id);
    return id;
  }

  function borderIdFor(style) {
    const source = style?.border || {};
    const sides = {};
    for (const side of ['left', 'right', 'top', 'bottom']) {
      const value = source[side] || {};
      if (value.style) sides[side] = { style: String(value.style), color: normalizeColor(value.color) };
    }
    const key = JSON.stringify(sides);
    if (!Object.keys(sides).length) return 0;
    if (borderIds.has(key)) return borderIds.get(key);
    const id = borders.length;
    borders.push({ sides });
    borderIds.set(key, id);
    return id;
  }

  function alignmentFor(style) {
    const alignment = style?.alignment;
    if (!alignment) return null;
    const out = {};
    for (const key of ['horizontal', 'vertical', 'wrapText', 'textRotation', 'indent']) if (alignment[key] !== undefined) out[key] = alignment[key];
    return Object.keys(out).length ? out : null;
  }

  function dxfFor(style) {
    const font = style?.font || {};
    const fill = style?.fill || {};
    const fontColor = normalizeColor(font.color);
    const fillColor = normalizeColor(fill.fgColor || fill.fg || fill.color);
    const key = JSON.stringify({ bold: !!font.bold, italic: !!font.italic, underline: !!font.underline, fontColor, fillColor });
    const existing = dxfs.findIndex((entry) => entry.key === key);
    if (existing >= 0) return existing;
    const entry = { key, bold: !!font.bold, italic: !!font.italic, underline: !!font.underline, fontColor, fillColor };
    dxfs.push(entry);
    return dxfs.length - 1;
  }

  // Returns the 0-based cellXf index a cell's `s="N"` attribute should reference.
  function cellXfFor(code, style) {
    const numFmtId = !code || code === 'General' ? 0 : numFmtIdFor(code);
    const fontId = style ? fontIdFor(style) : 0;
    const fillId = style ? fillIdFor(style) : 0;
    const borderId = style ? borderIdFor(style) : 0;
    const alignment = style ? alignmentFor(style) : null;
    const alignmentKey = alignment ? JSON.stringify(alignment) : '';
    const key = `${numFmtId}|${fontId}|${fillId}|${borderId}|${alignmentKey}`;
    if (cellXfIds.has(key)) return cellXfIds.get(key);
    const idx = cellXfs.length;
    cellXfs.push({ numFmtId, fontId, fillId, borderId, alignment });
    cellXfIds.set(key, idx);
    return idx;
  }

  function colorXml(color) { return color ? `<color rgb="${color}"/>` : ''; }
  function sideXml(side, value) { const item = value || {}; return `<${side}${item.style ? ` style="${xmlAttr(item.style)}"` : ''}>${colorXml(item.color)}</${side}>`; }

  function build() {
    const numFmtsXml = customNumFmts.length
      ? `<numFmts count="${customNumFmts.length}">${customNumFmts
          .map((f) => `<numFmt numFmtId="${f.id}" formatCode="${xmlAttr(f.code)}"/>`)
          .join('')}</numFmts>`
      : '';
    const fontXml = fonts.map((font) => `<font>${font.bold ? '<b/>' : ''}${font.italic ? '<i/>' : ''}${font.underline ? '<u/>' : ''}<sz val="${font.size}"/><name val="${xmlAttr(font.name)}"/>${colorXml(font.color)}</font>`).join('');
    const fillXml = fills.map((fill) => fill.pattern === 'solid' ? `<fill><patternFill patternType="solid"><fgColor rgb="${fill.color}"/><bgColor indexed="64"/></patternFill></fill>` : '<fill><patternFill patternType="none"/></fill>').join('');
    const borderXml = borders.map((border) => `<border>${sideXml('left', border.sides.left)}${sideXml('right', border.sides.right)}${sideXml('top', border.sides.top)}${sideXml('bottom', border.sides.bottom)}<diagonal/></border>`).join('');
    const dxfXml = dxfs.map((dxf) => { const font = dxf.bold || dxf.italic || dxf.underline || dxf.fontColor ? `<font>${dxf.bold ? '<b/>' : ''}${dxf.italic ? '<i/>' : ''}${dxf.underline ? '<u/>' : ''}${dxf.fontColor ? `<color rgb="${dxf.fontColor}"/>` : ''}</font>` : ''; const fill = dxf.fillColor ? `<fill><patternFill patternType="solid"><fgColor rgb="${dxf.fillColor}"/><bgColor indexed="64"/></patternFill></fill>` : ''; return `<dxf>${font}${fill}</dxf>`; }).join('');
    const xfs = cellXfs.map((xf) => { const alignment = xf.alignment ? `<alignment${xf.alignment.horizontal ? ` horizontal="${xmlAttr(xf.alignment.horizontal)}"` : ''}${xf.alignment.vertical ? ` vertical="${xmlAttr(xf.alignment.vertical)}"` : ''}${xf.alignment.wrapText !== undefined ? ` wrapText="${xf.alignment.wrapText ? 1 : 0}"` : ''}${xf.alignment.textRotation !== undefined ? ` textRotation="${Number(xf.alignment.textRotation)}"` : ''}${xf.alignment.indent !== undefined ? ` indent="${Number(xf.alignment.indent)}"` : ''}/>` : ''; return `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0"${alignment ? ' applyAlignment="1"' : ''}>${alignment}</xf>`; }).join('');
    const dxfsXml = dxfs.length ? `<dxfs count="${dxfs.length}">${dxfXml}</dxfs>` : '';
    return (
      XML_DECL +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      numFmtsXml +
      `<fonts count="${fonts.length}">${fontXml}</fonts>` +
      `<fills count="${fills.length}">${fillXml}</fills>` +
      `<borders count="${borders.length}">${borderXml}</borders>` +
      dxfsXml +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${cellXfs.length}">${xfs}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleMedium9"/>' +
      '</styleSheet>'
    );
  }

  return { cellXfFor, dxfFor, build };
}

// ---- worksheet body: cells, rows, merges, hidden rows/cols, dimension ----

const CELL_REF_RE = /^[A-Z]+[0-9]+$/;

// Flattens a WorkSheet's populated cells into `[{r, c, cell}]` (0-based), regardless of
// whether the caller used sparse (cell-ref-keyed object, e.g. aoa_to_sheet's default) or
// dense (`Array.isArray(ws)`, e.g. sheet_add_aoa's dense mode) storage — one shape the row
// -grouping logic below builds `<row>`/`<c>` XML from either way. A worksheet in neither
// shape (not an object, not an array) is rejected explicitly rather than silently
// producing an empty sheet — this project's own rule for any input it doesn't recognize.
function collectCells(ws) {
  const out = [];
  if (Array.isArray(ws)) {
    for (let r = 0; r < ws.length; r++) {
      const row = ws[r];
      if (row == null) continue;
      if (!Array.isArray(row)) {
        throw unsupported(ELIXCEE_UNSUPPORTED_SHEET_SHAPE, `dense worksheet row ${r} is not an array`);
      }
      for (let c = 0; c < row.length; c++) {
        if (row[c] != null) out.push({ r, c, cell: row[c] });
      }
    }
    return out;
  }
  if (ws == null || typeof ws !== 'object') {
    throw unsupported(ELIXCEE_UNSUPPORTED_SHEET_SHAPE, 'worksheet must be an object (sparse) or an array (dense)');
  }
  for (const key of Object.keys(ws)) {
    if (key.charCodeAt(0) === 33 /* '!' */) continue; // !ref, !merges, !cols, !rows, ...
    if (!CELL_REF_RE.test(key)) continue;
    const { r, c } = decodeCell(key);
    out.push({ r, c, cell: ws[key] });
  }
  return out;
}

// One `<c>` element, or '' for a cell with nothing to write (a `{t:'z'}` stub, or any
// cell whose `.v` is `undefined` and which carries no formula — real Excel omits `<c>`
// for a truly empty cell, and reader.rs's own `xlsx_sheet_cells` never records one that
// isn't there, so writing nothing here is a round-trippable choice, not a gap).
//
// Type dispatch order matters: a formula (`.f` present) is checked before value-type
// inference, since a formula cell's `.t` often reflects its *cached result's* type
// (matching read()'s own convention) rather than selecting a different XML shape — the
// cached value still needs a `t="str"`/`t="b"`/bare-numeric `<v>`, just inside the same
// `<f>`-bearing `<c>`.
function cellXml(ref, cell, styleTable) {
  if (cell == null) return '';
  const t = cell.t;
  const v = cell.v;

  if (cell.f) {
    const formula = String(cell.f).replace(/^=/, '');
    let inner = `<f>${xmlText(formula)}</f>`;
    // A missing cached value defaults to 0 rather than omitting <v> entirely — see this
    // module's top doc comment: reader.rs only records a cell when it sees a <v>/<is>
    // child, so a formula with no cached value would silently vanish from a subsequent
    // read(), breaking the "own write -> own read" round trip for no benefit (real Excel
    // recalculates on open regardless of what the cached value says).
    const cached = v === undefined ? 0 : v;
    let typeAttr = '';
    if (typeof cached === 'string') {
      typeAttr = ' t="str"';
      inner += `<v>${xmlText(cached)}</v>`;
    } else if (typeof cached === 'boolean') {
      typeAttr = ' t="b"';
      inner += `<v>${cached ? 1 : 0}</v>`;
    } else if (typeof cached === 'number' && isFinite(cached)) {
      inner += `<v>${cached}</v>`;
    } else {
      throw unsupported(
        ELIXCEE_UNSUPPORTED_CELL_TYPE,
        `formula cell ${ref} has an unsupported cached value (must be a string, boolean, or finite number)`
      );
    }
    const sAttr = cell.z || cell.s ? ` s="${styleTable.cellXfFor(cell.z, cell.s)}"` : '';
    return `<c r="${ref}"${typeAttr}${sAttr}>${inner}</c>`;
  }

  if (v instanceof Date || t === 'd') {
    const serial = v instanceof Date ? datenum(v) : v;
    if (typeof serial !== 'number' || !isFinite(serial)) {
      throw unsupported(ELIXCEE_UNSUPPORTED_CELL_TYPE, `date cell ${ref} has neither a Date nor a finite serial value`);
    }
    const sAttr = ` s="${styleTable.cellXfFor(cell.z || 'm/d/yy', cell.s)}"`;
    return `<c r="${ref}"${sAttr}><v>${serial}</v></c>`;
  }

  const sAttr = cell.z || cell.s ? ` s="${styleTable.cellXfFor(cell.z, cell.s)}"` : '';

  if (t === 'n' || (t === undefined && typeof v === 'number')) {
    if (typeof v !== 'number' || !isFinite(v)) {
      throw unsupported(ELIXCEE_UNSUPPORTED_CELL_TYPE, `numeric cell ${ref} value is not a finite number`);
    }
    return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
  }

  if (t === 'b' || (t === undefined && typeof v === 'boolean')) {
    return `<c r="${ref}" t="b"${sAttr}><v>${v ? 1 : 0}</v></c>`;
  }

  if (t === 's' || (t === undefined && typeof v === 'string')) {
    return `<c r="${ref}" t="inlineStr"${sAttr}><is><t xml:space="preserve">${xmlText(v)}</t></is></c>`;
  }

  if (t === 'z' || v === undefined) return '';

  throw unsupported(ELIXCEE_UNSUPPORTED_CELL_TYPE, `cell ${ref} has unsupported type '${t}'`);
}

// `ws['!rows']`/`ws['!cols']` are 0-based sparse arrays of `{hidden:true}|undefined` —
// read()'s own opts.cellStyles output shape (see internal/read-shape.cjs's
// expandHiddenIntervals), reproduced here in reverse: a hidden index becomes a `<row
// hidden="1">` attribute, or a run of hidden column indices becomes one `<col min max
// hidden="1">` element (real Excel's own interval-run representation, not one `<col>` per
// column).
function hiddenRowSet(rowsMeta) {
  const set = new Set();
  if (Array.isArray(rowsMeta)) {
    for (let i = 0; i < rowsMeta.length; i++) {
      if (rowsMeta[i] && rowsMeta[i].hidden) set.add(i);
    }
  }
  return set;
}

function buildColsXml(colsMeta) {
  if (!Array.isArray(colsMeta)) return '';
  const inner = [];
  for (let i = 0; i < colsMeta.length; i += 1) {
    const meta = colsMeta[i]; if (!meta) continue;
    const width = Number(meta.wch ?? meta.width);
    const widthAttr = Number.isFinite(width) && width >= 0.5 && width <= 255 ? ` width="${width}" customWidth="1"` : '';
    const hiddenAttr = meta.hidden ? ' hidden="1"' : '';
    if (widthAttr || hiddenAttr) inner.push(`<col min="${i + 1}" max="${i + 1}"${widthAttr || ' width="9.140625"'}${hiddenAttr}/>`);
  }
  if (!inner.length) return '';
  return `<cols>${inner}</cols>`;
}

function buildSheetViewsXml(ws) {
  const freeze = Array.isArray(ws) ? undefined : ws?.['!freezePane'];
  const rows = Number(freeze?.rows);
  const cols = Number(freeze?.cols);
  if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 0 || cols < 0 || (rows === 0 && cols === 0) || rows > 1048575 || cols > 16383) return '';
  const topLeftCell = encodeCell({ r: rows, c: cols });
  const activePane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
  const split = `${cols ? ` xSplit="${cols}"` : ''}${rows ? ` ySplit="${rows}"` : ''}`;
  return `<sheetViews><sheetView workbookViewId="0"><pane${split} topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/><selection pane="${activePane}" activeCell="${topLeftCell}" sqref="${topLeftCell}"/></sheetView></sheetViews>`;
}

function buildMergesXml(merges) {
  if (!Array.isArray(merges) || !merges.length) return '';
  const inner = merges
    .map((m) => `<mergeCell ref="${xmlAttr(encodeCell(m.s) + ':' + encodeCell(m.e))}"/>`)
    .join('');
  return `<mergeCells count="${merges.length}">${inner}</mergeCells>`;
}

function buildDataValidationsXml(validations) {
  if (!Array.isArray(validations) || !validations.length) return '';
  const allowedTypes = new Set(['list', 'whole', 'decimal', 'date', 'time', 'textLength', 'custom']);
  const entries = validations.map((validation) => {
    if (!validation || typeof validation !== 'object') return '';
    const type = String(validation.type || 'list');
    const ranges = Array.isArray(validation.sqref) ? validation.sqref : [validation.sqref];
    if (!allowedTypes.has(type) || !ranges.length || ranges.some((range) => typeof range !== 'string' || !range.trim())) return '';
    const attrs = [`type="${xmlAttr(type)}"`, `sqref="${xmlAttr(ranges.join(' '))}"`];
    if (validation.operator) attrs.push(`operator="${xmlAttr(validation.operator)}"`);
    if (validation.allowBlank !== false) attrs.push('allowBlank="1"');
    if (validation.showInputMessage) attrs.push('showInputMessage="1"');
    if (validation.showErrorMessage !== false) attrs.push('showErrorMessage="1"');
    if (validation.errorTitle) attrs.push(`errorTitle="${xmlAttr(validation.errorTitle)}"`);
    if (validation.error) attrs.push(`error="${xmlAttr(validation.error)}"`);
    const formula1 = validation.formula1 ?? validation.formula;
    const formula2 = validation.formula2;
    return `<dataValidation ${attrs.join(' ')}>${formula1 !== undefined ? `<formula1>${xmlText(formula1)}</formula1>` : ''}${formula2 !== undefined ? `<formula2>${xmlText(formula2)}</formula2>` : ''}</dataValidation>`;
  }).filter(Boolean);
  return entries.length ? `<dataValidations count="${entries.length}">${entries.join('')}</dataValidations>` : '';
}

function buildTableInfo(ws, table, sheetIndex, sheetName, tableNumber) {
  if (!table) return null;
  const ref = typeof table.ref === 'string' ? table.ref.trim() : '';
  let range;
  try { range = safeDecodeRange(ref); } catch { return null; }
  if (range.s.r >= range.e.r || range.s.c > range.e.c) return null;
  const displayName = String(table.displayName || table.name || `Table${sheetIndex}`).trim();
  if (!/^[A-Za-z_][A-Za-z0-9_.]{0,254}$/.test(displayName) || /^[A-Za-z]+[0-9]+$/.test(displayName)) return null;
  const headers = [];
  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const cell = ws[encodeCell({ r: range.s.r, c: col })];
    const label = String(cell?.v ?? '').trim() || `Column${col - range.s.c + 1}`;
    headers.push(`${label}`);
  }
  const projectedColumns = Array.isArray(table.columns) ? table.columns.map((column) => String(column?.name ?? '').trim()) : [];
  if (projectedColumns.length === headers.length && projectedColumns.every(Boolean)) for (let index = 0; index < headers.length; index += 1) headers[index] = projectedColumns[index];
  const columns = headers.map((label, index) => `<tableColumn id="${index + 1}" name="${xmlAttr(label)}"/>`).join('');
  const styleName = String(table.styleName || 'TableStyleMedium2');
  const safeStyle = /^[A-Za-z_][A-Za-z0-9_.-]{0,254}$/.test(styleName) ? styleName : 'TableStyleMedium2';
  let autoFilterRef = ref;
  if (typeof table.autoFilterRef === 'string' && table.autoFilterRef.trim()) {
    try {
      const candidate = table.autoFilterRef.trim();
      const filterRange = safeDecodeRange(candidate);
      if (filterRange.s.r <= filterRange.e.r && filterRange.s.c <= filterRange.e.c) autoFilterRef = candidate;
    } catch { /* retain the table range for malformed projected metadata */ }
  }
  // Keep the established single-table filename (`table1.xml`) for compatibility;
  // additional tables on the same sheet receive a deterministic suffix.
  const partName = tableNumber === 1 ? `${sheetIndex}` : `${sheetIndex}_${tableNumber}`;
  const tableId = tableNumber === 1 ? sheetIndex : sheetIndex * 1000 + tableNumber;
  return { ref, displayName, partName, xml: XML_DECL + `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${tableId}" name="${xmlAttr(displayName)}" displayName="${xmlAttr(displayName)}" ref="${xmlAttr(ref)}">${buildTableAutoFilterXml(autoFilterRef, table, headers.length)}<tableColumns count="${headers.length}">${columns}</tableColumns><tableStyleInfo name="${xmlAttr(safeStyle)}" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>` };
}

function buildTableInfos(ws, sheetIndex, sheetName) {
  if (Array.isArray(ws) || !Array.isArray(ws?.['!tables'])) return [];
  return ws['!tables'].map((table, index) => buildTableInfo(ws, table, sheetIndex, sheetName, index + 1)).filter(Boolean);
}

function buildTableAutoFilterXml(ref, table, width) {
  const columns = Array.isArray(table.autoFilterColumns) ? table.autoFilterColumns : [];
  const rendered = columns.map((column) => {
    const colId = Number(column?.colId);
    if (!Number.isInteger(colId) || colId < 0 || colId >= width) return '';
    const attrs = [`colId="${colId}"`];
    if (column.hiddenButton === true) attrs.push('hiddenButton="1"');
    if (column.showButton === false) attrs.push('showButton="0"');
    const criteria = column.criteria;
    if (!criteria || typeof criteria !== 'object') return '';
    if (criteria.kind === 'values' && Array.isArray(criteria.values) && criteria.values.length) {
      const values = criteria.values.slice(0, 1000).map((value) => `<filter val="${xmlAttr(String(value))}"/>`).join('');
      return `<filterColumn ${attrs.join(' ')}><filters>${values}</filters></filterColumn>`;
    }
    if (criteria.kind === 'blank') return `<filterColumn ${attrs.join(' ')}><filters blank="1"/></filterColumn>`;
    if (criteria.kind === 'custom' && criteria.op1 && criteria.val1 !== undefined) {
      const custom = [`<customFilter operator="${xmlAttr(String(criteria.op1))}" val="${xmlAttr(String(criteria.val1))}"/>`];
      if (criteria.op2 && criteria.val2 !== undefined) custom.push(`<customFilter operator="${xmlAttr(String(criteria.op2))}" val="${xmlAttr(String(criteria.val2))}"/>`);
      return `<filterColumn ${attrs.join(' ')}><customFilters${criteria.and === true ? ' and="1"' : ''}>${custom.join('')}</customFilters></filterColumn>`;
    }
    if (criteria.kind === 'top10' && Number.isFinite(Number(criteria.val))) {
      return `<filterColumn ${attrs.join(' ')}><top10 top="${criteria.top === false ? '0' : '1'}" percent="${criteria.percent === true ? '1' : '0'}" val="${Number(criteria.val)}"/></filterColumn>`;
    }
    return '';
  }).filter(Boolean).join('');
  return `<autoFilter ref="${xmlAttr(ref)}">${rendered}</autoFilter>`;
}

function buildConditionalFormattingXml(rules, styleTable) {
  if (!Array.isArray(rules) || !rules.length) return '';
  const allowedTypes = new Set(['cellIs', 'expression']);
  const allowedOperators = new Set(['equal', 'notEqual', 'lessThan', 'lessThanOrEqual', 'greaterThan', 'greaterThanOrEqual', 'between', 'notBetween']);
  return rules.map((rule) => {
    if (!rule || typeof rule !== 'object' || !allowedTypes.has(String(rule.type || 'cellIs'))) return '';
    const ranges = Array.isArray(rule.sqref) ? rule.sqref : [rule.sqref];
    if (!ranges.length || ranges.some((range) => typeof range !== 'string' || !range.trim())) return '';
    const formula = String(rule.formula ?? '').trim();
    if (!formula || formula.length > 4096 || /[\u0000-\u001f]/.test(formula)) return '';
    const operator = String(rule.operator || 'greaterThan');
    if (rule.type !== 'expression' && !allowedOperators.has(operator)) return '';
    const dxf = rule.dxf && typeof rule.dxf === 'object' ? rule.dxf : {};
    const dxfId = styleTable.dxfFor(dxf);
    const attrs = [`type="${xmlAttr(rule.type || 'cellIs')}"`, `priority="${Number.isInteger(rule.priority) && rule.priority > 0 ? rule.priority : 1}"`, `dxfId="${dxfId}"`];
    if (rule.stopIfTrue === true) attrs.push('stopIfTrue="1"');
    if (rule.type !== 'expression') attrs.push(`operator="${xmlAttr(operator)}"`);
    return `<conditionalFormatting sqref="${xmlAttr(ranges.join(' '))}"><cfRule ${attrs.join('')}><formula>${xmlText(formula.replace(/^=/, ''))}</formula>${rule.formula2 ? `<formula>${xmlText(String(rule.formula2).replace(/^=/, ''))}</formula>` : ''}</cfRule></conditionalFormatting>`;
  }).filter(Boolean).join('');
}

function buildHyperlinkInfo(cells) {
  const links = [];
  const relationships = [];
  for (const { r, c, cell } of cells) {
    if (!cell || !cell.l || !cell.l.Target) continue;
    const target = String(cell.l.Target);
    const ref = encodeCell({ r, c });
    const tooltip = cell.l.Tooltip ? ` tooltip="${xmlAttr(cell.l.Tooltip)}"` : '';
    if (target.charAt(0) === '#') {
      links.push(`<hyperlink ref="${ref}" location="${xmlAttr(target.slice(1))}"${tooltip}/>`);
    } else {
      const id = `rId${relationships.length + 1}`;
      relationships.push(
        `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${xmlAttr(target)}" TargetMode="External"/>`
      );
      links.push(`<hyperlink ref="${ref}" r:id="${id}"${tooltip}/>`);
    }
  }
  return { links, relationships };
}

function commentEntries(cells) {
  const comments = [];
  const authors = [];
  const authorIds = new Map();
  for (const { r, c, cell } of cells) {
    if (!cell || !Array.isArray(cell.c) || !cell.c.length) continue;
    const items = cell.c.map((comment) => {
      const author = comment && comment.a ? String(comment.a) : 'SheetJS';
      let authorId = authorIds.get(author);
      if (authorId === undefined) {
        authorId = authors.length;
        authors.push(author);
        authorIds.set(author, authorId);
      }
      return { authorId, text: comment && comment.t };
    });
    comments.push({ ref: encodeCell({ r, c }), row: r, col: c, items });
  }
  return { comments, authors };
}

function commentText(items) {
  if (items.length === 1) return String(items[0].text);
  return items
    .map((item, i) => `${i ? 'Reply' : 'Comment'}:\n    ${String(item.text)}`)
    .join('\n');
}

function buildCommentsXml(cells) {
  const { comments, authors } = commentEntries(cells);
  if (!comments.length) return null;
  const authorsXml = authors.map((author) => `<author>${xmlText(author)}</author>`).join('');
  const commentsXml = comments
    .map(
      ({ ref, items }) =>
        `<comment ref="${ref}" authorId="${items[0].authorId}"><text><t${items.length > 1 ? ' xml:space="preserve"' : ''}>${xmlText(commentText(items))}</t></text></comment>`
    )
    .join('');
  return {
    comments,
    xml:
      XML_DECL +
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      `<authors>${authorsXml}</authors><commentList>${commentsXml}</commentList></comments>`,
  };
}

function buildVmlXml(comments) {
  const shapes = comments
    .map(
      ({ row, col }, i) =>
        `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" style="position:absolute; margin-left:80pt;margin-top:5pt;width:104pt;height:64pt;z-index:10" fillcolor="#ECFAD4" strokecolor="#edeaa1"><v:fill color2="#BEFF82" type="gradient" angle="-180"><o:fill type="gradientUnscaled" v:ext="view"/></v:fill><v:shadow on="t" obscured="t"/><v:path o:connecttype="none"/><v:textbox><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1,0,1,0,3,20,5,20</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column><x:Visible/></x:ClientData></v:shape>`
    )
    .join('');
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:mv="http://macVmlSchemaUri"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><v:shapetype id="_x0000_t202" o:spt="202" coordsize="21600,21600" path="m0,0l0,21600,21600,21600,21600,0xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${shapes}</xml>`;
}

function quoteChartSheet(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function chartForWorksheet(ws, sheetName, chartIndex) {
  const chartOrdinal = arguments.length > 3 ? arguments[3] : 0;
  const source = !Array.isArray(ws) && Array.isArray(ws?.['!charts']) ? ws['!charts'][chartOrdinal] : undefined;
  if (!source || typeof source !== 'object' || typeof source.ref !== 'string') return null;
  let range;
  try { range = safeDecodeRange(source.ref); checkRangeSize(range); } catch { return null; }
  if (range.e.c - range.s.c < 1 || range.e.r - range.s.r < 1) return null;
  const series = [];
  for (let col = range.s.c + 1; col <= Math.min(range.e.c, range.s.c + 65); col += 1) {
    const categories = []; const values = [];
    for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
      const number = Number(ws[encodeCell({ r: row, c: col })]?.v);
      if (!Number.isFinite(number)) continue;
      categories.push(String(ws[encodeCell({ r: row, c: range.s.c })]?.v ?? '')); values.push(number);
    }
    if (!values.length) continue;
    const valueCol = encodeCell({ r: range.s.r, c: col }).replace(/\d+$/, '');
    series.push({
      name: String(ws[encodeCell({ r: range.s.r, c: col })]?.v ?? `Series ${series.length + 1}`),
      nameRef: `${quoteChartSheet(sheetName)}!$${valueCol}$${range.s.r + 1}`,
      valueRef: `${quoteChartSheet(sheetName)}!$${valueCol}$${range.s.r + 2}:$${valueCol}$${range.e.r + 1}`,
      color: chartSeriesColor(source.colors?.[series.length], series.length),
      categories,
      values,
    });
  }
  if (!series.length) return null;
  const categoryRef = `${quoteChartSheet(sheetName)}!$${encodeCell({ r: range.s.r + 1, c: range.s.c }).replace(/\d+$/, '')}$${range.s.r + 2}:$${encodeCell({ r: range.e.r, c: range.s.c }).replace(/\d+$/, '')}$${range.e.r + 1}`;
  const widthCols = Number(source.widthCols);
  const heightRows = Number(source.heightRows);
  return { index: chartIndex * 100 + chartOrdinal + 1, partName: chartOrdinal === 0 ? String(chartIndex) : `${chartIndex}_${chartOrdinal + 1}`, ordinal: chartOrdinal, type: source.type === 'line' ? 'line' : 'bar', title: String(source.title || 'Chart'), xAxisTitle: typeof source.xAxisTitle === 'string' ? source.xAxisTitle.trim().slice(0, 255) : '', yAxisTitle: typeof source.yAxisTitle === 'string' ? source.yAxisTitle.trim().slice(0, 255) : '', categoryRef, series, legend: source.legend !== false && series.length > 1, widthCols: Number.isFinite(widthCols) ? Math.max(4, Math.min(20, widthCols)) : 7, heightRows: Number.isFinite(heightRows) ? Math.max(8, Math.min(40, heightRows)) : 16 };
}

function chartsForWorksheet(ws, sheetName, sheetIndex) {
  const count = !Array.isArray(ws) && Array.isArray(ws?.['!charts']) ? ws['!charts'].length : 0;
  return Array.from({ length: count }, (_, ordinal) => chartForWorksheet(ws, sheetName, sheetIndex, ordinal)).filter(Boolean);
}

const CHART_SERIES_COLORS = ['3856D9', '1C8B52', 'D97706', 'B33A8A', '64748B'];
function chartSeriesColor(value, index) {
  const candidate = String(value || '').replace(/^#/, '');
  return /^[0-9A-F]{6}$/i.test(candidate) ? candidate.toUpperCase() : CHART_SERIES_COLORS[index % CHART_SERIES_COLORS.length];
}

function chartCategoryXml(chartSeries, categoryRef) {
  return `<c:cat><c:strRef><c:f>${xmlText(categoryRef)}</c:f></c:strRef></c:cat>`;
}

function chartSeriesXml(chart, chartSeries, index) {
  const name = `<c:tx><c:strRef><c:f>${xmlText(chartSeries.nameRef)}</c:f></c:strRef></c:tx>`;
  const cat = chartCategoryXml(chartSeries, chart.categoryRef);
  const val = `<c:val><c:numRef><c:f>${xmlText(chartSeries.valueRef)}</c:f></c:numRef></c:val>`;
  const properties = chart.type === 'line'
    ? `<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="${chartSeries.color}"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>`
    : `<c:spPr><a:solidFill><a:srgbClr val="${chartSeries.color}"/></a:solidFill></c:spPr>`;
  // CT_Ser orders the series properties before category/value references.
  const invertIfNegative = chart.type === 'bar' ? '<c:invertIfNegative val="0"/>' : '';
  return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/>${name}${properties}${invertIfNegative}${cat}${val}</c:ser>`;
}

function chartAxisTitleXml(title) {
  if (!title) return '';
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr/><a:r><a:rPr lang="en-US"/><a:t>${xmlText(title)}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></c:rich></c:tx><c:layout/></c:title>`;
}

function buildChartXml(chart) {
  const seriesXml = chart.series.map((entry, index) => chartSeriesXml(chart, entry, index)).join('');
  const legend = chart.legend ? '<c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/></c:legend>' : '';
  const series = chart.type === 'line'
    ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesXml}<c:axId val="202100"/><c:axId val="202099"/></c:lineChart>`
    : `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${seriesXml}<c:gapWidth val="150"/><c:axId val="202100"/><c:axId val="202099"/></c:barChart>`;
  return XML_DECL + `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:lang val="en-US"/><c:roundedCorners val="0"/><c:style val="2"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr/><a:r><a:rPr lang="en-US"/><a:t>${xmlText(chart.title)}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${series}<c:catAx><c:axId val="202100"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${chartAxisTitleXml(chart.xAxisTitle)}<c:crossAx val="202099"/><c:crosses val="autoZero"/></c:catAx><c:valAx><c:axId val="202099"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>${chartAxisTitleXml(chart.yAxisTitle)}<c:crossAx val="202100"/><c:crosses val="autoZero"/></c:valAx></c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}

function buildDrawingXml(charts) {
  const anchors = charts.map((chart, ordinal) => {
    const fromCol = ordinal % 2 === 0 ? 3 : 3 + chart.widthCols + 1;
    const fromRow = 1 + Math.floor(ordinal / 2) * (chart.heightRows + 1);
    // Use the same stable anchor shape emitted by Excel/openpyxl: the chart is
    // positioned by its top-left cell and sized explicitly in EMUs. This avoids
    // relying on an empty xfrm inside a two-cell anchor, which Excel may repair.
    const widthEmu = Math.round(chart.widthCols * 675000);
    const heightEmu = Math.round(chart.heightRows * 168750);
    return `<xdr:oneCellAnchor><xdr:from><xdr:col>${fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${widthEmu}" cy="${heightEmu}"/><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="${chart.index}" name="Chart ${chart.index}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId${ordinal + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`;
  }).join('');
  return XML_DECL + `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${anchors}</xdr:wsDr>`;
}

function buildSheetRels(relationships) {
  if (!relationships.length) return '';
  return (
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    relationships.join('') +
    '</Relationships>'
  );
}

function buildSheetDataXml(ws, cells, styleTable) {
  const rows = new Map();
  let minR = Infinity;
  let maxR = -Infinity;
  let minC = Infinity;
  let maxC = -Infinity;
  for (const entry of cells) {
    if (!rows.has(entry.r)) rows.set(entry.r, []);
    rows.get(entry.r).push(entry);
    minR = Math.min(minR, entry.r);
    maxR = Math.max(maxR, entry.r);
    minC = Math.min(minC, entry.c);
    maxC = Math.max(maxC, entry.c);
  }

  const wsRowsMeta = Array.isArray(ws) ? undefined : ws['!rows'];
  const hiddenRows = hiddenRowSet(wsRowsMeta);
  for (const row of hiddenRows) {
    if (!rows.has(row)) rows.set(row, []);
  }

  let xml = '';
  for (const row of [...rows.keys()].sort((a, b) => a - b)) {
    const rowCells = rows.get(row).slice().sort((a, b) => a.c - b.c);
    const rowMeta = wsRowsMeta?.[row];
    const height = Number(rowMeta?.hpt ?? rowMeta?.height);
    const heightAttr = Number.isFinite(height) && height >= 1 && height <= 409
      ? ` ht="${height}" customHeight="1"`
      : '';
    const hiddenAttr = hiddenRows.has(row) ? ' hidden="1"' : '';
    const cellsXml = rowCells
      .map(({ r, c, cell }) => cellXml(encodeCell({ r, c }), cell, styleTable))
      .join('');
    xml += `<row r="${row + 1}"${heightAttr}${hiddenAttr}>${cellsXml}</row>`;
  }

  const dimensionRef = minR !== Infinity
    ? `${encodeCell({ r: minR, c: minC })}:${encodeCell({ r: maxR, c: maxC })}`
    : 'A1';
  return { xml, dimensionRef };
}

function buildSheetXml(ws, styleTable, sheetIndex, sheetName) {
  const cells = collectCells(ws);
  const hyperlinkInfo = buildHyperlinkInfo(cells);
  const commentsInfo = buildCommentsXml(cells);
  const chartInfos = chartsForWorksheet(ws, sheetName, sheetIndex);
  const hasCharts = chartInfos.length > 0;
  const tableInfos = buildTableInfos(ws, sheetIndex, sheetName);
  const commentRelOffset = hyperlinkInfo.relationships.length + (hasCharts ? 1 : 0) + tableInfos.length;
  const relationships = hyperlinkInfo.relationships.slice();
  if (hasCharts) relationships.push(`<Relationship Id="rId${hyperlinkInfo.relationships.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${sheetIndex}.xml"/>`);
  const tableRelStart = hyperlinkInfo.relationships.length + (hasCharts ? 2 : 1);
  tableInfos.forEach((table, index) => relationships.push(`<Relationship Id="rId${tableRelStart + index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table${table.partName}.xml"/>`));
  if (commentsInfo) {
    relationships.push(
      `<Relationship Id="rId${commentRelOffset + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing${sheetIndex}.vml"/>`,
      `<Relationship Id="rId${commentRelOffset + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments${sheetIndex}.xml"/>`
    );
  }

  // A declared `!ref` is validated (not iterated — the loop below only ever visits cells
  // the caller actually populated, sparse-safe by construction) via the same
  // ELIXCEE_RANGE_TOO_LARGE guard sheet_to_csv/sheet_to_json/sheet_to_formulae already use
  // (internal/range-guard.cjs), rather than a second bespoke limit — see
  // compat/differential/classify.mjs's SAFETY_DIVERGENCE_REGISTRY, keyed by that exact code.
  let declaredRef = null;
  if (ws && !Array.isArray(ws) && typeof ws['!ref'] === 'string') {
    checkRangeSize(safeDecodeRange(ws['!ref']));
    declaredRef = ws['!ref'];
  }

  const wsRowsMeta = Array.isArray(ws) ? undefined : ws['!rows'];
  const wsColsMeta = Array.isArray(ws) ? undefined : ws['!cols'];
  const sheetData = buildSheetDataXml(ws, cells, styleTable);
  const dimensionRef = declaredRef || sheetData.dimensionRef;

  const hyperlinkXml = hyperlinkInfo.links.length
    ? `<hyperlinks>${hyperlinkInfo.links.join('')}</hyperlinks>`
    : '';
  return {
    xml: (
    XML_DECL +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"${relationships.length ? ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' : ''}>` +
    `<dimension ref="${xmlAttr(dimensionRef)}"/>` +
    buildSheetViewsXml(ws) +
    buildColsXml(wsColsMeta) +
    `<sheetData>${sheetData.xml}</sheetData>` +
    buildMergesXml(Array.isArray(ws) ? undefined : ws['!merges']) +
    buildDataValidationsXml(Array.isArray(ws) ? undefined : ws['!dataValidations']) +
    buildConditionalFormattingXml(Array.isArray(ws) ? undefined : ws['!conditionalFormats'], styleTable) +
    hyperlinkXml +
    (hasCharts ? `<drawing r:id="rId${hyperlinkInfo.relationships.length + 1}"/>` : '') +
    (commentsInfo
      ? `<legacyDrawing r:id="rId${commentRelOffset + 1}"/>`
      : '') +
    (tableInfos.length ? `<tableParts count="${tableInfos.length}">${tableInfos.map((_, index) => `<tablePart r:id="rId${tableRelStart + index}"/>`).join('')}</tableParts>` : '') +
    '</worksheet>'
    ),
    relationships,
    comments: commentsInfo,
    vml: commentsInfo ? buildVmlXml(commentsInfo.comments) : null,
    charts: chartInfos,
    tables: tableInfos,
    drawing: hasCharts ? buildDrawingXml(chartInfos) : null,
    drawingRels: hasCharts ? buildSheetRels(chartInfos.map((chart, index) => `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chart.partName}.xml"/>`)) : null,
  };
}

// ---- top-level orchestration: WorkBook -> ZIP entries ----

function buildXlsxZipEntries(wb) {
  if (!wb || typeof wb !== 'object' || !Array.isArray(wb.SheetNames) || typeof wb.Sheets !== 'object') {
    throw unsupported(ELIXCEE_UNSUPPORTED_SHEET_SHAPE, 'write() requires a WorkBook ({ SheetNames, Sheets })');
  }
  if (!wb.SheetNames.length) {
    throw unsupported(ELIXCEE_UNSUPPORTED_SHEET_SHAPE, 'write() requires at least one worksheet');
  }
  const vba = wb['!vbaProject'];
  const hasVba = vba instanceof Uint8Array || vba instanceof ArrayBuffer;
  if (vba != null && !hasVba) throw unsupported(ELIXCEE_UNSUPPORTED_SHEET_SHAPE, "'!vbaProject' must be Uint8Array or ArrayBuffer");
  const vbaBytes = hasVba ? (vba instanceof Uint8Array ? vba : new Uint8Array(vba)) : undefined;

  const styleTable = createStyleTable();
  const sheetOutputs = wb.SheetNames.map((name, i) => {
    const ws = wb.Sheets[name];
    if (ws == null) {
      throw unsupported(ELIXCEE_UNSUPPORTED_SHEET_SHAPE, `sheet '${name}' is listed in SheetNames but missing from Sheets`);
    }
    return buildSheetXml(ws, styleTable, i + 1, name);
  });
  const commentSheets = sheetOutputs
    .map((output, i) => (output.comments ? i + 1 : null))
    .filter((i) => i !== null);
  const chartSheets = sheetOutputs
    .map((output, i) => (output.charts?.length ? { drawing: String(i + 1), charts: output.charts.map((chart) => chart.partName) } : null))
    .filter(Boolean);
  const tableSheets = sheetOutputs.flatMap((output) => output.tables.map((table) => table.partName));

  // TextEncoder, not Buffer.from(s, 'utf8') — a standard Web/Node API both platforms have
  // natively (Node 11+, every real browser), unlike Buffer which is Node-only and, unlike
  // read()'s own Buffer-free toBytes() in index.browser.mjs, is NOT polyfilled by esbuild
  // for `platform: 'browser'` by default — confirmed live (not assumed): a real Chrome tab
  // running this package's own browser-smoke.mjs bundle threw `ReferenceError: Buffer is
  // not defined` here before this fix. index.cjs's writeBuffer() wraps this module's
  // Uint8Array-returning output back into a real Node Buffer for API compatibility with
  // the oracle's own type:'buffer' contract — see its own doc comment.
  const utf8 = (s) => new TextEncoder().encode(s);
  const entries = [
    { name: '[Content_Types].xml', data: utf8(buildContentTypes(wb.SheetNames.length, commentSheets, chartSheets, tableSheets, hasVba)) },
    { name: '_rels/.rels', data: utf8(buildRootRels()) },
    { name: 'docProps/core.xml', data: utf8(buildCoreXml()) },
    { name: 'docProps/app.xml', data: utf8(buildAppXml(wb.SheetNames)) },
    { name: 'xl/workbook.xml', data: utf8(buildWorkbookXml(wb.SheetNames, wb.Workbook)) },
    { name: 'xl/_rels/workbook.xml.rels', data: utf8(buildWorkbookRels(wb.SheetNames.length, hasVba)) },
  ];
  wb.SheetNames.forEach((_, i) => {
    entries.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: utf8(sheetOutputs[i].xml) });
    if (sheetOutputs[i].relationships.length) {
      entries.push({
        name: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
        data: utf8(buildSheetRels(sheetOutputs[i].relationships)),
      });
    }
    if (sheetOutputs[i].comments) {
      entries.push(
        { name: `xl/comments${i + 1}.xml`, data: utf8(sheetOutputs[i].comments.xml) },
        { name: `xl/drawings/vmlDrawing${i + 1}.vml`, data: utf8(sheetOutputs[i].vml) }
      );
    }
    if (sheetOutputs[i].charts?.length) {
      entries.push(
        { name: `xl/drawings/drawing${i + 1}.xml`, data: utf8(sheetOutputs[i].drawing) },
        { name: `xl/drawings/_rels/drawing${i + 1}.xml.rels`, data: utf8(sheetOutputs[i].drawingRels) },
        ...sheetOutputs[i].charts.map((chart) => ({ name: `xl/charts/chart${chart.partName}.xml`, data: utf8(buildChartXml(chart)) }))
      );
    }
    for (const table of sheetOutputs[i].tables) entries.push({ name: `xl/tables/table${table.partName}.xml`, data: utf8(table.xml) });
  });
  entries.push({ name: 'xl/styles.xml', data: utf8(styleTable.build()) });
  if (hasVba) entries.push({ name: 'xl/vbaProject.bin', data: vbaBytes });

  return entries;
}

module.exports = {
  buildXlsxZipEntries,
  xmlText,
  xmlAttr,
  XML_DECL,
  buildContentTypes,
  buildRootRels,
  buildWorkbookRels,
  buildCoreXml,
  buildAppXml,
  buildWorkbookXml,
  buildSheetRels,
  buildHyperlinkInfo,
  createStyleTable,
  unsupported,
  ELIXCEE_UNSUPPORTED_CELL_TYPE,
  ELIXCEE_UNSUPPORTED_SHEET_SHAPE,
  checkRangeSize,
  safeDecodeRange,
  datenum,
  encodeCell,
  decodeCell,
};
