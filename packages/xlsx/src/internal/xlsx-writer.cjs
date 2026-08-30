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
// bookType: 'xlsx' only (see write() in index.cjs for the explicit-unsupported-error
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

function buildContentTypes(sheetCount) {
  const overrides = [
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
  ];
  for (let i = 1; i <= sheetCount; i++) {
    overrides.push(
      `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    );
  }
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

function buildWorkbookRels(sheetCount) {
  const rels = [];
  for (let i = 1; i <= sheetCount; i++) {
    rels.push(
      `<Relationship Id="rId${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i}.xml"/>`
    );
  }
  rels.push(
    `<Relationship Id="rId${sheetCount + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
  );
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
  const cellXfByNumFmtId = new Map(); // numFmtId -> cellXf index (0 reserved for General)
  let nextCustomId = 164;
  cellXfByNumFmtId.set(0, 0);

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

  // Returns the 0-based cellXf index a cell's `s="N"` attribute should reference.
  function cellXfFor(code) {
    if (!code || code === 'General') return 0;
    const numFmtId = numFmtIdFor(code);
    if (cellXfByNumFmtId.has(numFmtId)) return cellXfByNumFmtId.get(numFmtId);
    const idx = cellXfByNumFmtId.size;
    cellXfByNumFmtId.set(numFmtId, idx);
    return idx;
  }

  function build() {
    const numFmtsXml = customNumFmts.length
      ? `<numFmts count="${customNumFmts.length}">${customNumFmts
          .map((f) => `<numFmt numFmtId="${f.id}" formatCode="${xmlAttr(f.code)}"/>`)
          .join('')}</numFmts>`
      : '';
    // cellXfByNumFmtId is insertion-ordered (index 0 = General, inserted first above) —
    // Map iteration order matches insertion order, so this reproduces the same index
    // assignment cellXfFor() already handed out.
    const xfs = [...cellXfByNumFmtId.keys()]
      .map((numFmtId) => `<xf numFmtId="${numFmtId}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`)
      .join('');
    return (
      XML_DECL +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      numFmtsXml +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${cellXfByNumFmtId.size}">${xfs}</cellXfs>` +
      '</styleSheet>'
    );
  }

  return { cellXfFor, build };
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
    const sAttr = cell.z ? ` s="${styleTable.cellXfFor(cell.z)}"` : '';
    return `<c r="${ref}"${typeAttr}${sAttr}>${inner}</c>`;
  }

  if (v instanceof Date || t === 'd') {
    const serial = v instanceof Date ? datenum(v) : v;
    if (typeof serial !== 'number' || !isFinite(serial)) {
      throw unsupported(ELIXCEE_UNSUPPORTED_CELL_TYPE, `date cell ${ref} has neither a Date nor a finite serial value`);
    }
    const sAttr = ` s="${styleTable.cellXfFor(cell.z || 'm/d/yy')}"`;
    return `<c r="${ref}"${sAttr}><v>${serial}</v></c>`;
  }

  const sAttr = cell.z ? ` s="${styleTable.cellXfFor(cell.z)}"` : '';

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
  const runs = [];
  let runStart = null;
  for (let i = 0; i <= colsMeta.length; i++) {
    const isHidden = i < colsMeta.length && colsMeta[i] && colsMeta[i].hidden;
    if (isHidden && runStart === null) {
      runStart = i;
    } else if (!isHidden && runStart !== null) {
      runs.push([runStart, i - 1]);
      runStart = null;
    }
  }
  if (!runs.length) return '';
  const inner = runs
    .map(([s, e]) => `<col min="${s + 1}" max="${e + 1}" width="9.140625" hidden="1" customWidth="1"/>`)
    .join('');
  return `<cols>${inner}</cols>`;
}

function buildMergesXml(merges) {
  if (!Array.isArray(merges) || !merges.length) return '';
  const inner = merges
    .map((m) => `<mergeCell ref="${xmlAttr(encodeCell(m.s) + ':' + encodeCell(m.e))}"/>`)
    .join('');
  return `<mergeCells count="${merges.length}">${inner}</mergeCells>`;
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

function buildSheetRels(relationships) {
  if (!relationships.length) return '';
  return (
    XML_DECL +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    relationships.join('') +
    '</Relationships>'
  );
}

function buildSheetXml(ws, styleTable) {
  const cells = collectCells(ws);
  const hyperlinkInfo = buildHyperlinkInfo(cells);

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

  const rows = new Map(); // 0-based row -> [{r,c,cell}]
  let minR = Infinity;
  let maxR = -Infinity;
  let minC = Infinity;
  let maxC = -Infinity;
  for (const entry of cells) {
    if (!rows.has(entry.r)) rows.set(entry.r, []);
    rows.get(entry.r).push(entry);
    if (entry.r < minR) minR = entry.r;
    if (entry.r > maxR) maxR = entry.r;
    if (entry.c < minC) minC = entry.c;
    if (entry.c > maxC) maxC = entry.c;
  }

  const wsRowsMeta = Array.isArray(ws) ? undefined : ws['!rows'];
  const wsColsMeta = Array.isArray(ws) ? undefined : ws['!cols'];
  const hiddenRows = hiddenRowSet(wsRowsMeta);
  // A hidden row with no populated cells still needs its own <row> element to carry the
  // hidden flag — real Excel does the same for a hidden-but-empty row.
  for (const r of hiddenRows) {
    if (!rows.has(r)) rows.set(r, []);
  }

  let sheetDataXml = '';
  for (const r of [...rows.keys()].sort((a, b) => a - b)) {
    const rowCells = rows.get(r).slice().sort((a, b) => a.c - b.c);
    const hiddenAttr = hiddenRows.has(r) ? ' hidden="1"' : '';
    let cellsXml = '';
    for (const { r: cr, c, cell } of rowCells) {
      cellsXml += cellXml(encodeCell({ r: cr, c }), cell, styleTable);
    }
    sheetDataXml += `<row r="${r + 1}"${hiddenAttr}>${cellsXml}</row>`;
  }

  let dimensionRef;
  if (declaredRef) dimensionRef = declaredRef;
  else if (minR !== Infinity) dimensionRef = `${encodeCell({ r: minR, c: minC })}:${encodeCell({ r: maxR, c: maxC })}`;
  else dimensionRef = 'A1';

  const hyperlinkXml = hyperlinkInfo.links.length
    ? `<hyperlinks>${hyperlinkInfo.links.join('')}</hyperlinks>`
    : '';
  return {
    xml: (
    XML_DECL +
    `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"${hyperlinkInfo.relationships.length ? ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' : ''}>` +
    `<dimension ref="${xmlAttr(dimensionRef)}"/>` +
    buildColsXml(wsColsMeta) +
    `<sheetData>${sheetDataXml}</sheetData>` +
    buildMergesXml(Array.isArray(ws) ? undefined : ws['!merges']) +
    hyperlinkXml +
    '</worksheet>'
    ),
    relationships: hyperlinkInfo.relationships,
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

  const styleTable = createStyleTable();
  const sheetOutputs = wb.SheetNames.map((name) => {
    const ws = wb.Sheets[name];
    if (ws == null) {
      throw unsupported(ELIXCEE_UNSUPPORTED_SHEET_SHAPE, `sheet '${name}' is listed in SheetNames but missing from Sheets`);
    }
    return buildSheetXml(ws, styleTable);
  });

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
    { name: '[Content_Types].xml', data: utf8(buildContentTypes(wb.SheetNames.length)) },
    { name: '_rels/.rels', data: utf8(buildRootRels()) },
    { name: 'docProps/core.xml', data: utf8(buildCoreXml()) },
    { name: 'docProps/app.xml', data: utf8(buildAppXml(wb.SheetNames)) },
    { name: 'xl/workbook.xml', data: utf8(buildWorkbookXml(wb.SheetNames, wb.Workbook)) },
    { name: 'xl/_rels/workbook.xml.rels', data: utf8(buildWorkbookRels(wb.SheetNames.length)) },
  ];
  wb.SheetNames.forEach((_, i) => {
    entries.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: utf8(sheetOutputs[i].xml) });
    if (sheetOutputs[i].relationships.length) {
      entries.push({
        name: `xl/worksheets/_rels/sheet${i + 1}.xml.rels`,
        data: utf8(buildSheetRels(sheetOutputs[i].relationships)),
      });
    }
  });
  entries.push({ name: 'xl/styles.xml', data: utf8(styleTable.build()) });

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
