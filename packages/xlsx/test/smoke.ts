// TypeScript smoke test: compiles a small consumer snippet against src/index.d.ts.
// Not a runtime test (see compat/differential/xlsx-utils.test.mjs for that) — this only
// proves the type declarations are usable the way real xlsx-typed consumer code is.
import * as XLSX from '../src/index';
import { CellAddress, Range, WorkBook, WorkSheet } from '../src/index';

const addr: CellAddress = { c: 0, r: 0 };
const cellStr: string = XLSX.encode_cell(addr);
const decoded: CellAddress = XLSX.decode_cell(cellStr);

const range: Range = { s: { c: 0, r: 0 }, e: { c: 1, r: 1 } };
const rangeStr: string = XLSX.encode_range(range);
const decodedRange: Range = XLSX.decode_range(rangeStr);
const rangeStr2: string = XLSX.encode_range({ c: 0, r: 0 }, { c: 1, r: 1 });

const col: number = XLSX.decode_col('A');
const colStr: string = XLSX.encode_col(0);
const row: number = XLSX.decode_row('1');
const rowStr: string = XLSX.encode_row(0);
const parts: [string, string] = XLSX.split_cell('A1');

const wb: WorkBook = XLSX.book_new();
const ws: WorkSheet = XLSX.aoa_to_sheet([[1, 2], [3, 4]]);
const sheetName: string = XLSX.book_append_sheet(wb, ws, 'Sheet1');
XLSX.book_set_sheet_visibility(wb, 0, 1);
XLSX.book_set_sheet_visibility(wb, 'Sheet1', 0);

const ws2: WorkSheet = XLSX.sheet_add_aoa(ws, [[5, 6]], { origin: 'A3', dense: false });
const ws3: WorkSheet = XLSX.json_to_sheet([{ a: 1, b: 'x' }, { a: 2, b: 'y' }]);
const ws4: WorkSheet = XLSX.sheet_add_json(ws3, [{ a: 3, b: 'z' }], { origin: -1, skipHeader: true });

const cell = { t: 'n' as const, v: 1234.5 };
XLSX.cell_set_number_format(cell, 'General');
const formatted: string = XLSX.format_cell(cell);
const formatted2: string = XLSX.format_cell(cell, 99, { dateNF: 'm/d/yy' });

const formulae: string[] = XLSX.sheet_to_formulae(ws4);

XLSX.cell_set_hyperlink(cell, 'https://example.com', 'tip');
XLSX.cell_set_internal_link(cell, 'Sheet1!A1');
XLSX.cell_add_comment(cell, 'a comment', 'author');
const ws5: WorkSheet = XLSX.sheet_set_array_formula(ws4, 'A1:A1', 'SUM(1,2)', true);

const csv: string = XLSX.sheet_to_csv(ws5, { FS: ',', blankrows: false });
const txt: string = XLSX.sheet_to_txt(ws5, { type: 'string' });

// sheet_get_cell: all 3 call shapes.
const gotByRef: XLSX.CellObject = XLSX.sheet_get_cell(ws4, 'A1');
const gotByAddr: XLSX.CellObject = XLSX.sheet_get_cell(ws4, { r: 0, c: 0 });
const gotByRC: XLSX.CellObject = XLSX.sheet_get_cell(ws4, 0, 0);
const gotByRow: XLSX.CellObject = XLSX.sheet_get_cell(ws4, 0);

console.log(
  decoded, decodedRange, col, colStr, row, rowStr, parts, sheetName, rangeStr2, ws2, ws4, ws5,
  formatted, formatted2, formulae, csv, txt, gotByRef, gotByAddr, gotByRC, gotByRow
);
