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

console.log(decoded, decodedRange, col, colStr, row, rowStr, parts, sheetName, rangeStr2);
