// Type shapes mirror xlsx@0.18.5's own types/index.d.ts (CellAddress, Range) so existing
// `xlsx`-typed consumer code keeps compiling unchanged. Phase 1A only implements the
// functions below — see docs/xlsx-architecture.md's Phase 1A plan for what's next.

export interface CellAddress {
  c: number;
  r: number;
}

export interface Range {
  s: CellAddress;
  e: CellAddress;
}

export interface CellObject {
  v?: string | number | boolean | Date;
  t: 'b' | 'n' | 'e' | 's' | 'd' | 'z';
  f?: string;
  [key: string]: unknown;
}

export type WorkSheet = { [address: string]: CellObject | unknown } & { '!ref'?: string };

export interface WorkBook {
  SheetNames: string[];
  Sheets: { [name: string]: WorkSheet };
  Workbook?: { Sheets?: Array<{ Hidden?: 0 | 1 | 2 }> };
}

export interface AOA2SheetOpts {
  dense?: boolean;
}

export function encode_col(col: number): string;
export function decode_col(colstr: string): number;
export function encode_row(row: number): string;
export function decode_row(rowstr: string): number;
export function encode_cell(cell: CellAddress): string;
export function decode_cell(cellstr: string): CellAddress;
export function encode_range(range: Range): string;
export function encode_range(start: CellAddress | string, end: CellAddress | string): string;
export function decode_range(range: string): Range;
export function split_cell(cellstr: string): [string, string];

export function book_new(): WorkBook;
export function book_append_sheet(
  workbook: WorkBook,
  worksheet: WorkSheet,
  name?: string,
  roll?: boolean
): string;
export function book_set_sheet_visibility(
  workbook: WorkBook,
  sheet: number | string,
  visibility: 0 | 1 | 2
): void;

export function aoa_to_sheet<T>(data: T[][], opts?: AOA2SheetOpts): WorkSheet;
