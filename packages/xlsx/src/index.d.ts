// Type shapes mirror xlsx@0.18.5's own types/index.d.ts (CellAddress, Range, Origin/AOA/
// JSON option interfaces) so existing `xlsx`-typed consumer code keeps compiling
// unchanged. Phase 1B-1 added worksheet mutation (sheet_add_aoa/sheet_add_json) and a
// narrow number-format subset (format_cell/cell_set_number_format). Phase 1B-2A adds
// formula extraction (sheet_to_formulae) and cell metadata (hyperlink/comment/array
// formula) — see docs/xlsx-architecture.md's Phase 1B plan for what's next.

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

export interface OriginOption {
  /** Top-Left cell for the operation (CellAddress, "A1"-style string, or row number) */
  origin?: number | string | CellAddress;
}

export interface AOA2SheetOpts {
  dense?: boolean;
  sheetStubs?: boolean;
  cellDates?: boolean;
  nullError?: boolean;
  /** Use specified date format (only 'm/d/yy', numFmtId 14's default, is implemented) */
  dateNF?: string | number;
}

export interface SheetAOAOpts extends AOA2SheetOpts, OriginOption {}

export interface JSON2SheetOpts {
  /** Use specified column order */
  header?: string[];
  /** Skip header row in generated sheet */
  skipHeader?: boolean;
  cellDates?: boolean;
  nullError?: boolean;
  dateNF?: string | number;
}

export interface SheetJSONOpts extends JSON2SheetOpts, OriginOption {}

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
export function sheet_add_aoa<T>(ws: WorkSheet, data: T[][], opts?: SheetAOAOpts): WorkSheet;
export function json_to_sheet<T>(data: T[], opts?: SheetJSONOpts): WorkSheet;
export function sheet_add_json<T>(ws: WorkSheet, data: T[], opts?: SheetJSONOpts): WorkSheet;

export function format_cell(cell: CellObject, v?: unknown, opts?: { dateNF?: string | number }): string;
export function cell_set_number_format(cell: CellObject, fmt: string | number): CellObject;

export function sheet_to_formulae(worksheet: WorkSheet): string[];

export function cell_set_hyperlink(cell: CellObject, target: string, tooltip?: string): CellObject;
export function cell_set_internal_link(cell: CellObject, target: string, tooltip?: string): CellObject;
export function cell_add_comment(cell: CellObject, text: string, author?: string): void;
export function sheet_set_array_formula(
  ws: WorkSheet,
  range: Range | string,
  formula: string,
  dynamic?: boolean
): WorkSheet;
