// Type shapes mirror xlsx@0.18.5's own types/index.d.ts (CellAddress, Range, Origin/AOA/
// JSON/CSV option interfaces) so existing `xlsx`-typed consumer code keeps compiling
// unchanged. Phase 1B-1 added worksheet mutation (sheet_add_aoa/sheet_add_json) and a
// number-format subset (format_cell/cell_set_number_format, now full — Phase 1B-2B).
// Phase 1B-2A added formula extraction (sheet_to_formulae) and cell metadata (hyperlink/
// comment/array formula). Phase 1B-2B adds text export (sheet_to_csv/sheet_to_txt) — see
// docs/xlsx-architecture.md's Phase 1B plan for what's next.

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
  /** Use specified date format */
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

export interface Sheet2CSVOpts {
  /** Field Separator ("delimiter") */
  FS?: string;
  /** Record Separator ("row separator") */
  RS?: string;
  /** Remove trailing field separators in each record */
  strip?: boolean;
  /** Include blank lines in the CSV output */
  blankrows?: boolean;
  /** Skip hidden rows and columns in the CSV output */
  skipHidden?: boolean;
  /** Force quotes around fields */
  forceQuotes?: boolean;
  /** if true, return raw numbers; if false, return formatted numbers */
  rawNumbers?: boolean;
  dateNF?: string | number;
}

export interface Sheet2TXTOpts extends Sheet2CSVOpts {
  /** If 'string', return a plain string instead of BOM + UTF-16LE encoding */
  type?: 'string';
}

export interface Sheet2JSONOpts {
  /** Output format: 1 -> 0-based index-number keys, "A" -> column-letter keys, an
   * array -> explicit header names, omitted -> infer from row 0's formatted text */
  header?: 'A' | 1 | string[];
  /** Override worksheet range */
  range?: any;
  /** Include or omit blank lines in the output */
  blankrows?: boolean;
  /** Default value for null/undefined values */
  defval?: any;
  /** if true, return raw data; if false, return formatted text */
  raw?: boolean;
  /** if true, skip hidden rows and columns */
  skipHidden?: boolean;
  /** if true, return raw numbers; if false, return formatted numbers */
  rawNumbers?: boolean;
  dateNF?: string | number;
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
export function sheet_add_aoa<T>(ws: WorkSheet, data: T[][], opts?: SheetAOAOpts): WorkSheet;
export function json_to_sheet<T>(data: T[], opts?: SheetJSONOpts): WorkSheet;
export function sheet_add_json<T>(ws: WorkSheet, data: T[], opts?: SheetJSONOpts): WorkSheet;

// Not present in xlsx@0.18.5's own types/index.d.ts at all (confirmed: no `get_cell`
// entry there) even though it's a real runtime export (`sheet_get_cell: ws_get_cell_stub`
// in the oracle's own source) — this is pure addition, not a narrowing of any existing
// oracle declaration. Mirrors the runtime's 3 call shapes exactly.
export function sheet_get_cell(ws: WorkSheet, ref: string): CellObject;
export function sheet_get_cell(ws: WorkSheet, cell: CellAddress): CellObject;
export function sheet_get_cell(ws: WorkSheet, row: number, col?: number): CellObject;

export function format_cell(cell: CellObject, v?: unknown, opts?: { dateNF?: string | number }): string;
export function cell_set_number_format(cell: CellObject, fmt: string | number): CellObject;

// Mirrors xlsx@0.18.5's own overload set verbatim (types/index.d.ts) — including the two
// non-generic overloads below the generic one, even though normal TS overload resolution
// makes them largely unreachable in practice, so any call site pattern the real oracle's
// types accept is still accepted here.
export function sheet_to_json<T>(worksheet: WorkSheet, opts?: Sheet2JSONOpts): T[];
export function sheet_to_json(worksheet: WorkSheet, opts?: Sheet2JSONOpts): any[][];
export function sheet_to_json(worksheet: WorkSheet, opts?: Sheet2JSONOpts): any[];

export function sheet_to_formulae(worksheet: WorkSheet): string[];
export function sheet_to_csv(worksheet: WorkSheet, options?: Sheet2CSVOpts): string;
export function sheet_to_txt(worksheet: WorkSheet, options?: Sheet2TXTOpts): string;

export function cell_set_hyperlink(cell: CellObject, target: string, tooltip?: string): CellObject;
export function cell_set_internal_link(cell: CellObject, target: string, tooltip?: string): CellObject;
export function cell_add_comment(cell: CellObject, text: string, author?: string): void;
export function sheet_set_array_formula(
  ws: WorkSheet,
  range: Range | string,
  formula: string,
  dynamic?: boolean
): WorkSheet;
