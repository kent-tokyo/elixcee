// TypeScript compile test: proves a real HTMLTableElement (not just `any`) is accepted by
// sheet_add_dom/table_to_sheet/table_to_book, since the oracle's own types declare these
// as `data: any` rather than `data: HTMLTableElement`. This file's presence under the
// DEFAULT tsconfig.json (whose implicit lib already includes "DOM" — TypeScript's default
// lib inference pulls in DOM unless explicitly restricted, confirmed by checking this
// project's own tsconfig.json, which has no `lib` override) is the "DOM-lib TypeScript
// setting" compile check. See tsconfig.no-dom.json / `npm run typecheck:no-dom` for the
// complementary check that packages/xlsx's own .d.ts does NOT require DOM lib at all —
// narrowing sheet_add_dom/table_to_sheet/table_to_book to accept ONLY HTMLTableElement
// would force every consumer (including non-browser, non-DOM-lib TypeScript projects) to
// pull in DOM lib just to import this package's types, which is why the oracle types
// these `any` and this package matches that exactly (a SAFE_EXTENSION would be narrowing
// further; this file proves elixcee did NOT do that).
import * as XLSX from '../src/index';

const table: HTMLTableElement = document.createElement('table');
const ws1: XLSX.WorkSheet = XLSX.table_to_sheet(table);
const wb1: XLSX.WorkBook = XLSX.table_to_book(table, { sheet: 'FromDOM' });
const ws2: XLSX.WorkSheet = XLSX.aoa_to_sheet([[1]]);
const ws3: XLSX.WorkSheet = XLSX.sheet_add_dom(ws2, table, { origin: -1, cellDates: true });

console.log(ws1, wb1, ws3);
