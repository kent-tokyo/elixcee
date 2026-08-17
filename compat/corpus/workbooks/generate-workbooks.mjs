// Generates the small set of named base workbooks scenarios.json refers to by name
// (see ../SCHEMA.md, "Base workbooks"). Uses the `xlsx` npm package already installed
// as a devDependency for the unrelated xlsx-oracle differential work (compat/oracle) —
// no new dependency. Run: `node workbooks/generate-workbooks.mjs` from compat/corpus/.
import XLSX from 'xlsx';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const DIR = path.dirname(fileURLToPath(import.meta.url));

function writeBook(name, sheetName, aoa) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, path.join(DIR, `${name}.xlsx`));
  console.log(`wrote ${name}.xlsx`);
}

// empty: one sheet, no cells at all — macros that only write have a clean slate.
writeBook('empty', 'Sheet1', [[]]);

// numeric_grid: a small block of numbers for read/aggregate scenarios.
writeBook('numeric_grid', 'Sheet1', [
  [1, 2, 3, 4],
  [5, 6, 7, 8],
  [9, 10, 11, 12],
  [13, 14, 15, 16],
]);

// mixed_types: numbers, strings, booleans side by side for type-handling scenarios.
writeBook('mixed_types', 'Sheet1', [
  [42, 'hello', true],
  [3.14, 'world', false],
  [0, '', null],
]);

// with_text: a small text-only block (string concatenation / length scenarios).
writeBook('with_text', 'Sheet1', [
  ['apple', 'banana', 'cherry'],
  ['Alpha', 'Beta', 'Gamma'],
  ['  padded  ', 'CAPS', 'lower'],
]);

// with_negatives: negatives and zero for arithmetic edge cases.
writeBook('with_negatives', 'Sheet1', [
  [-5, -1, 0],
  [10, -10, 5],
  [-3.5, 2.5, -0],
]);
