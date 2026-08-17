// Puts elixcee's and LibreOffice's raw per-scenario results into a comparable canonical
// shape. Mirrors this project's existing differential-testing convention (see
// ../differential/normalize.mjs's doc comment: normalize before classify, never compare
// raw/unnormalized shapes) but for the VBA-macro-vs-spreadsheet-engine domain instead of
// the xlsx-npm-API domain — the two normalizers solve different problems and aren't
// merged.
//
// elixcee's `cells` shape (see docs/agent-contract.md): [{sheet, address, value}], value
// is a JSON number/boolean/string/null, or the literal placeholder strings "[array]" /
// "[record]" for types the CLI doesn't serialize.
//
// LibreOffice's `cells` shape (see run-libreoffice.mjs's harness macro): [{address, type,
// value}], where `type` is one of "number" | "string" | "formula_number" — deliberately
// read via `getType()` branching rather than a bare `getValue()`, because `getValue()`
// returns 0 for a text cell (a real elixcee number could also legitimately be 0 — reading
// it via getValue() alone would silently manufacture a false MATCH on that address; see
// the harness's use of getType() for why this matters).
//
// KNOWN, NOT-YET-FIXED GAP: a VBA Boolean written via `Range(...).Value = True` has no
// distinct UNO CellContentType — LibreOffice reports it as an ordinary VALUE cell holding
// the number 1, indistinguishable at the getType() level from a real numeric 1 (telling
// them apart would need a NumberFormat check, not implemented here). elixcee, by
// contrast, serializes VBA Booleans as real JSON `true`/`false` (see
// docs/agent-contract.md). Comparing here (`true` vs. `1`) fails strict equality, so a
// scenario that legitimately matches on a Boolean cell would currently classify
// UNCLASSIFIED instead of MATCH — a false MISMATCH, not a false MATCH, so it fails safe,
// but it is real noise, not yet worked around. Not exercised by this milestone's actual
// run: every `boolean_logic`-category scenario timed out in run-libreoffice.mjs before
// reaching a cell dump (see README.md), so this gap has not yet produced an incorrect
// classification in a real result — it would on the first Boolean cell that reaches this
// comparison after the underlying LibreOffice hang (see README.md) is fixed.

/**
 * @param {{sheet?: string, address: string, value: unknown}} cell
 * @returns {{address: string, value: unknown}}
 */
function normalizeElixceeCell(cell) {
  let value = cell.value;
  // "[array]" / "[record]" are elixcee CLI serialization placeholders, not real
  // spreadsheet values — never let these participate in a numeric/string equality check;
  // callers should route these through classify()'s dedicated placeholder handling
  // instead (see classify.mjs).
  return { address: cell.address.toUpperCase(), value, isPlaceholder: value === '[array]' || value === '[record]' };
}

/**
 * @param {{address: string, type: string, value: unknown}} cell
 * @returns {{address: string, value: unknown}}
 */
function normalizeLibreOfficeCell(cell) {
  // LibreOffice's AbsoluteName is like "$Sheet1.$A$1" — reduce to the bare "A1" address
  // so it lines up with elixcee's `address` field (elixcee doesn't include the sheet
  // name in `address`, only alongside it in `sheet`).
  const found = cell.address.match(/\$?([A-Z]+)\$?(\d+)$/i);
  const address = found ? `${found[1].toUpperCase()}${found[2]}` : cell.address.toUpperCase();
  return { address, value: cell.value, isPlaceholder: false };
}

function roundFloat(value) {
  // Round to 6 decimal places: enough to catch a real divergence while absorbing binary
  // floating-point representation noise between two independent engines computing "the
  // same" arithmetic (e.g. 0.1 + 0.2, or LibreOffice's Str()-then-reparse round trip).
  return typeof value === 'number' && Number.isFinite(value) ? Math.round(value * 1e6) / 1e6 : value;
}

/**
 * @param {Array<{sheet?: string, address: string, value: unknown}>} cells
 */
export function normalizeElixceeCells(cells) {
  return cells
    .map(normalizeElixceeCell)
    .map((c) => ({ ...c, value: roundFloat(c.value) }))
    .sort((a, b) => a.address.localeCompare(b.address));
}

/**
 * @param {Array<{address: string, type: string, value: unknown}>} cells
 */
export function normalizeLibreOfficeCells(cells) {
  return cells
    .map(normalizeLibreOfficeCell)
    .map((c) => ({ ...c, value: roundFloat(c.value) }))
    .sort((a, b) => a.address.localeCompare(b.address));
}

/**
 * Structural equality over two normalized cell arrays. Placeholder cells (elixcee's
 * "[array]"/"[record]") never compare equal to anything, including each other — they're
 * meant to be filtered to a distinct classification path before this is called, not
 * silently matched.
 */
export function cellsEqual(a, b) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (a[i].isPlaceholder || b[i].isPlaceholder) return false;
    if (a[i].address !== b[i].address) return false;
    if (!Object.is(a[i].value, b[i].value) && a[i].value !== b[i].value) return false;
  }
  return true;
}

// Runnable self-check: `node normalize.mjs`
if (import.meta.url === `file://${process.argv[1]}`) {
  const assert = await import('node:assert/strict');

  const e = normalizeElixceeCells([
    { sheet: 'sheet1', address: 'B2', value: 3 },
    { sheet: 'sheet1', address: 'A1', value: 0.1 + 0.2 },
  ]);
  assert.deepEqual(e.map((c) => c.address), ['A1', 'B2'], 'sorted by address');
  assert.equal(e[0].value, 0.3, 'float rounding absorbs binary representation noise');

  const lo = normalizeLibreOfficeCells([
    { address: '$Sheet1.$B$2', type: 'number', value: 3 },
    { address: '$Sheet1.$A$1', type: 'number', value: 0.300000012 },
  ]);
  assert.deepEqual(lo.map((c) => c.address), ['A1', 'B2']);
  assert.equal(lo[0].value, 0.3, 'LO address reduced to bare A1 form and value rounded to match');

  assert.equal(cellsEqual(e, lo), true, 'elixcee and LibreOffice normalize to the same shape here');

  const withPlaceholder = normalizeElixceeCells([{ sheet: 's', address: 'A1', value: '[array]' }]);
  assert.equal(withPlaceholder[0].isPlaceholder, true);
  assert.equal(
    cellsEqual(withPlaceholder, normalizeLibreOfficeCells([{ address: 'A1', type: 'string', value: '[array]' }])),
    false,
    'a placeholder cell never compares equal, even to a literal string match of its own placeholder text'
  );

  console.log('normalize.mjs self-check: all assertions passed');
}
