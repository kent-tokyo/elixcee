function cellText(cell) {
  if (cell == null) return '';
  return String(cell.v ?? cell.w ?? '');
}

function encodeCell(row, col) {
  let value = col + 1;
  let letters = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    value = Math.floor((value - 1) / 26);
  }
  return `${letters}${row + 1}`;
}

/**
 * Adjust a summary's source range using the same insertion/deletion boundary
 * rules as tables and charts. Coordinates are zero-based internally.
 */
export function adjustPivotSourceRef(ref, axis, point, delta) {
  let range;
  try { range = JSON.parse(JSON.stringify({ s: decodeRef(ref).s, e: decodeRef(ref).e })); } catch { return undefined; }
  const index = axis === 'r' ? point.r : point.c;
  const start = axis === 'r' ? range.s.r : range.s.c;
  const end = axis === 'r' ? range.e.r : range.e.c;
  if (delta > 0) {
    if (index <= start) { range.s[axis] += delta; range.e[axis] += delta; }
    else if (index <= end) range.e[axis] += delta;
  } else if (index < start) {
    range.s[axis] = Math.max(0, range.s[axis] + delta);
    range.e[axis] = Math.max(0, range.e[axis] + delta);
  } else if (index <= end) {
    if (start === end || index === start) return undefined;
    range.e[axis] = Math.max(range.s[axis], range.e[axis] + delta);
  }
  return `${encodeCell(range.s.r, range.s.c)}:${encodeCell(range.e.r, range.e.c)}`;
}

function decodeRef(ref) {
  const match = /^\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?$/i.exec(String(ref).trim());
  if (!match) throw new Error('invalid range');
  const column = (value) => value.toUpperCase().split('').reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
  const start = { c: column(match[1]), r: Number(match[2]) - 1 };
  const end = { c: match[3] ? column(match[3]) : start.c, r: match[4] ? Number(match[4]) - 1 : start.r };
  if (start.r < 0 || end.r < start.r || end.c < start.c) throw new Error('invalid range');
  return { s: start, e: end };
}

/**
 * Build a bounded worksheet-style pivot summary from a rectangular source.
 * The first column is the category and the remaining columns are numeric
 * measures. This deliberately does not claim to be an OOXML PivotTable.
 */
export function buildPivotRows(ws, range, requested = 'sum', categoryFilter = '') {
  const aggregate = String(requested).trim().toLocaleLowerCase();
  if (!['sum', 'count', 'average'].includes(aggregate)) return undefined;
  if (!ws || !range || range.e.c - range.s.c < 1 || range.e.r - range.s.r < 2) return undefined;
  const filters = new Set(String(categoryFilter).split(',').map((value) => value.trim().toLocaleLowerCase()).filter(Boolean));
  const categoryHeader = cellText(ws[encodeCell(range.s.r, range.s.c)]) || 'Category';
  const valueColumns = [];
  for (let col = range.s.c + 1; col <= range.e.c; col += 1) {
    valueColumns.push({ col, header: cellText(ws[encodeCell(range.s.r, col)]) || `Value ${col - range.s.c}` });
  }
  const groups = new Map();
  for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
    const category = cellText(ws[encodeCell(row, range.s.c)]);
    if (!category || (filters.size && !filters.has(category.toLocaleLowerCase()))) continue;
    const totals = groups.get(category) || valueColumns.map(() => ({ sum: 0, count: 0 }));
    valueColumns.forEach(({ col }, index) => {
      const value = Number(ws[encodeCell(row, col)]?.v);
      if (Number.isFinite(value)) {
        totals[index].sum += value;
        totals[index].count += 1;
      }
    });
    groups.set(category, totals);
  }
  if (!groups.size || !Array.from(groups.values()).some((totals) => totals.some((entry) => entry.count))) return undefined;
  const label = aggregate[0].toUpperCase() + aggregate.slice(1);
  const valueFor = (entry) => entry.count ? (aggregate === 'count' ? entry.count : aggregate === 'average' ? entry.sum / entry.count : entry.sum) : '';
  return [[categoryHeader, ...valueColumns.map(({ header }) => `${label} of ${header}`)], ...Array.from(groups, ([category, totals]) => [category, ...totals.map(valueFor)])];
}
