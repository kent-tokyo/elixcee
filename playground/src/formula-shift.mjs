function columnLabel(index) {
  let label = '';
  let value = index + 1;
  while (value) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
}

function insideFormulaString(formula, offset) {
  let inside = false;
  for (let index = 0; index < offset; index += 1) {
    if (formula[index] !== '"') continue;
    if (formula[index + 1] === '"') {
      index += 1;
      continue;
    }
    inside = !inside;
  }
  return inside;
}

export function shiftFormulaReferences(formula, axis, index, delta) {
  return String(formula).replace(/(^|[^A-Za-z0-9_])([$]?)([A-Z]{1,3})([$]?)(\d+)/gi, (match, prefix, colAbsolute, colText, rowAbsolute, rowText, offset) => {
    if (insideFormulaString(String(formula), offset + prefix.length) || prefix.endsWith('!')) return match;
    const column = colText.toUpperCase().split('').reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0) - 1;
    const row = Number(rowText) - 1;
    const position = axis === 'r' ? row : column;
    if (delta < 0 && position === index) return `${prefix}#REF!`;
    if (position < index) return match;
    const nextPosition = position + delta;
    if (nextPosition < 0) return `${prefix}#REF!`;
    if (axis === 'r') return `${prefix}${colAbsolute ? '$' : ''}${colText.toUpperCase()}${rowAbsolute ? '$' : ''}${nextPosition + 1}`;
    return `${prefix}${colAbsolute ? '$' : ''}${columnLabel(nextPosition)}${rowAbsolute ? '$' : ''}${rowText}`;
  });
}

export function shiftQualifiedFormulaReferences(formula, targetSheet, axis, index, delta) {
  const escaped = String(targetSheet).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const quoted = `'${escaped.replace(/'/g, "''")}'`;
  const pattern = new RegExp(`((?:${quoted}|${escaped})!)((?:\\$)?[A-Z]{1,3}(?:\\$)?\\d+)`, 'gi');
  return String(formula).replace(pattern, (match, qualified, reference, offset) => {
    if (insideFormulaString(String(formula), offset)) return match;
    return `${qualified}${shiftFormulaReferences(reference, axis, index, delta)}`;
  });
}

export function shiftWorkbookFormulaReferences(workbook, targetSheet, axis, index, delta) {
  for (const [hostSheet, worksheet] of Object.entries(workbook?.Sheets || {})) {
    for (const cell of Object.values(worksheet || {})) {
      if (!cell?.f) continue;
      let formula = shiftQualifiedFormulaReferences(cell.f, targetSheet, axis, index, delta);
      if (hostSheet === targetSheet) formula = shiftFormulaReferences(formula, axis, index, delta);
      cell.f = formula;
    }
  }
}
