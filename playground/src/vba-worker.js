const forbidden = /\b(?:shell|createobject|open|kill|msgbox|userform|filesystemobject|http|declare|callbyname|for|do|while|loop|select|case)\b/i;

function refOf(row, col) {
  let name = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return `${name}${row}`;
}
function cellRef(token) {
  const match = /^Cells\s*\(\s*(\d+)\s*,\s*(\d+)\s*\)$/i.exec(token);
  if (!match) throw new Error(`unsupported cell reference: ${token}`);
  return refOf(Number(match[1]), Number(match[2]));
}
function tokens(source) {
  const result = source.match(/"(?:[^"]|"")*"|Cells\s*\(\s*\d+\s*,\s*\d+\s*\)(?:\.Value)?|\d+(?:\.\d+)?|[A-Za-z_][A-Za-z0-9_]*|[()+*/-]/gi) || [];
  if (result.join('').replace(/\s/g, '') !== source.replace(/\s/g, '')) throw new Error(`unsupported expression: ${source}`);
  return result;
}
function evaluate(source, cells, vars) {
  const input = tokens(source); let index = 0;
  const value = (token) => {
    if (/^Cells/i.test(token)) return Number(cells[cellRef(token.replace(/\.Value$/i, ''))] ?? 0) || 0;
    if (/^"/.test(token)) return token.slice(1, -1).replace(/""/g, '"');
    if (/^\d/.test(token)) return Number(token);
    if (/^[A-Za-z_]/.test(token)) return vars[token.toLowerCase()] ?? 0;
    throw new Error(`unsupported value: ${token}`);
  };
  const factor = () => { if (input[index] === '(') { index++; const result = expression(); if (input[index++] !== ')') throw new Error('unbalanced expression'); return result; } return value(input[index++]); };
  const term = () => { let result = factor(); while (input[index] === '*' || input[index] === '/') { const op = input[index++]; const right = factor(); result = op === '*' ? result * right : result / right; } return result; };
  const expression = () => { let result = term(); while (input[index] === '+' || input[index] === '-') { const op = input[index++]; const right = term(); result = op === '+' ? result + right : result - right; } return result; };
  const result = expression(); if (index !== input.length) throw new Error(`unsupported expression: ${source}`); return result;
}
function run({ source, macroName, cells }) {
  if (source.length > 10000) throw new Error('macro source exceeds 10 KiB');
  if (forbidden.test(source)) throw new Error('blocked VBA statement: only bounded assignments are allowed');
  const lines = source.split(/\r?\n/).map((line) => line.replace(/'.*$/, '').trim()).filter(Boolean);
  const header = lines.find((line) => /^Sub\s+/i.test(line));
  const match = header && /^Sub\s+([A-Za-z_]\w*)\s*\([^)]*\)\s*$/i.exec(header);
  if (!match || match[1].toLowerCase() !== macroName.toLowerCase()) throw new Error(`macro must contain Sub ${macroName}()`);
  const body = lines.slice(lines.indexOf(header) + 1).filter((line) => !/^End\s+Sub$/i.test(line) && !/^Dim\s+/i.test(line));
  if (body.length > 100) throw new Error('macro statement limit exceeded (100)');
  const vars = Object.create(null); const changes = [];
  body.forEach((line) => {
    const assignment = /^(.+?)\s*=\s*(.+)$/.exec(line); if (!assignment) throw new Error(`unsupported statement: ${line}`);
    const lhs = assignment[1].trim(); const evaluated = evaluate(assignment[2].trim(), cells, vars);
    if (/^Cells\s*\(/i.test(lhs)) { const ref = cellRef(lhs.replace(/\.Value$/i, '')); cells[ref] = evaluated; changes.push({ ref, value: evaluated }); }
    else if (/^[A-Za-z_]\w*$/.test(lhs)) vars[lhs.toLowerCase()] = evaluated;
    else throw new Error(`unsupported assignment target: ${lhs}`);
  });
  return { changes, statements: body.length };
}
self.onmessage = (event) => { try { self.postMessage({ ok: true, ...run(event.data) }); } catch (error) { self.postMessage({ ok: false, error: String(error.message || error) }); } };
