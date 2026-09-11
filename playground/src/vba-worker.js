const forbidden = /\b(?:shell|createobject|open|kill|msgbox|userform|filesystemobject|http|declare|callbyname)\b/i;
const MAX_INPUT_CELLS = 10000;
const MAX_INPUT_CELL_TEXT = 1_000_000;
const MAX_INPUT_TEXT = 4_000_000;

function validateInputCells(cells) {
  if (!cells || typeof cells !== 'object' || Array.isArray(cells)) throw new Error('macro input cells must be an object');
  const entries = Object.entries(cells);
  if (entries.length > MAX_INPUT_CELLS) throw new Error(`macro input cell limit exceeded (${MAX_INPUT_CELLS})`);
  let total = 0;
  for (const [ref, value] of entries) {
    if (!/^[A-Z]+[0-9]+$/i.test(ref)) throw new Error(`invalid macro input cell: ${ref}`);
    const length = String(value ?? '').length;
    if (length > MAX_INPUT_CELL_TEXT) throw new Error(`macro input cell exceeds ${MAX_INPUT_CELL_TEXT} characters`);
    total += length;
    if (total > MAX_INPUT_TEXT) throw new Error(`macro input exceeds ${MAX_INPUT_TEXT} characters`);
  }
}

function refOf(row, col) {
  let name = '';
  for (let n = col; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return `${name}${row}`;
}
function cellRef(token, vars = {}) {
  const match = /^Cells\s*\(\s*(.+?)\s*,\s*(.+?)\s*\)$/i.exec(token);
  if (!match) throw new Error(`unsupported cell reference: ${token}`);
  const row = /^\d+$/.test(match[1]) ? Number(match[1]) : Math.trunc(evaluate(match[1], {}, vars));
  const col = /^\d+$/.test(match[2]) ? Number(match[2]) : Math.trunc(evaluate(match[2], {}, vars));
  if (!Number.isInteger(row) || !Number.isInteger(col) || row < 1 || col < 1) throw new Error(`invalid cell coordinates: ${token}`);
  return refOf(row, col);
}
function tokens(source) {
  const result = source.match(/"(?:[^"]|"")*"|Cells\s*\(\s*\d+\s*,\s*\d+\s*\)(?:\.Value)?|\d+(?:\.\d+)?|[A-Za-z_][A-Za-z0-9_]*|[()+*/-]/gi) || [];
  if (result.join('').replace(/\s/g, '') !== source.replace(/\s/g, '')) throw new Error(`unsupported expression: ${source}`);
  return result;
}
function evaluate(source, cells, vars) {
  const input = tokens(source); let index = 0;
  const value = (token) => {
    if (/^Cells/i.test(token)) return Number(cells[cellRef(token.replace(/\.Value$/i, ''), vars)] ?? 0) || 0;
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
function condition(source, cells, vars) {
  const match = /^(.+?)\s*(>=|<=|<>|=|>|<)\s*(.+)$/.exec(source);
  if (!match) return Boolean(evaluate(source, cells, vars));
  const left = evaluate(match[1].trim(), cells, vars); const right = evaluate(match[3].trim(), cells, vars);
  if (match[2] === '=') return left === right;
  if (match[2] === '<>') return left !== right;
  if (match[2] === '>') return left > right;
  if (match[2] === '<') return left < right;
  if (match[2] === '>=') return left >= right;
  return left <= right;
}
function run({ source, macroName, cells }) {
  if (source.length > 10000) throw new Error('macro source exceeds 10 KiB');
  validateInputCells(cells);
  if (forbidden.test(source)) throw new Error('blocked VBA statement: only bounded assignments are allowed');
  const lines = source.split(/\r?\n/).map((line) => line.replace(/'.*$/, '').trim()).filter(Boolean);
  const headerIndex = lines.findIndex((line) => /^Sub\s+/i.test(line));
  const header = headerIndex >= 0 ? lines[headerIndex] : undefined;
  const match = header && /^Sub\s+([A-Za-z_]\w*)\s*\([^)]*\)\s*$/i.exec(header);
  if (!match || match[1].toLowerCase() !== macroName.toLowerCase()) throw new Error(`macro must contain Sub ${macroName}()`);
  const endSubIndices = lines.map((line, index) => /^End\s+Sub$/i.test(line) ? index : -1).filter((index) => index >= headerIndex);
  if (endSubIndices.length !== 1) throw new Error('macro must contain exactly one End Sub');
  const endSubIndex = endSubIndices[0];
  if (lines.slice(endSubIndex + 1).length) throw new Error('macro cannot contain code after End Sub');
  const body = lines.slice(headerIndex + 1, endSubIndex).filter((line) => !/^Dim\s+/i.test(line));
  if (body.length > 100) throw new Error('macro statement limit exceeded (100)');
  const vars = Object.create(null); const changes = []; let executed = 0;
  const assign = (line) => {
    executed += 1; if (executed > 2000) throw new Error('macro execution limit exceeded (2000)');
    const assignment = /^(.+?)\s*=\s*(.+)$/.exec(line); if (!assignment) throw new Error(`unsupported statement: ${line}`);
    const lhs = assignment[1].trim(); const evaluated = evaluate(assignment[2].trim(), cells, vars);
    if (/^Cells\s*\(/i.test(lhs)) { const ref = cellRef(lhs.replace(/\.Value$/i, ''), vars); cells[ref] = evaluated; changes.push({ ref, value: evaluated }); }
    else if (/^[A-Za-z_]\w*$/.test(lhs)) vars[lhs.toLowerCase()] = evaluated;
    else throw new Error(`unsupported assignment target: ${lhs}`);
  };
  const execute = (lines, depth = 0) => {
    if (depth > 8) throw new Error('nested control depth exceeded (8)');
    for (let index = 0; index < lines.length; index += 1) {
      const line = lines[index];
      const selectMatch = /^Select\s+Case\s+(.+)$/i.exec(line);
      if (selectMatch) {
        let level = 1; let endIndex = -1; const cases = [];
        for (let cursor = index + 1; cursor < lines.length; cursor += 1) {
          if (/^Select\s+Case\s+/i.test(lines[cursor])) level += 1;
          else if (/^End\s+Select$/i.test(lines[cursor])) { level -= 1; if (level === 0) { endIndex = cursor; break; } }
          else if (level === 1 && /^Case\s+/i.test(lines[cursor])) cases.push({ index: cursor, line: lines[cursor] });
        }
        if (endIndex < 0 || !cases.length) throw new Error('missing End Select or Case');
        const selector = evaluate(selectMatch[1], cells, vars); let chosen;
        for (let caseIndex = 0; caseIndex < cases.length; caseIndex += 1) {
          const entry = cases[caseIndex]; const next = cases[caseIndex + 1]?.index ?? endIndex; const expression = entry.line.replace(/^Case\s+/i, '').trim();
          const matches = /^Else$/i.test(expression) || expression.split(',').some((candidate) => selector === evaluate(candidate.trim(), cells, vars));
          if (matches) { chosen = lines.slice(entry.index + 1, next); break; }
        }
        if (chosen) execute(chosen, depth + 1); index = endIndex; continue;
      }
      const ifMatch = /^If\s+(.+?)\s+Then$/i.exec(line);
      if (ifMatch) {
        let level = 1; let elseIndex = -1; let endIndex = -1;
        for (let cursor = index + 1; cursor < lines.length; cursor += 1) { if (/^If\s+.+\s+Then$/i.test(lines[cursor])) level += 1; else if (/^End\s+If$/i.test(lines[cursor])) { level -= 1; if (level === 0) { endIndex = cursor; break; } } else if (level === 1 && /^Else$/i.test(lines[cursor])) elseIndex = cursor; }
        if (endIndex < 0) throw new Error('missing End If');
        const branch = condition(ifMatch[1], cells, vars) ? lines.slice(index + 1, elseIndex >= 0 ? elseIndex : endIndex) : (elseIndex >= 0 ? lines.slice(elseIndex + 1, endIndex) : []);
        execute(branch, depth + 1); index = endIndex; continue;
      }
      const forMatch = /^For\s+([A-Za-z_]\w*)\s*=\s*(.+?)\s+To\s+(.+)$/i.exec(line);
      if (forMatch) {
        let level = 1; let endIndex = -1;
        for (let cursor = index + 1; cursor < lines.length; cursor += 1) { if (/^For\s+/i.test(lines[cursor])) level += 1; else if (/^Next(?:\s+[A-Za-z_]\w*)?$/i.test(lines[cursor])) { level -= 1; if (level === 0) { endIndex = cursor; break; } } }
        if (endIndex < 0) throw new Error('missing Next');
        const start = Math.trunc(evaluate(forMatch[2], cells, vars)); const end = Math.trunc(evaluate(forMatch[3], cells, vars)); if (Math.abs(end - start) > 1000) throw new Error('For range exceeds 1000 iterations');
        const step = start <= end ? 1 : -1; for (let value = start; step > 0 ? value <= end : value >= end; value += step) { vars[forMatch[1].toLowerCase()] = value; execute(lines.slice(index + 1, endIndex), depth + 1); }
        index = endIndex; continue;
      }
      const doMatch = /^Do(?:\s+(While|Until)\s+(.+))?$/i.exec(line);
      if (doMatch) {
        let level = 1; let endIndex = -1;
        for (let cursor = index + 1; cursor < lines.length; cursor += 1) { if (/^Do\s+/i.test(lines[cursor])) level += 1; else if (/^Loop(?:\s+(?:While|Until)\s+.+)?$/i.test(lines[cursor])) { level -= 1; if (level === 0) { endIndex = cursor; break; } } }
        if (endIndex < 0) throw new Error('missing Loop');
        const loopMatch = /^Loop(?:\s+(While|Until)\s+(.+))?$/i.exec(lines[endIndex]); let iterations = 0;
        while (true) {
          if (doMatch[1] && (doMatch[1].toLowerCase() === 'while' ? !condition(doMatch[2], cells, vars) : condition(doMatch[2], cells, vars))) break;
          iterations += 1; if (iterations > 1000) throw new Error('Do loop exceeds 1000 iterations'); execute(lines.slice(index + 1, endIndex), depth + 1);
          if (loopMatch?.[1] && (loopMatch[1].toLowerCase() === 'while' ? !condition(loopMatch[2], cells, vars) : condition(loopMatch[2], cells, vars))) break;
          if (!doMatch[1] && !loopMatch?.[1]) throw new Error('unbounded Do loop requires While or Until');
        }
        index = endIndex; continue;
      }
      if (/^(?:Else|End\s+If|Next(?:\s+[A-Za-z_]\w*)?)$/i.test(line)) throw new Error(`unexpected control statement: ${line}`);
      assign(line);
    }
  };
  execute(body);
  return { changes, statements: executed };
}
self.onmessage = (event) => { try { self.postMessage({ ok: true, ...run(event.data) }); } catch (error) { self.postMessage({ ok: false, error: String(error.message || error) }); } };
