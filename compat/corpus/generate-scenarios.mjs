// Generates compat/corpus/scenarios.json: ~600 VBA-macro scenarios, from parameterized
// templates rather than hand-typed one at a time (600 hand-written fixtures would be
// unauditable and mostly-duplicate busywork; a template x parameter-grid expansion is
// both less code and easier to review category-by-category). See SCHEMA.md for the
// per-scenario shape and why there's no "expected" field.
//
// Run: `node generate-scenarios.mjs` from compat/corpus/. Deterministic — re-running
// regenerates byte-identical output for the same template list, so it's safe to commit
// scenarios.json itself (reviewable diff) rather than regenerating it on every run.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const DIR = path.dirname(fileURLToPath(import.meta.url));

/** @type {Array<{id:string, category:string, description:string, vbaSource:string, entrypoint:string, workbook:string|null}>} */
const scenarios = [];

function add(category, description, body, workbook = null) {
  const idx = scenarios.filter((s) => s.category === category).length + 1;
  const id = `${category}_${String(idx).padStart(4, '0')}`;
  scenarios.push({
    id,
    category,
    description,
    vbaSource: `Sub Scenario()\n${body}\nEnd Sub\n`,
    entrypoint: 'Scenario',
    workbook,
  });
}

// ---------------------------------------------------------------------------
// arithmetic: binary ops across literal operand pairs, several ops, into A1.
const ARITH_OPS = [
  ['+', 'add'],
  ['-', 'subtract'],
  ['*', 'multiply'],
  ['/', 'divide'],
  ['\\', 'integer_divide'],
  ['Mod', 'modulo'],
  ['^', 'power'],
];
const ARITH_PAIRS = [
  [2, 3], [10, 4], [7, 7], [100, 9], [1, 1000], [17, 5], [0, 5], [5, 0.5],
  [-4, 6], [6, -4], [-3, -8], [2.5, 4], [9.5, 2], [1000000, 3], [13, 6],
  [8, 8], [21, 3], [45, 6], [99, 11], [1, 2], [3, 4], [12, 5], [50, 7],
  [0.1, 0.2], [1.5, 1.5], [-1, -1], [1000, 1000], [2, -2], [-7, 3], [3, -7],
];
for (const [op, opName] of ARITH_OPS) {
  for (const [a, b] of ARITH_PAIRS) {
    if (op === '/' && b === 0) continue; // division by zero is its own category below
    if (op === 'Mod' && b === 0) continue;
    add('arithmetic', `A1 = ${a} ${op} ${b}`, `  Range("A1").Value = ${a} ${op} ${b}`);
  }
}
// division/mod by zero: elixcee's documented error-value behavior vs. LO's.
for (const a of [5, -5, 0, 3.5]) {
  add('arithmetic_div_by_zero', `A1 = ${a} / 0`, `  Range("A1").Value = ${a} / 0`);
}

// ---------------------------------------------------------------------------
// boolean_logic
const BOOL_PAIRS = [
  ['True', 'True'], ['True', 'False'], ['False', 'True'], ['False', 'False'],
];
for (const op of ['And', 'Or', 'Xor']) {
  for (const [a, b] of BOOL_PAIRS) {
    add('boolean_logic', `A1 = ${a} ${op} ${b}`, `  Range("A1").Value = ${a} ${op} ${b}`);
  }
}
for (const v of ['True', 'False']) {
  add('boolean_logic', `A1 = Not ${v}`, `  Range("A1").Value = Not ${v}`);
}
for (const [a, op, b] of [
  [3, '>', 5], [5, '>', 3], [3, '=', 3], [3, '<>', 3], [4, '>=', 4], [4, '<=', 3],
]) {
  add('boolean_logic', `A1 = (${a} ${op} ${b})`, `  Range("A1").Value = (${a} ${op} ${b})`);
}

// ---------------------------------------------------------------------------
// string_functions
const STRINGS = ['"hello"', '"World"', '"  padded  "', '"MiXeD CaSe"', '""', '"a"'];
for (const fn of ['UCase', 'LCase', 'Trim', 'Len', 'LTrim', 'RTrim']) {
  for (const s of STRINGS) {
    add('string_functions', `A1 = ${fn}(${s})`, `  Range("A1").Value = ${fn}(${s})`);
  }
}
for (const [s, start, len] of [
  ['"hello world"', 1, 5], ['"hello world"', 7, 5], ['"abc"', 2, 1],
]) {
  add('string_functions', `A1 = Mid(${s}, ${start}, ${len})`, `  Range("A1").Value = Mid(${s}, ${start}, ${len})`);
}
for (const [s, n] of [['"hello"', 2], ['"hello"', 0], ['"hi"', 10]]) {
  add('string_functions', `A1 = Left(${s}, ${n})`, `  Range("A1").Value = Left(${s}, ${n})`);
  add('string_functions', `A1 = Right(${s}, ${n})`, `  Range("A1").Value = Right(${s}, ${n})`);
}
for (const [a, b] of [
  ['"foo"', '"bar"'], ['"a"', '""'], ['""', '"b"'], ['"1"', '"2"'],
]) {
  add('string_functions', `A1 = ${a} & ${b}`, `  Range("A1").Value = ${a} & ${b}`);
}
for (const [hay, needle] of [
  ['"hello world"', '"world"'], ['"hello"', '"xyz"'], ['"aaa"', '"a"'],
]) {
  add('string_functions', `A1 = InStr(${hay}, ${needle})`, `  Range("A1").Value = InStr(${hay}, ${needle})`);
}
for (const [s, find, repl] of [
  ['"hello world"', '"world"', '"there"'], ['"aaa"', '"a"', '"b"'],
]) {
  add(
    'string_functions',
    `A1 = Replace(${s}, ${find}, ${repl})`,
    `  Range("A1").Value = Replace(${s}, ${find}, ${repl})`
  );
}

// ---------------------------------------------------------------------------
// range_readwrite: read from a base workbook, transform, write elsewhere.
const READWRITE_CASES = [
  ['B2', 'C2', 'x * 2'],
  ['A1', 'A2', 'x + 100'],
  ['C3', 'D3', 'x - 1'],
  ['B1', 'B4', 'x / 2'],
  ['D4', 'A4', '-x'],
];
for (const [src, dst, expr] of READWRITE_CASES) {
  add(
    'range_readwrite',
    `${dst} = f(${src}) over numeric_grid`,
    `  Dim x As Double\n  x = Range("${src}").Value\n  Range("${dst}").Value = ${expr.replace('x', 'x')}`,
    'numeric_grid'
  );
}
for (const cell of ['A1', 'B2', 'C3', 'D4', 'A4', 'D1']) {
  add(
    'range_readwrite',
    `Cells(...) equivalent write to ${cell}`,
    `  Range("${cell}").Value = Range("${cell}").Value + 1`,
    'numeric_grid'
  );
}
for (const cell of ['A1', 'B1', 'C2']) {
  add(
    'range_readwrite',
    `Cells indexing write near ${cell}`,
    `  Cells(1, 1).Value = Cells(1, 1).Value + 1`,
    'numeric_grid'
  );
}

// ---------------------------------------------------------------------------
// control_flow_for
for (const [n, step] of [[5, 1], [10, 1], [3, 1], [10, 2], [10, -1], [1, 1]]) {
  add(
    'control_flow_for',
    `Sum 1..${n} step ${step} into A1`,
    `  Dim i As Integer, total As Double\n  total = 0\n  For i = 1 To ${n} Step ${step}\n    total = total + i\n  Next i\n  Range("A1").Value = total`
  );
}
for (const n of [3, 5, 8]) {
  add(
    'control_flow_for',
    `Write squares 1..${n} down column A`,
    `  Dim i As Integer\n  For i = 1 To ${n}\n    Cells(i, 1).Value = i * i\n  Next i`
  );
}

// ---------------------------------------------------------------------------
// control_flow_do_while
for (const n of [5, 10, 1, 20]) {
  add(
    'control_flow_do_while',
    `Do While countdown from ${n} into A1`,
    `  Dim x As Integer\n  x = ${n}\n  Dim total As Double\n  total = 0\n  Do While x > 0\n    total = total + x\n    x = x - 1\n  Loop\n  Range("A1").Value = total`
  );
}

// ---------------------------------------------------------------------------
// control_flow_if
for (const v of [-5, 0, 5, 100, -100]) {
  add(
    'control_flow_if',
    `If/ElseIf/Else classify ${v} into A1`,
    `  Dim v As Double\n  v = ${v}\n  If v > 0 Then\n    Range("A1").Value = "positive"\n  ElseIf v < 0 Then\n    Range("A1").Value = "negative"\n  Else\n    Range("A1").Value = "zero"\n  End If`
  );
}

// ---------------------------------------------------------------------------
// control_flow_select_case
for (const v of [1, 2, 3, 4, 99]) {
  add(
    'control_flow_select_case',
    `Select Case ${v} into A1`,
    `  Dim v As Integer\n  v = ${v}\n  Select Case v\n    Case 1\n      Range("A1").Value = "one"\n    Case 2, 3\n      Range("A1").Value = "two_or_three"\n    Case Is > 10\n      Range("A1").Value = "big"\n    Case Else\n      Range("A1").Value = "other"\n  End Select`
  );
}

// ---------------------------------------------------------------------------
// arrays
for (const vals of [[1, 2, 3], [10, 20, 30, 40], [-1, 0, 1], [5]]) {
  add(
    'arrays',
    `Dim arr(${vals.length - 1}), sum into A1`,
    `  Dim arr(${vals.length - 1}) As Double\n${vals.map((v, i) => `  arr(${i}) = ${v}`).join('\n')}\n  Dim i As Integer, total As Double\n  total = 0\n  For i = 0 To ${vals.length - 1}\n    total = total + arr(i)\n  Next i\n  Range("A1").Value = total`
  );
}
for (const vals of [[1, 2, 3], [4, 5, 6, 7]]) {
  add(
    'arrays',
    `Array values written to column A`,
    `  Dim arr(${vals.length - 1}) As Double\n${vals.map((v, i) => `  arr(${i}) = ${v}`).join('\n')}\n  Dim i As Integer\n  For i = 0 To ${vals.length - 1}\n    Cells(i + 1, 1).Value = arr(i)\n  Next i`
  );
}

// ---------------------------------------------------------------------------
// type_conversion
for (const [fn, arg] of [
  ['CStr', '42'], ['CStr', '3.14'], ['CInt', '"42"'], ['CDbl', '"3.14"'],
  ['CBool', '1'], ['CBool', '0'], ['Int', '3.9'], ['Int', '-3.9'],
  ['Fix', '3.9'], ['Fix', '-3.9'], ['Abs', '-5'], ['Abs', '5'],
  ['Round', '3.14159'], ['Sgn', '-5'], ['Sgn', '5'], ['Sgn', '0'],
]) {
  add('type_conversion', `A1 = ${fn}(${arg})`, `  Range("A1").Value = ${fn}(${arg})`);
}
add('type_conversion', 'A1 = Sqr(16)', '  Range("A1").Value = Sqr(16)');
add('type_conversion', 'A1 = Sqr(2)', '  Range("A1").Value = Sqr(2)');

// ---------------------------------------------------------------------------
// worksheet_functions: Application.WorksheetFunction over the numeric_grid fixture.
for (const fn of ['Sum', 'Average', 'Min', 'Max', 'Count']) {
  add(
    'worksheet_functions',
    `A1 = WorksheetFunction.${fn}(A1:D4) over numeric_grid`,
    `  Range("A1").Value = Application.WorksheetFunction.${fn}(Range("A1:D4"))`,
    'numeric_grid'
  );
}

// ---------------------------------------------------------------------------
// nested_calls: Sub calling a Function, result written to a cell.
for (const [a, b] of [[2, 3], [10, -4], [0, 0], [7, 7]]) {
  scenarios.push({
    id: `nested_calls_${String(scenarios.filter((s) => s.category === 'nested_calls').length + 1).padStart(4, '0')}`,
    category: 'nested_calls',
    description: `Scenario calls Helper(${a}, ${b})`,
    vbaSource:
      `Function Helper(x As Double, y As Double) As Double\n  Helper = x * x + y\nEnd Function\n\n` +
      `Sub Scenario()\n  Range("A1").Value = Helper(${a}, ${b})\nEnd Sub\n`,
    entrypoint: 'Scenario',
    workbook: null,
  });
}

// ---------------------------------------------------------------------------
// error_handling: On Error Resume Next around a failing operation.
add(
  'error_handling',
  'On Error Resume Next swallows a divide-by-zero, A1 stays default then B1 set',
  `  On Error Resume Next\n  Range("A1").Value = 1 / 0\n  Range("B1").Value = "reached"`
);
add(
  'error_handling',
  'On Error GoTo label after a runtime error',
  `  On Error GoTo Handler\n  Dim arr(2) As Integer\n  arr(5) = 1\n  Range("A1").Value = "no error"\n  Exit Sub\nHandler:\n  Range("A1").Value = "caught"`
);

// ---------------------------------------------------------------------------
// unsupported_functions: deliberately exercise functions FUNCTIONS.md lists as not yet
// supported, so the classifier has real UNSUPPORTED-candidate material (see
// FUNCTIONS.md's "Not Yet Supported" section) rather than only ever hitting MATCH/BUG.
for (const expr of [
  'Application.WorksheetFunction.TextJoin(",", True, "a", "b")',
  'Application.WorksheetFunction.Xlookup(1, Range("A1:A4"), Range("B1:B4"))',
]) {
  add('unsupported_functions', `A1 = ${expr}`, `  Range("A1").Value = ${expr}`, 'numeric_grid');
}

// ---------------------------------------------------------------------------
// nondeterministic: intentionally time-dependent, included so the classifier has real
// material for the NONDETERMINISTIC verdict rather than never exercising it. Compared
// oracle-vs-itself (two elixcee runs, or two LibreOffice runs) rather than
// elixcee-vs-oracle, since a live clock makes cross-engine equality meaningless anyway.
add('nondeterministic', 'A1 = Now() (time-dependent)', '  Range("A1").Value = Now()');
add('nondeterministic', 'A1 = Timer() (time-dependent)', '  Range("A1").Value = Timer()');

// ---------------------------------------------------------------------------
// mixed_types fixture: type-introspection functions over a grid with numbers, strings,
// booleans, blanks side by side (see workbooks/generate-workbooks.mjs).
const MIXED_CELLS = ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3'];
for (const cell of MIXED_CELLS) {
  for (const fn of ['IsNumeric', 'IsEmpty', 'VarType']) {
    add(
      'type_handling_mixed',
      `A1 = ${fn}(${cell}) over mixed_types`,
      `  Range("A1").Value = ${fn}(Range("${cell}").Value)`,
      'mixed_types'
    );
  }
}

// ---------------------------------------------------------------------------
// with_text fixture: string transforms read from real cells (not literals).
const TEXT_CELLS = ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3'];
for (const cell of TEXT_CELLS) {
  for (const fn of ['UCase', 'LCase', 'Trim', 'Len']) {
    add(
      'text_fixture_ops',
      `A1 = ${fn}(${cell}) over with_text`,
      `  Range("A1").Value = ${fn}(Range("${cell}").Value)`,
      'with_text'
    );
  }
}

// ---------------------------------------------------------------------------
// with_negatives fixture: sign-sensitive functions read from real cells.
const NEG_CELLS = ['A1', 'B1', 'C1', 'A2', 'B2', 'C2', 'A3', 'B3', 'C3'];
for (const cell of NEG_CELLS) {
  for (const fn of ['Abs', 'Sgn', 'Int', 'Fix']) {
    add(
      'negatives_fixture_ops',
      `A1 = ${fn}(${cell}) over with_negatives`,
      `  Range("A1").Value = ${fn}(Range("${cell}").Value)`,
      'with_negatives'
    );
  }
}

// ---------------------------------------------------------------------------
// for_each: iterate a Range's cells (distinct VBA construct from indexed For, per
// FUNCTIONS.md's "For Each" row).
for (const rng of ['A1:D1', 'A1:A4', 'A1:D4', 'B1:C2']) {
  add(
    'for_each',
    `For Each cell In Range("${rng}"), sum into A1 (numeric_grid, may double-count A1)`,
    `  Dim c As Range, total As Double\n  total = 0\n  For Each c In Range("${rng}")\n    total = total + c.Value\n  Next c\n  Range("F1").Value = total`,
    'numeric_grid'
  );
}

// ---------------------------------------------------------------------------
// with_block: With ... End With over a Range, per FUNCTIONS.md's "With block" row.
for (const cell of ['A1', 'B2', 'D4']) {
  add(
    'with_block',
    `With Range("${cell}") ... .Value = ... over numeric_grid`,
    `  With Range("${cell}")\n    .Value = .Value + 1000\n  End With`,
    'numeric_grid'
  );
}

// ---------------------------------------------------------------------------
// named_ranges: Range(...).Name assignment then reference by name, per FUNCTIONS.md.
for (const [rng, name] of [
  ['A1:B2', 'TopLeft'], ['C1:D2', 'TopRight'], ['A1:D1', 'FirstRow'],
]) {
  add(
    'named_ranges',
    `Name ${rng} as ${name}, sum via WorksheetFunction over numeric_grid`,
    `  Range("${rng}").Name = "${name}"\n  Range("F1").Value = Application.WorksheetFunction.Sum(Range("${name}"))`,
    'numeric_grid'
  );
}

// ---------------------------------------------------------------------------
// worksheet_functions, expanded: same functions over a second fixture and a sub-range.
for (const fn of ['Sum', 'Average', 'Min', 'Max', 'Count']) {
  add(
    'worksheet_functions',
    `A1 = WorksheetFunction.${fn}(A1:B2) sub-range over numeric_grid`,
    `  Range("A1").Value = Application.WorksheetFunction.${fn}(Range("A1:B2"))`,
    'numeric_grid'
  );
}
for (const fn of ['Sum', 'Average', 'Max']) {
  add(
    'worksheet_functions',
    `A1 = WorksheetFunction.${fn}(A1:C3) over with_negatives`,
    `  Range("A1").Value = Application.WorksheetFunction.${fn}(Range("A1:C3"))`,
    'with_negatives'
  );
}

// ---------------------------------------------------------------------------
// arrays, expanded: a few more sizes / value sets, plus a 2D-ish accumulation pattern.
for (const vals of [[2, 4, 6, 8, 10], [1, 1, 2, 3, 5, 8], [-2, -4, -6], [100]]) {
  add(
    'arrays',
    `Dim arr(${vals.length - 1}), max into A1`,
    `  Dim arr(${vals.length - 1}) As Double\n${vals.map((v, i) => `  arr(${i}) = ${v}`).join('\n')}\n  Dim i As Integer, m As Double\n  m = arr(0)\n  For i = 1 To ${vals.length - 1}\n    If arr(i) > m Then m = arr(i)\n  Next i\n  Range("A1").Value = m`
  );
}

// ---------------------------------------------------------------------------
// control_flow_for, expanded: a few more n/step combinations.
for (const [n, step] of [[15, 3], [7, 1], [20, 5], [6, 2], [9, -3]]) {
  add(
    'control_flow_for',
    `Sum bounded loop ${n} step ${step} into A1`,
    `  Dim i As Integer, total As Double\n  total = 0\n  For i = 1 To ${n} Step ${step}\n    total = total + i\n  Next i\n  Range("A1").Value = total`
  );
}

// ---------------------------------------------------------------------------
// type_conversion, expanded.
for (const [fn, arg] of [
  ['Round', '2.71828'], ['Round', '-1.5'], ['CInt', '"7"'], ['CDbl', '"2.5"'],
  ['CBool', '"True"'], ['CBool', '"False"'], ['Abs', '-0.001'], ['Sgn', '-0.001'],
  ['Int', '5.999'], ['Fix', '5.999'],
]) {
  add('type_conversion', `A1 = ${fn}(${arg})`, `  Range("A1").Value = ${fn}(${arg})`);
}

// ---------------------------------------------------------------------------
// grid_transform: every cell in the 4x4 numeric_grid fixture, transformed in place —
// the shape of a real "apply a formula-like transform across a used range" macro.
const GRID_CELLS = ['A1', 'B1', 'C1', 'D1', 'A2', 'B2', 'C2', 'D2', 'A3', 'B3', 'C3', 'D3', 'A4', 'B4', 'C4', 'D4'];
for (const cell of GRID_CELLS) {
  for (const [label, expr] of [
    ['double', `Range("${cell}").Value * 2`],
    ['square', `Range("${cell}").Value ^ 2`],
    ['increment', `Range("${cell}").Value + 1`],
  ]) {
    add('grid_transform', `${cell} ${label} in place over numeric_grid`, `  Range("${cell}").Value = ${expr}`, 'numeric_grid');
  }
}

// ---------------------------------------------------------------------------
// comparison_chains: nested If over two variables, a common real-macro shape.
for (const [a, b] of [[1, 2], [5, 5], [10, -10], [0, 0], [-3, 7], [100, 1]]) {
  add(
    'comparison_chains',
    `Classify (${a}, ${b}) relationship into A1`,
    `  Dim a As Double, b As Double\n  a = ${a}\n  b = ${b}\n  If a > b Then\n    Range("A1").Value = "a_greater"\n  ElseIf a < b Then\n    Range("A1").Value = "b_greater"\n  Else\n    Range("A1").Value = "equal"\n  End If`
  );
}

// ---------------------------------------------------------------------------
// string_functions, expanded: a few more literal combinations for Mid/InStr/Replace.
for (const [s, start, len] of [['"VBA macro"', 1, 3], ['"VBA macro"', 5, 5], ['"x"', 1, 1]]) {
  add('string_functions', `A1 = Mid(${s}, ${start}, ${len})`, `  Range("A1").Value = Mid(${s}, ${start}, ${len})`);
}
for (const [hay, needle] of [['"the quick brown fox"', '"brown"'], ['"aaaa"', '"aa"']]) {
  add('string_functions', `A1 = InStr(${hay}, ${needle})`, `  Range("A1").Value = InStr(${hay}, ${needle})`);
}
for (const s of ['"StringLength"', '"a b c d e"']) {
  add('string_functions', `A1 = Len(${s})`, `  Range("A1").Value = Len(${s})`);
}

// ---------------------------------------------------------------------------
// control_flow_do_while, expanded, plus a Do ... Loop Until variant.
for (const n of [3, 7, 15, 25, 2]) {
  add(
    'control_flow_do_while',
    `Do While countdown from ${n} into A1 (extra)`,
    `  Dim x As Integer\n  x = ${n}\n  Dim total As Double\n  total = 0\n  Do While x > 0\n    total = total + x\n    x = x - 1\n  Loop\n  Range("A1").Value = total`
  );
}
for (const n of [1, 4, 9, 16, 25]) {
  add(
    'control_flow_do_while',
    `Do ... Loop Until x*x >= ${n}`,
    `  Dim x As Integer\n  x = 0\n  Do\n    x = x + 1\n  Loop Until x * x >= ${n}\n  Range("A1").Value = x`
  );
}

// ---------------------------------------------------------------------------
// boolean_logic, expanded: a few more relational/logical combinations chained together.
for (const [a, b, c] of [[1, 2, 3], [5, 5, 5], [-1, 0, 1], [10, 5, 1]]) {
  add(
    'boolean_logic',
    `A1 = (${a} < ${b}) And (${b} < ${c})`,
    `  Range("A1").Value = (${a} < ${b}) And (${b} < ${c})`
  );
}
for (const v of [-10, -1, 0, 1, 10]) {
  add('boolean_logic', `A1 = (${v} >= 0) Or (${v} Mod 2 = 0)`, `  Range("A1").Value = (${v} >= 0) Or (${v} Mod 2 = 0)`);
}

// ---------------------------------------------------------------------------
// harness_smoke: touches NO Range/Cells at all (pure local-variable arithmetic), so it
// does not hit the Range/Cells hang documented in README.md's "Known, reproducible
// limitation" section. Its only purpose is to prove the LibreOffice runner's full
// pipeline end to end — module insertion, nested invoke, cell dump, normalize, classify —
// on at least one real, non-timed-out data point, using the numeric_grid fixture so both
// engines dump the same 16 pre-existing cells untouched. See README.md for what this
// scenario found.
add(
  'harness_smoke',
  'Pure local-variable arithmetic, no Range/Cells touched, over numeric_grid',
  '  Dim x As Integer\n  x = 1 + 1',
  'numeric_grid'
);

fs.writeFileSync(path.join(DIR, 'scenarios.json'), JSON.stringify(scenarios, null, 2) + '\n');

const byCategory = new Map();
for (const s of scenarios) byCategory.set(s.category, (byCategory.get(s.category) || 0) + 1);
console.log(`wrote ${scenarios.length} scenarios to scenarios.json`);
for (const [cat, count] of byCategory) console.log(`  ${cat}: ${count}`);
