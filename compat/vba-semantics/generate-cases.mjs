// Generates cases.json (VBA scenario definitions) and expected-results.json (the
// documented-VBA-semantics ground truth for each, computed from reference/*.mjs rather
// than hand-typed) for the value-correctness suite. Mirrors compat/corpus/generate-
// scenarios.mjs's own "generated from templates, not hand-typed one at a time, but
// committed" precedent — re-run this script and commit the result when adding cases,
// don't hand-edit cases.json/expected-results.json directly.
//
// This suite answers a genuinely different question from compat/corpus/: not "does
// elixcee run without erroring" (that's compat/corpus/'s PASS/FAIL axis, and separately
// compat/corpus/'s own elixcee-vs-LibreOffice/Excel axis), but "is the VALUE elixcee
// produces the one real, documented VBA semantics says it should be". A function that
// runs without error and returns a plausible-but-wrong number is invisible to
// compat/corpus/'s own classifiers; that's exactly the failure mode Round()'s negative-
// digits bug, CInt/CLng's rounding mode, IsNumeric's string handling, Str() vs CStr(),
// and Val()'s whole-string-only parsing all were before this suite existed.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  vbaInt, vbaFix, vbaSgn, bankersRound, excelRound, vbaIntDiv, vbaMod,
} from './reference/numeric.mjs';
import { vbaStr, vbaIsNumeric, vbaVal } from './reference/strings.mjs';
import { vbaAnd, vbaOr, vbaXor, vbaNot } from './reference/logical.mjs';

const DIR = path.dirname(fileURLToPath(import.meta.url));

const cases = [];
const expected = {};

/** Registers one case + its expected outcome. `expectedSpec` is either
 * `{value}` (A1 must equal this after Sub Scenario runs) or
 * `{error: "exact message"}` (Sub Scenario must fail with exactly this message) —
 * `error` is always the DOCUMENTED REAL-VBA WORDING, even when elixcee's own current
 * message differs (see `knownLimitation` below for that case), so this file never quietly
 * launders elixcee's own behavior into being "the spec".
 * `knownLimitation`, if given, is a reason string: elixcee is known and disclosed (in
 * ROADMAP.md/CHANGELOG.md, or right here) to diverge from `expectedSpec` for this exact
 * case, so report.mjs classifies a mismatch here as KNOWN_LIMITATION, not BUG. Omit it for
 * every case that elixcee is expected to actually get right — that's the default, and
 * report.mjs's BUG/UNCLASSIFIED-must-be-0 gate exists specifically so a real regression on
 * one of those can't hide. */
function addCase(id, category, description, vbaBody, expectedSpec, reason, knownLimitation) {
  addCaseWithSource(id, category, description, `Sub Scenario()\n${vbaBody}\nEnd Sub\n`, expectedSpec, reason, knownLimitation);
}

/** Same as addCase, but takes the complete VBA module source verbatim (not wrapped in a
 * single `Sub Scenario() ... End Sub`) — for cases that need more than one top-level
 * declaration, e.g. a Scenario Sub plus a separate Function it calls. */
function addCaseWithSource(id, category, description, vbaSource, expectedSpec, reason, knownLimitation) {
  if (cases.some(c => c.id === id)) throw new Error(`duplicate case id: ${id}`);
  cases.push({ id, category, description, vbaSource, entrypoint: 'Scenario' });
  let spec;
  if ('error' in expectedSpec) {
    spec = { kind: 'error', errorMessage: expectedSpec.error };
  } else if ('computedBy' in expectedSpec) {
    // The expected value can't be baked in at generation time without going stale (e.g.
    // Date()'s serial changes daily) — report.mjs recomputes it fresh, at check time,
    // using the same reference function named here.
    spec = { kind: 'computed', address: expectedSpec.address ?? 'A1', computedBy: expectedSpec.computedBy };
  } else if (expectedSpec.nondeterministic) {
    // No fixed value is even meaningful (e.g. Now()'s sub-second component) — report.mjs
    // only checks that the scenario ran without erroring.
    spec = { kind: 'nondeterministic' };
  } else {
    spec = { kind: 'value', address: expectedSpec.address ?? 'A1', value: expectedSpec.value };
  }
  expected[id] = { ...spec, reason, ...(knownLimitation ? { knownLimitation } : {}) };
}

// Collision-free id fragment for an arbitrary string (used for Val()/IsNumeric() cases
// keyed by their literal string input) — unlike a blanket regex replace, distinct inputs
// that only differ by punctuation (".5" vs "+5") can't collapse to the same id, since
// addCase's own duplicate-id check only catches that AFTER it's too late to tell which
// two inputs collided.
function slugifyLiteral(s) {
  if (s === '') return 'empty';
  return s
    .replace(/^\s+/, 'lead_ws_')
    .replace(/-/g, 'neg')
    .replace(/\+/g, 'pos')
    .replace(/\./g, 'dot')
    .replace(/\s+/g, 'ws')
    .replace(/[^a-zA-Z0-9_]/g, 'x');
}

function vbaLiteral(n) {
  // Render a JS number as VBA source text. Negative numbers need parens in some
  // argument positions (VBA parses `Fix(-3.9)` fine, but `-3.9` as a bare second
  // Range index would not apply here) — plain literal is fine for every use in this file.
  return String(n);
}

// ── numeric_conversion_rounding ─────────────────────────────────────────────
{
  const CAT = 'numeric_conversion_rounding';
  const intFixSgnValues = [3.1, 3.9, -3.1, -3.9, 0, 0.5, -0.5, 7, -7, 100.99, -100.99];
  for (const v of intFixSgnValues) {
    addCase(`int_${v}`, CAT, `Int(${v})`, `  Range("A1").Value = Int(${vbaLiteral(v)})`,
      { value: vbaInt(v) }, 'Int() floors toward negative infinity.');
    addCase(`fix_${v}`, CAT, `Fix(${v})`, `  Range("A1").Value = Fix(${vbaLiteral(v)})`,
      { value: vbaFix(v) }, 'Fix() truncates toward zero, unlike Int().');
    addCase(`sgn_${v}`, CAT, `Sgn(${v})`, `  Range("A1").Value = Sgn(${vbaLiteral(v)})`,
      { value: vbaSgn(v) }, 'Sgn() returns -1/0/1.');
  }

  const roundTieValues = [0.5, 1.5, 2.5, 3.5, 4.5, -0.5, -1.5, -2.5, -3.5, -4.5];
  for (const v of roundTieValues) {
    addCase(`round_tie_${v}`, CAT, `Round(${v}) — exact .5 tie`,
      `  Range("A1").Value = Round(${vbaLiteral(v)})`,
      { value: bankersRound(v) },
      'Round() uses banker\'s rounding (round-half-to-even) on an exact tie, not away-from-zero.');
  }
  const roundNonTieValues = [3.14159, 2.71828, -1.23456, 0.001, -0.001, 9.99999];
  for (const v of roundNonTieValues) {
    addCase(`round_${v}`, CAT, `Round(${v}) — not a tie`,
      `  Range("A1").Value = Round(${vbaLiteral(v)})`,
      { value: bankersRound(v) }, 'Round() with no tie: nearest integer, both rounding modes agree.');
  }
  const roundDigitsCases = [[0.125, 2], [3.14159, 2], [3.14159, 4], [1234.5678, -1], [1234.5678, 0]];
  for (const [v, d] of roundDigitsCases) {
    if (d < 0) {
      addCase(`round_neg_digits_${v}_${d}`, CAT, `Round(${v}, ${d}) — negative digit count`,
        `  Range("A1").Value = Round(${vbaLiteral(v)}, ${d})`,
        { error: 'Invalid procedure call or argument' },
        'Real VBA\'s Round() errors on a negative digit count (unlike WorksheetFunction.Round/Excel\'s ROUND(), which both accept it).');
    } else {
      addCase(`round_digits_${v}_${d}`, CAT, `Round(${v}, ${d})`,
        `  Range("A1").Value = Round(${vbaLiteral(v)}, ${d})`,
        { value: bankersRound(v, d) }, `Round() to ${d} decimal digit(s), banker's rounding.`);
    }
  }

  const cintClngValues = [0.5, 1.5, 2.5, -0.5, -1.5, -2.5, 3.7, -3.7];
  for (const v of cintClngValues) {
    addCase(`cint_${v}`, CAT, `CInt(${v})`, `  Range("A1").Value = CInt(${vbaLiteral(v)})`,
      { value: bankersRound(v) }, 'CInt() uses banker\'s rounding, same as Round(), not away-from-zero.');
    addCase(`clng_${v}`, CAT, `CLng(${v})`, `  Range("A1").Value = CLng(${vbaLiteral(v)})`,
      { value: bankersRound(v) }, 'CLng() uses banker\'s rounding, same as Round(), not away-from-zero.');
  }

  // Round() vs WorksheetFunction.Round() must genuinely diverge on the same tied input.
  addCase('round_vs_wsf_round_2_5', CAT, 'Round(2.5) — VBA\'s own Round()',
    '  Range("A1").Value = Round(2.5)',
    { value: 2 }, 'Round() banker-rounds 2.5 to 2.');
  addCase('round_vs_wsf_round_2_5_wsf', CAT, 'WorksheetFunction.Round(2.5) — same tied input',
    '  Range("A1").Value = Application.WorksheetFunction.Round(2.5)',
    { value: 3 }, 'WorksheetFunction.Round()/Excel\'s ROUND() rounds half-away-from-zero: 2.5 -> 3, unlike VBA\'s own Round().');
}

// ── negative_intdiv_mod ──────────────────────────────────────────────────────
{
  const CAT = 'negative_intdiv_mod';
  const pairs = [
    [7, 2], [-7, 2], [7, -2], [-7, -2], [5, 3], [-5, 3],
    [10, 4], [-10, 4], [1, 3], [-1, 3],
  ];
  for (const [a, b] of pairs) {
    const div = vbaIntDiv(a, b);
    const mod = vbaMod(a, b);
    addCase(`intdiv_${a}_${b}`, CAT, `${a} \\ ${b}`,
      `  Range("A1").Value = ${vbaLiteral(a)} \\ ${vbaLiteral(b)}`,
      'value' in div ? { value: div.value } : { error: div.error },
      'Integer division of the rounded operands, truncated toward zero.');
    addCase(`mod_${a}_${b}`, CAT, `${a} Mod ${b}`,
      `  Range("A1").Value = ${vbaLiteral(a)} Mod ${vbaLiteral(b)}`,
      'value' in mod ? { value: mod.value } : { error: mod.error },
      'Mod takes the sign of the dividend, on the rounded operands.');
  }
  // Fractional operands that round to zero before dividing -- the exact bug class
  // Round()/CInt/CLng's fixes were about, tested here at the \\/Mod operator level.
  const fractionalPairs = [[5, 0.5], [0.1, 0.2], [3, 0.4], [-3, 0.4]];
  for (const [a, b] of fractionalPairs) {
    const div = vbaIntDiv(a, b);
    addCase(`intdiv_frac_${a}_${b}`, CAT, `${a} \\ ${b} — operand rounds before dividing`,
      `  Range("A1").Value = ${vbaLiteral(a)} \\ ${vbaLiteral(b)}`,
      'value' in div ? { value: div.value } : { error: div.error },
      `\\\\ rounds each operand to a whole number first (banker's rounding) before dividing; ${b} rounds to ${bankersRound(b)}.`);
  }
}

// ── logical_bitwise ──────────────────────────────────────────────────────────
{
  const CAT = 'logical_bitwise';
  const boolPairs = [[true, true], [true, false], [false, true], [false, false]];
  for (const [a, b] of boolPairs) {
    for (const [name, fn, vbaOp] of [['and', vbaAnd, 'And'], ['or', vbaOr, 'Or'], ['xor', vbaXor, 'Xor']]) {
      const r = fn(a, b);
      addCase(`${name}_bool_${a}_${b}`, CAT, `${a} ${vbaOp} ${b} (both Boolean)`,
        `  Range("A1").Value = (${a} ${vbaOp} ${b})`,
        { value: r.value }, `${vbaOp} on two genuine Booleans is logical, not bitwise.`);
    }
  }
  const intPairs = [[5, 3], [-5, 3], [12, 10], [0, 5], [-1, 0]];
  for (const [a, b] of intPairs) {
    for (const [name, fn, vbaOp] of [['and', vbaAnd, 'And'], ['or', vbaOr, 'Or'], ['xor', vbaXor, 'Xor']]) {
      const r = fn(a, b);
      addCase(`${name}_int_${a}_${b}`, CAT, `${a} ${vbaOp} ${b} (both Integer)`,
        `  Range("A1").Value = ${vbaLiteral(a)} ${vbaOp} ${vbaLiteral(b)}`,
        { value: r.value }, `${vbaOp} on non-Boolean operands is real bitwise math (banker-rounded to int first).`);
    }
  }
  for (const v of [5, -6, 0, -1, 100]) {
    const r = vbaNot(v);
    addCase(`not_int_${v}`, CAT, `Not ${v} (Integer)`,
      `  Range("A1").Value = Not ${vbaLiteral(v)}`,
      { value: r.value }, 'Not on a non-Boolean is a real bitwise complement, not truthy coercion.');
  }
  for (const v of [true, false]) {
    const r = vbaNot(v);
    addCase(`not_bool_${v}`, CAT, `Not ${v} (Boolean)`,
      `  Range("A1").Value = Not ${v}`,
      { value: r.value }, 'Not on a genuine Boolean is logical negation.');
  }
  addCase('not_and_precedence', CAT, 'Not 5 And 3 — operator precedence',
    '  Range("A1").Value = Not 5 And 3',
    { value: vbaAnd(vbaNot(5).value, 3).value },
    'Not binds tighter than And: (Not 5) And 3, both bitwise since neither operand is a genuine Boolean.');
}

// ── str_cstr_val ─────────────────────────────────────────────────────────────
{
  const CAT = 'str_cstr_val';
  for (const n of [459, 0, 1, 1000000, -459, -1]) {
    addCase(`str_${n}`, CAT, `Str(${n})`, `  Range("A1").Value = "[" & Str(${vbaLiteral(n)}) & "]"`,
      { value: `[${vbaStr(n)}]` }, 'Str() reserves a leading space for the sign position on a non-negative number.');
    addCase(`cstr_${n}`, CAT, `CStr(${n})`, `  Range("A1").Value = "[" & CStr(${vbaLiteral(n)}) & "]"`,
      { value: `[${n}]` }, 'CStr() has no leading-space reservation, unlike Str().');
  }
  const valInputs = ['123abc', '  42.5xyz', 'abc', '', '-5.5xyz', '.5', '5.', '5', '+5', '0abc'];
  for (const s of valInputs) {
    const id = `val_${slugifyLiteral(s)}`;
    addCase(id, CAT, `Val("${s}")`, `  Range("A1").Value = Val("${s.replace(/"/g, '""')}")`,
      { value: vbaVal(s) }, 'Val() parses a leading numeric prefix, stopping at the first non-fitting character.');
  }
}

// ── isnumeric ────────────────────────────────────────────────────────────────
{
  const CAT = 'isnumeric';
  const strInputs = ['123', '12.5', ' 42 ', 'abc', '12abc', '', '-5', '1e10', '3.14.15'];
  for (const s of strInputs) {
    const id = `isnumeric_str_${slugifyLiteral(s)}`;
    addCase(id, CAT, `IsNumeric("${s}")`, `  Range("A1").Value = IsNumeric("${s.replace(/"/g, '""')}")`,
      { value: vbaIsNumeric(s) }, 'IsNumeric() on a string: true iff it parses as a plain number after trimming.');
  }
  for (const n of [123, -5, 0, 3.14]) {
    addCase(`isnumeric_num_${n}`, CAT, `IsNumeric(${n})`, `  Range("A1").Value = IsNumeric(${vbaLiteral(n)})`,
      { value: true }, 'IsNumeric() on an already-numeric value is always true.');
  }
}

// ── typename_vartype ─────────────────────────────────────────────────────────
{
  const CAT = 'typename_vartype';
  const typeCases = [
    ['5', 'Long', 3, 'Integer literal'],
    ['5.5', 'Double', 5, 'Float literal'],
    ['"hi"', 'String', 8, 'String literal'],
    ['True', 'Boolean', 11, 'Boolean literal'],
  ];
  for (const [expr, expectedType, expectedVarType, desc] of typeCases) {
    addCase(`typename_${desc.replace(/\s+/g, '_')}`, CAT, `TypeName(${expr}) — ${desc}`,
      `  Range("A1").Value = TypeName(${expr})`,
      { value: expectedType }, `TypeName() of a ${desc} is "${expectedType}".`);
    addCase(`vartype_${desc.replace(/\s+/g, '_')}`, CAT, `VarType(${expr}) — ${desc}`,
      `  Range("A1").Value = VarType(${expr})`,
      { value: expectedVarType }, `VarType() of a ${desc} is ${expectedVarType} (vb${expectedType}).`);
  }
  addCase('typename_date', CAT, 'TypeName(Date()) — Date() returns a real Date value',
    '  Range("A1").Value = TypeName(Date())',
    { value: 'Date' }, 'Date() returns a genuine Date-typed value (whole-day Excel serial).');
}

// ── date_time_now ────────────────────────────────────────────────────────────
{
  const CAT = 'date_time_now';
  // Date()'s VALUE is deterministic given "today", but baking today's serial in at
  // generation time would go stale by the next calendar day -- `computedBy` tells
  // report.mjs to recompute it fresh (via reference/numeric.mjs's todaySerial(), added
  // below) at check time instead, which runs close enough to run-elixcee.mjs's own
  // invocation in practice (same CI run) to be reliable.
  addCase('date_matches_today', CAT, 'Date() matches today\'s real date',
    '  Range("A1").Value = Date()',
    { computedBy: 'todayDateString' },
    'Date() returns today\'s date; --json renders a Date value as its "YYYY-MM-DD" display string, not the raw Excel-serial number.');
  addCase('date_bare_no_parens_typename', CAT, 'TypeName(Date) — bare, no parens',
    '  Range("A1").Value = TypeName(Date)',
    { value: 'Date' }, 'Real VBA allows omitting () on this zero-arg function.');
  addCase('time_no_error', CAT, 'Time() runs without erroring',
    '  x = Time()\n  Range("A1").Value = "ran"',
    { value: 'ran' },
    'Time()\'s own value is sub-second-volatile and asserted separately as NONDETERMINISTIC below; this only confirms the call itself succeeds.');
  addCase('time_value_nondeterministic', CAT, 'Time() itself — value not asserted',
    '  Range("A1").Value = Time()',
    { nondeterministic: true },
    'Time()\'s value changes every call; only confirmed to run without erroring, per compat/corpus\'s own precedent for Timer().');
  addCase('now_value_nondeterministic', CAT, 'Now() itself — value not asserted',
    '  Range("A1").Value = Now()',
    { nondeterministic: true },
    'Now()\'s value changes every call; only confirmed to run without erroring.');
}

// ── empty_null_error ─────────────────────────────────────────────────────────
{
  const CAT = 'empty_null_error';
  addCase('isempty_uninitialized', CAT, 'IsEmpty on an uninitialized variable',
    '  Dim x\n  Range("A1").Value = IsEmpty(x)',
    { value: true }, 'An uninitialized Variant is Empty.');
  addCase('isempty_after_assignment', CAT, 'IsEmpty after assignment',
    '  Dim x\n  x = 5\n  Range("A1").Value = IsEmpty(x)',
    { value: false }, 'A variable is no longer Empty once assigned.');
  addCase('empty_numeric_coercion', CAT, 'Empty coerces to 0 in a numeric context',
    '  Dim x\n  Range("A1").Value = x + 5',
    { value: 5 }, 'Empty + 5 == 5: Empty coerces to 0 numerically.');
  addCase('isnumeric_empty', CAT, 'IsNumeric(Empty) is True',
    '  Dim x\n  Range("A1").Value = IsNumeric(x)',
    { value: true }, 'Empty coerces to 0 in a numeric context, so IsNumeric(Empty) is True.');
  addCase('division_by_zero_message', CAT, 'Division by zero error message',
    '  Range("A1").Value = 5 / 0',
    { error: 'Division by zero' }, 'A literal division by zero raises this exact runtime error.');
  addCase('array_oob_message', CAT, 'Array out-of-bounds error message',
    '  Dim arr(3)\n  Range("A1").Value = arr(10)',
    { error: 'Subscript out of range' },
    'Real VBA\'s documented runtime error for an out-of-bounds array index is "Subscript out of range" (error 9). Was elixcee\'s own diagnostic wording until this suite found and disclosed the gap; fixed in the same round (see CHANGELOG.md) — no knownLimitation needed anymore.');
}

// ── string_boundaries ────────────────────────────────────────────────────────
{
  const CAT = 'string_boundaries';
  addCase('left_zero_len', CAT, 'Left("hello", 0)', '  Range("A1").Value = "[" & Left("hello", 0) & "]"',
    { value: '[]' }, 'Left() with length 0 returns an empty string.');
  addCase('left_over_len', CAT, 'Left("hi", 10) — length exceeds string', '  Range("A1").Value = "[" & Left("hi", 10) & "]"',
    { value: '[hi]' }, 'Left() with a length longer than the string just returns the whole string.');
  addCase('right_zero_len', CAT, 'Right("hello", 0)', '  Range("A1").Value = "[" & Right("hello", 0) & "]"',
    { value: '[]' }, 'Right() with length 0 returns an empty string.');
  addCase('right_over_len', CAT, 'Right("hi", 10)', '  Range("A1").Value = "[" & Right("hi", 10) & "]"',
    { value: '[hi]' }, 'Right() with a length longer than the string just returns the whole string.');
  addCase('mid_past_end', CAT, 'Mid("hi", 10)', '  Range("A1").Value = "[" & Mid("hi", 10) & "]"',
    { value: '[]' }, 'Mid() starting past the end of the string returns an empty string.');
  addCase('mid_no_length', CAT, 'Mid("hello", 2)', '  Range("A1").Value = Mid("hello", 2)',
    { value: 'ello' }, 'Mid() with no length argument returns everything from the start position onward.');
  addCase('len_empty_string', CAT, 'Len("")', '  Range("A1").Value = Len("")',
    { value: 0 }, 'Len() of an empty string is 0.');
  addCase('instr_not_found', CAT, 'InStr("hello", "z")', '  Range("A1").Value = InStr("hello", "z")',
    { value: 0 }, 'InStr() returns 0 when the substring is not found.');
  addCase('instr_empty_needle', CAT, 'InStr("hello", "")', '  Range("A1").Value = InStr("hello", "")',
    { value: 1 }, 'InStr() with an empty search string returns the start position (1 by default).');
  addCase('ucase_lcase_mixed', CAT, 'UCase/LCase on mixed-case text',
    '  Range("A1").Value = UCase("Hello") & LCase("WORLD")',
    { value: 'HELLOworld' }, 'UCase()/LCase() convert the whole string.');
}

// ── array_indices ────────────────────────────────────────────────────────────
{
  const CAT = 'array_indices';
  addCase('ubound_basic', CAT, 'UBound of Dim arr(3)',
    '  Dim arr(3)\n  Range("A1").Value = UBound(arr)',
    { value: 3 }, 'Dim arr(3) creates a 0-to-3 array (4 elements); UBound is 3.');
  addCase('lbound_basic', CAT, 'LBound of Dim arr(3)',
    '  Dim arr(3)\n  Range("A1").Value = LBound(arr)',
    { value: 0 }, 'elixcee arrays are always 0-based; LBound is always 0.');
  addCase('isarray_true', CAT, 'IsArray on a real array', '  Dim arr(3)\n  Range("A1").Value = IsArray(arr)',
    { value: true }, 'IsArray() is true for an array variable.');
  addCase('isarray_false_scalar', CAT, 'IsArray on a scalar', '  x = 42\n  Range("A1").Value = IsArray(x)',
    { value: false }, 'IsArray() is false for a non-array value.');
  addCase('redim_preserve_grows_elem0', CAT, 'ReDim Preserve keeps arr(0)',
    '  Dim arr(2)\n  arr(0) = 10\n  arr(1) = 20\n  ReDim Preserve arr(4)\n  Range("A1").Value = arr(0)',
    { value: 10 }, 'ReDim Preserve keeps existing element values when growing an array.');
  addCase('redim_preserve_grows_elem1', CAT, 'ReDim Preserve keeps arr(1)',
    '  Dim arr(2)\n  arr(0) = 10\n  arr(1) = 20\n  ReDim Preserve arr(4)\n  Range("A1").Value = arr(1)',
    { value: 20 }, 'ReDim Preserve keeps existing element values when growing an array.');
}

// ── range_values ─────────────────────────────────────────────────────────────
{
  const CAT = 'range_values';
  addCase('range_write_read_integer', CAT, 'Write then read an Integer via Range',
    '  Range("A1").Value = 42\n  Range("A2").Value = Range("A1").Value',
    { value: 42, address: 'A2' }, 'A written Integer round-trips through Range unchanged.');
  addCase('range_write_read_string', CAT, 'Write then read a String via Range',
    '  Range("A1").Value = "hello"\n  Range("A2").Value = Range("A1").Value',
    { value: 'hello', address: 'A2' }, 'A written String round-trips through Range unchanged.');
  addCase('range_write_read_boolean', CAT, 'Write then read a Boolean via Range',
    '  Range("A1").Value = True\n  Range("A2").Value = Range("A1").Value',
    { value: true, address: 'A2' }, 'A written Boolean round-trips through Range unchanged.');
  addCase('range_write_read_float', CAT, 'Write then read a Float via Range',
    '  Range("A1").Value = 3.14\n  Range("A2").Value = Range("A1").Value',
    { value: 3.14, address: 'A2' }, 'A written Float round-trips through Range unchanged.');
  addCase('cells_write_read', CAT, 'Write via Range, read via Cells',
    '  Range("A1").Value = 99\n  Range("B1").Value = Cells(1, 1).Value',
    { value: 99, address: 'B1' }, 'Cells(1,1) and Range("A1") address the same cell.');
}

// ── error_kind ───────────────────────────────────────────────────────────────
{
  const CAT = 'error_kind';
  addCase('undefined_variable_error', CAT, 'Reading an undefined variable',
    '  Range("A1").Value = someUndefinedVariableXyz',
    { error: "Undefined variable: 'someundefinedvariablexyz'" },
    'VBA identifiers are case-insensitive; the error message reports the lowercased name.');
  // A case for "calling an unknown VBA function" was considered and dropped: real VBA
  // catches this at COMPILE time (the macro never starts running at all), a different
  // failure *category* from elixcee's runtime error here, not just different wording —
  // this suite's schema (one VBA source, one runtime outcome) can't represent "should
  // have failed to compile" cleanly. Noted in ROADMAP.md rather than forced into a bad fit.
}

// ── division_by_zero ─────────────────────────────────────────────────────────
// Broader than negative_intdiv_mod's own div-by-zero cases (which are specifically about
// \\/Mod's operand-rounding-to-zero path) — plain `/` across operand type/sign/magnitude
// combinations. All confirmed live to give exactly "Division by zero", matching real VBA.
{
  const CAT = 'division_by_zero';
  const pairs = [
    ['5', '0'], ['-5', '0'], ['0', '0'], ['3.5', '0'], ['-3.5', '0'],
    ['1000000', '0'], ['0.001', '0'], ['(5 + 3)', '0'],
  ];
  for (const [a, b] of pairs) {
    addCase(`divzero_${slugifyLiteral(a)}_${slugifyLiteral(b)}`, CAT, `${a} / ${b}`,
      `  Range("A1").Value = ${a} / ${b}`,
      { error: 'Division by zero' }, 'A literal division by zero raises this exact runtime error, regardless of operand sign/magnitude.');
  }
  addCase('divzero_int_typed_var', CAT, 'Integer-declared variable / 0',
    '  Dim x As Integer\n  x = 5\n  Range("A1").Value = x / 0',
    { error: 'Division by zero' }, 'Division by zero is unconditional regardless of the dividend\'s declared type.');
  addCase('divzero_float_typed_var', CAT, 'Double-declared variable / 0',
    '  Dim x As Double\n  x = 5.5\n  Range("A1").Value = x / 0',
    { error: 'Division by zero' }, 'Division by zero is unconditional regardless of the dividend\'s declared type.');
}

// ── invalid_procedure_argument ──────────────────────────────────────────────
// Real VBA raises "Invalid procedure call or argument" (error 5) for several of these;
// elixcee currently accepts all of them and produces a plausible-but-wrong value instead
// of erroring -- found while building this suite, not previously disclosed. All registered
// as KNOWN_LIMITATION with the specific wrong-vs-right values named, not silently dropped.
{
  const CAT = 'invalid_procedure_argument';
  addCase('left_negative_length', CAT, 'Left("hello", -1)',
    '  Range("A1").Value = Left("hello", -1)',
    { error: 'Invalid procedure call or argument' },
    'Real VBA errors on a negative Length argument to Left().',
    'elixcee returns "" instead of erroring (Left\'s own `.take(n)` with n computed via `as usize` from a negative float saturates to a huge value, then truncates to available chars -- i.e. it silently clamps rather than validating). Found while building this suite.');
  addCase('right_negative_length', CAT, 'Right("hello", -1)',
    '  Range("A1").Value = Right("hello", -1)',
    { error: 'Invalid procedure call or argument' },
    'Real VBA errors on a negative Length argument to Right().',
    'elixcee returns "" instead of erroring, same root cause as Left(-1) above.');
  addCase('mid_zero_start', CAT, 'Mid("hello", 0)',
    '  Range("A1").Value = Mid("hello", 0)',
    { error: 'Invalid procedure call or argument' },
    'Real VBA errors on a Start argument less than 1 (Mid is 1-indexed).',
    'elixcee returns the whole string instead of erroring (`start.saturating_sub(1)` maps 0 to the same value as 1). Found while building this suite.');
  addCase('mid_negative_start', CAT, 'Mid("hello", -1)',
    '  Range("A1").Value = Mid("hello", -1)',
    { error: 'Invalid procedure call or argument' },
    'Real VBA errors on a negative Start argument to Mid().',
    'elixcee returns the whole string instead of erroring, same root cause as Mid(0) above.');
  addCase('chr_negative_code', CAT, 'Chr(-5)',
    '  Range("A1").Value = Len(Chr(-5))',
    { error: 'Invalid procedure call or argument' },
    'Real VBA errors on a negative character code to Chr().',
    'elixcee saturates the negative code to 0 (float-to-u32 cast saturates, not wraps) and returns a 1-character null-byte string instead of erroring. Asserted via Len() rather than the raw character to keep the case JSON-representable.');
  addCase('chr_code_over_255', CAT, 'Chr(256)',
    '  Range("A1").Value = Asc(Chr(256))',
    { error: 'Invalid procedure call or argument' },
    'Real VBA\'s Chr() is documented for character codes 0-255 (use ChrW for the wider Unicode range); 256 is out of Chr()\'s own valid range.',
    'elixcee accepts any code up to u32::MAX via char::from_u32 and returns the corresponding Unicode character (U+0100) instead of erroring — confirmed via Asc(Chr(256)) round-tripping to 256. Found while building this suite.');
  addCase('instr_zero_start', CAT, 'InStr(0, "hello", "l")',
    '  Range("A1").Value = InStr(0, "hello", "l")',
    { error: 'Invalid procedure call or argument' },
    'Real VBA\'s 4-argument InStr requires Start >= 1; 0 is invalid.',
    'elixcee\'s `start.saturating_sub(1)` maps a Start of 0 to the same starting position as 1, silently returning 2 (the same answer InStr(1, "hello", "l") would give) instead of erroring. Found while building this suite.');
}

// ── overflow ─────────────────────────────────────────────────────────────────
// Real VBA's Integer type is -32768..32767, Long is -2147483648..2147483647; CInt/CLng of
// a value outside that range raises "Overflow" (error 6). elixcee's Variant::Integer is
// backed by i64 with no declared-type-width tracking at all -- there is no notion of
// "this value is Integer-typed vs Long-typed" once it's a Variant, so no overflow check is
// possible without a structural type-system change (the same class of change as the
// Date/Time ADR, not attempted here). All registered as KNOWN_LIMITATION.
{
  const CAT = 'overflow';
  addCase('cint_at_boundary_ok', CAT, 'CInt(32767) — the real Integer max, should NOT overflow',
    '  Range("A1").Value = CInt(32767)',
    { value: 32767 }, 'The exact boundary of real VBA\'s Integer range — still valid, not an overflow case.');
  addCase('cint_over_boundary', CAT, 'CInt(32768) — one past the real Integer max',
    '  Range("A1").Value = CInt(32768)',
    { error: 'Overflow' },
    'Real VBA\'s Integer type tops out at 32767; CInt(32768) raises Overflow.',
    'elixcee\'s Variant::Integer is an unconditional i64 with no declared-type-width tracking, so it returns 32768 instead of erroring. Found while building this suite.');
  addCase('cint_large', CAT, 'CInt(40000)',
    '  Range("A1").Value = CInt(40000)',
    { error: 'Overflow' },
    'Real VBA Integer overflow, same reasoning as CInt(32768).',
    'elixcee returns 40000 instead of erroring.');
  addCase('cint_large_negative', CAT, 'CInt(-40000)',
    '  Range("A1").Value = CInt(-40000)',
    { error: 'Overflow' },
    'Real VBA Integer underflow (below -32768) also raises Overflow.',
    'elixcee returns -40000 instead of erroring.');
  addCase('clng_large', CAT, 'CLng(3000000000)',
    '  Range("A1").Value = CLng(3000000000)',
    { error: 'Overflow' },
    'Real VBA\'s Long type tops out at 2147483647; CLng(3000000000) raises Overflow.',
    'elixcee returns 3000000000 instead of erroring, same root cause as the CInt cases above (no declared-type-width tracking).');
  addCase('clng_large_negative', CAT, 'CLng(-3000000000)',
    '  Range("A1").Value = CLng(-3000000000)',
    { error: 'Overflow' },
    'Real VBA Long underflow (below -2147483648) also raises Overflow.',
    'elixcee returns -3000000000 instead of erroring.');
}

// ── single_line_if_control_transfer ─────────────────────────────────────────
// Confirms single-line `If cond Then stmt` actually TRANSFERS CONTROL correctly, not just
// that it parses (which compat/corpus/ and this project's own unit tests already cover at
// the parse level) -- exactly the class of bug the Exit/GoTo safety fix earlier this round
// was about (routing Exit/GoTo through the generic identifier-statement parser silently
// turned them into no-ops instead of actually transferring control). All confirmed live to
// behave correctly; all MATCH.
function addNoCellWrittenCase(id, category, description, vbaBody, reason) {
  // A dedicated case shape for "this line must NEVER run" -- represented as absence of
  // any cell in the scenario's output (--json's `cells` array only ever lists non-empty
  // cells, so a guarded-off Range write simply never appears), distinct from the normal
  // {kind:'value'} shape which requires a specific address to be present. Requires the
  // scenario to still run successfully (ok:true) with zero cells -- a parse/runtime error
  // is a different outcome entirely and must use addCase's {error:...} shape instead.
  if (cases.some(c => c.id === id)) throw new Error(`duplicate case id: ${id}`);
  cases.push({ id, category, description, vbaSource: `Sub Scenario()\n${vbaBody}\nEnd Sub\n`, entrypoint: 'Scenario' });
  expected[id] = { kind: 'no_cells', reason };
}

{
  const CAT = 'single_line_if_control_transfer';
  addNoCellWrittenCase('single_line_exit_sub_true_exits', CAT,
    'If True Then Exit Sub actually exits (guarded write never runs)',
    '  x = 1\n  If x > 0 Then Exit Sub\n  Range("A1").Value = 99',
    'If the condition is true, Exit Sub must leave the Sub immediately -- the guarded Range write on the next line must never execute, so no cell should appear in the output at all.');
  addCase('single_line_exit_sub_false_continues', CAT, 'If False Then Exit Sub — Sub continues normally',
    '  x = -1\n  If x > 0 Then Exit Sub\n  Range("A1").Value = 99',
    { value: 99 }, 'When the condition is false, execution must fall through to the guarded line normally.');
  addNoCellWrittenCase('single_line_goto_true_jumps', CAT,
    'If True Then GoTo Skip actually jumps (guarded write never runs)',
    '  x = 1\n  If x > 0 Then GoTo Skip\n  Range("A1").Value = 99\nSkip:',
    'If the condition is true, GoTo must jump past the guarded write, so nothing is ever written. (Complementary case single_line_goto_true_jumps_target_runs uses the same jump but adds a write after the label, to confirm execution resumes normally there.)');
  addCase('single_line_goto_true_jumps_target_runs', CAT, 'If True Then GoTo Skip — code after the label still runs',
    '  x = 1\n  If x > 0 Then GoTo Skip\n  Range("A1").Value = 99\nSkip:\n  Range("B1").Value = 1',
    { value: 1, address: 'B1' }, 'Execution resumes normally at the label after the jump.');
  addNoCellWrittenCase('single_line_exit_sub_else_true_branch', CAT,
    'If True Then Exit Sub Else stmt — takes Exit, Sub ends immediately',
    '  x = 1\n  If x > 0 Then Exit Sub Else Range("A1").Value = 1\n  Range("B1").Value = 2',
    'The true branch takes Exit Sub, so neither the Else branch nor anything after the If should ever run.');
  addCase('single_line_exit_sub_else_false_branch', CAT, 'If False Then Exit Sub Else stmt — takes the Else',
    '  x = -1\n  If x > 0 Then Exit Sub Else Range("A1").Value = 1',
    { value: 1 }, 'The false branch takes the Else clause, which must actually run.');
  addCase('single_line_if_inside_loop_with_exit_for', CAT, 'Single-line If + Exit For inside a loop',
    '  total = 0\n  For i = 1 To 10\n    If i > 3 Then Exit For\n    total = total + i\n  Next i\n  Range("A1").Value = total',
    { value: 6 }, 'Single-line If correctly gates Exit For inside a loop body (1+2+3, stops before 4).');
}

// ── exit_statements ──────────────────────────────────────────────────────────
// Exit For/Do/Sub/Function's control-transfer correctness (not just that they parse) --
// nested-loop scoping in particular (Exit For must only exit the *nearest* enclosing loop),
// all confirmed live. All MATCH.
{
  const CAT = 'exit_statements';
  addCase('exit_for_stops_at_right_point', CAT, 'Exit For stops the loop at the expected count',
    '  total = 0\n  For i = 1 To 10\n    If i > 3 Then Exit For\n    total = total + i\n  Next i\n  Range("A1").Value = total',
    { value: 6 }, '1+2+3, stops before adding 4.');
  addCase('exit_do_stops_at_right_point', CAT, 'Exit Do stops the loop at the expected count',
    '  total = 0\n  i = 0\n  Do\n    i = i + 1\n    If i > 3 Then Exit Do\n    total = total + i\n  Loop\n  Range("A1").Value = total',
    { value: 6 }, 'Same accumulation as Exit For, via Do/Loop instead.');
  addNoCellWrittenCase('exit_sub_stops_execution', CAT,
    'Exit Sub stops execution immediately (guarded write never runs)',
    '  Exit Sub\n  Range("A1").Value = 99',
    'Exit Sub as the very first statement must prevent everything after it from running.');
  addCaseWithSource('exit_function_returns_already_set_value', CAT,
    'Exit Function returns the value already assigned',
    'Sub Scenario()\n  Range("A1").Value = F(10)\nEnd Sub\n\nFunction F(x)\n  If x > 5 Then\n    F = 99\n    Exit Function\n  End If\n  F = 1\nEnd Function\n',
    { value: 99 }, 'Exit Function immediately after setting the return value must preserve that value, not fall through to the later F = 1 default. Uses a top-level Function alongside Scenario (a Function can\'t be nested inside a Sub in real VBA), via addCaseWithSource rather than addCase\'s single-Sub-body wrapper.');
  addNoCellWrittenCase('nested_exit_for_only_exits_inner_loop', CAT,
    'Exit For in a nested loop only exits the nearest (inner) loop',
    '  count = 0\n  For i = 1 To 3\n    For j = 1 To 3\n      If j = 2 Then Exit For\n      count = count + 1\n    Next j\n  Next i\n  If count <> 3 Then Range("A1").Value = "WRONG: " & count',
    'Exit For must only exit the innermost For (j-loop), letting the outer i-loop keep running -- count should end at exactly 3 (one increment per outer iteration, from j=1, before each inner loop\'s Exit For at j=2). Asserted as "no cell written" so a wrong count is visible as a genuine A1 value rather than silently matching by coincidence.');
  addCase('nested_exit_for_count_value', CAT, 'Exit For in a nested loop — the actual accumulated count',
    '  count = 0\n  For i = 1 To 3\n    For j = 1 To 3\n      If j = 2 Then Exit For\n      count = count + 1\n    Next j\n  Next i\n  Range("A1").Value = count',
    { value: 3 }, 'One increment per outer iteration (j=1 only, before each inner Exit For at j=2) — 3 outer iterations, count=3.');
}

// ── object_nothing_access ────────────────────────────────────────────────────
// Real VBA: accessing a member of an unset (Nothing) object variable raises "Object
// variable or With block variable not set" (error 91). elixcee's actual behavior is a
// real, previously-undisclosed gap. (A third natural case here, `r Is Nothing`, needs the
// `Is` operator, which doesn't parse in elixcee at all -- a categorically different,
// parser-level gap already tracked in ROADMAP.md; not forced into this suite's per-value
// schema, which assumes the scenario at least parses.)
{
  const CAT = 'object_nothing_access';
  addNoCellWrittenCase('unset_range_variable_member_write_noop', CAT,
    'Writing through a never-Set Range variable',
    '  Dim r As Range\n  r.Value = 5',
    'Real VBA raises "Object variable or With block variable not set" here. elixcee instead silently no-ops (confirmed: the scenario runs to completion with ok:true and writes nothing at all) — found while building this suite, not previously disclosed.');
  addCase('set_nothing_does_not_clear_reference', CAT,
    'Set r = Nothing does not actually clear the reference',
    '  Dim r As Range\n  Set r = Range("A1")\n  Set r = Nothing\n  r.Value = 5',
    { error: 'Object variable or With block variable not set' },
    'Real VBA: after Set r = Nothing, r no longer refers to Range("A1") -- writing through it raises this error, same as the never-Set case above.',
    'elixcee\'s Set r = Nothing silently no-ops (Nothing isn\'t recognized as a valid Set target), so r still refers to its previous Range("A1") assignment -- confirmed live: this scenario succeeds and writes 5 to A1 instead of erroring. Found while building this suite.');
}

// ── write output ─────────────────────────────────────────────────────────────
fs.writeFileSync(path.join(DIR, 'cases.json'), JSON.stringify(cases, null, 2) + '\n');
fs.writeFileSync(path.join(DIR, 'expected-results.json'), JSON.stringify(expected, null, 2) + '\n');
console.log(`generated ${cases.length} cases -> cases.json, expected-results.json`);
