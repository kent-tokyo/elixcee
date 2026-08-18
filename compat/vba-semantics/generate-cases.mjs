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
  // A vbaBody containing its own Sub/Function/Option declaration means it needs a full
  // module (addCaseWithSource) instead -- wrapping it in another Sub Scenario() here would
  // silently produce a nested-Sub-inside-Sub mess that isn't the VBA the case claims to
  // test. Caught twice by review this round (once fixed before shipping, once not); this
  // guard makes a third instance a thrown error instead of a silently-wrong case.
  if (/^\s*(Sub|Function|Option)\s/mi.test(vbaBody)) {
    throw new Error(`addCase('${id}'): vbaBody looks like a full module (contains Sub/Function/Option) -- use addCaseWithSource instead`);
  }
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
// variable or With block variable not set" (error 91). The first two cases below were
// disclosed KNOWN_LIMITATIONs when this category was written; both are FIXED now (the VM
// grew a real ObjectRef::Nothing state, `Dim r As Range` registers it, `Set r = Nothing`
// assigns it, and every member-access path checks it -- see src/vm/mod.rs's
// `require_live_object`), so their knownLimitation annotations are gone rather than
// weakened. The `Is` operator now parses for the `Is Nothing` shape specifically, which
// is what makes the state-observing cases below expressible at all.
{
  const CAT = 'object_nothing_access';
  addCase('unset_range_variable_member_write_noop', CAT, 'Writing through a never-Set Range variable',
    '  Dim r As Range\n  r.Value = 5',
    { error: 'Object variable or With block variable not set' },
    'Real VBA raises this error for any member access through an object variable that was never Set — a Dim As Range/Object alone only declares the variable, it does not assign a live reference. (Fixed this round: elixcee used to silently no-op here.)');
  addCase('set_nothing_does_not_clear_reference', CAT,
    'Set r = Nothing actually clears the reference',
    '  Dim r As Range\n  Set r = Range("A1")\n  Set r = Nothing\n  r.Value = 5',
    { error: 'Object variable or With block variable not set' },
    'Real VBA: after Set r = Nothing, r no longer refers to Range("A1") -- writing through it raises this error, same as the never-Set case above. (Fixed this round: elixcee\'s Set r = Nothing used to silently no-op, leaving the previous reference live.)');
  addCase('unset_range_variable_member_read_errors', CAT,
    'Reading through a never-Set Range variable also raises error 91',
    '  Dim r As Range\n  x = r.Value\n  Range("A1").Value = x',
    { error: 'Object variable or With block variable not set' },
    'Error 91 is raised by *any* member access through an object variable holding no reference, not only by a write -- the read path must check the same state as the write path.');
  addCase('unset_range_variable_is_nothing_is_true', CAT,
    'A declared-but-never-Set Range variable Is Nothing',
    '  Dim r As Range\n  Range("A1").Value = (r Is Nothing)',
    { value: true },
    'A Dim As Range declaration creates the variable but assigns no reference, so it holds the null object reference -- `r Is Nothing` is True until a Set gives it one.');
  addCase('set_range_variable_is_not_nothing', CAT,
    'A Set object variable is not Nothing',
    '  Dim r As Range\n  Set r = Range("A1")\n  Range("A1").Value = (r Is Nothing)',
    { value: false },
    'Once Set assigns a live Range reference, `r Is Nothing` is False -- the complement of unset_range_variable_is_nothing_is_true, proving the state actually changes rather than being a constant.');
  addCase('set_nothing_makes_is_nothing_true_again', CAT,
    'Set r = Nothing makes Is Nothing True again',
    '  Dim r As Range\n  Set r = Range("B1")\n  Set r = Nothing\n  Range("A1").Value = (r Is Nothing)',
    { value: true },
    'Set r = Nothing assigns the null object reference, returning the variable to exactly the state a never-Set declaration leaves it in -- real VBA cannot distinguish the two either.');
  addCase('set_nothing_does_not_clear_an_alias', CAT,
    'Set r = Nothing does not affect a variable previously assigned from r',
    '  Dim r As Range\n  Dim r2 As Range\n  Set r = Range("B1")\n  Set r2 = r\n  Set r = Nothing\n  Range("A1").Value = (r2 Is Nothing)',
    { value: false },
    'Set copies the *reference* into r2\'s own variable slot; clearing r afterwards rebinds only r. Real VBA has no way for one variable\'s assignment to reach into another\'s -- r2 still refers to the same Range("B1") object.');
  addCase('alias_still_reads_through_after_original_cleared', CAT,
    'An alias still reads its object after the original was Set to Nothing',
    '  Range("B1").Value = 42\n  Dim r As Range\n  Dim r2 As Range\n  Set r = Range("B1")\n  Set r2 = r\n  Set r = Nothing\n  Range("A1").Value = r2.Value',
    { value: 42 },
    'The stronger form of set_nothing_does_not_clear_an_alias: not only is r2 not Nothing, it still resolves to the same cell -- a member access through it returns that cell\'s value rather than raising error 91.');
  addCase('alias_still_writes_through_after_original_cleared', CAT,
    'An alias still writes through after the original was Set to Nothing',
    '  Dim r As Range\n  Dim r2 As Range\n  Set r = Range("A1")\n  Set r2 = r\n  Set r = Nothing\n  r2.Value = 7',
    { value: 7 },
    'Third independent confirmation of the same alias rule, on the write path: r2 must still be a live reference to Range("A1") after r was cleared.');
  addCase('set_from_an_unset_variable_stays_nothing', CAT,
    'Set r2 = r where r is unset leaves r2 Nothing too',
    '  Dim r As Range\n  Dim r2 As Range\n  Set r2 = r\n  Range("A1").Value = (r2 Is Nothing)',
    { value: true },
    'Assigning from a variable that holds the null object reference is legal VBA and simply copies that null reference -- it is not itself an error, and r2 ends up Nothing.');
  addCase('scalar_variable_assignment_is_unaffected_by_object_tracking', CAT,
    'A plain (non-object) variable still assigns with = and is untouched by Nothing tracking',
    '  Dim x As Long\n  x = 5\n  x = x + 1\n  Range("A1").Value = x',
    { value: 6 },
    'Scalar and object variables have genuinely different assignment semantics in VBA (`x = 5` vs `Set r = ...`), and they live in separate namespaces here. A Dim As Long must never acquire object-variable state -- if it did, `x = x + 1` would start raising error 91.');
  addCase('udt_field_assignment_without_a_dim_still_works', CAT,
    'A `.field = value` write on a name that is not an object variable still auto-creates a record',
    '  p.x = 3\n  Range("A1").Value = p.x',
    { value: 3 },
    'Guard against over-reaching: only a name registered as a declared *object* variable may raise error 91. A name that is not an object variable at all keeps its pre-existing record behavior, so ordinary UDT-style field writes are unaffected.');
}

// ── operator_coercion ────────────────────────────────────────────────────────
// The + vs & type-coercion rules, sourced directly from Microsoft's own VBA language
// reference (learn.microsoft.com/.../plus-operator and .../ampersand-operator, fetched
// live while building this category, not recalled from memory) rather than folklore.
// Every case verified live against elixcee before being encoded. Found (and, where a
// narrow well-understood fix, corrected rather than just disclosed) one real bug this way:
// VBA represents Boolean True as -1 internally (CInt(True) = -1), but elixcee's VBA-side
// numeric coercion (src/vm/mod.rs's to_f64) was treating it as 1 -- a one-line, unambiguous
// constant fix, distinct from Excel *worksheet formula* semantics (where TRUE genuinely is
// 1 in arithmetic; src/formula/eval.rs's own separate to_float is correct as-is and was not
// touched).
{
  const CAT = 'operator_coercion';
  addCase('plus_variant_string_plus_variant_number', CAT,
    'Variant holding a numeric string + Variant holding a number adds numerically',
    '  Dim Var1, Var2\n  Var1 = "34"\n  Var2 = 6\n  Range("A1").Value = Var1 + Var2',
    { value: 40 },
    'Per the + operator\'s documented Variant rules: "One Variant expression is numeric and the other is a string -> Add." Matches Microsoft\'s own worked example exactly.');
  addCase('plus_both_variant_strings_concatenates', CAT,
    'Two Variants that both hold strings concatenate with +, even if both look numeric',
    '  Dim Var1, Var2\n  Var1 = "34"\n  Var2 = "6"\n  Range("A1").Value = Var1 + Var2',
    { value: '346' },
    'Per the + operator\'s documented Variant rules: "Both Variant expressions are strings -> Concatenate" -- real VBA returns the *string* "346" here, not the number 40, precisely because + checks the operands\' stored type before their content. Matches Microsoft\'s own worked example exactly.',
    'elixcee has no per-Variant stored-type tag distinguishing "declared Variant, currently holding a numeric-looking string" from "declared Variant, currently holding a genuine number" -- its + operator always numeric-parses both operands when they look numeric, giving the number 40 instead of real VBA\'s string "346". Same class of gap as the CInt/CLng overflow limitations: no declared/runtime type-tag tracking beyond the Variant enum\'s own value shape. Found while building this suite via Microsoft\'s own documented example, not previously disclosed.');
  addCase('ampersand_mixed_types_always_concatenates', CAT,
    '& always concatenates regardless of operand type, unlike +',
    '  Dim Var1, Var2\n  Var1 = "34"\n  Var2 = 6\n  Range("A1").Value = Var1 & Var2',
    { value: '346' },
    'The & operator always converts both operands to string and concatenates -- it never numeric-adds, regardless of whether the operands look numeric. This is the documented reason to prefer & over + for concatenation: no ambiguity.');
  addCase('plus_empty_and_number_returns_number_unchanged', CAT,
    'Empty + a number returns that number unchanged',
    '  Dim r\n  x = r + 5\n  Range("A1").Value = x',
    { value: 5 },
    'Documented: "if only one expression is Empty, the other expression is returned unchanged as result." r is Dim\'d but never assigned, so it holds Empty.');
  addCase('plus_empty_and_empty_is_integer_zero', CAT,
    'Empty + Empty is 0 (Integer)',
    '  Dim r1, r2\n  x = r1 + r2\n  Range("A1").Value = x',
    { value: 0 },
    'Documented: "If both expressions are Empty, result is an Integer." Empty numerically behaves as 0, so 0 + 0 = 0.');
  addCase('ampersand_empty_treated_as_zero_length_string', CAT,
    'Empty & a string treats Empty as a zero-length string',
    '  Dim r\n  x = r & "hi"\n  Range("A1").Value = x',
    { value: 'hi' },
    'Documented (ampersand operator): "Any expression that is Empty is also treated as a zero-length string."');
  addCase('ampersand_null_and_string_treats_null_as_empty_string', CAT,
    '& with one Null operand treats Null as a zero-length string',
    '  Dim r\n  r = Null\n  Range("A1").Value = r & "x"',
    { value: 'x' },
    'Documented (ampersand operator): "if only one expression is Null, that expression is treated as a zero-length string ("") when concatenated with the other expression."');
  addCase('ampersand_both_null_propagates_null', CAT,
    '& with both operands Null produces Null (unlike the one-Null case)',
    '  Dim r1, r2\n  r1 = Null\n  r2 = Null\n  Range("A1").Value = IsNull(r1 & r2)',
    { value: true },
    'Documented (ampersand operator): "If both expressions are Null, result is Null." Distinct from the only-one-Null case, which treats the Null side as "". Asserted via IsNull() since a Null cell value is not itself a well-formed --json cell value. (Fixed this round: elixcee used to stringify both operands to "" first, giving IsNull => False.)');
  addCase('plus_null_propagates_null', CAT,
    '+ with a Null operand always produces Null',
    '  Dim r\n  r = Null\n  Range("A1").Value = IsNull(r + 5)',
    { value: true },
    'Documented (+ operator): "If one or both expressions are Null, result is Null" -- unlike + with Empty (which returns the other operand unchanged) or & with one Null (which treats it as ""). Asserted via IsNull() for the same JSON-representability reason as the & case above. (Fixed this round: elixcee used to treat Null exactly like Empty and coerce it to 0.)');
  addCase('plus_boolean_true_uses_negative_one', CAT,
    'True + a number arithmetic-coerces True as -1, not 1',
    '  Range("A1").Value = True + 5',
    { value: 4 },
    'VBA represents Boolean True as -1 internally (CInt(True) = -1) -- + is documented to Add when both operands are numeric data types, and Boolean is explicitly listed as numeric. Fixed live this round: previously elixcee\'s to_f64 coerced True to 1.0, giving 6 instead of 4 (see src/vm/mod.rs).');
  addCase('plus_two_booleans_negative_two', CAT,
    'True + True is -2, matching the -1-per-True rule',
    '  Range("A1").Value = True + True',
    { value: -2 },
    'Direct consequence of True = -1 internally: -1 + -1 = -2. A second, independent confirmation of the same fixed coercion path as plus_boolean_true_uses_negative_one.');
  addCase('cint_of_true_is_negative_one', CAT,
    'CInt(True) is -1, not 1',
    '  Range("A1").Value = CInt(True)',
    { value: -1 },
    'Same documented fact (True internally is -1) verified through an explicit conversion function rather than an arithmetic operator, confirming the fix applies uniformly to to_f64\'s callers.');
  addCase('plus_nonnumeric_string_and_number_errors', CAT,
    '+ between a non-numeric-string Variant and a numeric Variant raises Type mismatch',
    '  Dim Var1, Var2\n  Var1 = "abc"\n  Var2 = 3\n  Range("A1").Value = Var1 + Var2',
    { error: 'Type mismatch' },
    'Documented (+ operator): "One expression is a numeric data type and the other is a String -> A Type mismatch error occurs." (Fixed this round, and narrowly: the wording is applied by a new arith_to_f64 wrapper used only by eval_binop\'s arithmetic arms, NOT by changing the shared to_f64 helper -- its ~53 other call sites, each with its own correct real-VBA wording for its own failure, are untouched. That shared-helper blast radius was the exact reason this stayed disclosed rather than fixed when the category was first written.)');
}

// ── comparison_coercion ──────────────────────────────────────────────────────
// Comparison-operator (</>/<=/>=/=/<>) type-coercion rules, sourced from Microsoft's own
// VBA language reference (.../comparison-operators, fetched live) rather than folklore.
// Every case verified live before being encoded. Found and fixed one real, narrowly-scoped
// bug this way (vba_eq's missing Empty arm, see operator_coercion above); found one
// deliberately-not-fixed divergence (see compare_variant_numeric_always_less_than_string_variant
// below) where the documented rule is a rarely-hit pedantic edge case and "fixing" it would
// invert vba_cmp's much more commonly relied-on numeric-string-vs-number magnitude
// comparison (used by ordinary Select Case / threshold-check code) for every caller, not
// just this one -- a worse trade than leaving it disclosed.
{
  const CAT = 'comparison_coercion';
  addCase('compare_variant_string_and_number_worked_example', CAT,
    'Variant numeric string is numeric-compared against a Variant number',
    '  Dim Var1, Var2\n  Var1 = "5"\n  Var2 = 4\n  Range("A1").Value = (Var1 > Var2)',
    { value: true },
    'Matches Microsoft\'s own worked example exactly: Var1="5", Var2=4, Var1>Var2 is True (numeric comparison, since Var2 is numeric and Var1 can be converted to a number).');
  addCase('compare_variant_numeric_always_less_than_string_variant', CAT,
    'A numeric Variant is documented as always less than a string Variant, regardless of value',
    '  Dim Var1, Var2\n  Var1 = 100\n  Var2 = "5"\n  Range("A1").Value = (Var1 < Var2)',
    { value: true },
    'Documented: "One Variant expression is numeric and the other is a string -> The numeric expression is less than the string expression" -- unconditionally, not by comparing 100 to 5 numerically. A genuinely surprising, well-documented quirk distinct from the case above (where the numeric side successfully being a number-that-can-be-derived-from-a-string still triggers ordinary numeric comparison per a different rule row).',
    'elixcee\'s vba_cmp always attempts a numeric comparison first when both operands parse as numbers (via to_f64), giving false (100 < 5 is false) instead of true. Deliberately not fixed: vba_cmp is also used for Select Case value/range matching, where numeric-string-vs-number magnitude comparison is the overwhelmingly more common and more useful real-world behavior -- inverting it to match this pedantic edge case would break far more than it fixes. Found while building this suite via Microsoft\'s own documented rule, not previously disclosed.');
  addCase('compare_empty_and_number_uses_zero', CAT,
    'Empty numeric-compares as 0 against a number',
    '  Dim Var1, Var2\n  Var1 = 5\n  Var2 = Empty\n  Range("A1").Value = (Var1 > Var2)',
    { value: true },
    'Matches Microsoft\'s own worked example: Var1=5, Var2=Empty, Var1>Var2 is True (numeric comparison using 0 for Empty).');
  addCase('compare_empty_equals_zero', CAT,
    'Empty equals the number 0',
    '  Dim Var1, Var2\n  Var1 = 0\n  Var2 = Empty\n  Range("A1").Value = (Var1 = Var2)',
    { value: true },
    'Matches Microsoft\'s own worked example: Var1=0, Var2=Empty, Var1=Var2 is True. Fixed live this round -- see the vba_eq fix in operator_coercion\'s notes.');
  addCase('compare_empty_equals_empty_string', CAT,
    'Empty equals the zero-length string',
    '  Dim r\n  Range("A1").Value = (r = "")',
    { value: true },
    'Documented: "One expression is Empty and the other is a String -> Perform a string comparison, using a zero-length string ("") as the Empty expression" -- so Empty = "" is True. A second, independent confirmation of the same vba_eq fix via its string-coercion arm rather than its numeric one.');
  addCase('compare_both_empty_are_equal', CAT,
    'Two never-assigned variables (both Empty) are equal',
    '  Dim r1, r2\n  Range("A1").Value = (r1 = r2)',
    { value: true },
    'Documented: "Both Variant expressions are Empty -> The expressions are equal."');
  addCase('compare_null_operand_propagates_null', CAT,
    'Any comparison with a Null operand produces Null, not True or False',
    '  Dim r\n  r = Null\n  Range("A1").Value = IsNull(5 < r)',
    { value: true },
    'Documented (comparison operators table): every comparison operator lists "Null if expression1 or expression2 = Null" as a third, separate outcome alongside True/False. Asserted via IsNull() since a Null cell value is not itself a well-formed --json cell value, same convention as the Null cases in operator_coercion. (Fixed this round: Null used to numeric-coerce to 0, making 5 < r evaluate as 5 < 0 = False.)');
  addCase('compare_two_variant_strings_lexical_not_numeric', CAT,
    'Two string Variants compare lexically even when both look numeric',
    '  Dim Var1, Var2\n  Var1 = "10"\n  Var2 = "9"\n  Range("A1").Value = (Var1 < Var2)',
    { value: true },
    'Documented: "Both Variant expressions are strings -> Perform a string comparison." Lexically, "10" < "9" (the character \'1\' sorts before \'9\'), even though 10 > 9 numerically -- the classic string-vs-numeric-compare gotcha, and correctly implemented in elixcee already (vba_cmp only attempts numeric comparison when at least one side isn\'t a string).');
  addCase('compare_boolean_true_less_than_false', CAT,
    'True < False is True, since True is -1 and False is 0',
    '  Range("A1").Value = (True < False)',
    { value: true },
    'Direct consequence of True\'s documented internal value of -1: -1 < 0 is True. A second, independent confirmation (via vba_cmp -> to_f64 this time, rather than direct arithmetic) that the Boolean-coercion fix in operator_coercion applies uniformly across to_f64\'s callers.');
}

// ── select_case_matching ─────────────────────────────────────────────────────
// Select Case's own documented matching rules: comma-separated value lists, `To` ranges
// (including a reversed, never-matching range), `Is <comparison>` clauses, mixed
// list+range within one Case line, string matching, no-match-no-Else fall-through, and
// first-match-wins semantics (a later Case that would also match is never reached once an
// earlier one already matched). All well-established, unambiguous VBA control-flow
// semantics -- no operator-coercion ambiguity here, unlike the two categories above. All
// confirmed live; all MATCH (no divergences found in this category).
{
  const CAT = 'select_case_matching';
  addCase('select_case_comma_list_matches', CAT, 'Case with a comma-separated value list matches any listed value',
    '  Dim x\n  x = 7\n  Select Case x\n    Case 1 To 5\n      Range("A1").Value = "low"\n    Case 6, 7, 8\n      Range("A1").Value = "mid"\n    Case Is > 8\n      Range("A1").Value = "high"\n    Case Else\n      Range("A1").Value = "none"\n  End Select',
    { value: 'mid' }, '7 is listed explicitly in "Case 6, 7, 8", independent of the To-range and Is-comparison Cases above and below it.');
  addCase('select_case_to_range_matches', CAT, 'Case with a To range matches any value in the inclusive range',
    '  Dim x\n  x = 3\n  Select Case x\n    Case 1 To 5\n      Range("A1").Value = "low"\n    Case 6, 7, 8\n      Range("A1").Value = "mid"\n  End Select',
    { value: 'low' }, '3 falls within the inclusive range 1 To 5.');
  addCase('select_case_is_comparison_matches', CAT, 'Case Is <op> <value> matches via the given comparison',
    '  Dim x\n  x = 20\n  Select Case x\n    Case Is > 8\n      Range("A1").Value = "high"\n    Case 1 To 5\n      Range("A1").Value = "low"\n  End Select',
    { value: 'high' }, '20 > 8, so the Is clause matches; it is checked in the order written, ahead of the range Case below it.');
  addCase('select_case_no_match_no_else_falls_through', CAT, 'No matching Case and no Case Else simply skips the block',
    '  Dim x\n  x = 99\n  Select Case x\n    Case 1 To 5\n      Range("A1").Value = "low"\n  End Select\n  Range("A2").Value = "after"',
    { value: 'after', address: 'A2' }, 'Select Case with no matching Case and no Case Else is not an error -- execution simply continues after End Select, and A1 is never written.');
  addNoCellWrittenCase('select_case_no_match_writes_nothing', CAT,
    'The unmatched Case body above genuinely never runs (A1 stays unwritten)',
    '  Dim x\n  x = 99\n  Select Case x\n    Case 1 To 5\n      Range("A1").Value = "low"\n  End Select',
    'Companion to select_case_no_match_no_else_falls_through, isolating the "guarded write never happens" half of the same scenario as its own no_cells assertion.');
  addCase('select_case_first_match_wins', CAT, 'The first matching Case runs; a later Case that would also match is never reached',
    '  Dim x\n  x = 3\n  Select Case x\n    Case 1, 2, 3\n      Range("A1").Value = "first"\n    Case 3, 4, 5\n      Range("A1").Value = "second"\n  End Select',
    { value: 'first' }, 'x=3 matches both Case clauses, but Select Case evaluates Cases in written order and stops at the first match -- "second" must never be written.');
  addCase('select_case_reversed_range_never_matches', CAT, 'A backwards To range (high To low) matches nothing',
    '  Dim x\n  x = 3\n  Select Case x\n    Case 5 To 1\n      Range("A1").Value = "reversed-matched"\n    Case Else\n      Range("A1").Value = "else"\n  End Select',
    { value: 'else' }, 'VBA\'s Case ... To ... requires the first bound to be the lower one; "5 To 1" is not an error, it is simply a range containing no values, so Case Else runs instead.');
  addCase('select_case_mixed_list_and_range_in_one_case', CAT, 'A single Case line can mix discrete values and a To range',
    '  Dim x\n  x = 15\n  Select Case x\n    Case 1, 10 To 20\n      Range("A1").Value = "mixed"\n  End Select',
    { value: 'mixed' }, '"Case 1, 10 To 20" matches if x is 1 OR within 10 To 20 -- 15 satisfies the range half.');
  addCase('select_case_string_matching', CAT, 'Case matches strings the same way as the = operator',
    '  Dim x\n  x = "banana"\n  Select Case x\n    Case "apple", "banana"\n      Range("A1").Value = "fruit1"\n    Case Else\n      Range("A1").Value = "other"\n  End Select',
    { value: 'fruit1' }, 'Select Case dispatches via the same equality semantics as the = operator (case-insensitive string match, VBA\'s default Option Compare Binary notwithstanding for ASCII letters -- both operands are already same-case here so this case does not probe that separately).');
}

// ── with_block_resolution ────────────────────────────────────────────────────
// With...End With's own documented semantics: bare .member resolves against the With
// target, nested With blocks restore the outer target on exit, a Sub call inside the body
// doesn't disturb the target, and .member works both as an assignment target and inside an
// arbitrary expression. All confirmed live. Also surfaces two genuine, previously-
// undisclosed structural gaps: elixcee's With-target resolution is a parse-time textual
// rewrite keyed to either a literal Range("...") address string or a bare UDT variable name
// (src/parser/mod.rs's parse_with/parse_with_dot_stmt, via with_range_target/with_target
// parser-level state) rather than a runtime-resolved "current With target" stack. That one
// root cause shows up two ways: a computed target like Cells(r, c) can't be represented as
// the compile-time-known string the mechanism needs, and a bare .member nested inside
// another block construct (If/For/Do/Select Case) within the body never reaches it, since
// only parse_with_body's own direct statement loop special-cases a leading Dot token.
// Structural, not a narrow fix -- same class of gap as the Date/Time Variant model and
// CInt/CLng overflow limitations already disclosed elsewhere in this suite.
{
  const CAT = 'with_block_resolution';
  addCase('with_range_resolves_bare_dot_value', CAT, 'A bare .Value inside With Range(...) resolves against that range',
    '  With Range("A1")\n    .Value = 5\n  End With',
    { value: 5 }, 'The most basic documented With usage: .Value inside the body is shorthand for Range("A1").Value.');
  addCase('with_nested_restores_outer_target_on_exit', CAT, 'A nested With restores the outer target once its own End With runs',
    '  With Range("A1")\n    .Value = 1\n    With Range("B1")\n      .Value = 2\n    End With\n    .Value = .Value + 10\n  End With',
    { value: 11 }, 'After the inner With Range("B1")...End With completes, a bare .Value in the outer body must resolve against A1 again, not B1 -- confirms proper target save/restore around nested With blocks.');
  addCase('with_dot_value_usable_in_an_expression', CAT, '.Value can appear inside an expression, not just as a bare assignment target',
    '  With Range("A1")\n    .Value = 3\n    Range("B1").Value = .Value * 2\n  End With',
    { value: 6, address: 'B1' }, 'Confirms .Value resolves correctly when read as part of a larger expression (.Value * 2), not only when it is the entire right-hand side or a standalone assignment target.');
  addCaseWithSource('with_sub_call_does_not_disturb_target', CAT, 'Calling another Sub from inside a With body does not change the With target',
    'Sub Helper()\n  Range("C1").Value = 100\nEnd Sub\nSub Scenario()\n  With Range("A1")\n    .Value = 1\n    Helper\n    .Value = .Value + 1\n  End With\nEnd Sub\n',
    { value: 2 }, 'A Sub call is not itself a block construct that redefines the enclosing With target -- .Value after the Helper call must still mean Range("A1").Value.');
  addCase('with_computed_cells_target_unsupported', CAT, 'With Cells(row, col) as a target does not parse',
    '  With Cells(1, 3)\n    .Value = 42\n  End With',
    { value: 42, address: 'C1' },
    'Real VBA supports any object expression as a With target, including Cells(r, c) -- this line is valid VBA that writes 42 to C1.',
    'elixcee\'s With Range("...") target is a parse-time literal string, not a general expression -- parse_with only recognizes a Range("literal") or Sheets/Worksheets("name") call shape, or a bare identifier (UDT target). Cells(1, 3) falls into none of those and fails to parse entirely (actual: parse_error "expected newline, got LParen"), rather than merely producing a wrong value. Found while building this suite, not previously disclosed.');
  addCase('with_dot_member_inside_nested_if_unsupported', CAT, 'A bare .member inside an If block nested in a With body does not parse',
    '  With Range("A1")\n    If .Value = 0 Then\n      .Value = 7\n    End If\n  End With',
    { value: 7 },
    'Real VBA resolves .Value against the enclosing With target no matter how deeply it is nested inside other block constructs in the body -- this line is valid VBA that writes 7 to A1 (a fresh cell reads as 0).',
    'elixcee\'s bare-.member rewrite only fires in parse_with_body\'s own direct statement loop (which special-cases a leading Dot token before delegating to parse_stmt) -- once execution descends into a nested block\'s own body (parse_if/parse_for/parse_do_loop/parse_select_case, each parsed via ordinary parse_stmt), a leading Dot is simply unrecognized (actual: parse_error "unexpected token starting statement: Dot"), rather than merely producing a wrong value. Same root cause as with_computed_cells_target_unsupported above: With-target resolution is parse-time-textual, not a runtime-resolved stack. Found while building this suite, not previously disclosed.');
}

// ── array_bounds ─────────────────────────────────────────────────────────────
// Array declaration/resize/bounds semantics: default (Option Base 0) LBound, ReDim
// Preserve vs plain ReDim, Erase on a fixed-size array, IsArray, and multi-dimensional
// bounds. All confirmed live. Surfaces several genuine, previously-undisclosed gaps, each
// independent (not one shared root cause like operator_coercion/with_block_resolution):
// `Dim arr(lo To hi)` and `Dim arr()` (dynamic, no size) both fail to parse; `Option Base 1`
// is silently not honored; `UBound(arr, 2)` ignores its dimension argument and always
// returns dimension 1's bound; `Erase` on a fixed Variant array doesn't reset elements to
// Empty; the `Array(...)` builtin function is not implemented at all.
{
  const CAT = 'array_bounds';
  addCase('dim_array_default_lower_bound_is_zero', CAT, 'Dim arr(5) defaults to a zero-based lower bound',
    '  Dim arr(5)\n  Range("A1").Value = LBound(arr)',
    { value: 0 }, 'Documented default: absent an Option Base 1 statement, array lower bounds default to 0.');
  addCase('dim_array_upper_bound_matches_declared_size', CAT, 'Dim arr(5) sets the upper bound to the declared size',
    '  Dim arr(5)\n  Range("A1").Value = UBound(arr)',
    { value: 5 }, 'Dim arr(5) declares indices 0 through 5 inclusive (6 elements) -- UBound is the literal size given, not size-minus-one.');
  addCase('redim_preserve_keeps_first_element', CAT, 'ReDim Preserve keeps existing elements at their original indices',
    '  Dim arr(3)\n  arr(0) = 10\n  arr(1) = 20\n  arr(2) = 30\n  arr(3) = 40\n  ReDim Preserve arr(5)\n  Range("A1").Value = arr(0)',
    { value: 10 }, 'Preserve keeps element 0 unchanged through the resize.');
  addCase('redim_preserve_keeps_last_original_element', CAT, 'ReDim Preserve keeps every original element, not just the first',
    '  Dim arr(3)\n  arr(0) = 10\n  arr(1) = 20\n  arr(2) = 30\n  arr(3) = 40\n  ReDim Preserve arr(5)\n  Range("A1").Value = arr(3)',
    { value: 40 }, 'Companion to redim_preserve_keeps_first_element, checking the last pre-resize element (index 3) instead of the first.');
  addCase('redim_preserve_grows_upper_bound', CAT, 'ReDim Preserve actually grows the array\'s bound',
    '  Dim arr(3)\n  ReDim Preserve arr(5)\n  Range("A1").Value = UBound(arr)',
    { value: 5 }, 'After ReDim Preserve arr(5), UBound must report the new size (5), not the original (3).');
  addCase('redim_without_preserve_clears_elements', CAT, 'A plain ReDim (no Preserve) resets every element back to Empty',
    '  Dim arr(3)\n  arr(0) = 10\n  arr(1) = 20\n  ReDim arr(5)\n  Range("A1").Value = IsEmpty(arr(0))',
    { value: true }, 'Without Preserve, ReDim discards all prior contents -- element 0, previously 10, is Empty again after the resize.');
  addCase('erase_fixed_array_preserves_bounds', CAT, 'Erase on a fixed-size array does not change its bounds',
    '  Dim arr(3)\n  arr(0) = 5\n  Erase arr\n  Range("A1").Value = UBound(arr)',
    { value: 3 }, 'Documented: Erase on a fixed-size (statically declared) array resets element values but does not deallocate or resize it -- UBound stays 3.');
  addCase('is_array_true_for_declared_array', CAT, 'IsArray returns True for a declared array variable',
    '  Dim arr(3)\n  Range("A1").Value = IsArray(arr)',
    { value: true }, 'Basic IsArray usage on an actual array.');
  addCase('is_array_false_for_scalar', CAT, 'IsArray returns False for an ordinary scalar variable',
    '  Dim x\n  x = 5\n  Range("A1").Value = IsArray(x)',
    { value: false }, 'Companion to is_array_true_for_declared_array, confirming IsArray does not just always return True.');
  addCase('two_dimensional_array_write_and_read_round_trips', CAT, 'A 2D array element written at (row, col) reads back correctly at the same indices',
    '  Dim arr(3, 2)\n  arr(2, 1) = 77\n  Range("A1").Value = arr(2, 1)',
    { value: 77 }, 'Basic round-trip correctness check for two-dimensional array storage, independent of the UBound(arr, dimension) gap disclosed below.');
  addCase('lbound_and_ubound_combined_boolean_check', CAT, 'LBound and UBound combined via And, on a differently-sized array than the other cases',
    '  Dim arr(3)\n  Range("A1").Value = (LBound(arr) = 0) And (UBound(arr) = 3)',
    { value: true }, 'A second, independent confirmation of default array bounds (0 To 3) via a differently-sized declaration than dim_array_default_lower_bound_is_zero/dim_array_upper_bound_matches_declared_size (which use arr(5)).');
  addCase('write_past_upper_bound_does_not_corrupt_in_bounds_data', CAT, 'Attempting to write past the declared upper bound (with On Error Resume Next) does not corrupt existing in-bounds elements',
    '  Dim arr(3)\n  arr(3) = 1\n  On Error Resume Next\n  arr(4) = 2\n  Range("A1").Value = (arr(3) = 1)',
    { value: true }, 'Regardless of exactly how the out-of-bounds write at index 4 is handled, the previously-written in-bounds element at index 3 must remain intact.');
  addCase('dim_array_explicit_lower_bound_unsupported', CAT, 'Dim arr(lo To hi) with an explicit non-zero lower bound does not parse',
    '  Dim arr(2 To 8)\n  Range("A1").Value = LBound(arr)',
    { value: 2 },
    'Real VBA supports an explicit lower bound in a Dim/ReDim size clause -- Dim arr(2 To 8) declares indices 2 through 8 inclusive, so LBound(arr) is 2.',
    'elixcee\'s array-declarator parser only accepts a single upper-bound expression per dimension (Dim arr(5)), not a lo To hi pair -- Dim arr(2 To 8) fails to parse entirely (actual: parse_error "expected RParen, got Ident(\\"to\\")"). Found while building this suite, not previously disclosed.');
  addCaseWithSource('option_base_one_not_respected', CAT, 'Option Base 1 does not change the default lower bound',
    'Option Base 1\nSub Scenario()\n  Dim arr(5)\n  Range("A1").Value = LBound(arr)\nEnd Sub\n',
    { value: 1 },
    'Documented: an Option Base 1 statement at module level changes the default lower bound (for Dim declarations that don\'t give an explicit lower bound) from 0 to 1.',
    'elixcee parses Option Base without erroring but does not appear to feed it into array-bound calculation -- LBound(arr) is still 0 after Option Base 1. Found while building this suite, not previously disclosed.');
  addCase('ubound_second_dimension_argument_ignored', CAT, 'UBound(arr, 2) ignores its dimension argument and returns dimension 1\'s bound',
    '  Dim arr(3, 2)\n  Range("A1").Value = UBound(arr, 2)',
    { value: 2 },
    'Documented: UBound\'s optional second argument selects which dimension to report the bound for -- Dim arr(3, 2) declares dimension 1 as 0 To 3 and dimension 2 as 0 To 2, so UBound(arr, 2) is 2.',
    'elixcee\'s UBound(arr, 2) returns 3 (dimension 1\'s bound) instead of 2 -- the dimension argument does not appear to be used to select which stored bound to report, even though the array\'s own storage genuinely is two-dimensional (see two_dimensional_array_write_and_read_round_trips, which confirms independent read/write addressing already works). Found while building this suite, not previously disclosed.');
  addCase('erase_fixed_variant_array_does_not_reset_to_empty', CAT, 'Erase on a fixed Variant array does not reset elements to Empty',
    '  Dim arr(3)\n  arr(0) = 5\n  arr(1) = 10\n  Erase arr\n  Range("A1").Value = IsEmpty(arr(0))',
    { value: true },
    'Documented: Erase on a fixed-size array resets each element to its type\'s default -- for an (implicitly Variant) array, that default is Empty, so IsEmpty(arr(0)) is True immediately after Erase.',
    'elixcee\'s Erase does not appear to reset element values on a fixed-size array -- IsEmpty(arr(0)) is False after Erase (element 0 still holds its previously-assigned 5). Found while building this suite, not previously disclosed.');
  addCase('dim_array_empty_parens_dynamic_declaration_unsupported', CAT, 'Dim arr() (no size, for a later ReDim) does not parse',
    '  Dim arr()\n  ReDim arr(5)\n  Range("A1").Value = UBound(arr)',
    { value: 5 },
    'Real VBA supports declaring a dynamic array with empty parentheses and sizing it later via ReDim -- Dim arr() followed by ReDim arr(5) is valid VBA that gives UBound(arr) = 5.',
    'elixcee\'s array-declarator parser requires at least one dimension-size expression inside the parens -- Dim arr() fails to parse entirely (actual: parse_error "unexpected token in expression: RParen"). Only the ReDim-without-a-prior-sized-Dim spelling works around this. Found while building this suite, not previously disclosed.');
  addCase('array_builtin_function_unsupported', CAT, 'The Array(...) builtin function is not implemented',
    '  Dim arr\n  arr = Array(10, 20, 30)\n  Range("A1").Value = arr(1)',
    { value: 20 },
    'Real VBA\'s Array() builtin constructs a zero-based Variant array from its arguments -- Array(10, 20, 30) has arr(1) = 20 (the middle element).',
    'elixcee has no Array() builtin at all (actual: undefined_sub_or_function "Unknown VBA function: \'array\'") -- the only way to build an array is a Dim/ReDim declaration plus individual element assignments. Found while building this suite, not previously disclosed.');
}

// ── null_propagation ─────────────────────────────────────────────────────────
// VBA's Null ("no valid data", as from a database NULL) is a genuinely different value
// from Empty (an uninitialized Variant), and every operator has its own documented rule
// for it. All of them were fetched live from Microsoft's VBA language reference while
// building this category, not recalled: the + operator page ("If one or both expressions
// are Null expressions, result is Null"), the minus operator page (same sentence), the
// ampersand operator page ("If both expressions are Null, result is Null. However, if only
// one expression is Null, that expression is treated as a zero-length string"), the
// comparison-operators table ("Null if expression1 or expression2 = Null" for all six),
// the And/Or/Xor/Not pages' three-valued truth tables, and the If...Then...Else statement
// page ("If condition is Null, condition is treated as False").
//
// Null is asserted through IsNull()/TypeName()/VarType() rather than by writing it to a
// cell -- the convention operator_coercion's own Null cases already established, since a
// Null cell value is not a well-formed --json cell value.
//
// Deliberately NOT covered: Select Case with a Null test expression. The Select Case
// reference documents only that testexpression is "matched" against each expressionlist,
// and says nothing about Null; deriving an answer from that would be a guess, and this
// suite does not encode guesses. Left uncovered rather than covered wrongly.
{
  const CAT = 'null_propagation';
  // ── Null is not Empty ──────────────────────────────────────────────────────
  addCase('isnull_of_null_is_true', CAT, 'IsNull(Null) is True',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(n)',
    { value: true }, 'IsNull() reports whether a Variant holds the Null value.');
  addCase('isnull_of_empty_is_false', CAT, 'IsNull(Empty) is False',
    '  Dim e\n  Range("A1").Value = IsNull(e)',
    { value: false },
    'Null and Empty are different VBA values: an uninitialized Variant is Empty, and Empty is not Null. IsNull() must distinguish them, not answer for "has no useful value".');
  addCase('isempty_of_null_is_false', CAT, 'IsEmpty(Null) is False',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsEmpty(n)',
    { value: false },
    'The mirror image of isnull_of_empty_is_false: a Variant explicitly assigned Null is no longer uninitialized, so IsEmpty() is False. Together the two cases pin that neither predicate answers for the other value.');
  addCase('isempty_of_empty_is_true', CAT, 'IsEmpty(Empty) is True',
    '  Dim e\n  Range("A1").Value = IsEmpty(e)',
    { value: true },
    'The unchanged baseline of the pair above — splitting IsNull from IsEmpty must not break the ordinary uninitialized-variable answer.');
  addCase('typename_of_null_is_null', CAT, 'TypeName(Null) is "Null"',
    '  Dim n\n  n = Null\n  Range("A1").Value = TypeName(n)',
    { value: 'Null' },
    'TypeName() names the Variant\'s own subtype: a Null-valued Variant reports "Null", not "Empty".');
  addCase('vartype_of_null_is_vbnull', CAT, 'VarType(Null) is 1 (vbNull)',
    '  Dim n\n  n = Null\n  Range("A1").Value = VarType(n)',
    { value: 1 },
    'vbNull is 1 and vbEmpty is 0 — two distinct documented VarType constants, which is itself evidence the language treats them as different values.');
  addCase('vartype_of_empty_is_vbempty', CAT, 'VarType(Empty) is 0 (vbEmpty)',
    '  Dim e\n  Range("A1").Value = VarType(e)',
    { value: 0 },
    'The complement of vartype_of_null_is_vbnull, confirming the split did not simply relabel Empty.');

  // ── Arithmetic propagates Null from either side ────────────────────────────
  addCase('null_plus_number_is_null', CAT, 'Null + number is Null',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(n + 5)',
    { value: true }, 'Documented (+ operator): "If one or both expressions are Null expressions, result is Null."');
  addCase('number_plus_null_is_null', CAT, 'number + Null is Null (right-hand side)',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(5 + n)',
    { value: true },
    'The rule is "one or both expressions", so it must fire from either side — a left-operand-only check would pass the sibling case and fail this one.');
  addCase('null_minus_number_is_null', CAT, 'Null - number is Null',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(n - 1)',
    { value: true }, 'Documented (minus operator): "If one or both expressions are Null expressions, result is Null."');
  addCase('null_times_number_is_null', CAT, 'Null * number is Null',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(n * 3)',
    { value: true },
    'The multiplication operator carries the same documented sentence as + and -, so a third arithmetic operator confirms the rule is applied per-operator-class rather than hard-coded for +.');
  addCase('null_divided_by_number_is_null_not_a_division_error', CAT,
    'Null / number is Null',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(n / 2)',
    { value: true },
    'Null propagation is decided before the operands are coerced to numbers, so division never sees a 0-coerced Null. A Null-coercing implementation would instead compute 0 / 2 = 0.');
  addCase('number_divided_by_null_is_null_not_division_by_zero', CAT,
    'number / Null is Null, not a Division by zero error',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(5 / n)',
    { value: true },
    'The sharpest case for "propagate before coercing": an implementation that coerced Null to 0 would raise Division by zero here instead of returning Null.');
  addCase('null_arithmetic_result_is_not_zero', CAT,
    'A Null arithmetic result is Null, not the number 0',
    '  Dim n\n  n = Null\n  Range("A1").Value = (n + 5 = 0)\n  Range("A1").Value = IsNull(n + 5)',
    { value: true },
    'Second, independent confirmation that n + 5 is genuinely Null rather than a 0 that merely happens to be falsy: TypeName/IsNull, not truthiness, is what distinguishes them.');
  addCase('typename_of_a_null_arithmetic_result_is_null', CAT,
    'TypeName(Null + 5) is "Null"',
    '  Dim n\n  n = Null\n  Range("A1").Value = TypeName(n + 5)',
    { value: 'Null' },
    'Asserts the propagated *value*\'s own type rather than a predicate over it — a Null-coercing implementation would report "Long" here.');
  addCase('empty_plus_number_is_still_the_number_not_null', CAT,
    'Empty + number returns the number unchanged (Empty does NOT propagate)',
    '  Dim e\n  Range("A1").Value = e + 5',
    { value: 5 },
    'Documented (+ operator): "if only one expression is Empty, the other expression is returned unchanged as result." The load-bearing contrast case: adding Null propagation must not accidentally make Empty propagate too.');

  // ── & concatenation: only both-Null propagates ─────────────────────────────
  addCase('ampersand_null_on_the_left_is_empty_string', CAT,
    'Null & string treats the Null as ""',
    '  Dim n\n  n = Null\n  Range("A1").Value = n & "abc"',
    { value: 'abc' },
    'Documented (ampersand operator): "if only one expression is Null, that expression is treated as a zero-length string ("") when concatenated with the other expression."');
  addCase('ampersand_null_on_the_right_is_empty_string', CAT,
    'string & Null treats the Null as ""',
    '  Dim n\n  n = Null\n  Range("A1").Value = "abc" & n',
    { value: 'abc' },
    'Same documented rule from the other side — the one-Null exception is symmetric, and must not render as the text "Null".');
  addCase('ampersand_both_null_is_null_via_typename', CAT,
    'TypeName(Null & Null) is "Null"',
    '  Dim n1, n2\n  n1 = Null\n  n2 = Null\n  Range("A1").Value = TypeName(n1 & n2)',
    { value: 'Null' },
    'Documented (ampersand operator): "If both expressions are Null, result is Null." Asserted via TypeName rather than IsNull for independence from the IsNull cases above — & is the single operator where one Null does not propagate but two do.');

  // ── Comparison operators: every one propagates ─────────────────────────────
  for (const [op, id] of [['<', 'lt'], ['<=', 'le'], ['>', 'gt'], ['>=', 'ge'], ['=', 'eq'], ['<>', 'ne']]) {
    addCase(`compare_${id}_with_null_right_operand_is_null`, CAT,
      `5 ${op} Null is Null`,
      `  Dim n\n  n = Null\n  Range("A1").Value = IsNull(5 ${op} n)`,
      { value: true },
      `Documented (comparison operators table): the ${op} row lists "Null if expression1 or expression2 = Null" as a third outcome alongside True and False. Each of the six operators is checked separately because the table states the rule per-operator.`);
  }
  addCase('compare_with_null_left_operand_is_null', CAT,
    'Null < 5 is Null (left-hand side)',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(n < 5)',
    { value: true },
    '"expression1 or expression2 = Null" — either side triggers it, so the left-operand position needs its own case.');
  addCase('null_equals_null_is_null_not_true', CAT,
    'Null = Null is Null, not True',
    '  Dim n1, n2\n  n1 = Null\n  n2 = Null\n  Range("A1").Value = IsNull(n1 = n2)',
    { value: true },
    'The most counter-intuitive row of the table, and the one an ordinary equality implementation gets wrong: two Nulls are not "equal", because there is no data to compare — the result is Null again.');
  addCase('ordinary_comparison_still_returns_a_boolean', CAT,
    'A comparison with no Null operand still returns a plain Boolean',
    '  Range("A1").Value = (3 < 5)',
    { value: true },
    'Regression guard: adding a third possible comparison outcome must not disturb the ordinary two-valued case.');

  // ── Logical operators: documented three-valued tables ──────────────────────
  addCase('false_and_null_is_false_not_null', CAT, 'False And Null is False',
    '  Dim n\n  n = Null\n  Range("A1").Value = (False And n)',
    { value: false },
    'Documented (And operator truth table): the "False / Null / False" row. Null does NOT uniformly propagate through And — the answer is already determined without knowing the missing operand. This is the case a blanket "any Null makes Null" rule gets wrong.');
  addCase('true_and_null_is_null', CAT, 'True And Null is Null',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(True And n)',
    { value: true },
    'Documented (And operator truth table): the "True / Null / Null" row — here the missing operand genuinely decides the answer, so the result is Null.');
  addCase('true_or_null_is_true_not_null', CAT, 'True Or Null is True',
    '  Dim n\n  n = Null\n  Range("A1").Value = (True Or n)',
    { value: true },
    'Documented (Or operator truth table): the "True / Null / True" row — the Or mirror of False And Null.');
  addCase('false_or_null_is_null', CAT, 'False Or Null is Null',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(False Or n)',
    { value: true },
    'Documented (Or operator truth table): the "False / Null / Null" row. Together with true_or_null_is_true_not_null this pins both halves of Or\'s three-valued behavior.');
  addCase('xor_with_a_null_operand_is_null', CAT, 'True Xor Null is Null',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(True Xor n)',
    { value: true },
    'Documented (Xor operator): "However, if either expression is Null, result is also Null." Unlike And/Or, Xor propagates unconditionally — one operand can never determine an exclusive-or.');
  addCase('not_null_is_null', CAT, 'Not Null is Null',
    '  Dim n\n  n = Null\n  Range("A1").Value = IsNull(Not n)',
    { value: true },
    'Documented (Not operator truth table): the third row, "Null -> Null".');

  // ── Null as a condition ────────────────────────────────────────────────────
  addCase('if_null_condition_is_treated_as_false', CAT,
    'If Null Then takes the Else branch (Null condition is treated as False)',
    '  Dim n\n  n = Null\n  If n Then\n    Range("A1").Value = 1\n  Else\n    Range("A1").Value = 2\n  End If',
    { value: 2 },
    'Documented explicitly on the If...Then...Else statement page: "If condition is Null, condition is treated as False." Not an error, and not True — verified from the reference precisely because this is the one place VBA gives Null a Boolean reading.');
  addCase('if_null_comparison_condition_is_treated_as_false', CAT,
    'A comparison that evaluates to Null is also treated as False by If',
    '  Dim n\n  n = Null\n  If 5 > n Then\n    Range("A1").Value = 1\n  Else\n    Range("A1").Value = 2\n  End If',
    { value: 2 },
    'The realistic form of the rule: the Null does not appear literally in the condition, it arrives as the *result* of a comparison against a Null operand. Combines the comparison rule and the If rule in one scenario.');
  addNoCellWrittenCase('do_while_null_condition_never_enters_the_loop', CAT,
    'A Do While loop with a Null condition never runs its body',
    '  Dim n\n  n = Null\n  Do While n\n    Range("A1").Value = 99\n  Loop',
    'A Null condition being treated as False must apply to loop conditions too, not just If — the body must never execute, so no cell is written at all.');

  // ── Null in a genuinely numeric context ────────────────────────────────────
  addCase('null_passed_to_a_numeric_function_raises_invalid_use_of_null', CAT,
    'Abs(Null) raises "Invalid use of Null"',
    '  Dim n\n  n = Null\n  Range("A1").Value = Abs(n)',
    { error: 'Invalid use of Null' },
    'Where Null cannot propagate — a function argument that must genuinely be a number — real VBA raises run-time error 94, "Invalid use of Null". Unlike Empty, which is documented to behave as 0 in a numeric context, Null has no numeric value at all.');
}

// ── colon_statement_separator ────────────────────────────────────────────────
// Real VBA's `:` multi-statement-per-line separator. Two independent Microsoft
// citations underpin this whole category, both fetched live from the VBA language
// reference while building it rather than recalled: (a) the If...Then...Else statement
// page documents a single-line If's `statements` part as "One or more statements
// separated by colons; executed if condition is True", with the worked example
// `If A > 10 Then A = A + 1 : B = B + A : C = C + B`; (b) the comparison-operators page's
// own example code uses the separator outside any If at all --
// `Var1 = "5": Var2 = 4    ' Initialize variables.` -- confirming it's a general
// statement separator, not an If-only affordance.
// The three cases that a naive "replace `:` with a newline before tokenizing"
// implementation would silently get wrong are pinned deliberately: a `:` inside a string
// literal, a `label:` declaration (one statement, not two), and a single-line If's own
// Then/Else clause boundary (where the colon extends the *branch*, not the enclosing
// statement list). Every case below was verified live against elixcee before being
// encoded, same as every other category here.
{
  const CAT = 'colon_statement_separator';
  addCase('colon_two_statements_one_line', CAT, 'Two assignments separated by a colon on one line',
    '  a = 1: b = 2\n  Range("A1").Value = a + b',
    { value: 3 },
    'Both statements on the colon-separated line must execute, in order -- exactly the form Microsoft\'s own comparison-operators example uses (`Var1 = "5": Var2 = 4`). 1 + 2 = 3 proves neither was dropped.');
  addCase('colon_three_statements_one_line', CAT, 'Three assignments separated by colons on one line',
    '  a = 1: b = 2: c = 3\n  Range("A1").Value = a + b + c',
    { value: 6 },
    'The separator chains: a line is a list of statements, not a pair. 1 + 2 + 3 = 6 proves the third statement ran too, not just the first two.');
  addCase('colon_inside_string_literal_is_not_a_separator', CAT,
    'A colon inside a string literal does not split the statement',
    '  s = "10:30": Range("A1").Value = s',
    { value: '10:30' },
    'A `:` inside a string literal is literal text, not a statement separator -- the string must survive intact AND the statement after the real separator must still run. This is the case a pre-tokenize `:`-to-newline rewrite corrupts.');
  addCase('colon_after_msgbox_with_colon_in_message', CAT,
    'MsgBox "x:y": Exit-adjacent statement on the same line',
    '  MsgBox "10:30"\n  Range("A1").Value = 1',
    { value: 1 },
    'Baseline companion to colon_inside_string_literal_is_not_a_separator: the same colon-bearing string literal as a MsgBox argument on its own line, confirming the literal is not itself the thing under test.');
  addCase('colon_msgbox_then_statement_same_line', CAT,
    'MsgBox "10:30": Range write on the same line',
    '  MsgBox "10:30": Range("A1").Value = 1',
    { value: 1 },
    'The colon after a MsgBox whose argument itself contains a colon must separate the two statements at the *real* separator only -- the Range write runs, and the message text is unaffected.');
  addCase('colon_label_then_statement_same_line', CAT,
    'label: statement — a label and a statement on one line',
    '  GoTo Skip\n  Range("B1").Value = 99\nSkip: Range("A1").Value = 7',
    { value: 7 },
    '`Skip:` is a line-label declaration whose own trailing colon is part of the label syntax, not a separator between two statements -- but a statement may still follow it on the same line, and must execute when the label is jumped to.');
  addNoCellWrittenCase('colon_label_line_still_a_valid_jump_target', CAT,
    'A label with a statement after it on the same line is still jumped over correctly',
    '  GoTo Skip\n  Range("A1").Value = 99\nSkip:',
    'The GoTo must skip the guarded write entirely; nothing is written. Confirms treating `label:` as a label (rather than as an empty statement plus a separator) did not break the label as a jump target.');
  addCase('colon_single_line_if_then_first_statement', CAT,
    'Single-line If Then with a colon-separated statement list — first statement runs',
    '  x = 5\n  If x > 0 Then Range("A1").Value = 1: Range("B1").Value = 2',
    { value: 1 },
    'Documented (If...Then...Else statement): `statements` is "One or more statements separated by colons; executed if condition is True". The first of the two runs when the condition is True.');
  addCase('colon_single_line_if_then_second_statement', CAT,
    'Single-line If Then with a colon-separated statement list — second statement also runs',
    '  x = 5\n  If x > 0 Then Range("A1").Value = 1: Range("B1").Value = 2',
    { value: 2, address: 'B1' },
    'Same documented rule, asserted on the second statement of the Then list -- Microsoft\'s own worked example is `If A > 10 Then A = A + 1 : B = B + A : C = C + B`, where every colon-separated statement is gated by the one condition.');
  addNoCellWrittenCase('colon_single_line_if_false_skips_the_whole_then_list', CAT,
    'A false single-line If skips every colon-separated statement in its Then list',
    '  x = -5\n  If x > 0 Then Range("A1").Value = 1: Range("B1").Value = 2',
    'The colon-separated statements belong to the Then branch, so a False condition must skip ALL of them -- if the second one leaked out into the enclosing statement list it would run unconditionally and write B1.');
  addCase('colon_single_line_if_else_takes_rest_of_line', CAT,
    'Single-line If ... Else with a colon-separated Else list',
    '  x = -5\n  If x > 0 Then Range("C1").Value = 9 Else Range("A1").Value = 1: Range("B1").Value = 2',
    { value: 2, address: 'B1' },
    'A single-line If ends only at end-of-line, and `elsestatements` is documented as "One or more statements" -- so everything after `Else` on the line, including past a colon, belongs to the Else branch and runs together when the condition is False.');
  addNoCellWrittenCase('colon_single_line_if_true_skips_the_whole_else_list', CAT,
    'A true single-line If skips every colon-separated statement in its Else list',
    '  x = 5\n  If x > 0 Then Exit Sub Else Range("A1").Value = 1: Range("B1").Value = 2',
    'The true branch exits the Sub, so neither Else-list statement may run -- the complement of colon_single_line_if_else_takes_rest_of_line, proving the second Else statement is genuinely gated by the condition rather than being an unconditional trailing statement.');
  addNoCellWrittenCase('colon_exit_sub_after_a_statement_on_the_same_line', CAT,
    'Exit Sub as the second colon-separated statement really exits',
    '  x = 1: Exit Sub\n  Range("A1").Value = 99',
    'A control-transfer statement in the second colon position must transfer control for real -- the guarded write on the following line must never run, so no cell appears at all.');
  addCase('colon_inside_for_loop_body_on_the_header_line', CAT,
    'For header, body and Next all on one colon-separated line',
    '  total = 0\n  For i = 1 To 3: total = total + i: Next i\n  Range("A1").Value = total',
    { value: 6 },
    'The colon terminates a block-construct header and each body statement just as a newline does: 1 + 2 + 3 = 6 proves the loop ran its body three times rather than being mis-parsed into a single flat statement list.');
  addCase('colon_inside_do_loop_on_one_line', CAT,
    'Do While header, body and Loop all on one colon-separated line',
    '  i = 0\n  Do While i < 3: i = i + 1: Loop\n  Range("A1").Value = i',
    { value: 3 },
    'A second, independent block construct (Do...Loop rather than For...Next) confirming the separator terminates block headers and bodies generally, not just in the one construct.');
  addCase('colon_inside_with_block_on_one_line', CAT,
    'With header, body and End With all on one colon-separated line',
    '  With Range("A1"): .Value = 5: End With',
    { value: 5 },
    'The separator must work inside a With body too, where a leading `.member` statement -- not an identifier -- follows the colon.');
  addCase('colon_run_of_empty_statements', CAT, 'A run of consecutive colons is an empty statement, not an error',
    '  a = 1:: b = 2\n  Range("A1").Value = a + b',
    { value: 3 },
    'Consecutive separators delimit an empty statement, which does nothing -- both real statements must still run. Guards against an off-by-one in separator consumption.');
  addCase('colon_dim_then_assignment_same_line', CAT, 'Dim and an assignment separated by a colon',
    '  Dim x: x = 5\n  Range("A1").Value = x',
    { value: 5 },
    'A declaration and an assignment on one colon-separated line -- the shape most commonly written by hand in real VBA modules.');
  addCase('colon_named_argument_is_not_a_separator', CAT,
    'A `:=` named argument is not a statement separator',
    '  Range("A1").Value = 1\n  Range("A1").Copy Destination:=Range("B1")\n  Range("C1").Value = Range("B1").Value',
    { value: 1, address: 'C1' },
    'VBA\'s named-argument syntax `Name:=value` contains a colon that is part of the `:=` token, not a statement separator -- the Copy must still receive its Destination, proving the separator logic did not split the argument off.');
}

// ── write output ─────────────────────────────────────────────────────────────
fs.writeFileSync(path.join(DIR, 'cases.json'), JSON.stringify(cases, null, 2) + '\n');
fs.writeFileSync(path.join(DIR, 'expected-results.json'), JSON.stringify(expected, null, 2) + '\n');
console.log(`generated ${cases.length} cases -> cases.json, expected-results.json`);
