// Pure-JS reference implementations of real VBA's documented numeric semantics, used by
// generate-cases.mjs to COMPUTE expected values programmatically rather than hand-typing
// (and risking arithmetic mistakes in) hundreds of individual expected values. Each
// function here is deliberately small and independently checkable against its own doc
// comment — the correctness of the whole suite rests on these being right.
//
// todaySerial() is the one exception to "computed at generation time": report.mjs calls
// it fresh, at check time, for the date_matches_today case — baking today's serial into
// expected-results.json at generation time would silently go stale the very next day.

// Whole-day Excel serial for "today" (UTC), same 25569 epoch offset (Excel serial of
// 1970-01-01) as src/vm/mod.rs's own unix_epoch_days()/"date" arm and the formula
// engine's func_now() -- kept independent rather than shared, same reasoning as the
// Rust side's own comment on this: one small constant duplicated is cheaper than a new
// cross-language dependency.
export function todaySerial() {
  const unixDays = Math.floor(Date.now() / 86400000);
  return unixDays + 25569;
}

// "YYYY-MM-DD" (UTC) for today — matches elixcee's own --json cell-value representation
// of a Date value (serial_to_display's format!("{:04}-{:02}-{:02}", ...)), which is a
// formatted display string, not the raw serial number. UTC, not local time, to match
// unix_epoch_days()'s own UTC-epoch-seconds-based computation.
export function todayDateString() {
  const d = new Date(Date.now());
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

// VBA's Int(): floors toward negative infinity. Int(-3.1) == -4.
export function vbaInt(f) {
  return Math.floor(f);
}

// VBA's Fix(): truncates toward zero. Fix(-3.9) == -3, not -4.
export function vbaFix(f) {
  return Math.trunc(f);
}

// VBA's Sgn(): -1 / 0 / 1.
export function vbaSgn(f) {
  if (f > 0) return 1;
  if (f < 0) return -1;
  return 0;
}

// Round-half-to-even ("banker's rounding") at a given decimal digit count.
// Real VBA's Round() and CInt/CLng both use this — NOT JavaScript/Rust's
// own default round-half-away-from-zero.
export function bankersRound(f, digits = 0) {
  const factor = 10 ** digits;
  const scaled = f * factor;
  const floor = Math.floor(scaled);
  const diff = scaled - floor;
  let rounded;
  if (diff < 0.5) rounded = floor;
  else if (diff > 0.5) rounded = floor + 1;
  else rounded = floor % 2 === 0 ? floor : floor + 1; // exact tie -> nearest even
  return rounded / factor;
}

// Excel's own ROUND() worksheet formula (and WorksheetFunction.Round): round-half-
// away-from-zero. A genuinely different function from VBA's own Round(), not an alias.
export function excelRound(f, digits = 0) {
  const factor = 10 ** digits;
  const scaled = f * factor;
  const rounded = f >= 0 ? Math.floor(scaled + 0.5) : Math.ceil(scaled - 0.5);
  return rounded / factor;
}

// CInt/CLng: banker's rounding to a whole number (same rule as Round(), digits=0).
export function vbaCIntCLng(f) {
  return bankersRound(f, 0);
}

// VBA's \\ and Mod operators round each operand to a whole number first (banker's
// rounding), THEN perform integer division / modulus on the rounded operands.
export function vbaIntDivOperands(a, b) {
  return [bankersRound(a, 0), bankersRound(b, 0)];
}

export function vbaIntDiv(a, b) {
  const [ra, rb] = vbaIntDivOperands(a, b);
  if (rb === 0) return { error: 'Division by zero' };
  return { value: Math.trunc(ra / rb) };
}

export function vbaMod(a, b) {
  const [ra, rb] = vbaIntDivOperands(a, b);
  if (rb === 0) return { error: 'Division by zero' };
  // VBA's Mod result takes the sign of the dividend (ra), same as Rust's `%` and JS's `%`.
  return { value: ra % rb };
}

// VBA's numeric-context coercion for And/Or/Xor/Not on non-Boolean operands: round to a
// whole number first (same banker's rounding), then real bitwise math on that integer.
// A Boolean's own bit pattern is True == -1 (all-ones), False == 0.
export function toVbaBitwiseInt(v) {
  if (typeof v === 'boolean') return v ? -1 : 0;
  return bankersRound(v, 0);
}
