// Pure-JS reference implementations of real VBA's documented string/conversion semantics.
// See numeric.mjs's header for why these exist as small, independently-checkable functions
// rather than hand-typed expected values.

// VBA's Str(): like CStr, but reserves a leading space for the sign position on a
// non-negative NUMBER. Str(459) == " 459", Str(-459) == "-459", CStr(459) == "459".
export function vbaStr(n) {
  return n >= 0 ? ` ${n}` : `${n}`;
}

// VBA's IsNumeric(): true for an already-numeric value, Empty (coerces to 0), or a string
// that parses as a plain decimal/scientific-notation number after trimming whitespace.
export function vbaIsNumeric(v) {
  if (v === null) return true; // Empty, in this suite's JS-side modeling
  if (typeof v === 'number') return true;
  if (typeof v === 'string') return v.trim() !== '' && !Number.isNaN(Number(v.trim()));
  return false;
}

// VBA's Val(): parses a leading numeric prefix (optional sign, digits, one decimal
// point + digits) and stops at the first character that doesn't fit. Returns 0 if no
// valid numeric prefix exists at all. Does NOT attempt VBA's documented embedded-
// whitespace-stripping inside the prefix -- matches this suite's own elixcee-side scope.
export function vbaVal(s) {
  const trimmed = s.replace(/^\s+/, '');
  const m = trimmed.match(/^[+-]?(\d+(\.\d+)?|\.\d+)/);
  if (!m) return 0;
  return Number(m[0]);
}
