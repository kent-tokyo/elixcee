// Real VBA's And/Or/Xor/Not: a genuine Boolean pair (or, for Not, a lone genuine Boolean)
// gets logical evaluation with a Boolean result; anything else numeric-coerces (banker's
// rounding to a whole number first, same as CInt/CLng) and gets real bitwise math with an
// Integer result. This is VBA's own documented distinction, not a truthy/falsy coercion.
import { toVbaBitwiseInt } from './numeric.mjs';

function toI32(n) {
  // JS bitwise ops operate on 32-bit ints; elixcee's own And/Or/Xor/Not operate on i64.
  // Test values in this suite are kept within i32 range so both agree exactly.
  return n | 0;
}

export function vbaAnd(a, b) {
  if (typeof a === 'boolean' && typeof b === 'boolean') return { value: a && b, isBoolean: true };
  return { value: toI32(toVbaBitwiseInt(a)) & toI32(toVbaBitwiseInt(b)), isBoolean: false };
}

export function vbaOr(a, b) {
  if (typeof a === 'boolean' && typeof b === 'boolean') return { value: a || b, isBoolean: true };
  return { value: toI32(toVbaBitwiseInt(a)) | toI32(toVbaBitwiseInt(b)), isBoolean: false };
}

export function vbaXor(a, b) {
  if (typeof a === 'boolean' && typeof b === 'boolean') return { value: a !== b, isBoolean: true };
  return { value: toI32(toVbaBitwiseInt(a)) ^ toI32(toVbaBitwiseInt(b)), isBoolean: false };
}

export function vbaNot(a) {
  if (typeof a === 'boolean') return { value: !a, isBoolean: true };
  return { value: ~toI32(toVbaBitwiseInt(a)), isBoolean: false };
}
