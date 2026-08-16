// Recursively converts a value into a comparable, JSON.stringify-safe shape WITHOUT
// using JSON.stringify — a naive `JSON.parse(JSON.stringify(v))` round-trip silently
// erases distinctions that matter for compatibility testing:
//   - a real Array hole vs. an explicit `undefined` element vs. an explicit value
//   - a non-index own property on an Array (e.g. a dense worksheet's "!ref") — JSON only
//     serializes numeric indices, dropping anything else
//   - `undefined` (JSON drops the key entirely, or turns it to `null` inside an array)
//   - `NaN` / `Infinity` / `-Infinity` (JSON.stringify silently turns all three into the
//     literal text "null", making them indistinguishable from each other AND from a real
//     null)
//   - `-0` vs `+0` (JSON.stringify renders both as "0")
// This file exists because exactly the first bug above (array + "!ref") shipped
// undetected in Phase 1A's differential harness — see docs/xlsx-architecture.md's
// changelog / compat/differential/xlsx-utils.test.mjs history. Every case below is a
// fixed regression guard, not speculative coverage.
//
// Phase 1B-3: the oracle's own sheet_to_json can hand back a row object whose OWN
// [[Prototype]] has been reassigned to a Date instance (confirmed live: `header:
// ['__proto__']` + a Date-typed cell value hits the `__proto__` accessor's "assign an
// object" branch). `v instanceof Date` is true for such a row too — `instanceof` walks
// the prototype chain looking for `Date.prototype`, and it finds it one hop further out,
// via the corrupted prototype itself. Calling `.toISOString()` on `v` in that branch then
// throws (`this is not a Date object`) because `v` never actually went through `new
// Date()` and has no own [[DateValue]] internal slot — internal slots aren't inherited
// through the prototype chain. `Object.prototype.toString.call(v)` checks that internal
// slot directly, so it's unaffected by prototype tampering and correctly reports such a
// row as `'[object Object]'`, not `'[object Date]'`.
const MAX_DEPTH = 64;

export function normalize(v, seen = new WeakSet(), depth = 0) {
  if (v === undefined) return { __tag: 'undefined' };
  if (v === null) return null;
  if (typeof v === 'number') {
    if (Number.isNaN(v)) return { __tag: 'NaN' };
    if (v === Infinity) return { __tag: 'Infinity' };
    if (v === -Infinity) return { __tag: '-Infinity' };
    if (Object.is(v, -0)) return { __tag: '-0' };
    return v;
  }
  if (typeof v !== 'object') return v; // string, boolean
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return { __tag: 'Date', iso: v.toISOString() };
  }

  // Circular-reference guard: this is a differential-testing harness, not the code under
  // test, so it must never infinite-recurse just because a fixture (accidentally or as a
  // crafted probe) hands it a self-referential structure. A depth cap is a backstop for
  // very deep (but non-circular) structures the WeakSet alone wouldn't catch in time.
  if (seen.has(v)) return { __tag: 'circular' };
  if (depth >= MAX_DEPTH) return { __tag: 'max-depth-exceeded' };
  seen.add(v);

  if (Array.isArray(v)) {
    const out = [];
    const len = v.length;
    for (let i = 0; i < len; ++i) {
      out[i] = i in v ? normalize(v[i], seen, depth + 1) : { __tag: 'hole' };
    }
    // Non-index own enumerable properties (e.g. "!ref" on a dense worksheet array) —
    // Object.keys() on an array returns index keys first (in numeric order) then
    // non-index string keys in insertion order; the index ones are already handled
    // above, so skip them here to avoid double-processing.
    for (const key of Object.keys(v)) {
      if (/^\d+$/.test(key) && Number(key) < len) continue;
      out[key] = normalize(v[key], seen, depth + 1);
    }
    return out;
  }

  const out = {};
  for (const key of Object.keys(v)) out[key] = normalize(v[key], seen, depth + 1);
  return out;
}

// Runnable self-check (no test framework): `node compat/differential/normalize.mjs`
if (import.meta.url === `file://${process.argv[1]}`) {
  const assert = await import('node:assert/strict');

  // The exact regression case that motivated this file: a dense worksheet is an Array
  // with a non-index "!ref" string property.
  {
    const sheet = [];
    sheet[0] = [{ t: 's', v: 'A' }];
    sheet['!ref'] = 'A1:A1';
    const normalized = normalize(sheet);
    assert.equal(normalized['!ref'], 'A1:A1');
    assert.deepEqual(Object.keys(normalized), ['0', '!ref']);
  }

  // undefined
  assert.deepEqual(normalize(undefined), { __tag: 'undefined' });
  assert.deepEqual(normalize({ a: undefined }), { a: { __tag: 'undefined' } });

  // NaN / Infinity / -Infinity are distinguishable from each other and from null
  assert.deepEqual(normalize(NaN), { __tag: 'NaN' });
  assert.deepEqual(normalize(Infinity), { __tag: 'Infinity' });
  assert.deepEqual(normalize(-Infinity), { __tag: '-Infinity' });
  assert.notDeepEqual(normalize(NaN), normalize(null));
  assert.notDeepEqual(normalize(Infinity), normalize(-Infinity));

  // -0 is distinguishable from +0
  assert.deepEqual(normalize(-0), { __tag: '-0' });
  assert.notDeepEqual(normalize(-0), normalize(0));
  assert.equal(normalize(0), 0);

  // Date is preserved as a tagged, comparable value (not silently stringified/dropped)
  const d = new Date('2020-01-01T00:00:00.000Z');
  assert.deepEqual(normalize(d), { __tag: 'Date', iso: '2020-01-01T00:00:00.000Z' });

  // Array non-index property, independent of the dense-worksheet case above
  {
    const arr = [1, 2];
    arr.extra = 'x';
    const n = normalize(arr);
    assert.equal(n[0], 1);
    assert.equal(n[1], 2);
    assert.equal(n.extra, 'x');
    assert.deepEqual(Object.keys(n), ['0', '1', 'extra']);
  }

  // Plain enumerable object property
  assert.deepEqual(normalize({ a: 1, b: 'x' }), { a: 1, b: 'x' });

  // Sparse array hole vs. explicit undefined vs. explicit value are three distinct results
  {
    const sparse = [1];
    sparse[2] = 3; // hole at index 1
    const n = normalize(sparse);
    assert.deepEqual(n[1], { __tag: 'hole' });
    assert.equal(n[0], 1);
    assert.equal(n[2], 3);

    const withUndefined = [1, undefined, 3];
    const n2 = normalize(withUndefined);
    assert.deepEqual(n2[1], { __tag: 'undefined' });
    assert.notDeepEqual(n2[1], n[1]); // hole !== explicit undefined
  }

  // A plain object whose OWN prototype has been reassigned to a Date instance (the exact
  // shape the oracle's sheet_to_json can produce via a "__proto__" header + Date value)
  // must not be misidentified as a Date and must not throw — it has no own [[DateValue]]
  // slot, so it normalizes as an ordinary (empty) object, not a { __tag: 'Date', ... }.
  {
    const corrupted = {};
    Object.setPrototypeOf(corrupted, new Date('2020-01-01T00:00:00.000Z'));
    assert.equal(corrupted instanceof Date, true); // confirms the hazard is real
    assert.doesNotThrow(() => normalize(corrupted));
    assert.deepEqual(normalize(corrupted), {});
  }

  // Circular references must not infinite-recurse.
  {
    const a = { name: 'a' };
    const b = { name: 'b', a };
    a.b = b;
    let n;
    assert.doesNotThrow(() => { n = normalize(a); });
    assert.equal(n.name, 'a');
    assert.equal(n.b.name, 'b');
    assert.deepEqual(n.b.a, { __tag: 'circular' });

    const self = {};
    self.self = self;
    assert.deepEqual(normalize(self), { self: { __tag: 'circular' } });
  }

  // Very deep (non-circular) structures hit the depth cap instead of overflowing the
  // call stack or running unbounded.
  {
    let deep = { __tag: 'leaf' };
    for (let i = 0; i < 1000; ++i) deep = { next: deep };
    let n;
    assert.doesNotThrow(() => { n = normalize(deep); });
    let cur = n;
    let hitCap = false;
    for (let i = 0; i < 1000; ++i) {
      if (cur && cur.__tag === 'max-depth-exceeded') { hitCap = true; break; }
      cur = cur.next;
    }
    assert.equal(hitCap, true);
  }

  console.log('normalize.mjs self-check: all assertions passed');
}
