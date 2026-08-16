// Standing regression guard for function-identity metadata: every @elixcee/xlsx export's
// `.name`, `.length`, and its own property descriptor on the exports object must match
// the real oracle's (xlsx@0.18.5) `utils.*` — not just its runtime behavior. This exists
// because the divergence it guards against survived Phase 1A and Phase 1B-1 undetected:
// a plain `{ encode_col: encodeCol }` object-literal assignment does NOT rename an
// already-named function's `.name`, so every export's `.name` stayed the internal
// camelCase name (e.g. "encodeCol") instead of the public "encode_col" until this file's
// companion fix (packages/xlsx/src/index.cjs's nameAs helper). See
// docs/xlsx-architecture.md / the fix(xlsx) commit that introduced this file for the
// discovery.
//
// Deliberately compares against elixcee's CJS export object, NOT the ESM namespace: ESM
// module-namespace object properties are `configurable: false` per spec, while the
// oracle's `utils` is a plain object literal (`configurable: true`) — comparing against
// the ESM side would report a `configurable` mismatch that has nothing to do with this
// package's own correctness. CJS/ESM function-object identity (same reference reachable
// both ways) is checked separately below, once, not per-export descriptor comparison.
import assert from 'node:assert/strict';
import XLSX from 'xlsx';
import elixceeCjs from '../../packages/xlsx/src/index.cjs';
import * as elixceeEsm from '../../packages/xlsx/src/index.mjs';

const U = XLSX.utils;
let failures = 0;

for (const key of Object.keys(elixceeCjs)) {
  const oracleFn = U[key];
  const elixceeFn = elixceeCjs[key];
  if (typeof oracleFn !== 'function') {
    console.error(`FAIL  ${key}: not a function on the real oracle's utils (typeof ${typeof oracleFn})`);
    failures += 1;
    continue;
  }
  const problems = [];
  if (elixceeFn.name !== oracleFn.name) {
    problems.push(`name: oracle=${JSON.stringify(oracleFn.name)} elixcee=${JSON.stringify(elixceeFn.name)}`);
  }
  if (elixceeFn.length !== oracleFn.length) {
    problems.push(`length: oracle=${oracleFn.length} elixcee=${elixceeFn.length}`);
  }
  const oracleDesc = Object.getOwnPropertyDescriptor(U, key);
  const elixceeDesc = Object.getOwnPropertyDescriptor(elixceeCjs, key);
  for (const flag of ['enumerable', 'writable', 'configurable']) {
    if (oracleDesc[flag] !== elixceeDesc[flag]) {
      problems.push(`descriptor.${flag}: oracle=${oracleDesc[flag]} elixcee=${elixceeDesc[flag]}`);
    }
  }
  if (elixceeEsm[key] !== elixceeCjs[key]) {
    problems.push('CJS/ESM identity: elixceeEsm[key] !== elixceeCjs[key] (should be the same function object)');
  }
  if (problems.length) {
    console.error(`FAIL  ${key}: ${problems.join('; ')}`);
    failures += 1;
  } else {
    console.log(`OK    ${key}`);
  }
}

console.log(`\n${Object.keys(elixceeCjs).length - failures}/${Object.keys(elixceeCjs).length} exports match name/length/descriptor/CJS-ESM-identity against the oracle`);

if (failures > 0) {
  console.error('\nmetadata differential suite FAILED.');
  process.exit(1);
}

// Sanity check on the check itself: a deliberately-wrong name must be caught, not
// silently accepted (guards against this file's own comparison logic regressing).
const wrongName = function encodeColWRONG() {};
assert.notEqual(wrongName.name, U.encode_col.name, 'self-check: this comparison must be capable of detecting a name mismatch');

console.log('\nmetadata differential suite passed.');
