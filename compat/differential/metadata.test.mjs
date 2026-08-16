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
  const oracleVal = U[key];
  const elixceeVal = elixceeCjs[key];
  if (typeof oracleVal !== typeof elixceeVal) {
    console.error(`FAIL  ${key}: type mismatch (oracle=${typeof oracleVal} elixcee=${typeof elixceeVal})`);
    failures += 1;
    continue;
  }
  const problems = [];
  if (typeof oracleVal === 'function') {
    // .name/.length are function-only metadata — plain data exports (e.g. `consts`) have
    // neither and are compared purely by descriptor + value shape below instead.
    if (elixceeVal.name !== oracleVal.name) {
      problems.push(`name: oracle=${JSON.stringify(oracleVal.name)} elixcee=${JSON.stringify(elixceeVal.name)}`);
    }
    if (elixceeVal.length !== oracleVal.length) {
      problems.push(`length: oracle=${oracleVal.length} elixcee=${elixceeVal.length}`);
    }
  } else {
    // Data export (e.g. `consts`): compare own enumerable keys and their values/
    // descriptors directly, since there's no .name/.length to check.
    const oracleKeys = Object.keys(oracleVal).sort();
    const elixceeKeys = Object.keys(elixceeVal).sort();
    if (JSON.stringify(oracleKeys) !== JSON.stringify(elixceeKeys)) {
      problems.push(`own keys: oracle=${JSON.stringify(oracleKeys)} elixcee=${JSON.stringify(elixceeKeys)}`);
    } else {
      for (const k of oracleKeys) {
        if (oracleVal[k] !== elixceeVal[k]) problems.push(`${key}.${k}: oracle=${JSON.stringify(oracleVal[k])} elixcee=${JSON.stringify(elixceeVal[k])}`);
      }
    }
  }
  const oracleDesc = Object.getOwnPropertyDescriptor(U, key);
  const elixceeDesc = Object.getOwnPropertyDescriptor(elixceeCjs, key);
  for (const flag of ['enumerable', 'writable', 'configurable']) {
    if (oracleDesc[flag] !== elixceeDesc[flag]) {
      problems.push(`descriptor.${flag}: oracle=${oracleDesc[flag]} elixcee=${elixceeDesc[flag]}`);
    }
  }
  if (elixceeEsm[key] !== elixceeCjs[key]) {
    problems.push('CJS/ESM identity: elixceeEsm[key] !== elixceeCjs[key] (should be the same object/function)');
  }
  if (problems.length) {
    console.error(`FAIL  ${key}: ${problems.join('; ')}`);
    failures += 1;
  } else {
    console.log(`OK    ${key}`);
  }
}

console.log(`\n${Object.keys(elixceeCjs).length - failures}/${Object.keys(elixceeCjs).length} exports match name/length/descriptor/CJS-ESM-identity against the oracle`);

// Key ORDER, not just per-key content — Phase 1C discovered this had never been checked
// (every key matched individually, but elixcee's module.exports literal had been ordered
// by this file's own section comments, not the oracle's actual Object.keys() insertion
// order, undetected across Phases 1A-1B-3). Compares elixcee's own key order against the
// oracle's FULL key order filtered down to only the keys elixcee currently implements —
// this must hold at every point during the utils-completion phases, not just once every
// oracle key is implemented.
{
  const oracleOrder = Object.keys(U);
  const elixceeOrder = Object.keys(elixceeCjs);
  const oracleFiltered = oracleOrder.filter((k) => elixceeOrder.includes(k));
  if (JSON.stringify(elixceeOrder) !== JSON.stringify(oracleFiltered)) {
    console.error('FAIL  key order: elixcee\'s Object.keys() does not match the oracle\'s own relative order');
    console.error(`      elixcee: ${JSON.stringify(elixceeOrder)}`);
    console.error(`      oracle:  ${JSON.stringify(oracleFiltered)}`);
    failures += 1;
  } else {
    console.log('OK    key order matches the oracle\'s own Object.keys(XLSX.utils) relative order');
  }
}

if (failures > 0) {
  console.error('\nmetadata differential suite FAILED.');
  process.exit(1);
}

// Sanity check on the check itself: a deliberately-wrong name must be caught, not
// silently accepted (guards against this file's own comparison logic regressing).
const wrongName = function encodeColWRONG() {};
assert.notEqual(wrongName.name, U.encode_col.name, 'self-check: this comparison must be capable of detecting a name mismatch');

console.log('\nmetadata differential suite passed.');
