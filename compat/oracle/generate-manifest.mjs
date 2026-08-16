// Generates compat/oracle/api-manifest.json: a machine-derived record of what the real
// `xlsx@0.18.5` package (the oracle) actually exposes at runtime, for both its CJS and
// ESM entrypoints. Never hand-edit api-manifest.json — regenerate it with this script.
//
// Bump GENERATOR_SCRIPT_VERSION whenever this script's introspection logic changes, so a
// stale manifest can be told apart from a freshly-regenerated one at a glance.
const GENERATOR_SCRIPT_VERSION = '0.2.0';

import { createRequire } from 'node:module';
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const here = path.dirname(fileURLToPath(import.meta.url));
const compatDir = path.dirname(here);
const require = createRequire(import.meta.url);

let pkgJson;
try {
  pkgJson = JSON.parse(readFileSync(path.join(compatDir, 'node_modules', 'xlsx', 'package.json'), 'utf8'));
} catch (err) {
  console.error('Could not read node_modules/xlsx/package.json under compat/.');
  console.error('Run `npm install` inside compat/ first.');
  console.error(String(err));
  process.exit(1);
}

if (pkgJson.version !== '0.18.5') {
  console.error(`Expected xlsx@0.18.5 exactly, but node_modules/xlsx is version ${pkgJson.version}.`);
  console.error('compat/package.json must pin an exact version (no ^ or ~) — check it has not drifted.');
  process.exit(1);
}

function describePrototype(ns) {
  if (ns == null) return null;
  const proto = Object.getPrototypeOf(ns);
  if (proto === null) return 'null';
  if (proto === Object.prototype) return 'Object.prototype';
  if (proto === Function.prototype) return 'Function.prototype';
  return proto.constructor?.name ? `other (constructor: ${proto.constructor.name})` : 'other';
}

// Derives enumerableKeys/allOwnKeys from the SAME Object.getOwnPropertyDescriptors call
// (not two independent Object.keys/getOwnPropertyDescriptors calls that could disagree),
// captures symbol-keyed own properties via Reflect.ownKeys (invisible to Object.keys),
// and records each own property's prototype, and, for functions, name/arity.
function introspect(ns) {
  if (ns == null) return null;
  const descs = Object.getOwnPropertyDescriptors(ns);
  const allOwnKeys = Object.keys(descs); // string keys, both enumerable and non-enumerable
  const enumerableKeys = allOwnKeys.filter((k) => descs[k].enumerable);
  const symbolKeys = Reflect.ownKeys(ns)
    .filter((k) => typeof k === 'symbol')
    .map((s) => s.toString());
  const descriptors = {};
  for (const key of allOwnKeys) {
    const d = descs[key];
    const value = d.value;
    const kind = typeof value;
    const entry = {
      kind,
      enumerable: d.enumerable,
      writable: d.writable ?? null,
      configurable: d.configurable,
    };
    if (kind === 'function') {
      entry.name = value.name;
      entry.arity = value.length; // fn.length: declared (non-rest, non-default) param count
    }
    if (kind === 'string' || kind === 'number' || kind === 'boolean') entry.value = value;
    descriptors[key] = entry;
  }
  return { enumerableKeys, allOwnKeys, symbolKeys, prototype: describePrototype(ns), descriptors };
}

function introspectEntrypoint(XLSX) {
  return {
    topLevel: introspect(XLSX),
    utils: introspect(XLSX.utils),
    stream: introspect(XLSX.stream),
  };
}

// No subprocess spawn (avoids exec()-family command-injection footguns for no real
// benefit here): `npm run ...` sets npm_config_user_agent, e.g.
// "npm/11.5.1 node/v24.5.0 darwin arm64 workspaces/false". If the script is invoked
// directly via `node` (no npm parent), this is simply null.
function detectNpmVersion() {
  const ua = process.env.npm_config_user_agent;
  const match = ua && ua.match(/npm\/(\S+)/);
  return match ? match[1] : null;
}

function readLockIntegrity() {
  try {
    const lock = JSON.parse(readFileSync(path.join(compatDir, 'package-lock.json'), 'utf8'));
    const entry = lock.packages?.['node_modules/xlsx'];
    if (!entry) return null;
    return { resolved: entry.resolved ?? null, integrity: entry.integrity ?? null };
  } catch (err) {
    console.warn('Could not read compat/package-lock.json for integrity info:', String(err));
    return null;
  }
}

const xlsxDir = path.join(compatDir, 'node_modules', 'xlsx');
const distFiles = {};
for (const rel of ['xlsx.js', 'xlsx.mjs', 'dist/xlsx.full.min.js', 'bin/xlsx.njs']) {
  distFiles[rel] = existsSync(path.join(xlsxDir, rel));
}

const XLSX_CJS = require('xlsx');

let XLSX_ESM;
let esmResolvedVia;
try {
  XLSX_ESM = await import('xlsx/xlsx.mjs');
  esmResolvedVia = 'xlsx/xlsx.mjs';
} catch {
  XLSX_ESM = await import('xlsx');
  esmResolvedVia = 'xlsx';
}

const manifest = {
  generatedAt: new Date().toISOString(),
  generatorNodeVersion: process.version,
  generatorNpmVersion: detectNpmVersion(),
  generatorScriptVersion: GENERATOR_SCRIPT_VERSION,
  package: {
    requestedVersion: '0.18.5', // exact-pinned in compat/package.json — never ^ or ~
    installedVersion: pkgJson.version,
    main: pkgJson.main ?? null,
    module: pkgJson.module ?? null,
    types: pkgJson.types ?? null,
    bin: pkgJson.bin ?? null,
    browser: pkgJson.browser ?? null,
    exports: pkgJson.exports ?? null,
    sideEffects: pkgJson.sideEffects ?? null,
    dependencies: pkgJson.dependencies ?? null,
    distFiles,
    lock: readLockIntegrity(),
  },
  entrypoints: {
    cjs: { resolvedVia: "require('xlsx')", ...introspectEntrypoint(XLSX_CJS) },
    esm: { resolvedVia: esmResolvedVia, ...introspectEntrypoint(XLSX_ESM) },
  },
};

const outPath = path.join(here, 'api-manifest.json');
writeFileSync(outPath, JSON.stringify(manifest, null, 2) + '\n');

console.log(`Wrote ${outPath}`);
console.log(`  xlsx version: ${manifest.package.installedVersion} (lock integrity: ${manifest.package.lock?.integrity ? 'present' : 'MISSING'})`);
console.log(`  cjs top-level keys: ${manifest.entrypoints.cjs.topLevel.enumerableKeys.length}`);
console.log(`  cjs utils keys: ${manifest.entrypoints.cjs.utils?.enumerableKeys.length ?? 'n/a'}`);
console.log(`  esm resolved via: ${manifest.entrypoints.esm.resolvedVia}`);
