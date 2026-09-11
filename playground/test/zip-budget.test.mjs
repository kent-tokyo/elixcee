import assert from 'node:assert/strict';
import { makeZip } from '../../packages/xlsx/src/internal/zip-writer.cjs';
import { inspectZipBudget } from '../src/zip-budget.js';

const bytes = (text) => new TextEncoder().encode(text);
const safe = inspectZipBudget(makeZip([{ name: 'xl/workbook.xml', data: bytes('<workbook/>') }]));
assert.equal(safe.hasExternalLinks, false);

const external = inspectZipBudget(makeZip([
  { name: 'xl/workbook.xml', data: bytes('<workbook/>') },
  { name: 'xl/externalLinks/externalLink1.xml', data: bytes('<externalLink/>') },
]));
assert.equal(external.hasExternalLinks, true);
const unsupported = inspectZipBudget(makeZip([
  { name: 'xl/workbook.xml', data: bytes('<workbook/>') },
  { name: 'xl/pivotCache/pivotCacheDefinition1.xml', data: bytes('<pivotCacheDefinition/>') },
  { name: 'xl/media/image1.png', data: bytes('image') },
]));
assert.deepEqual(unsupported.unsupportedParts, ['media', 'pivotCache']);
console.log('zip budget external-link detection passed');
