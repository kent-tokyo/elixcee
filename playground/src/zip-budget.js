const MAX_ZIP_ENTRIES = 5000;
const MAX_ZIP_UNCOMPRESSED_BYTES = 256 * 1024 * 1024;
const MAX_ZIP_COMPRESSION_RATIO = 200;

/**
 * Validate the bounded ZIP shape before handing untrusted bytes to the XLSX reader.
 * This is deliberately a preflight guard, not a complete ZIP implementation.
 */
export function inspectZipBudget(bytes) {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const lower = Math.max(0, data.byteLength - 0xffff - 22);
  let eocd = -1;
  for (let offset = data.byteLength - 22; offset >= lower; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) { eocd = offset; break; }
  }
  if (eocd < 0 || eocd + 22 > data.byteLength) throw new Error('missing ZIP end record');
  const entries = view.getUint16(eocd + 10, true);
  const centralSize = view.getUint32(eocd + 12, true);
  const centralOffset = view.getUint32(eocd + 16, true);
  if (entries === 0xffff || centralSize === 0xffffffff || centralOffset === 0xffffffff) throw new Error('ZIP64 is outside browser budget');
  if (entries > MAX_ZIP_ENTRIES || centralOffset > data.byteLength || centralSize > data.byteLength - centralOffset) throw new Error('ZIP entry budget exceeded');
  let offset = centralOffset;
  let uncompressed = 0;
  let hasExternalLinks = false;
  const unsupportedParts = new Set();
  const decoder = new TextDecoder();
  for (let index = 0; index < entries; index += 1) {
    if (offset + 46 > data.byteLength || view.getUint32(offset, true) !== 0x02014b50) throw new Error('invalid ZIP central directory');
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const name = decoder.decode(data.subarray(offset + 46, offset + 46 + nameLength));
    if (name.startsWith('xl/externalLinks/')) hasExternalLinks = true;
    if (name.startsWith('xl/pivotTables/')) unsupportedParts.add('pivotTables');
    if (name.startsWith('xl/pivotCache/')) unsupportedParts.add('pivotCache');
    if (name.startsWith('xl/media/')) unsupportedParts.add('media');
    if (name.startsWith('xl/threadedComments/')) unsupportedParts.add('threadedComments');
    if (name.startsWith('xl/slicer') || name.startsWith('xl/slicers/')) unsupportedParts.add('slicers');
    if (name.startsWith('xl/customXml/')) unsupportedParts.add('customXml');
    if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff) throw new Error('ZIP64 entry is outside browser budget');
    uncompressed += uncompressedSize;
    if (uncompressed > MAX_ZIP_UNCOMPRESSED_BYTES || (compressedSize > 0 && uncompressedSize > compressedSize * MAX_ZIP_COMPRESSION_RATIO)) throw new Error('ZIP expansion budget exceeded');
    offset += 46 + nameLength + extraLength + commentLength;
    if (offset > centralOffset + centralSize) throw new Error('invalid ZIP central directory length');
  }
  if (offset !== centralOffset + centralSize) throw new Error('ZIP central directory length mismatch');
  return { entries, uncompressed, hasExternalLinks, unsupportedParts: [...unsupportedParts].sort() };
}

/**
 * Extract one bounded ZIP entry after inspectZipBudget() has approved the archive.
 * This intentionally supports only the two methods used by OOXML: stored and
 * raw deflate. It is used for the opaque XLSM vbaProject.bin payload; the VBA
 * project is preserved, never parsed or executed in the browser.
 */
export async function extractZipEntry(bytes, wantedName) {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  inspectZipBudget(data);
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const decoder = new TextDecoder();
  const lower = Math.max(0, data.byteLength - 0xffff - 22);
  let eocd = -1;
  for (let offset = data.byteLength - 22; offset >= lower; offset -= 1) {
    if (view.getUint32(offset, true) === 0x06054b50) { eocd = offset; break; }
  }
  if (eocd < 0) throw new Error('missing ZIP end record');
  const count = view.getUint16(eocd + 10, true);
  const centralOffset = view.getUint32(eocd + 16, true);
  let offset = centralOffset;
  for (let index = 0; index < count; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) throw new Error('invalid ZIP central directory');
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = decoder.decode(data.subarray(offset + 46, offset + 46 + nameLength));
    offset += 46 + nameLength + extraLength + commentLength;
    if (name !== wantedName) continue;
    if (method !== 0 && method !== 8) throw new Error(`ZIP compression method ${method} is unsupported`);
    if (localOffset + 30 > data.byteLength || view.getUint32(localOffset, true) !== 0x04034b50) throw new Error('invalid ZIP local entry');
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const start = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = data.subarray(start, start + compressedSize);
    if (start < 0 || start + compressedSize > data.byteLength) throw new Error('invalid ZIP entry bounds');
    if (method === 0) return new Uint8Array(compressed);
    if (typeof DecompressionStream !== 'function') throw new Error('browser deflate decompression is unavailable');
    const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  return undefined;
}
