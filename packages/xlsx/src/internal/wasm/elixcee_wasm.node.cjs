/* @ts-self-types="./elixcee_wasm.d.ts" */

/**
 * Read an in-memory XLSX/XLSM buffer, returning a JSON string shaped like xlsx@0.18.5's
 * `WorkBook` (`{SheetNames, Sheets}`; each `WorkSheet` a sparse `{"A1": {t,v}, ...,
 * "!ref": "A1:C3", "!merges": [...] }` object — see `packages/xlsx/src/index.d.ts`'s
 * `WorkBook`/`WorkSheet` types). The JS side (`packages/xlsx/src/index.cjs`'s `read()`)
 * does `JSON.parse` on the result — no `serde`/`serde_json` dependency needed for a shape
 * this small; reuses `elixcee::diagnostics::json_string`'s existing hand-rolled escaper
 * (src/diagnostics.rs) rather than duplicating a JSON writer or adding a dependency.
 *
 * Deliberately not feature-complete with the oracle's `read()`: no cell formulas (`.f` —
 * `reader.rs` never parses `<f>`), no formatted display text (`.w` — would need
 * `styles.xml` number-format parsing `reader.rs` doesn't do), no date-typed cells (`t:
 * 'd'` — same `styles.xml` gap, so a date serial reads back as a plain number, matching
 * every other `reader.rs` consumer today), no hidden-row/col `!rows`/`!cols` mapping
 * (`reader.rs` has that as coalesced intervals; `WorkSheet` wants a per-index array —
 * left for a follow-up once a caller needs `skipHidden`). Merged ranges ARE mapped
 * (`!merges`) since `reader.rs` already parses them and `sheet_to_html` already consumes
 * that shape.
 * @param {Uint8Array} bytes
 * @returns {string}
 */
function readWorkbook(bytes) {
    let deferred3_0;
    let deferred3_1;
    try {
        const ptr0 = passArray8ToWasm0(bytes, wasm.__wbindgen_malloc);
        const len0 = WASM_VECTOR_LEN;
        const ret = wasm.readWorkbook(ptr0, len0);
        var ptr2 = ret[0];
        var len2 = ret[1];
        if (ret[3]) {
            ptr2 = 0; len2 = 0;
            throw takeFromExternrefTable0(ret[2]);
        }
        deferred3_0 = ptr2;
        deferred3_1 = len2;
        return getStringFromWasm0(ptr2, len2);
    } finally {
        wasm.__wbindgen_free(deferred3_0, deferred3_1, 1);
    }
}
exports.readWorkbook = readWorkbook;
function __wbg_get_imports() {
    const import0 = {
        __proto__: null,
        __wbindgen_cast_0000000000000001: function(arg0, arg1) {
            // Cast intrinsic for `Ref(String) -> Externref`.
            const ret = getStringFromWasm0(arg0, arg1);
            return ret;
        },
        __wbindgen_init_externref_table: function() {
            const table = wasm.__wbindgen_externrefs;
            const offset = table.grow(4);
            table.set(0, undefined);
            table.set(offset + 0, undefined);
            table.set(offset + 1, null);
            table.set(offset + 2, true);
            table.set(offset + 3, false);
        },
    };
    return {
        __proto__: null,
        "./elixcee_wasm_bg.js": import0,
    };
}

function getStringFromWasm0(ptr, len) {
    return decodeText(ptr >>> 0, len);
}

let cachedUint8ArrayMemory0 = null;
function getUint8ArrayMemory0() {
    if (cachedUint8ArrayMemory0 === null || cachedUint8ArrayMemory0.byteLength === 0) {
        cachedUint8ArrayMemory0 = new Uint8Array(wasm.memory.buffer);
    }
    return cachedUint8ArrayMemory0;
}

function passArray8ToWasm0(arg, malloc) {
    const ptr = malloc(arg.length * 1, 1) >>> 0;
    getUint8ArrayMemory0().set(arg, ptr / 1);
    WASM_VECTOR_LEN = arg.length;
    return ptr;
}

function takeFromExternrefTable0(idx) {
    const value = wasm.__wbindgen_externrefs.get(idx);
    wasm.__externref_table_dealloc(idx);
    return value;
}

let cachedTextDecoder = new TextDecoder('utf-8', { ignoreBOM: true, fatal: true });
cachedTextDecoder.decode();
function decodeText(ptr, len) {
    return cachedTextDecoder.decode(getUint8ArrayMemory0().subarray(ptr, ptr + len));
}

let WASM_VECTOR_LEN = 0;

const wasmPath = `${__dirname}/elixcee_wasm_bg.wasm`;
const wasmBytes = require('fs').readFileSync(wasmPath);
const wasmModule = new WebAssembly.Module(wasmBytes);
let wasmInstance = new WebAssembly.Instance(wasmModule, __wbg_get_imports());
let wasm = wasmInstance.exports;
wasm.__wbindgen_start();
