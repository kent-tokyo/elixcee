// ESM entrypoint — thin re-export of the CJS implementation (the single source of
// truth). Node's ESM loader synthesizes named imports from a CJS module's static
// `module.exports = {...}` shape via cjs-module-lexer, so this works without a build
// step or duplicated logic.
export {
  encode_col,
  decode_col,
  encode_row,
  decode_row,
  encode_cell,
  decode_cell,
  encode_range,
  decode_range,
  split_cell,
  book_new,
  book_append_sheet,
  book_set_sheet_visibility,
  aoa_to_sheet,
} from './index.cjs';
