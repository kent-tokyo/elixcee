// Minimal XLSX/ODS reader — replaces calamine as a runtime dependency.
// Supports: .xlsx, .xlsm (Office Open XML ZIP), .ods (OpenDocument ZIP).
// Row/col indices are 1-based, matching the VM's convention.

use std::collections::HashMap;
use std::io::{Cursor, Read, Seek};
use zip::ZipArchive;

// ── Public types ──────────────────────────────────────────────────────────────

/// A 1-based inclusive `((row1,col1),(row2,col2))` rect (Milestone B6c2) —
/// a private per-module alias, not a shared type, matching this codebase's
/// existing per-module `col_to_letters` duplication convention rather than
/// a cross-module `utils` dependency.
type MergeRect = ((u32, u32), (u32, u32));

pub struct WorkbookSheet {
    pub name: String,
    pub cells: HashMap<(u32, u32), SheetCell>,
    /// The XLSX `<sheet sheetId="...">` attribute, when read from a real
    /// `.xlsx`/`.xlsm` file — `None` for `.ods` (no equivalent attribute) or
    /// if the attribute was missing. Not VBA's `CodeName` (that lives in
    /// `vbaProject.bin`, an OLE binary format this reader doesn't parse).
    pub sheet_id: Option<String>,
    /// The XLSX `xl/_rels/workbook.xml.rels` relationship id (`<sheet r:id="...">` in
    /// `xl/workbook.xml`) that resolved to this sheet's own part — `None` for `.ods` (no
    /// relationship-id concept) or if the attribute was missing. Currently computed and
    /// then discarded in `read_workbook_from_archive` (used only transiently to resolve
    /// `source_part_name` below); captured here instead so a future writer can preserve a
    /// sheet's original identity across a save rather than always renumbering positionally
    /// — see `docs/xlsx-worksheet-preservation-0.10.0-design.md` §6 (`WorksheetOrigin`).
    pub workbook_rel_id: Option<String>,
    /// The zip entry path this sheet's XML was actually read from (e.g.
    /// `"xl/worksheets/sheet3.xml"`) — `None` for `.ods`. The single most unstable of the
    /// three origin fields on this struct: `save_xlsx_impl` (`src/lib.rs`) renumbers
    /// worksheet parts sequentially by current position on every save, so this reflects
    /// the SOURCE file's naming, not necessarily what a prior elixcee save produced. See
    /// `workbook_rel_id`'s doc comment for why this is captured at all.
    pub source_part_name: Option<String>,
    /// Merged cell ranges, 1-based inclusive (Milestone B6c2) — from XLSX's
    /// `<mergeCells><mergeCell ref="..."/>` or ODS's
    /// `table:number-columns-spanned`/`table:number-rows-spanned` on the
    /// anchor cell. Empty if the sheet has no merges.
    pub merged_ranges: Vec<MergeRect>,
    /// Hidden row intervals, 1-based inclusive `(start, end)` (Milestone
    /// B7b) — from XLSX's `<row hidden="1">`. Always empty for `.ods`
    /// (deferred — see `docs/agent-contract.md`).
    pub hidden_rows: Vec<(u32, u32)>,
    /// Hidden column intervals, 1-based inclusive `(start, end)`
    /// (Milestone B7b) — from XLSX's `<col min=".." max=".." hidden="1">`.
    /// Always empty for `.ods` (deferred).
    pub hidden_columns: Vec<(u32, u32)>,
    /// Per-cell raw `s="N"` index (0-based position in `<cellXfs>`), 1-based
    /// `(row, col)` keys — kept whenever the attribute is present and parses,
    /// regardless of whether that `<xf>`'s own `numFmtId` is 0 (unlike
    /// `BufferSheet::style_ids`, which only keeps a non-zero *resolved*
    /// format id; a style index can carry font/fill/border info under a
    /// General number format, which still needs to survive a save). Lets
    /// `save_xlsx_impl` (`src/lib.rs`) re-emit each surviving cell's
    /// original `s="N"` unchanged — see `docs/xlsx-architecture.md`. Always
    /// empty for `.ods` (no `s`-index concept).
    pub raw_style_indices: HashMap<(u32, u32), u32>,
    /// Per-cell raw `<f>...</f>` formula text, 1-based `(row, col)` keys matching
    /// `cells` — the formula string exactly as written in the XML (no leading `=`),
    /// mirroring `BufferSheet::formulas` (see that field's doc comment for the shared-
    /// formula-follower-cell caveat, which applies here too). Lets `populate_from_sheets`
    /// (`src/vm/mod.rs`) keep a loaded cell's formula alive instead of flattening it to a
    /// bare cached value, which `save_xlsx_impl` would otherwise silently re-emit as a
    /// permanent literal on the very next save. Always empty for `.ods` (not parsed there).
    pub formulas: HashMap<(u32, u32), String>,
}

pub enum SheetCell {
    Integer(i64),
    Float(f64),
    Str(String),
    Bool(bool),
}

/// Read a spreadsheet file into sheets. Supports .xlsx, .xlsm, .ods.
pub fn read_workbook(path: &str) -> Result<Vec<WorkbookSheet>, String> {
    let lower = path.to_lowercase();
    if lower.ends_with(".ods") {
        read_ods(path)
    } else if lower.ends_with(".xlsx") || lower.ends_with(".xlsm") {
        read_xlsx(path)
    } else {
        Err(format!("unsupported file format: {}", path))
    }
}

/// Read an in-memory XLSX/XLSM (Office Open XML ZIP) buffer into sheets — the buffer-
/// first entry point the WASM bridge (`crates/elixcee-wasm`) and `@elixcee/xlsx`'s
/// `XLSX.read()` are built on (see `docs/xlsx-architecture.md`'s "reader.rs buffer-API
/// resolution"). ODS is intentionally not handled here: it's not part of the xlsx-compat
/// surface this entry point exists for, and `read_workbook(path)` above still handles it
/// unchanged for path-based callers.
///
/// Returns `BufferWorkbook`, not `Vec<WorkbookSheet>` — see that type's doc comment for
/// why: the per-cell formula text, declared `<dimension>`, and now (Milestone read-item 6)
/// per-cell number-format-id and workbook-level custom number formats / date1904 this
/// buffer-first API exposes have no home on `WorkbookSheet` itself without touching every
/// one of its other construction sites (`src/vm/mod.rs`'s tests, `src/snapshot.rs`), which
/// are out of scope this phase.
pub fn read_workbook_from_bytes(bytes: &[u8]) -> Result<BufferWorkbook, String> {
    let archive = ZipArchive::new(Cursor::new(bytes)).map_err(|e| e.to_string())?;
    read_workbook_from_archive(archive)
}

/// The buffer-API-only output of `read_workbook_from_bytes`: per-sheet data plus the two
/// workbook-level pieces item 6 needs (custom number formats, date1904) that don't belong
/// on any single sheet. See `BufferSheet`'s doc comment for why this whole tree is kept
/// separate from `WorkbookSheet` rather than growing it.
pub struct BufferWorkbook {
    pub sheets: Vec<BufferSheet>,
    /// Custom number-format definitions from `xl/styles.xml`'s `<numFmts><numFmt
    /// numFmtId="N" formatCode="..."/></numFmts>` — ids below 164 are reserved for
    /// built-ins the oracle's own SSF engine already knows (not duplicated here); this map
    /// only ever holds the file's OWN custom entries, exactly what `xlsx_styles` parsed.
    /// Empty when the sheet has no custom formats or `styles.xml` is absent.
    pub number_formats: HashMap<u32, String>,
    /// Whether the workbook declared `<workbookPr date1904="1"/>` (the 1904 date system) —
    /// from `xl/workbook.xml`, read once for the whole workbook (all sheets share it, this
    /// isn't a per-sheet setting). `false` (the default 1900 system) when absent.
    pub date1904: bool,
}

/// A `WorkbookSheet` plus buffer-API-only data (`read_workbook_from_bytes`) that has no
/// home on `WorkbookSheet` itself: adding a field there would force every existing
/// `WorkbookSheet { .. }` construction site — including `src/vm/mod.rs`'s direct test
/// literals and `src/snapshot.rs` — to list it too, none of which are in this phase's scope
/// (`src/vm/` is frozen/owned elsewhere) or even want this data (the path-based VM/CLI flow
/// has no use for `!ref`/formula text). Kept as a thin wrapper instead, used only by
/// `read_workbook_from_bytes` and its WASM-bridge caller.
pub struct BufferSheet {
    pub sheet: WorkbookSheet,
    /// Per-cell raw `<f>...</f>` formula text, 1-based `(row, col)` keys matching
    /// `sheet.cells` — the formula string exactly as written in the XML (no leading `=`,
    /// matching the oracle's own `.f` convention), unescaped the same way cell/shared-string
    /// text already is. Shared/array-formula follower cells (`<f t="shared" si="N"/>`, no
    /// inline text) are absent here, same as a cell with no `<f>` at all — reader.rs doesn't
    /// resolve/shift a shared formula's text for non-master cells (a substantially larger
    /// feature); this only ever captures literal inline formula text, which is exactly what
    /// every writer this codebase's own tests exercise (`aoa_to_sheet` + `XLSX.write`)
    /// produces — confirmed live it never emits shared formulas.
    pub formulas: HashMap<(u32, u32), String>,
    /// The worksheet's declared `<dimension ref="..."/>` range, 1-based inclusive, when
    /// present AND trusted — see `parse_dimension_ref`'s doc comment for the oracle's own
    /// colon-required-in-ref quirk this replicates exactly. `None` when the tag is absent,
    /// unparseable, degenerate/reversed, or (matching the oracle) a colon-less single-cell
    /// ref like `ref="A1"`.
    pub dimension: Option<MergeRect>,
    /// Per-cell resolved `numFmtId` (Milestone read-item 6), 1-based `(row, col)` keys —
    /// only cells with a NON-ZERO id are present (0 == "General", the same as no entry at
    /// all — matching the oracle's own `fmtid = 0` default when a cell has no `s`
    /// attribute, an out-of-range one, or an `<xf>` with no `numFmtId`). Resolving this id
    /// to an actual format STRING (built-in or, via `BufferWorkbook::number_formats`,
    /// custom) and deciding whether it's date-like is deliberately left to the JS layer
    /// (`packages/xlsx/src/internal/read-shape.cjs`), which already depends on the real
    /// `ssf` engine — see that file's doc comment for why porting SSF's own
    /// format-code-to-date-format heuristic into Rust would be a second, unverified
    /// implementation of logic already proven correct.
    pub style_ids: HashMap<(u32, u32), u32>,
}

// ── Minimal pull XML parser ───────────────────────────────────────────────────

#[derive(Debug)]
struct Attr {
    name: String,
    value: String,
}

#[derive(Debug)]
enum Ev {
    Open(String, Vec<Attr>),
    Close(String),
    SelfClose(String, Vec<Attr>),
    /// Raw, unescaped text preserved verbatim.
    Text(String),
}

struct XmlIter<'a> {
    s: &'a str,
}

impl<'a> XmlIter<'a> {
    fn new(s: &'a str) -> Self {
        XmlIter { s }
    }

    fn next_ev(&mut self) -> Option<Ev> {
        loop {
            if self.s.is_empty() {
                return None;
            }

            if !self.s.starts_with('<') {
                // Text node — preserve verbatim (trim happens at call site for leaf nodes)
                let end = self.s.find('<').unwrap_or(self.s.len());
                let raw = &self.s[..end];
                self.s = &self.s[end..];
                let text = xml_unescape(raw);
                if text.is_empty() {
                    continue;
                }
                return Some(Ev::Text(text));
            }

            self.s = &self.s[1..]; // consume '<'

            // Closing tag
            if self.s.starts_with('/') {
                self.s = &self.s[1..];
                let end = self.s.find('>').unwrap_or(self.s.len());
                let name = self.s[..end].trim().to_string();
                self.s = &self.s[(end + 1).min(self.s.len())..];
                return Some(Ev::Close(name));
            }

            // Comment
            if self.s.starts_with("!--") {
                let end = self.s.find("-->").map(|p| p + 3).unwrap_or(self.s.len());
                self.s = &self.s[end..];
                continue;
            }

            // CDATA
            if self.s.starts_with("![CDATA[") {
                self.s = &self.s[8..];
                let end = self.s.find("]]>").unwrap_or(self.s.len());
                let text = self.s[..end].to_string();
                self.s = &self.s[(end + 3).min(self.s.len())..];
                if !text.is_empty() {
                    return Some(Ev::Text(text));
                }
                continue;
            }

            // Processing instruction or DOCTYPE
            if self.s.starts_with('?') || self.s.starts_with('!') {
                let end = self.s.find('>').map(|p| p + 1).unwrap_or(self.s.len());
                self.s = &self.s[end..];
                continue;
            }

            // Opening / self-closing tag
            let tag_end = find_tag_close(self.s);
            let tag_inner = self.s[..tag_end].trim_end();
            let self_close = tag_inner.ends_with('/');
            let tag_body = if self_close {
                tag_inner[..tag_inner.len() - 1].trim_end()
            } else {
                tag_inner
            };
            self.s = &self.s[(tag_end + 1).min(self.s.len())..];

            let name_end = tag_body
                .find(|c: char| c.is_ascii_whitespace())
                .unwrap_or(tag_body.len());
            let name = tag_body[..name_end].to_string();
            let attrs = parse_attrs(&tag_body[name_end..]);

            if self_close {
                return Some(Ev::SelfClose(name, attrs));
            }
            return Some(Ev::Open(name, attrs));
        }
    }
}

/// Find the byte position of the unquoted `>` that closes the current tag body.
fn find_tag_close(s: &str) -> usize {
    let mut in_quote = false;
    let mut qchar = '"';
    for (i, c) in s.char_indices() {
        if in_quote {
            if c == qchar {
                in_quote = false;
            }
        } else {
            match c {
                '"' | '\'' => {
                    in_quote = true;
                    qchar = c;
                }
                '>' => return i,
                _ => {}
            }
        }
    }
    s.len()
}

/// Parse ` name="value" ...` attribute string.
fn parse_attrs(mut s: &str) -> Vec<Attr> {
    let mut attrs = vec![];
    loop {
        s = s.trim_start();
        if s.is_empty() {
            break;
        }
        let Some(eq) = s.find('=') else { break };
        let name = s[..eq].trim().to_string();
        if name.is_empty() {
            break;
        }
        s = s[eq + 1..].trim_start();
        let Some(quote) = s.chars().next() else { break };
        if quote != '"' && quote != '\'' {
            break;
        }
        s = &s[1..]; // skip opening quote
        let end = s.find(quote).unwrap_or(s.len());
        let value = xml_unescape(&s[..end]);
        s = &s[(end + 1).min(s.len())..];
        attrs.push(Attr { name, value });
    }
    attrs
}

fn attr_get<'a>(attrs: &'a [Attr], name: &str) -> Option<&'a str> {
    attrs
        .iter()
        .find(|a| a.name == name || a.name.split(':').next_back() == Some(name))
        .map(|a| a.value.as_str())
}

/// True when a named attribute is present and its value is a "true" xsd:boolean literal —
/// OOXML types attributes like `<row hidden="...">`/`<col hidden="...">` as xsd:boolean,
/// whose valid lexical space is BOTH "1"/"0" and "true"/"false" (not "1"/"0" only). A
/// hardcoded `== Some("1")` check missed real files: confirmed live that the oracle's own
/// writer emits `hidden="1"` for `<row>` but `hidden="true"` for `<col>` (an asymmetry in
/// the oracle's own writer, not a hypothetical) — so a "1"-only check silently never
/// recognized an oracle-written hidden column at all. Used for both `<row>` and `<col>`
/// so the two stay consistent rather than each hardcoding its own literal.
fn attr_is_true(attrs: &[Attr], name: &str) -> bool {
    matches!(
        attr_get(attrs, name),
        Some("1") | Some("true") | Some("TRUE")
    )
}

// Longest real entity is a numeric ref like "&#x10FFFF;" (10 chars between
// '&' and ';') — bounding the ';' search to this window keeps a run of
// many unterminated '&' characters O(n) instead of O(n^2) (each `find`
// would otherwise rescan to the end of the string).
const MAX_ENTITY_BODY_LEN: usize = 12;

fn xml_unescape(s: &str) -> String {
    if !s.contains('&') {
        return s.to_string();
    }
    // Single forward pass, each '&...;' consumed at most once — chained
    // .replace() calls (the previous implementation) double-unescape
    // input like the literal text "&amp;lt;", which must stay "&lt;", not
    // become "<": replacing "&amp;" first turns it into "&lt;", and the
    // very next replace pass then corrupts that into "<".
    let mut out = String::with_capacity(s.len());
    let mut rest = s;
    while let Some(amp) = rest.find('&') {
        out.push_str(&rest[..amp]);
        let after = &rest[amp + 1..];
        let window_end = after.len().min(MAX_ENTITY_BODY_LEN);
        let decoded = after[..window_end].find(';').and_then(|semi| {
            let entity = &after[..semi];
            let ch = match entity {
                "amp" => Some('&'),
                "lt" => Some('<'),
                "gt" => Some('>'),
                "quot" => Some('"'),
                "apos" => Some('\''),
                _ => entity.strip_prefix('#').and_then(|numeric| {
                    let code = if let Some(hex) = numeric
                        .strip_prefix('x')
                        .or_else(|| numeric.strip_prefix('X'))
                    {
                        u32::from_str_radix(hex, 16).ok()
                    } else {
                        numeric.parse::<u32>().ok()
                    };
                    code.and_then(char::from_u32)
                }),
            };
            ch.map(|c| (c, semi))
        });
        match decoded {
            Some((c, semi)) => {
                out.push(c);
                rest = &after[semi + 1..];
            }
            None => {
                // Not a recognized entity (or no ';' nearby) — keep the
                // '&' literal, matching the previous implementation's
                // tolerance for bare/unrecognized '&' in real-world input.
                out.push('&');
                rest = after;
            }
        }
    }
    out.push_str(rest);
    out
}

// ── Helper: read a ZIP entry into a String ────────────────────────────────────

/// 64 MB decompressed cap per entry — enough for any real spreadsheet XML.
const ZIP_ENTRY_MAX_BYTES: u64 = 64 * 1024 * 1024;

fn zip_read_text<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<String, String> {
    let mut entry = archive
        .by_name(name)
        .map_err(|e| format!("{}: {}", name, e))?;
    let mut s = String::new();
    entry
        .by_ref()
        .take(ZIP_ENTRY_MAX_BYTES)
        .read_to_string(&mut s)
        .map_err(|e| e.to_string())?;
    Ok(s)
}

// ── Raw ZIP passthrough (Milestone: safe round-trip) ───────────────────────────

/// Every ZIP entry's decompressed bytes, keyed by entry name — used only by
/// `save_xlsx_impl` (`src/lib.rs`) at save time, to pass through OOXML parts this
/// reader doesn't parse (`xl/vbaProject.bin`, tables, named ranges, full styles,
/// etc.) unchanged instead of losing them on every save. Not called from any
/// read-only path (`check`/`snapshot`/`diagnose`/`test-workbook` never write a
/// workbook back out), so those paths never pay this cost — see
/// `docs/xlsx-architecture.md`.
pub(crate) fn read_raw_zip_entries(path: &str) -> Result<HashMap<String, Vec<u8>>, String> {
    let file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;
    let mut out = HashMap::new();
    for i in 0..archive.len() {
        let mut entry = archive.by_index(i).map_err(|e| e.to_string())?;
        if entry.is_dir() {
            continue;
        }
        let name = entry.name().to_string();
        let mut buf = Vec::new();
        entry
            .by_ref()
            .take(ZIP_ENTRY_MAX_BYTES)
            .read_to_end(&mut buf)
            .map_err(|e| e.to_string())?;
        out.insert(name, buf);
    }
    Ok(out)
}

/// `(defaults, overrides)` — see `content_type_decls`.
pub(crate) type ContentTypeDecls = (Vec<(String, String)>, Vec<(String, String)>);

/// Parses `[Content_Types].xml`'s `Default`/`Override` declarations, in document
/// order — `(extension, content_type)` for `Default`, `(part_name, content_type)`
/// for `Override`. Used by `save_xlsx_impl` to carry over a passed-through part's
/// real declared content type instead of guessing one.
pub(crate) fn content_type_decls(xml: &str) -> ContentTypeDecls {
    let mut defaults = vec![];
    let mut overrides = vec![];
    let mut iter = XmlIter::new(xml);
    while let Some(ev) = iter.next_ev() {
        if let Ev::Open(ref tag, ref attrs) | Ev::SelfClose(ref tag, ref attrs) = ev {
            let local = tag.split(':').next_back().unwrap_or(tag);
            match local {
                "Default" => {
                    if let (Some(ext), Some(ct)) =
                        (attr_get(attrs, "Extension"), attr_get(attrs, "ContentType"))
                    {
                        defaults.push((ext.to_string(), ct.to_string()));
                    }
                }
                "Override" => {
                    if let (Some(part), Some(ct)) =
                        (attr_get(attrs, "PartName"), attr_get(attrs, "ContentType"))
                    {
                        overrides.push((part.to_string(), ct.to_string()));
                    }
                }
                _ => {}
            }
        }
    }
    (defaults, overrides)
}

/// Returns `xml`'s root element's raw attribute string (everything between the tag name
/// and the closing `>`/`/>` of its start tag, trimmed, self-closing `/` stripped) iff the
/// root element's local name (namespace prefix ignored) matches `local_name`; `None`
/// otherwise, including when the root has no attributes at all. Used to carry a source
/// worksheet's `<worksheet xmlns=".." mc:Ignorable=".." xr:uid="..">` namespace
/// declarations verbatim into a regenerated root tag, rather than reconstructing them
/// selectively — see docs/xlsx-worksheet-preservation-0.10.0-design.md §8.
pub(crate) fn extract_root_attrs(xml: &str, local_name: &str) -> Option<String> {
    let (start, tag_close_rel, full_name) = find_next_open_tag(xml, 0)?;
    if full_name.rsplit(':').next().unwrap_or(&full_name) != local_name {
        return None;
    }
    let after_name = &xml[start + 1 + full_name.len()..];
    let trimmed = after_name[..tag_close_rel].trim();
    let attrs = trimmed.strip_suffix('/').unwrap_or(trimmed).trim_end();
    if attrs.is_empty() {
        None
    } else {
        Some(attrs.to_string())
    }
}

/// Extracts the raw, byte-for-byte substring of the first `<local_name ..>...</local_name>`
/// or `<local_name ../>` top-level element found in `xml` (matched by local name, namespace
/// prefix ignored), including its own start/end tags — `None` if absent. Deliberately not a
/// full XML parser: opaque-fragment passthrough only needs one element's boundaries and its
/// untouched bytes, not a parsed tree — see
/// docs/xlsx-worksheet-preservation-0.10.0-design.md §7(b). The closing-tag search is a
/// literal string match on `</local_name>`, not tag-depth tracking, so this assumes
/// `local_name` never nests an element of the same name — true for every 0.10.0-B target
/// (`sheetViews`, `sheetPr`, `sheetFormatPr`, `dataValidations`, `autoFilter`,
/// `pageMargins` don't self-nest).
pub(crate) fn extract_raw_element(xml: &str, local_name: &str) -> Option<String> {
    let mut search_from = 0;
    loop {
        let (tag_start, tag_close_rel, full_name) = find_next_open_tag(xml, search_from)?;
        if full_name.rsplit(':').next().unwrap_or(&full_name) != local_name {
            search_from = tag_start + 1;
            continue;
        }
        let name_end = tag_start + 1 + full_name.len();
        let start_tag_end = name_end + tag_close_rel + 1;
        let self_closing = xml[name_end..name_end + tag_close_rel]
            .trim_end()
            .ends_with('/');
        if self_closing {
            return Some(xml[tag_start..start_tag_end].to_string());
        }
        let close_tag = format!("</{}>", full_name);
        let end_rel = xml[start_tag_end..].find(&close_tag)?;
        let end = start_tag_end + end_rel + close_tag.len();
        return Some(xml[tag_start..end].to_string());
    }
}

/// Shared scan primitive for `extract_root_attrs`/`extract_raw_element`: finds the next
/// opening or self-closing tag at or after byte offset `from` (skipping closing tags,
/// comments, CDATA, and processing instructions/XML declarations), returning
/// `(tag_start, tag_close_rel, local_name)` — `tag_start` is the byte offset of the tag's
/// `<`, `tag_close_rel` is the offset of its terminating unquoted `>` relative to just
/// after the tag name, and `local_name` has any namespace prefix stripped. `None` if no
/// more tags exist.
fn find_next_open_tag(xml: &str, mut search_from: usize) -> Option<(usize, usize, String)> {
    loop {
        let rel = xml[search_from..].find('<')?;
        let tag_start = search_from + rel;
        let after_lt = &xml[tag_start + 1..];
        if after_lt.starts_with(['/', '!', '?']) {
            search_from = tag_start + 1;
            continue;
        }
        let name_end = after_lt
            .find(|c: char| c.is_ascii_whitespace() || c == '>' || c == '/')
            .unwrap_or(after_lt.len());
        let full_name = after_lt[..name_end].to_string();
        let rest = &after_lt[name_end..];
        let tag_close_rel = find_tag_close(rest);
        return Some((tag_start, tag_close_rel, full_name));
    }
}

/// Extracts the raw, byte-for-byte `<hyperlink .../>` spans inside `xml`'s
/// `<hyperlinks>...</hyperlinks>` container. Same-workbook, relationship-free
/// `location=` hyperlinks are always kept. `r:id`-bearing children are only kept when
/// `include_relationship_backed` is true — the caller's job to pass that only when this
/// sheet's own worksheet-level `.rels` genuinely survived into the same save's output
/// (see `save_xlsx_impl`'s `rels_survived`): an r:id-bearing hyperlink is meaningless,
/// or worse a dangling reference, without its `.rels` entry surviving alongside it.
/// Empty if `<hyperlinks>` is absent from `xml`, or if every child was excluded.
///
/// Unlike `extract_raw_element`, this does NOT return the source bytes verbatim as one
/// blob — the container is reconstructed by the caller from a filtered child subset, so
/// each child's raw span is preserved individually rather than the whole container being
/// byte-copied. `attr_get(&attrs, "id")` (not a literal `"r:id"` string match) is reused
/// deliberately: `CT_Hyperlink`'s own XSD definition has exactly one id-shaped attribute
/// (`r:id`, namespace-prefixed) and no bare `id`, so this is precise, not a shortcut.
///
/// `CT_Hyperlink` has no child elements (only attributes, confirmed against the real
/// ECMA-376 XSD) — every real `<hyperlink>` is self-closing, so this doesn't attempt to
/// handle a non-self-closing form.
pub(crate) fn extract_hyperlinks(xml: &str, include_relationship_backed: bool) -> Vec<String> {
    let Some(container) = extract_raw_element(xml, "hyperlinks") else {
        return Vec::new();
    };
    let mut out = Vec::new();
    let mut search_from = 0;
    while let Some((tag_start, tag_close_rel, full_name)) =
        find_next_open_tag(&container, search_from)
    {
        if full_name.rsplit(':').next().unwrap_or(&full_name) != "hyperlink" {
            search_from = tag_start + 1;
            continue;
        }
        let name_end = tag_start + 1 + full_name.len();
        let start_tag_end = name_end + tag_close_rel + 1;
        let attrs = parse_attrs(&container[name_end..name_end + tag_close_rel]);
        let has_rid = attr_get(&attrs, "id").is_some();
        if !has_rid || include_relationship_backed {
            out.push(container[tag_start..start_tag_end].to_string());
        }
        search_from = start_tag_end;
    }
    out
}

#[cfg(test)]
mod opaque_fragment_tests {
    use super::*;

    #[test]
    fn extract_root_attrs_captures_namespaces_and_xr_uid_verbatim() {
        let xml = concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" ",
            "mc:Ignorable=\"x14ac xr xr2 xr3\" xr:uid=\"{ACCE0F6A-5070-C341-A245-A04D433D82F2}\">\n",
            "<sheetData/></worksheet>",
        );
        let attrs = extract_root_attrs(xml, "worksheet").unwrap();
        assert!(
            attrs
                .starts_with("xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"")
        );
        assert!(attrs.contains("xr:uid=\"{ACCE0F6A-5070-C341-A245-A04D433D82F2}\""));
        assert!(
            !attrs.ends_with('/'),
            "self-closing slash must not leak in: {attrs:?}"
        );
    }

    #[test]
    fn extract_root_attrs_returns_none_for_a_bare_no_attribute_root() {
        let xml = "<?xml version=\"1.0\"?><worksheet><sheetData/></worksheet>";
        assert_eq!(extract_root_attrs(xml, "worksheet"), None);
    }

    #[test]
    fn extract_root_attrs_returns_none_on_local_name_mismatch() {
        let xml = "<?xml version=\"1.0\"?><workbook foo=\"bar\"><sheets/></workbook>";
        assert_eq!(extract_root_attrs(xml, "worksheet"), None);
    }

    #[test]
    fn extract_raw_element_returns_the_full_subtree_verbatim() {
        let xml = concat!(
            "<?xml version=\"1.0\"?><worksheet>",
            "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">",
            "<pane xSplit=\"1\" ySplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/>",
            "<selection pane=\"bottomRight\" activeCell=\"B2\" sqref=\"B2\"/>",
            "</sheetView></sheetViews>",
            "<sheetData/></worksheet>",
        );
        let frag = extract_raw_element(xml, "sheetViews").unwrap();
        assert_eq!(
            frag,
            concat!(
                "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\">",
                "<pane xSplit=\"1\" ySplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozen\"/>",
                "<selection pane=\"bottomRight\" activeCell=\"B2\" sqref=\"B2\"/>",
                "</sheetView></sheetViews>",
            )
        );
    }

    #[test]
    fn extract_raw_element_handles_a_self_closing_form() {
        let xml = "<worksheet><sheetViews/><sheetData/></worksheet>";
        assert_eq!(
            extract_raw_element(xml, "sheetViews"),
            Some("<sheetViews/>".to_string())
        );
    }

    #[test]
    fn extract_raw_element_returns_none_when_absent() {
        let xml = "<worksheet><sheetData/></worksheet>";
        assert_eq!(extract_raw_element(xml, "sheetViews"), None);
    }

    #[test]
    fn extract_raw_element_does_not_match_a_differently_named_element() {
        // Regression guard for a naive substring search: "sheetView" (singular, the CHILD
        // element) must not be matched when asking for "sheetViews" (plural, the container).
        let xml = "<worksheet><sheetView tabSelected=\"1\"/><sheetData/></worksheet>";
        assert_eq!(extract_raw_element(xml, "sheetViews"), None);
    }

    #[test]
    fn extract_raw_element_ignores_a_namespace_prefix_on_the_target_element() {
        let xml = "<worksheet><x:sheetViews><x:sheetView/></x:sheetViews></worksheet>";
        assert_eq!(
            extract_raw_element(xml, "sheetViews"),
            Some("<x:sheetViews><x:sheetView/></x:sheetViews>".to_string())
        );
    }

    #[test]
    fn hyperlinks_returns_none_when_hyperlinks_absent() {
        let xml = "<worksheet><sheetData/></worksheet>";
        assert_eq!(extract_hyperlinks(xml, false), Vec::<String>::new());
        assert_eq!(extract_hyperlinks(xml, true), Vec::<String>::new());
    }

    #[test]
    fn hyperlinks_all_location_form_all_kept_either_way() {
        // fixture6_internal_hyperlink.xlsm's real shape: a single location-only hyperlink.
        let xml = concat!(
            "<worksheet><sheetData/>",
            "<hyperlinks><hyperlink ref=\"A1\" location=\"Sheet2!B2\" display=\"Sheet2!B2\" ",
            "xr:uid=\"{7239724E-8623-EB4C-A548-F5CFD578FC11}\"/></hyperlinks>",
            "</worksheet>",
        );
        let expected = vec![
            "<hyperlink ref=\"A1\" location=\"Sheet2!B2\" display=\"Sheet2!B2\" \
             xr:uid=\"{7239724E-8623-EB4C-A548-F5CFD578FC11}\"/>"
                .to_string(),
        ];
        assert_eq!(extract_hyperlinks(xml, false), expected);
        assert_eq!(extract_hyperlinks(xml, true), expected);
    }

    #[test]
    fn hyperlinks_all_rid_form_excluded_unless_relationship_backed_requested() {
        // fixture4_hyperlink_comment_name.xlsm's real shape: a single r:id (external URL)
        // hyperlink.
        let xml = concat!(
            "<worksheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
            "<sheetData/><hyperlinks><hyperlink ref=\"D6\" r:id=\"rId1\"/></hyperlinks>",
            "</worksheet>",
        );
        assert_eq!(extract_hyperlinks(xml, false), Vec::<String>::new());
        assert_eq!(
            extract_hyperlinks(xml, true),
            vec!["<hyperlink ref=\"D6\" r:id=\"rId1\"/>".to_string()]
        );
    }

    #[test]
    fn hyperlinks_mixed_container_respects_the_flag_per_child() {
        // Synthetic -- no real fixture has a mixed <hyperlinks> container yet (see
        // docs/xlsx-worksheet-preservation-0.10.0-design.md's B4 entry). Two r:id-bearing
        // hyperlinks sandwich one location-only hyperlink to prove position doesn't matter.
        let xml = concat!(
            "<worksheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">",
            "<sheetData/><hyperlinks>",
            "<hyperlink ref=\"A1\" r:id=\"rId1\"/>",
            "<hyperlink ref=\"B1\" location=\"Sheet2!A1\"/>",
            "<hyperlink ref=\"C1\" r:id=\"rId2\"/>",
            "</hyperlinks></worksheet>",
        );
        assert_eq!(
            extract_hyperlinks(xml, false),
            vec!["<hyperlink ref=\"B1\" location=\"Sheet2!A1\"/>".to_string()]
        );
        assert_eq!(
            extract_hyperlinks(xml, true),
            vec![
                "<hyperlink ref=\"A1\" r:id=\"rId1\"/>".to_string(),
                "<hyperlink ref=\"B1\" location=\"Sheet2!A1\"/>".to_string(),
                "<hyperlink ref=\"C1\" r:id=\"rId2\"/>".to_string(),
            ]
        );
    }

    #[test]
    fn hyperlinks_returns_multiple_in_document_order() {
        let xml = concat!(
            "<worksheet><sheetData/><hyperlinks>",
            "<hyperlink ref=\"A1\" location=\"Sheet2!A1\"/>",
            "<hyperlink ref=\"B1\" location=\"Sheet3!A1\"/>",
            "</hyperlinks></worksheet>",
        );
        let expected = vec![
            "<hyperlink ref=\"A1\" location=\"Sheet2!A1\"/>".to_string(),
            "<hyperlink ref=\"B1\" location=\"Sheet3!A1\"/>".to_string(),
        ];
        assert_eq!(extract_hyperlinks(xml, false), expected);
        assert_eq!(extract_hyperlinks(xml, true), expected);
    }
}

/// `xl/_rels/workbook.xml.rels`'s own `<Relationship Type=".." Target=".."/>` entries —
/// `(Type, Target)` pairs, `Target` exactly as written (relative to `xl/`, no leading `/`).
/// Ids are dropped: callers assign fresh ones when carrying a relationship into a
/// regenerated rels file (see `save_xlsx_impl`'s `carried_rels`), since the writer's own
/// sequential-id scheme for worksheets/sharedStrings/styles/vbaProject would otherwise
/// collide with whatever ids the source happened to use.
pub(crate) fn workbook_rels_decls(xml: &str) -> Vec<(String, String)> {
    let mut rels = vec![];
    let mut iter = XmlIter::new(xml);
    while let Some(ev) = iter.next_ev() {
        if let Ev::Open(ref tag, ref attrs) | Ev::SelfClose(ref tag, ref attrs) = ev {
            let local = tag.split(':').next_back().unwrap_or(tag);
            if local == "Relationship"
                && let (Some(ty), Some(target)) =
                    (attr_get(attrs, "Type"), attr_get(attrs, "Target"))
            {
                rels.push((ty.to_string(), target.to_string()));
            }
        }
    }
    rels
}

// ── XLSX reader ───────────────────────────────────────────────────────────────

fn read_xlsx(path: &str) -> Result<Vec<WorkbookSheet>, String> {
    let file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let archive = ZipArchive::new(file).map_err(|e| e.to_string())?;
    // Path-based read_workbook doesn't expose formulas/!ref/style ids (see BufferSheet's
    // doc comment) — discard that half here rather than changing WorkbookSheet itself.
    Ok(read_workbook_from_archive(archive)?
        .sheets
        .into_iter()
        .map(|bs| bs.sheet)
        .collect())
}

/// The body of the XLSX reader, generalized over any `R: Read + Seek` archive source
/// (a `std::fs::File` for path-based reads, a `Cursor<&[u8]>` for `read_workbook_from_bytes`)
/// — see `docs/xlsx-architecture.md`'s "reader.rs buffer-API resolution". Pure extraction
/// from the former `read_xlsx`, no behavior change.
fn read_workbook_from_archive<R: Read + Seek>(
    mut archive: ZipArchive<R>,
) -> Result<BufferWorkbook, String> {
    let wb_xml = zip_read_text(&mut archive, "xl/workbook.xml")?;
    let sheet_refs = xlsx_workbook_sheets(&wb_xml);
    let date1904 = xlsx_workbook_date1904(&wb_xml);

    let rels_xml = zip_read_text(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels = xlsx_rels(&rels_xml);

    let shared: Vec<String> = match zip_read_text(&mut archive, "xl/sharedStrings.xml") {
        Ok(xml) => xlsx_shared_strings(&xml),
        Err(_) => vec![],
    };

    let styles = match zip_read_text(&mut archive, "xl/styles.xml") {
        Ok(xml) => xlsx_styles(&xml),
        Err(_) => XlsxStyles::default(),
    };

    let mut sheets = vec![];
    for (name, rid, sheet_id) in sheet_refs {
        let Some(target) = rels.get(&rid) else {
            continue;
        };
        let zip_path = if let Some(rest) = target.strip_prefix('/') {
            rest.to_string()
        } else {
            format!("xl/{}", target)
        };
        let sheet_xml = match zip_read_text(&mut archive, &zip_path) {
            Ok(s) => s,
            Err(_) => continue,
        };
        let sheet_data = xlsx_sheet_cells(&sheet_xml, &shared, &styles.cell_xfs);
        sheets.push(BufferSheet {
            sheet: WorkbookSheet {
                name,
                cells: sheet_data.cells,
                sheet_id,
                workbook_rel_id: Some(rid),
                source_part_name: Some(zip_path.clone()),
                merged_ranges: sheet_data.merged_ranges,
                hidden_rows: sheet_data.hidden_rows,
                hidden_columns: sheet_data.hidden_columns,
                raw_style_indices: sheet_data.raw_style_indices,
                formulas: sheet_data.formulas.clone(),
            },
            formulas: sheet_data.formulas,
            dimension: sheet_data.dimension,
            style_ids: sheet_data.style_ids,
        });
    }
    Ok(BufferWorkbook {
        sheets,
        number_formats: styles.number_formats,
        date1904,
    })
}

/// Whether `xl/workbook.xml` declares `<workbookPr date1904="1"/>` (the 1904 date
/// system) — mirrors the oracle's own `parsexmlbool`-based check on this exact attribute
/// (confirmed by reading xlsx.js's WBPropsDef/date1904 handling directly), via
/// `attr_is_true`'s same xsd:boolean lexical space.
fn xlsx_workbook_date1904(xml: &str) -> bool {
    let mut iter = XmlIter::new(xml);
    while let Some(ev) = iter.next_ev() {
        if let Ev::Open(ref tag, ref attrs) | Ev::SelfClose(ref tag, ref attrs) = ev
            && tag.split(':').next_back() == Some("workbookPr")
        {
            return attr_is_true(attrs, "date1904");
        }
    }
    false
}

/// Returns `[(sheet_name, rId, sheetId)]` in document order.
fn xlsx_workbook_sheets(xml: &str) -> Vec<(String, String, Option<String>)> {
    let mut iter = XmlIter::new(xml);
    let mut result = vec![];
    while let Some(ev) = iter.next_ev() {
        if let Ev::SelfClose(ref tag, ref attrs) = ev {
            let local = tag.split(':').next_back().unwrap_or(tag);
            if local == "sheet"
                && let (Some(name), Some(rid)) = (attr_get(attrs, "name"), attr_get(attrs, "id"))
            {
                let sheet_id = attr_get(attrs, "sheetId").map(|s| s.to_string());
                result.push((name.to_string(), rid.to_string(), sheet_id));
            }
        }
    }
    result
}

/// Returns `{rId → target_path}` for worksheet relationships.
fn xlsx_rels(xml: &str) -> HashMap<String, String> {
    let mut iter = XmlIter::new(xml);
    let mut map = HashMap::new();
    while let Some(ev) = iter.next_ev() {
        if let Ev::SelfClose(ref tag, ref attrs) = ev {
            let local = tag.split(':').next_back().unwrap_or(tag);
            if local == "Relationship"
                && let (Some(id), Some(ty), Some(target)) = (
                    attr_get(attrs, "Id"),
                    attr_get(attrs, "Type"),
                    attr_get(attrs, "Target"),
                )
                && ty.ends_with("/worksheet")
            {
                map.insert(id.to_string(), target.to_string());
            }
        }
    }
    map
}

/// Builds the shared-strings table.
fn xlsx_shared_strings(xml: &str) -> Vec<String> {
    let mut iter = XmlIter::new(xml);
    let mut strings = vec![];
    let mut in_si = false;
    let mut in_t = false;
    let mut current = String::new();

    while let Some(ev) = iter.next_ev() {
        match &ev {
            Ev::Open(tag, _) | Ev::SelfClose(tag, _) => {
                let local = tag.split(':').next_back().unwrap_or(tag);
                match local {
                    "si" => {
                        in_si = true;
                        current.clear();
                    }
                    "t" => {
                        in_t = true;
                    }
                    _ => {}
                }
            }
            Ev::Close(tag) => {
                let local = tag.split(':').next_back().unwrap_or(tag);
                match local {
                    "si" => {
                        strings.push(current.clone());
                        in_si = false;
                    }
                    "t" => {
                        in_t = false;
                    }
                    _ => {}
                }
            }
            Ev::Text(text) => {
                if in_si && in_t {
                    current.push_str(text);
                }
            }
        }
    }
    strings
}

/// `xl/styles.xml`, parsed down to exactly the two pieces read()'s `.w`/`.z`/date-typed-cell
/// support (Milestone read-item 6) needs — see `BufferWorkbook::number_formats` and
/// `BufferSheet::style_ids`'s doc comments. Deliberately not a general styles.xml parser:
/// fonts/fills/borders/cellStyles/cellStyleXfs are never read, matching the oracle's own
/// cell-format resolution (`cf = styles.CellXf[tag.s]; if (cf.numFmtId != null) ...`,
/// confirmed by reading xlsx.js directly), which never consults them either.
#[derive(Default)]
struct XlsxStyles {
    /// Custom `<numFmt numFmtId="N" formatCode="...">` definitions — see
    /// `BufferWorkbook::number_formats`.
    number_formats: HashMap<u32, String>,
    /// `<cellXfs><xf numFmtId="N".../></cellXfs>` entries in document order — a cell's
    /// `s="N"` attribute is a 0-based index into this Vec (`None` when an `<xf>` has no
    /// `numFmtId` attribute at all, matching the oracle's own `cf.numFmtId != null` check).
    cell_xfs: Vec<Option<u32>>,
}

fn xlsx_styles(xml: &str) -> XlsxStyles {
    let mut iter = XmlIter::new(xml);
    let mut number_formats: HashMap<u32, String> = HashMap::new();
    let mut cell_xfs: Vec<Option<u32>> = Vec::new();
    let mut in_cell_xfs = false;

    while let Some(ev) = iter.next_ev() {
        match &ev {
            Ev::Open(tag, attrs) | Ev::SelfClose(tag, attrs) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "numFmt" => {
                        if let (Some(id), Some(code)) = (
                            attr_get(attrs, "numFmtId").and_then(|s| s.parse::<u32>().ok()),
                            attr_get(attrs, "formatCode"),
                        ) {
                            number_formats.insert(id, code.to_string());
                        }
                    }
                    // A self-closing <cellXfs/> (zero entries) never produces a matching
                    // Close event — only an actual Open sets in_cell_xfs, mirroring how
                    // xlsx_sheet_cells already guards <f/>.
                    "cellXfs" if matches!(ev, Ev::Open(_, _)) => {
                        in_cell_xfs = true;
                    }
                    "xf" if in_cell_xfs => {
                        cell_xfs
                            .push(attr_get(attrs, "numFmtId").and_then(|s| s.parse::<u32>().ok()));
                    }
                    _ => {}
                }
            }
            Ev::Close(tag) => {
                if tag.split(':').next_back() == Some("cellXfs") {
                    in_cell_xfs = false;
                }
            }
            Ev::Text(_) => {}
        }
    }
    XlsxStyles {
        number_formats,
        cell_xfs,
    }
}

/// Parses a single worksheet XML into a 1-based (row, col) → SheetCell map,
/// plus any `<mergeCells><mergeCell ref="..."/></mergeCells>` ranges
/// (Milestone B6c2) and hidden row/column metadata (Milestone B7b).
/// A small return struct, not a growing bare tuple — B6c2 hit a
/// `clippy::type_complexity` error the first time this function's return
/// type grew, so this sidesteps a repeat of that churn.
struct XlsxSheetData {
    cells: HashMap<(u32, u32), SheetCell>,
    merged_ranges: Vec<MergeRect>,
    /// Hidden row intervals, 1-based inclusive `(start, end)` — coalesced
    /// from consecutive `<row r=".." hidden="1">` tags (Milestone B7b).
    hidden_rows: Vec<(u32, u32)>,
    /// Hidden column intervals, 1-based inclusive `(start, end)` — read
    /// directly from `<col min=".." max=".." hidden="1">` (Milestone
    /// B7b), already interval-shaped in the XML, no coalescing needed.
    hidden_columns: Vec<(u32, u32)>,
    /// Per-cell raw `<f>` formula text — see `BufferSheet::formulas`.
    formulas: HashMap<(u32, u32), String>,
    /// The worksheet's declared `<dimension>`, when present and trusted —
    /// see `BufferSheet::dimension` / `parse_dimension_ref`.
    dimension: Option<MergeRect>,
    /// Per-cell resolved non-zero numFmtId — see `BufferSheet::style_ids`.
    style_ids: HashMap<(u32, u32), u32>,
    /// Per-cell raw `s="N"` index — see `WorkbookSheet::raw_style_indices`.
    raw_style_indices: HashMap<(u32, u32), u32>,
}

fn xlsx_sheet_cells(xml: &str, shared: &[String], cell_xfs: &[Option<u32>]) -> XlsxSheetData {
    let mut iter = XmlIter::new(xml);
    let mut cells: HashMap<(u32, u32), SheetCell> = HashMap::new();
    let mut merged_ranges: Vec<MergeRect> = Vec::new();
    let mut hidden_rows: Vec<(u32, u32)> = Vec::new();
    let mut hidden_columns: Vec<(u32, u32)> = Vec::new();
    let mut pending_hidden_row_run: Option<(u32, u32)> = None;
    let mut formulas: HashMap<(u32, u32), String> = HashMap::new();
    let mut dimension: Option<MergeRect> = None;
    let mut style_ids: HashMap<(u32, u32), u32> = HashMap::new();
    let mut raw_style_indices: HashMap<(u32, u32), u32> = HashMap::new();
    let mut cur_row: u32 = 0;
    let mut cur_col: u32 = 0;
    let mut cur_type = String::new();
    let mut in_v = false;
    // `<v xml:space="preserve">` marks significant leading/trailing
    // whitespace in a t="str" cell's literal text, same as any XML element
    // — confirmed live against compat/corpus/workbooks/with_text.xlsx's raw
    // sheet1.xml, where cell A3 is `<c t="str"><v xml:space="preserve">
    // padded  </v></c>`. Only `<v>`'s own attribute matters; `<c>` never
    // carries it in that fixture.
    let mut v_preserve_space = false;
    let mut in_f = false;
    let mut cur_formula = String::new();
    let mut in_is_t = false; // inside <is><t>
    let mut is_text = String::new();

    while let Some(ev) = iter.next_ev() {
        match ev {
            Ev::Open(ref tag, ref attrs) | Ev::SelfClose(ref tag, ref attrs) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "row" => {
                        if let Some(r) = attr_get(attrs, "r") {
                            cur_row = r.parse().unwrap_or(0);
                        }
                        let hidden = attr_is_true(attrs, "hidden");
                        if hidden {
                            pending_hidden_row_run = Some(match pending_hidden_row_run {
                                Some((start, end)) if end + 1 == cur_row => (start, cur_row),
                                _ => {
                                    if let Some(run) = pending_hidden_row_run {
                                        hidden_rows.push(run);
                                    }
                                    (cur_row, cur_row)
                                }
                            });
                        } else if let Some(run) = pending_hidden_row_run.take() {
                            hidden_rows.push(run);
                        }
                    }
                    "col" => {
                        if attr_is_true(attrs, "hidden") {
                            let min = attr_get(attrs, "min").and_then(|s| s.parse().ok());
                            let max = attr_get(attrs, "max").and_then(|s| s.parse().ok());
                            if let (Some(min), Some(max)) = (min, max) {
                                hidden_columns.push((min, max));
                            }
                        }
                    }
                    "c" => {
                        cur_type = attr_get(attrs, "t").unwrap_or("").to_string();
                        in_v = false;
                        if let Some(r) = attr_get(attrs, "r")
                            && let Some((row, col)) = parse_cell_ref(r)
                        {
                            cur_row = row;
                            cur_col = col;
                        }
                        is_text.clear();
                        in_f = false;
                        cur_formula.clear();
                        // s="N" is a 0-based index into <cellXfs> (Milestone read-item 6)
                        // — mirrors the oracle's own `cf = styles.CellXf[tag.s]` resolution
                        // exactly: an absent/out-of-range index, or an <xf> whose own
                        // numFmtId attribute was absent, all fall back to 0 (General) —
                        // matching the oracle's `fmtid = 0` default — so only a resolved
                        // NON-zero id is worth recording (0 == "no entry" downstream) in
                        // `style_ids`. `raw_style_indices` below keeps the index itself
                        // unconditionally — a style can carry font/fill/border info under
                        // a General number format, which still needs to survive a save
                        // (see `WorkbookSheet::raw_style_indices`).
                        let s_idx = if cur_row > 0 && cur_col > 0 {
                            attr_get(attrs, "s").and_then(|s| s.parse::<usize>().ok())
                        } else {
                            None
                        };
                        if let Some(idx) = s_idx {
                            raw_style_indices.insert((cur_row, cur_col), idx as u32);
                            if let Some(Some(fmt_id)) = cell_xfs.get(idx)
                                && *fmt_id != 0
                            {
                                style_ids.insert((cur_row, cur_col), *fmt_id);
                            }
                        }
                    }
                    "v" => {
                        in_v = true;
                        v_preserve_space = attr_get(attrs, "xml:space") == Some("preserve");
                    }
                    "f" => {
                        // A self-closing <f/> (or a shared-formula follower cell,
                        // <f t="shared" si="N"/>, no inline text) never produces a
                        // matching Close("f") event — nothing to capture, leave in_f
                        // false so no stray Text event gets misattributed to it.
                        if !matches!(ev, Ev::SelfClose(_, _)) {
                            in_f = true;
                            cur_formula.clear();
                        }
                    }
                    "t" => {
                        // inside <is> for inline strings
                        in_is_t = true;
                        is_text.clear();
                    }
                    "mergeCell" => {
                        if let Some(rect) = attr_get(attrs, "ref").and_then(parse_merge_ref) {
                            merged_ranges.push(rect);
                        }
                    }
                    "dimension" if dimension.is_none() => {
                        dimension = attr_get(attrs, "ref").and_then(parse_dimension_ref);
                    }
                    _ => {}
                }
            }
            Ev::Close(ref tag) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "v" => {
                        // A zero-character <v></v> never produces an Ev::Text event (there's
                        // no text to emit), so `in_v` is still true here — the Text-event
                        // handler below never ran for this cell. Route the empty string
                        // through the same xlsx_parse_cell used for the non-empty path
                        // (rather than hardcoding a value) so type-specific behavior falls
                        // out for free: t="str"/"e" -> Str(""), numeric -> parse fails -> no
                        // cell, t="s" -> index parse fails -> no cell. Confirmed live: the
                        // oracle's own writer emits exactly this shape for an empty-string
                        // aoa cell (`<c t="str"><v></v></c>`), reporting {t:"s", v:""}.
                        if in_v
                            && cur_row > 0
                            && cur_col > 0
                            && let Some(c) = xlsx_parse_cell("", &cur_type, shared)
                        {
                            cells.insert((cur_row, cur_col), c);
                        }
                        in_v = false;
                    }
                    "t" => {
                        in_is_t = false;
                    }
                    "f" => {
                        if in_f && cur_row > 0 && cur_col > 0 && !cur_formula.is_empty() {
                            formulas.insert((cur_row, cur_col), cur_formula.clone());
                        }
                        in_f = false;
                    }
                    _ => {}
                }
            }
            Ev::Text(ref text) => {
                if in_v && cur_row > 0 && cur_col > 0 {
                    let raw = if v_preserve_space {
                        text.as_str()
                    } else {
                        text.trim()
                    };
                    let cell = xlsx_parse_cell(raw, &cur_type, shared);
                    if let Some(c) = cell {
                        cells.insert((cur_row, cur_col), c);
                    }
                    in_v = false;
                } else if in_is_t {
                    is_text.push_str(text);
                } else if in_f {
                    cur_formula.push_str(text);
                }
            }
        }

        // Emit inline string on </c>
        if let Ev::Close(ref tag) = ev
            && tag.split(':').next_back() == Some("c")
            && cur_type == "inlineStr"
            && !is_text.is_empty()
            && cur_row > 0
            && cur_col > 0
        {
            cells.insert((cur_row, cur_col), SheetCell::Str(is_text.clone()));
            is_text.clear();
        }
    }
    if let Some(run) = pending_hidden_row_run.take() {
        hidden_rows.push(run);
    }
    XlsxSheetData {
        cells,
        merged_ranges,
        hidden_rows,
        hidden_columns,
        formulas,
        dimension,
        style_ids,
        raw_style_indices,
    }
}

fn xlsx_parse_cell(v: &str, t: &str, shared: &[String]) -> Option<SheetCell> {
    match t {
        "s" => {
            let idx: usize = v.parse().ok()?;
            Some(SheetCell::Str(shared.get(idx)?.clone()))
        }
        "b" => Some(SheetCell::Bool(v == "1")),
        "str" | "e" => Some(SheetCell::Str(v.to_string())),
        _ => {
            // Numeric (default, no type attr)
            let f: f64 = v.parse().ok()?;
            Some(num_to_cell(f))
        }
    }
}

fn num_to_cell(f: f64) -> SheetCell {
    if f.fract() == 0.0 && f >= i64::MIN as f64 && f <= i64::MAX as f64 {
        SheetCell::Integer(f as i64)
    } else {
        SheetCell::Float(f)
    }
}

/// Parse an XLSX cell reference like "A1", "AB12" → (row, col), both 1-based.
fn parse_cell_ref(r: &str) -> Option<(u32, u32)> {
    let r = r.trim().to_uppercase();
    let alpha_end = r.find(|c: char| c.is_ascii_digit())?;
    if alpha_end == 0 {
        return None;
    }
    let col = r[..alpha_end]
        .chars()
        .fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1));
    let row: u32 = r[alpha_end..].parse().ok()?;
    Some((row, col))
}

/// Parses an XLSX `<mergeCell ref="A1:C1"/>` address into a 1-based
/// inclusive `(top-left, bottom-right)` pair (Milestone B6c2). Mirrors
/// `vm::parse_range_addr`'s logic locally rather than importing it, since
/// only `vm` depends on `reader` today, not the reverse.
fn parse_merge_ref(s: &str) -> Option<MergeRect> {
    let i = s.find(':')?;
    Some((parse_cell_ref(&s[..i])?, parse_cell_ref(&s[i + 1..])?))
}

/// Parses a worksheet's `<dimension ref="A1:C3"/>` into a 1-based inclusive rect — mirrors
/// the oracle's own dimension parsing EXACTLY, including a quirk confirmed by reading
/// xlsx.js directly (not assumed): its `dimregex = /"(\w*:\w*)"/` requires a literal colon
/// inside the quoted ref value, so a single-cell dimension like `ref="A1"` (no colon) never
/// matches at all and is silently NOT trusted — the oracle falls back to its own
/// populated-cell bounding box in that case, same as reader.rs's existing fallback.
/// Delegating to `parse_merge_ref` (which already requires a colon via `s.find(':')?`)
/// replicates this for free rather than needing a second implementation. A
/// degenerate/reversed range (start > end on either axis) is rejected too, matching the
/// oracle's own `parse_ws_xml_dim`'s `d.s.r<=d.e.r && d.s.c<=d.e.c` guard.
fn parse_dimension_ref(s: &str) -> Option<MergeRect> {
    let (start, end) = parse_merge_ref(s)?;
    if start.0 <= end.0 && start.1 <= end.1 {
        Some((start, end))
    } else {
        None
    }
}

// ── ODS reader ────────────────────────────────────────────────────────────────

fn read_ods(path: &str) -> Result<Vec<WorkbookSheet>, String> {
    let file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;
    let xml = zip_read_text(&mut archive, "content.xml")?;
    Ok(ods_parse(&xml))
}

fn ods_parse(xml: &str) -> Vec<WorkbookSheet> {
    let mut iter = XmlIter::new(xml);
    let mut sheets: Vec<WorkbookSheet> = vec![];
    let mut in_sheet = false;
    let mut row: u32 = 0;
    let mut col: u32 = 0;
    let mut in_text_p = false;
    let mut cell_text = String::new();
    let mut pending_cell: Option<OdsCellState> = None;
    // `table:number-*-repeated`: ODS's sparse-representation mechanism —
    // one <table-row>/<table-cell> element stands for N identical rows/
    // columns (LibreOffice uses this heavily, not just for trailing empty
    // runs but for any horizontal/vertical run of matching cells, so
    // real data routinely follows a repeated-empty block). Only the first
    // copy's content is ever written (matching emit_ods_cell's existing
    // convention); these track how far to advance row/col for the *next*
    // element so later real cells land at the correct coordinates instead
    // of being shifted left/up by the width of the skipped repeat. Kept
    // as an arithmetic skip, not a literal expansion loop, so a
    // pathological number-rows-repeated="1048576" costs O(1), not O(n).
    let mut row_repeat: u32 = 1;
    let mut col_repeat: u32 = 1;

    while let Some(ev) = iter.next_ev() {
        match &ev {
            Ev::Open(tag, attrs) | Ev::SelfClose(tag, attrs) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "table" => {
                        let name = attr_get(attrs, "name").unwrap_or("sheet1").to_lowercase();
                        sheets.push(WorkbookSheet {
                            name,
                            cells: HashMap::new(),
                            sheet_id: None,
                            workbook_rel_id: None,
                            source_part_name: None,
                            merged_ranges: Vec::new(),
                            hidden_rows: Vec::new(),
                            hidden_columns: Vec::new(),
                            raw_style_indices: HashMap::new(),
                            formulas: HashMap::new(),
                        });
                        in_sheet = true;
                        row = 0;
                        col = 0;
                        row_repeat = 1;
                    }
                    "table-row" if in_sheet => {
                        row += row_repeat;
                        col = 0;
                        col_repeat = 1;
                        pending_cell = None;
                        row_repeat = attr_get(attrs, "number-rows-repeated")
                            .and_then(|v| v.parse().ok())
                            .filter(|n| *n >= 1)
                            .unwrap_or(1);
                    }
                    "table-cell" | "covered-table-cell" if in_sheet => {
                        if let Some(state) = pending_cell.take() {
                            emit_ods_cell(&mut sheets, state);
                        }
                        col += col_repeat;
                        col_repeat = attr_get(attrs, "number-columns-repeated")
                            .and_then(|v| v.parse().ok())
                            .filter(|n| *n >= 1)
                            .unwrap_or(1);
                        let cell_type = attr_get(attrs, "value-type").unwrap_or("").to_string();
                        let val_attr = attr_get(attrs, "value").unwrap_or("").to_string();
                        let bool_attr = attr_get(attrs, "boolean-value").unwrap_or("").to_string();
                        cell_text.clear();
                        in_text_p = false;

                        // Merge span attrs only ever appear on the anchor
                        // `table-cell`, never `covered-table-cell`
                        // (Milestone B6c2).
                        if local == "table-cell" {
                            let cols_spanned: u32 = attr_get(attrs, "number-columns-spanned")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(1);
                            let rows_spanned: u32 = attr_get(attrs, "number-rows-spanned")
                                .and_then(|v| v.parse().ok())
                                .unwrap_or(1);
                            if (cols_spanned > 1 || rows_spanned > 1)
                                && let Some(sheet) = sheets.last_mut()
                            {
                                sheet.merged_ranges.push((
                                    (row, col),
                                    (row + rows_spanned - 1, col + cols_spanned - 1),
                                ));
                            }
                        }

                        let make_state = || OdsCellState {
                            row,
                            col,
                            cell_type,
                            val_attr,
                            bool_attr,
                            text: String::new(),
                        };
                        if matches!(ev, Ev::SelfClose(_, _)) {
                            emit_ods_cell(&mut sheets, make_state());
                            pending_cell = None;
                        } else {
                            pending_cell = Some(make_state());
                        }
                    }
                    "p" if in_sheet => {
                        in_text_p = true;
                    }
                    _ => {}
                }
            }
            Ev::Close(tag) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "table" => {
                        in_sheet = false;
                    }
                    "table-cell" | "covered-table-cell" if in_sheet => {
                        if let Some(ref mut state) = pending_cell {
                            state.text.clone_from(&cell_text);
                        }
                        if let Some(state) = pending_cell.take() {
                            emit_ods_cell(&mut sheets, state);
                        }
                        in_text_p = false;
                    }
                    "p" => {
                        in_text_p = false;
                    }
                    _ => {}
                }
            }
            Ev::Text(text) => {
                if in_sheet && in_text_p {
                    cell_text.push_str(text);
                }
            }
        }
    }
    sheets
}

struct OdsCellState {
    row: u32,
    col: u32,
    cell_type: String,
    val_attr: String,
    bool_attr: String,
    text: String,
}

fn emit_ods_cell(sheets: &mut [WorkbookSheet], state: OdsCellState) {
    let sheet = match sheets.last_mut() {
        Some(s) => s,
        None => return,
    };
    let cell = ods_make_cell(&state);
    if let Some(c) = cell {
        // Only write the first column for repeated cells (the rest are assumed identical/empty)
        sheet.cells.insert((state.row, state.col), c);
    }
    // Additional repeated columns: skip (usually trailing empties)
}

fn ods_make_cell(s: &OdsCellState) -> Option<SheetCell> {
    match s.cell_type.as_str() {
        "float" | "percentage" | "currency" => {
            let f: f64 = s.val_attr.parse().ok()?;
            Some(num_to_cell(f))
        }
        "string" => {
            if s.text.is_empty() {
                None
            } else {
                Some(SheetCell::Str(s.text.clone()))
            }
        }
        "boolean" => Some(SheetCell::Bool(s.bool_attr == "true")),
        _ => None, // empty / formula result not available / etc.
    }
}

#[cfg(test)]
mod sheet_id_tests {
    use super::*;

    #[test]
    fn xlsx_workbook_sheets_captures_non_contiguous_sheet_ids() {
        // sheetIds "1" and "5" (not "1"/"2") prove sheet_id is read from the
        // attribute itself, not inferred from document position.
        let xml = r#"<?xml version="1.0"?>
<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
<sheet name="Sheet2" sheetId="5" r:id="rId2"/>
</sheets>
</workbook>"#;
        let result = xlsx_workbook_sheets(xml);
        assert_eq!(
            result,
            vec![
                (
                    "Sheet1".to_string(),
                    "rId1".to_string(),
                    Some("1".to_string())
                ),
                (
                    "Sheet2".to_string(),
                    "rId2".to_string(),
                    Some("5".to_string())
                ),
            ]
        );
    }

    #[test]
    fn xlsx_workbook_sheets_handles_a_missing_sheet_id() {
        let xml = r#"<sheets><sheet name="Sheet1" r:id="rId1"/></sheets>"#;
        let result = xlsx_workbook_sheets(xml);
        assert_eq!(
            result,
            vec![("Sheet1".to_string(), "rId1".to_string(), None)]
        );
    }

    #[test]
    fn ods_sheets_always_have_no_sheet_id() {
        let xml = r#"<office:body><office:spreadsheet>
<table:table table:name="Sheet1"></table:table>
<table:table table:name="Sheet2"></table:table>
</office:spreadsheet></office:body>"#;
        let sheets = ods_parse(xml);
        assert_eq!(sheets.len(), 2);
        assert!(sheets.iter().all(|s| s.sheet_id.is_none()));
    }
}

// ── Milestone B6c2: merged-range parsing ────────────────────────────────
#[cfg(test)]
mod merge_tests {
    use super::*;

    #[test]
    fn parse_merge_ref_reads_top_left_and_bottom_right() {
        assert_eq!(parse_merge_ref("A1:C1"), Some(((1, 1), (1, 3))));
        assert_eq!(parse_merge_ref("B3:B4"), Some(((3, 2), (4, 2))));
    }

    #[test]
    fn parse_merge_ref_rejects_a_single_cell_with_no_colon() {
        assert_eq!(parse_merge_ref("A1"), None);
    }

    #[test]
    fn xlsx_sheet_cells_reads_merge_cells() {
        let xml = r#"<worksheet>
<sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
</sheetData>
<mergeCells count="2">
<mergeCell ref="A1:C1"/>
<mergeCell ref="B3:B4"/>
</mergeCells>
</worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.cells.len(), 1);
        assert_eq!(data.merged_ranges, vec![((1, 1), (1, 3)), ((3, 2), (4, 2))]);
    }

    #[test]
    fn xlsx_sheet_cells_with_no_merge_cells_element_has_empty_merged_ranges() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert!(data.merged_ranges.is_empty());
        assert!(data.hidden_rows.is_empty());
        assert!(data.hidden_columns.is_empty());
    }

    // ── Milestone B7b: hidden row/column parsing ────────────────────────────

    #[test]
    fn xlsx_sheet_cells_coalesces_consecutive_hidden_rows_into_intervals() {
        let xml = r#"<worksheet>
<cols>
<col min="2" max="2" hidden="1"/>
</cols>
<sheetData>
<row r="1"><c r="A1"><v>1</v></c></row>
<row r="11" hidden="1"/>
<row r="12" hidden="1"/>
<row r="13" hidden="1"/>
<row r="14" hidden="1"/>
<row r="20"><c r="A20"><v>2</v></c></row>
<row r="30" hidden="1"/>
<row r="31" hidden="1"/>
</sheetData>
</worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.hidden_rows, vec![(11, 14), (30, 31)]);
        assert_eq!(data.hidden_columns, vec![(2, 2)]);
    }

    #[test]
    fn xlsx_sheet_cells_starts_a_new_interval_across_a_row_number_gap() {
        // Row 6 is entirely absent from <sheetData> (no <row r="6"> element
        // at all) — row 5 and row 7 being hidden must NOT coalesce into a
        // single (5,7) interval just because no explicit non-hidden row
        // separates them.
        let xml = r#"<worksheet><sheetData>
<row r="5" hidden="1"/>
<row r="7" hidden="1"/>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.hidden_rows, vec![(5, 5), (7, 7)]);
    }

    #[test]
    fn xlsx_sheet_cells_reads_a_multi_column_hidden_col_span_without_coalescing() {
        let xml = r#"<worksheet><cols>
<col min="2" max="4" hidden="1"/>
<col min="6" max="6"/>
</cols><sheetData></sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.hidden_columns, vec![(2, 4)]);
    }

    #[test]
    fn xlsx_sheet_cells_accepts_the_xsd_boolean_true_literal_for_hidden() {
        // Confirmed live: the oracle's own writer emits hidden="true" (not "1") for
        // <col>, while emitting hidden="1" (not "true") for <row> — both are valid
        // xsd:boolean lexical forms per the OOXML spec, so both must be recognized on
        // both elements rather than each hardcoding the one literal the writer happened
        // to use for it.
        let xml = r#"<worksheet><cols>
<col min="1" max="1" hidden="true"/>
</cols><sheetData>
<row r="1" hidden="true"><c r="A1"><v>1</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.hidden_columns, vec![(1, 1)]);
        assert_eq!(data.hidden_rows, vec![(1, 1)]);
    }

    #[test]
    fn ods_parse_reads_column_and_row_span_into_a_merged_range() {
        let xml = r#"<office:spreadsheet>
<table:table table:name="Sheet1">
<table:table-row>
<table:table-cell table:number-columns-spanned="3" office:value-type="float" office:value="1"/>
<table:covered-table-cell/>
<table:covered-table-cell/>
</table:table-row>
</table:table>
</office:spreadsheet>"#;
        let sheets = ods_parse(xml);
        assert_eq!(sheets[0].merged_ranges, vec![((1, 1), (1, 3))]);
    }

    #[test]
    fn ods_parse_ordinary_cells_have_no_merged_ranges() {
        let xml = r#"<office:spreadsheet>
<table:table table:name="Sheet1">
<table:table-row>
<table:table-cell office:value-type="float" office:value="1"/>
</table:table-row>
</table:table>
</office:spreadsheet>"#;
        let sheets = ods_parse(xml);
        assert!(sheets[0].merged_ranges.is_empty());
    }

    #[test]
    fn ods_parse_skips_column_position_past_a_repeated_empty_cell_run() {
        // LibreOffice represents a run of empty cells as ONE <table-cell
        // table:number-columns-repeated="N"/> rather than N elements — a
        // real value following that run must land at column 6, not 2.
        let xml = r#"<office:spreadsheet>
<table:table table:name="Sheet1">
<table:table-row>
<table:table-cell table:number-columns-repeated="5"/>
<table:table-cell office:value-type="float" office:value="42"/>
</table:table-row>
</table:table>
</office:spreadsheet>"#;
        let sheets = ods_parse(xml);
        assert!(!sheets[0].cells.contains_key(&(1, 2)));
        match sheets[0].cells.get(&(1, 6)) {
            Some(SheetCell::Integer(v)) => assert_eq!(*v, 42),
            other => panic!("expected Integer(42) at (1,6), got {:?}", other.is_some()),
        }
    }

    #[test]
    fn ods_parse_skips_row_position_past_a_repeated_empty_row_run() {
        let xml = r#"<office:spreadsheet>
<table:table table:name="Sheet1">
<table:table-row table:number-rows-repeated="4">
<table:table-cell office:value-type="string"><text:p>skip</text:p></table:table-cell>
</table:table-row>
<table:table-row>
<table:table-cell office:value-type="float" office:value="7"/>
</table:table-row>
</table:table>
</office:spreadsheet>"#;
        let sheets = ods_parse(xml);
        assert!(!sheets[0].cells.contains_key(&(2, 1)));
        match sheets[0].cells.get(&(5, 1)) {
            Some(SheetCell::Integer(v)) => assert_eq!(*v, 7),
            other => panic!("expected Integer(7) at (5,1), got {:?}", other.is_some()),
        }
    }

    #[test]
    fn xml_unescape_decodes_numeric_character_references() {
        assert_eq!(xml_unescape("&#65;&#x42;&#X43;"), "ABC");
    }

    #[test]
    fn xml_unescape_does_not_double_unescape_a_literal_escaped_entity() {
        // The text "&lt;" (a literal, already-escaped less-than sign)
        // written into an XML value must itself be escaped as "&amp;lt;".
        // Unescaping it once must yield "&lt;", not "<" — a chained
        // .replace("&amp;","&") then .replace("&lt;","<") would corrupt
        // this by unescaping twice.
        assert_eq!(xml_unescape("&amp;lt;"), "&lt;");
    }

    #[test]
    fn xml_unescape_leaves_an_unterminated_ampersand_literal() {
        assert_eq!(xml_unescape("a & b"), "a & b");
        assert_eq!(
            xml_unescape("a &notarealentity forever"),
            "a &notarealentity forever"
        );
    }

    // ── read() item 1: empty-string cell fix ────────────────────────────────

    #[test]
    fn xlsx_sheet_cells_records_a_zero_length_string_cell() {
        // <v></v> with zero characters between the tags — confirmed live this is exactly
        // what the oracle's own writer emits for an empty-string aoa cell (see
        // compat/differential/xlsx-read.test.mjs's dedicated case). Previously silently
        // absent (no Ev::Text event ever fires for an empty element).
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1" t="str"><v></v></c><c r="B1" t="str"><v>after</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        match data.cells.get(&(1, 1)) {
            Some(SheetCell::Str(s)) => assert_eq!(s, ""),
            other => panic!("expected Str(\"\") at A1, got {:?}", other.is_some()),
        }
        assert_eq!(data.cells.len(), 2);
    }

    #[test]
    fn xlsx_sheet_cells_honors_xml_space_preserve_on_v() {
        // Real shape confirmed live from compat/corpus/workbooks/with_text.xlsx's own raw
        // sheet1.xml (cell A3) — see compat/differential/classify.mjs's now-removed
        // XML_SPACE_PRESERVE_DEFECT entry for the defect this fixes. B1 (no xml:space) also
        // confirms v_preserve_space is read fresh per-<v> rather than sticky from A1, and
        // that the default (still-trimming) behavior is unaffected by the fix.
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1" t="str"><v xml:space="preserve">  padded  </v></c><c r="B1" t="str"><v>  not preserved  </v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        match data.cells.get(&(1, 1)) {
            Some(SheetCell::Str(s)) => assert_eq!(s, "  padded  "),
            other => panic!(
                "expected Str(\"  padded  \") at A1, got {:?}",
                other.is_some()
            ),
        }
        match data.cells.get(&(1, 2)) {
            Some(SheetCell::Str(s)) => assert_eq!(s, "not preserved"),
            other => panic!(
                "expected Str(\"not preserved\") at B1, got {:?}",
                other.is_some()
            ),
        }
    }

    #[test]
    fn xlsx_sheet_cells_xml_space_preserve_on_a_numeric_v_still_parses_when_untrimmed() {
        // Real Excel/SheetJS writers never emit this combination (xml:space="preserve" only
        // ever marks up literal string text) — this just confirms the fix doesn't newly
        // break a numeric cell that happens to carry the attribute without surrounding
        // whitespace, since Rust's f64::parse rejects leading/trailing whitespace.
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><v xml:space="preserve">42</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        match data.cells.get(&(1, 1)) {
            Some(SheetCell::Integer(n)) => assert_eq!(*n, 42),
            other => panic!("expected Integer(42) at A1, got {:?}", other.is_some()),
        }
    }

    #[test]
    fn xlsx_sheet_cells_empty_v_on_a_numeric_cell_yields_no_cell() {
        // No t= attribute -> numeric parsing. "".parse::<f64>() fails, so (matching a
        // cell with no <v> content at all) no cell is inserted — the fix must not invent
        // a numeric value out of an empty string.
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><v></v></c></row></sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert!(data.cells.is_empty());
    }

    // ── read() item 2: <dimension> parsing ──────────────────────────────────

    #[test]
    fn parse_dimension_ref_reads_a_colon_separated_range() {
        assert_eq!(parse_dimension_ref("A1:C3"), Some(((1, 1), (3, 3))));
        assert_eq!(parse_dimension_ref("A1:A1"), Some(((1, 1), (1, 1))));
    }

    #[test]
    fn parse_dimension_ref_rejects_a_colon_less_single_cell_ref() {
        // Mirrors the oracle's own dimregex (/"(\w*:\w*)"/), which requires a literal
        // colon — a bare "A1" never matches on the oracle either, confirmed by reading
        // xlsx.js's parse_ws_xml_dim call site directly.
        assert_eq!(parse_dimension_ref("A1"), None);
    }

    #[test]
    fn parse_dimension_ref_rejects_a_reversed_range() {
        assert_eq!(parse_dimension_ref("C3:A1"), None);
    }

    #[test]
    fn xlsx_sheet_cells_reads_a_dimension_wider_than_the_populated_cells() {
        let xml = r#"<worksheet>
<dimension ref="A1:E10"/>
<sheetData>
<row r="1"><c r="A1" t="str"><v>a</v></c></row>
</sheetData>
</worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.dimension, Some(((1, 1), (10, 5))));
    }

    #[test]
    fn xlsx_sheet_cells_dimension_is_none_when_the_tag_is_absent() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.dimension, None);
    }

    // ── read() item 4: formula (<f>) capture ────────────────────────────────

    #[test]
    fn xlsx_sheet_cells_captures_inline_formula_text() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><f>SUM(B1:B2)</f><v>3</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(
            data.formulas.get(&(1, 1)).map(String::as_str),
            Some("SUM(B1:B2)")
        );
        match data.cells.get(&(1, 1)) {
            Some(SheetCell::Integer(v)) => assert_eq!(*v, 3),
            other => panic!("expected Integer(3) at A1, got {:?}", other.is_some()),
        }
    }

    #[test]
    fn xlsx_sheet_cells_shared_formula_follower_with_no_inline_text_captures_nothing() {
        // <f t="shared" si="0"/> (self-closing, no formula text) — the master cell of a
        // shared-formula group carries the real text; a follower cell doesn't. reader.rs
        // doesn't resolve/shift shared-formula text, so this cell simply has no captured
        // formula (an honest gap, not a wrong value) — see BufferSheet::formulas's doc
        // comment.
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1</f><v>1</v></c>
<c r="A2"><f t="shared" si="0"/><v>2</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.formulas.get(&(1, 1)).map(String::as_str), Some("B1"));
        assert_eq!(data.formulas.get(&(1, 2)), None);
    }

    #[test]
    fn xlsx_sheet_cells_formula_text_is_xml_unescaped() {
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1"><f>A1&amp;"x"</f><v>1</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(
            data.formulas.get(&(1, 1)).map(String::as_str),
            Some(r#"A1&"x""#)
        );
    }

    // ── read() item 6: styles.xml (numFmts/cellXfs), date1904 ──────────────

    #[test]
    fn xlsx_styles_reads_custom_number_formats_and_cell_xfs_in_order() {
        let xml = r#"<styleSheet>
<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00&quot;kg&quot;"/></numFmts>
<cellXfs count="3">
<xf numFmtId="0"/>
<xf numFmtId="2"/>
<xf numFmtId="164"/>
</cellXfs>
</styleSheet>"#;
        let styles = xlsx_styles(xml);
        assert_eq!(
            styles.number_formats.get(&164).map(String::as_str),
            Some(r#"0.00"kg""#)
        );
        assert_eq!(styles.cell_xfs, vec![Some(0), Some(2), Some(164)]);
    }

    #[test]
    fn xlsx_styles_an_xf_with_no_numfmtid_attribute_resolves_to_none() {
        let xml = r#"<styleSheet><cellXfs count="1"><xf fontId="0"/></cellXfs></styleSheet>"#;
        let styles = xlsx_styles(xml);
        assert_eq!(styles.cell_xfs, vec![None]);
    }

    #[test]
    fn xlsx_styles_ignores_xf_entries_outside_cell_xfs() {
        // <cellStyleXfs>'s <xf> entries must NOT leak into cell_xfs — only <cellXfs>'s
        // children are the ones a cell's s="N" attribute indexes into (matching the
        // oracle's own styles.CellXf, built from <cellXfs> alone).
        let xml = r#"<styleSheet>
<cellStyleXfs count="1"><xf numFmtId="9"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="14"/></cellXfs>
</styleSheet>"#;
        let styles = xlsx_styles(xml);
        assert_eq!(styles.cell_xfs, vec![Some(14)]);
    }

    #[test]
    fn xlsx_styles_handles_an_empty_self_closing_cell_xfs() {
        let xml = r#"<styleSheet><cellXfs count="0"/></styleSheet>"#;
        let styles = xlsx_styles(xml);
        assert!(styles.cell_xfs.is_empty());
    }

    #[test]
    fn xlsx_sheet_cells_resolves_a_cells_s_attribute_through_cell_xfs() {
        let cell_xfs = vec![Some(0u32), Some(14u32)];
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1" s="1"><v>45444</v></c><c r="B1" s="0"><v>1</v></c><c r="C1"><v>2</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &cell_xfs);
        assert_eq!(data.style_ids.get(&(1, 1)), Some(&14));
        // s="0" (General) and no s= at all both resolve to "no entry" (0 == absent).
        assert_eq!(data.style_ids.get(&(1, 2)), None);
        assert_eq!(data.style_ids.get(&(1, 3)), None);
    }

    #[test]
    fn xlsx_sheet_cells_an_out_of_range_s_index_resolves_to_no_style() {
        let cell_xfs = vec![Some(14u32)];
        let xml = r#"<worksheet><sheetData><row r="1"><c r="A1" s="99"><v>1</v></c></row></sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &cell_xfs);
        assert_eq!(data.style_ids.get(&(1, 1)), None);
    }

    #[test]
    fn xlsx_workbook_date1904_defaults_to_false_when_absent() {
        let xml = r#"<workbook><sheets></sheets></workbook>"#;
        assert!(!xlsx_workbook_date1904(xml));
    }

    #[test]
    fn xlsx_workbook_date1904_reads_the_declared_flag() {
        let xml = r#"<workbook><workbookPr date1904="1"/><sheets></sheets></workbook>"#;
        assert!(xlsx_workbook_date1904(xml));
        // The oracle's own writer/reader accepts "true" too (xsd:boolean), not just "1".
        let xml2 = r#"<workbook><workbookPr date1904="true"/></workbook>"#;
        assert!(xlsx_workbook_date1904(xml2));
        let xml3 = r#"<workbook><workbookPr date1904="0"/></workbook>"#;
        assert!(!xlsx_workbook_date1904(xml3));
    }
}

// ── Buffer-API resolution: read_workbook_from_bytes ─────────────────────────
#[cfg(test)]
mod from_bytes_tests {
    use super::*;

    // The path-based and bytes-based entry points must read the exact same real .xlsx
    // fixture into equal sheet data — read_workbook_from_bytes is meant to be a pure
    // buffer-input alternative to read_workbook(path), not a second implementation with
    // its own drift (see docs/xlsx-architecture.md's "reader.rs buffer-API resolution").
    fn cell_map_eq(a: &HashMap<(u32, u32), SheetCell>, b: &HashMap<(u32, u32), SheetCell>) -> bool {
        if a.len() != b.len() {
            return false;
        }
        a.iter().all(|(k, v)| match (v, b.get(k)) {
            (SheetCell::Integer(x), Some(SheetCell::Integer(y))) => x == y,
            (SheetCell::Float(x), Some(SheetCell::Float(y))) => x == y,
            (SheetCell::Str(x), Some(SheetCell::Str(y))) => x == y,
            (SheetCell::Bool(x), Some(SheetCell::Bool(y))) => x == y,
            _ => false,
        })
    }

    #[test]
    fn read_workbook_from_bytes_matches_read_workbook_on_a_real_xlsx_fixture() {
        let path = concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/tests/fixtures/e2e/source.xlsx"
        );
        let from_path = read_workbook(path).expect("read_workbook(path) should succeed");
        let bytes = std::fs::read(path).expect("fixture should be readable");
        let from_bytes =
            read_workbook_from_bytes(&bytes).expect("read_workbook_from_bytes should succeed");

        assert_eq!(from_path.len(), from_bytes.sheets.len());
        for (a, bs) in from_path.iter().zip(from_bytes.sheets.iter()) {
            let b = &bs.sheet;
            assert_eq!(a.name, b.name);
            assert_eq!(a.sheet_id, b.sheet_id);
            assert_eq!(a.merged_ranges, b.merged_ranges);
            assert_eq!(a.hidden_rows, b.hidden_rows);
            assert_eq!(a.hidden_columns, b.hidden_columns);
            assert!(cell_map_eq(&a.cells, &b.cells));
        }
    }

    #[test]
    fn read_workbook_from_bytes_rejects_a_non_zip_buffer() {
        assert!(read_workbook_from_bytes(b"not a zip file").is_err());
    }
}
