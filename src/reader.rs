// Minimal XLSX/ODS reader — replaces calamine as a runtime dependency.
// Supports: .xlsx, .xlsm (Office Open XML ZIP), .ods (OpenDocument ZIP).
// Row/col indices are 1-based, matching the VM's convention.

use std::collections::HashMap;
use std::io::Read;
use zip::ZipArchive;

// ── Public types ──────────────────────────────────────────────────────────────

pub struct WorkbookSheet {
    pub name: String,
    pub cells: HashMap<(u32, u32), SheetCell>,
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
    fn new(s: &'a str) -> Self { XmlIter { s } }

    fn next_ev(&mut self) -> Option<Ev> {
        loop {
            if self.s.is_empty() { return None; }

            if !self.s.starts_with('<') {
                // Text node — preserve verbatim (trim happens at call site for leaf nodes)
                let end = self.s.find('<').unwrap_or(self.s.len());
                let raw = &self.s[..end];
                self.s = &self.s[end..];
                let text = xml_unescape(raw);
                if text.is_empty() { continue; }
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
                if !text.is_empty() { return Some(Ev::Text(text)); }
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
            if c == qchar { in_quote = false; }
        } else {
            match c {
                '"' | '\'' => { in_quote = true; qchar = c; }
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
        if s.is_empty() { break; }
        let Some(eq) = s.find('=') else { break };
        let name = s[..eq].trim().to_string();
        if name.is_empty() { break; }
        s = s[eq + 1..].trim_start();
        let Some(quote) = s.chars().next() else { break };
        if quote != '"' && quote != '\'' { break; }
        s = &s[1..]; // skip opening quote
        let end = s.find(quote).unwrap_or(s.len());
        let value = xml_unescape(&s[..end]);
        s = &s[(end + 1).min(s.len())..];
        attrs.push(Attr { name, value });
    }
    attrs
}

fn attr_get<'a>(attrs: &'a [Attr], name: &str) -> Option<&'a str> {
    attrs.iter()
        .find(|a| a.name == name || a.name.split(':').next_back() == Some(name))
        .map(|a| a.value.as_str())
}

fn xml_unescape(s: &str) -> String {
    if !s.contains('&') { return s.to_string(); }
    s.replace("&amp;",  "&")
     .replace("&lt;",   "<")
     .replace("&gt;",   ">")
     .replace("&quot;", "\"")
     .replace("&apos;", "'")
}

// ── Helper: read a ZIP entry into a String ────────────────────────────────────

/// 64 MB decompressed cap per entry — enough for any real spreadsheet XML.
const ZIP_ENTRY_MAX_BYTES: u64 = 64 * 1024 * 1024;

fn zip_read_text(archive: &mut ZipArchive<std::fs::File>, name: &str) -> Result<String, String> {
    let mut entry = archive.by_name(name).map_err(|e| format!("{}: {}", name, e))?;
    let mut s = String::new();
    entry.by_ref().take(ZIP_ENTRY_MAX_BYTES).read_to_string(&mut s).map_err(|e| e.to_string())?;
    Ok(s)
}

// ── XLSX reader ───────────────────────────────────────────────────────────────

fn read_xlsx(path: &str) -> Result<Vec<WorkbookSheet>, String> {
    let file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;

    let wb_xml = zip_read_text(&mut archive, "xl/workbook.xml")?;
    let sheet_refs = xlsx_workbook_sheets(&wb_xml);

    let rels_xml = zip_read_text(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels = xlsx_rels(&rels_xml);

    let shared: Vec<String> = match zip_read_text(&mut archive, "xl/sharedStrings.xml") {
        Ok(xml) => xlsx_shared_strings(&xml),
        Err(_)  => vec![],
    };

    let mut sheets = vec![];
    for (name, rid) in sheet_refs {
        let Some(target) = rels.get(&rid) else { continue };
        let zip_path = if target.starts_with('/') {
            target[1..].to_string()
        } else {
            format!("xl/{}", target)
        };
        let sheet_xml = match zip_read_text(&mut archive, &zip_path) {
            Ok(s) => s, Err(_) => continue,
        };
        let cells = xlsx_sheet_cells(&sheet_xml, &shared);
        sheets.push(WorkbookSheet { name, cells });
    }
    Ok(sheets)
}

/// Returns `[(sheet_name, rId)]` in document order.
fn xlsx_workbook_sheets(xml: &str) -> Vec<(String, String)> {
    let mut iter = XmlIter::new(xml);
    let mut result = vec![];
    while let Some(ev) = iter.next_ev() {
        if let Ev::SelfClose(ref tag, ref attrs) = ev {
            let local = tag.split(':').next_back().unwrap_or(tag);
            if local == "sheet" {
                if let (Some(name), Some(rid)) = (
                    attr_get(attrs, "name"),
                    attr_get(attrs, "id"),
                ) {
                    result.push((name.to_string(), rid.to_string()));
                }
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
            if local == "Relationship" {
                if let (Some(id), Some(ty), Some(target)) = (
                    attr_get(attrs, "Id"),
                    attr_get(attrs, "Type"),
                    attr_get(attrs, "Target"),
                ) {
                    if ty.ends_with("/worksheet") {
                        map.insert(id.to_string(), target.to_string());
                    }
                }
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
                    "si" => { in_si = true; current.clear(); }
                    "t"  => { in_t = true; }
                    _ => {}
                }
            }
            Ev::Close(tag) => {
                let local = tag.split(':').next_back().unwrap_or(tag);
                match local {
                    "si" => { strings.push(current.clone()); in_si = false; }
                    "t"  => { in_t = false; }
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

/// Parses a single worksheet XML into a 1-based (row, col) → SheetCell map.
fn xlsx_sheet_cells(xml: &str, shared: &[String]) -> HashMap<(u32, u32), SheetCell> {
    let mut iter = XmlIter::new(xml);
    let mut cells: HashMap<(u32, u32), SheetCell> = HashMap::new();
    let mut cur_row: u32 = 0;
    let mut cur_col: u32 = 0;
    let mut cur_type = String::new();
    let mut in_v = false;
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
                    }
                    "c" => {
                        cur_type = attr_get(attrs, "t").unwrap_or("").to_string();
                        in_v = false;
                        if let Some(r) = attr_get(attrs, "r") {
                            if let Some((row, col)) = parse_cell_ref(r) {
                                cur_row = row;
                                cur_col = col;
                            }
                        }
                        is_text.clear();
                    }
                    "v" => { in_v = true; }
                    "t" => {
                        // inside <is> for inline strings
                        in_is_t = true;
                        is_text.clear();
                    }
                    _ => {}
                }
            }
            Ev::Close(ref tag) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "v"  => { in_v = false; }
                    "t"  => { in_is_t = false; }
                    _ => {}
                }
            }
            Ev::Text(ref text) => {
                if in_v && cur_row > 0 && cur_col > 0 {
                    let cell = xlsx_parse_cell(text.trim(), &cur_type, shared);
                    if let Some(c) = cell {
                        cells.insert((cur_row, cur_col), c);
                    }
                    in_v = false;
                } else if in_is_t {
                    is_text.push_str(text);
                }
            }
        }

        // Emit inline string on </c>
        if let Ev::Close(ref tag) = ev {
            if tag.split(':').next_back() == Some("c")
                && cur_type == "inlineStr"
                && !is_text.is_empty()
                && cur_row > 0 && cur_col > 0
            {
                cells.insert((cur_row, cur_col), SheetCell::Str(is_text.clone()));
                is_text.clear();
            }
        }
    }
    cells
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
    if alpha_end == 0 { return None; }
    let col = r[..alpha_end]
        .chars()
        .fold(0u32, |acc, c| acc * 26 + (c as u32 - 'A' as u32 + 1));
    let row: u32 = r[alpha_end..].parse().ok()?;
    Some((row, col))
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
    let mut cell_type = String::new();
    let mut cell_val_attr = String::new();
    let mut cell_bool_attr = String::new();
    let mut col_repeat: u32 = 1;
    let mut row_repeat: u32 = 1;
    let mut in_text_p = false;
    let mut cell_text = String::new();
    let mut pending_cell: Option<OdsCellState> = None;

    while let Some(ev) = iter.next_ev() {
        match &ev {
            Ev::Open(tag, attrs) | Ev::SelfClose(tag, attrs) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "table" => {
                        let name = attr_get(attrs, "name")
                            .unwrap_or("sheet1")
                            .to_lowercase();
                        sheets.push(WorkbookSheet { name, cells: HashMap::new() });
                        in_sheet = true;
                        row = 0;
                        col = 0;
                    }
                    "table-row" if in_sheet => {
                        row_repeat = attr_get(attrs, "number-rows-repeated")
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(1);
                        // Cap row_repeat to avoid OOM on trailing empty rows
                        if row_repeat > 1024 { row_repeat = 1; }
                        row += 1;
                        col = 0;
                        pending_cell = None;
                    }
                    "table-cell" | "covered-table-cell" if in_sheet => {
                        // Flush previous pending cell before starting a new one
                        if let Some(state) = pending_cell.take() {
                            emit_ods_cell(&mut sheets, state);
                        }
                        col_repeat = attr_get(attrs, "number-columns-repeated")
                            .and_then(|s| s.parse().ok())
                            .unwrap_or(1);
                        // Cap to avoid OOM on trailing empty columns
                        if col_repeat > 1024 { col_repeat = 1; }
                        col += 1;
                        cell_type = attr_get(attrs, "value-type").unwrap_or("").to_string();
                        cell_val_attr = attr_get(attrs, "value").unwrap_or("").to_string();
                        cell_bool_attr = attr_get(attrs, "boolean-value").unwrap_or("").to_string();
                        cell_text.clear();
                        in_text_p = false;

                        if matches!(ev, Ev::SelfClose(_, _)) {
                            // Self-closing: no text content, emit immediately
                            let state = OdsCellState {
                                row, col, col_repeat,
                                cell_type: cell_type.clone(),
                                val_attr: cell_val_attr.clone(),
                                bool_attr: cell_bool_attr.clone(),
                                text: String::new(),
                            };
                            emit_ods_cell(&mut sheets, state);
                            pending_cell = None;
                        } else {
                            pending_cell = Some(OdsCellState {
                                row, col, col_repeat,
                                cell_type: cell_type.clone(),
                                val_attr: cell_val_attr.clone(),
                                bool_attr: cell_bool_attr.clone(),
                                text: String::new(),
                            });
                        }
                    }
                    "p" if in_sheet => { in_text_p = true; }
                    _ => {}
                }
            }
            Ev::Close(tag) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "table" => { in_sheet = false; }
                    "table-cell" | "covered-table-cell" if in_sheet => {
                        if let Some(ref mut state) = pending_cell {
                            state.text.clone_from(&cell_text);
                        }
                        if let Some(state) = pending_cell.take() {
                            emit_ods_cell(&mut sheets, state);
                        }
                        in_text_p = false;
                    }
                    "p" => { in_text_p = false; }
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
    row: u32, col: u32, col_repeat: u32,
    cell_type: String, val_attr: String, bool_attr: String, text: String,
}

fn emit_ods_cell(sheets: &mut Vec<WorkbookSheet>, state: OdsCellState) {
    let sheet = match sheets.last_mut() { Some(s) => s, None => return };
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
            if s.text.is_empty() { None } else { Some(SheetCell::Str(s.text.clone())) }
        }
        "boolean" => {
            Some(SheetCell::Bool(s.bool_attr == "true"))
        }
        _ => None, // empty / formula result not available / etc.
    }
}
