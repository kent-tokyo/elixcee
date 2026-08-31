// Minimal XLSX/ODS reader — replaces calamine as a runtime dependency.
// Supports: .xlsx, .xlsm (Office Open XML ZIP), .ods (OpenDocument ZIP).
// Row/col indices are 1-based, matching the VM's convention.

use std::collections::HashMap;
use std::io::{Cursor, Read, Seek};
use std::str::FromStr;

use elixcee_types::ExcelError;
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
    /// Per-cell resolved number-format code string (GitHub #4: a date-formatted cell's
    /// serial number gave no way for a Python caller to tell it was a date at all) --
    /// e.g. `"m/d/yyyy"` for numFmtId 14, or a custom `<numFmt formatCode="...">` string.
    /// Resolved once at read time via `resolve_number_format` (custom formats checked
    /// first, then the built-in table) from each cell's `raw_style_indices` entry through
    /// `xl/styles.xml`'s `<cellXfs>`. Absent for a cell with no format, the General
    /// format (numFmtId 0), or an unresolvable id -- exposing `None`/no-entry rather
    /// than guessing is the same convention `BufferSheet::style_ids` already uses. Always
    /// empty for `.ods` (no equivalent parsed there).
    pub cell_number_formats: HashMap<(u32, u32), String>,
    /// The XLSX `<sheet state="...">` attribute's raw value, exactly as written --
    /// `Some("hidden")`/`Some("veryHidden")`, or `None` when the attribute is absent
    /// (the default, meaning visible) or unrecognized. Kept as the raw string here,
    /// the same way `cell_number_formats` keeps raw format-code strings rather than a
    /// resolved type -- `Vm::populate_from_sheets` is what turns this into a proper
    /// `SheetState`. Always `None` for `.ods` (no equivalent attribute).
    pub sheet_state: Option<String>,
    /// Per-row explicit height in points (P2), from `<row r=".." ht=".."
    /// customHeight="1">` -- sparse, only rows with an explicit height get an
    /// entry. `customHeight="1"` is required for `ht` to actually apply in real
    /// Excel, so a bare `ht` without it is not recorded. Always empty for `.ods`
    /// (deferred, same as `hidden_rows`).
    pub row_heights: HashMap<u32, f64>,
    /// Column width ranges in "characters" (P2), 1-based inclusive
    /// `(min, max, width)`, from `<col min=".." max=".." width=".."
    /// customWidth="1"/>` -- same `customWidth="1"`-required caveat as
    /// `row_heights`. Always empty for `.ods` (deferred, same as `hidden_columns`).
    pub column_widths: Vec<(u32, u32, f64)>,
    /// Per-row default style index (0.15.0-C2), from `<row r=".." s=".."
    /// customFormat="1">` -- `customFormat="1"` is required for `s` to mean
    /// "this row's own default style", same required-flag caveat as
    /// `row_heights`' `customHeight`. Always empty for `.ods`.
    pub row_styles: HashMap<u32, u32>,
    /// Column default-style ranges (0.15.0-C2), 1-based inclusive `(min, max,
    /// style_index)`, from `<col min=".." max=".." style=".."/>`. Always
    /// empty for `.ods`.
    pub column_styles: Vec<(u32, u32, u32)>,
    /// Tables defined on this sheet (0.16.0-A1), parsed from each `xl/tables/tableN.xml`
    /// part the sheet's own `<tableParts>` links to. Read-only: this is a pure `Vm`-side
    /// projection, not a writer input -- an unmodified table's bytes keep surviving via
    /// the existing generic unknown-part passthrough + `OpaqueWorksheetFragments::
    /// table_parts` splice exactly as before this field existed (see
    /// `internal_docs/tables-0.16.0-a-design.md` Finding 2). Structured references
    /// (`Table1[@Qty]`) are out of scope entirely -- `TableColumn::
    /// calculated_column_formula` is captured as raw, unparsed text. Always empty for
    /// `.ods` (no table concept there).
    pub tables: Vec<TableDef>,
    /// Data-validation rules defined directly inside this sheet's own XML (0.16.0-C,
    /// no separate part/relationship, unlike `tables` above). Read-only at parse time,
    /// like `tables`; `Vm::add_data_validation_on_sheet`/`remove_data_validation_on_sheet`
    /// mutate the `Vm`-side copy afterward. Always empty for `.ods`.
    pub data_validations: Vec<DataValidationRule>,
    /// Standalone (worksheet-level) `<autoFilter>`, when present (0.16.0-B) -- `None` for
    /// a sheet with no autofilter at all, distinct from `Some` with an empty `columns`
    /// list (a bare `<autoFilter ref="...">` with no active criteria, real Excel's own
    /// shape right after turning AutoFilter on but before filtering anything). At most
    /// one per sheet, per `CT_Worksheet`'s own `maxOccurs="1"` on this child -- unlike
    /// `tables`, never a `Vec`. A table's own NESTED `<autoFilter>` is a completely
    /// separate thing (`TableDef::auto_filter_ref`/`autofilter_columns`, its own
    /// storage and write path, 0.16.0-B2) -- not this field. Always `None` for `.ods`.
    pub autofilter: Option<AutoFilterDef>,
}

/// One `<filterColumn>`'s criteria payload (0.16.0-B) -- `ST_FilterColumn`'s child
/// choice, restricted to the shapes the roadmap's own filter-type list names.
/// `colorFilter`/`iconFilter`/`dynamicFilter`/`extLst` are out of scope entirely.
#[derive(Clone, Debug, PartialEq)]
pub enum FilterCriteria {
    /// `<filters><filter val="..."/>...</filters>` -- passes if the cell's text matches
    /// ANY value in the list (Excel's checkbox-list filter).
    Values(Vec<String>),
    /// `<customFilters and="0|1"><customFilter operator="..." val="..."/>...</customFilters>`
    /// -- one or two conditions. `ST_FilterOperator` has no `between`/`notBetween` (that's
    /// a data-validation-only operator set, confirmed against real generated bytes, not
    /// assumed from `DataValidationRule`'s own operator list) -- a numeric/date "between"
    /// range is two conditions joined by `and=true` instead.
    Custom {
        op1: String,
        val1: String,
        and: bool,
        op2: Option<String>,
        val2: Option<String>,
    },
    /// `<filters blank="1"/>`.
    Blank,
    /// `<top10 top="0|1" percent="0|1" val="N"/>` -- `top=false` means bottom-N, not
    /// "not top-N" (real Excel's own "Top 10" dialog has a Top/Bottom toggle).
    Top10 { top: bool, percent: bool, val: f64 },
    /// `<filters calendarType="..."><dateGroupItem .../>...</filters>` -- passes if the
    /// cell's date decomposes to match ANY item (each item's absent fields are
    /// wildcards, which is also how `dateTimeGrouping`'s own granularity is implied
    /// without needing to interpret that attribute during evaluation).
    DateGroup(Vec<DateGroupItem>),
}

/// One `<dateGroupItem>` -- absent fields are wildcards for matching (0.16.0-B).
#[derive(Clone, Debug, PartialEq)]
pub struct DateGroupItem {
    pub year: Option<i32>,
    pub month: Option<u32>,
    pub day: Option<u32>,
    pub hour: Option<u32>,
    pub minute: Option<u32>,
    pub second: Option<u32>,
    pub date_time_grouping: String,
}

/// One `<filterColumn>` record (0.16.0-B), parsed from a raw span `extract_records`
/// isolates inside an `<autoFilter>`. `raw_span`/`dirty` mirror `DataValidationRule`'s
/// own write-time-source-of-truth pattern: `Some(raw_span)` with `dirty=false` means
/// "untouched since load, re-emit these exact bytes" (preserves anything this struct
/// doesn't model, e.g. a real `extLst`); `Some(raw_span)` with `dirty=true` means "only
/// `colId` moved, patch just that attribute via `with_attr`"; `None` means "built fresh
/// from `criteria` by a `set_*_filter` call, nothing to preserve."
#[derive(Clone, Debug, PartialEq)]
pub struct FilterColumn {
    /// 0-based, relative to the autofilter's own `ref` left edge -- NOT the same
    /// convention as VBA's `RangeAutoFilter` `Field` (1-based). Confirmed against real
    /// generated bytes: `colId="0"` is the leftmost column of `ref`.
    pub col_offset: u32,
    pub hidden_button: bool,
    /// Defaults to `true` -- every real `<filterColumn>` seen has `showButton="1"`
    /// (or omits it, which real Excel/openpyxl both treat as shown).
    pub show_button: bool,
    pub criteria: FilterCriteria,
    pub raw_span: Option<String>,
    pub dirty: bool,
}

/// A parsed standalone `<autoFilter>` (0.16.0-B), read from a sheet's own XML (no
/// separate part/relationship, unlike `TableDef`).
#[derive(Clone, Debug, PartialEq)]
pub struct AutoFilterDef {
    /// 1-based inclusive `(top-left, bottom-right)`, header row (`ref_range.0.0`)
    /// included -- same convention as `TableDef::ref_range`. AutoFilter never hides the
    /// header row itself (`Stmt::RangeAutoFilter`'s own real VBA behavior, carried
    /// forward here).
    pub ref_range: MergeRect,
    pub columns: Vec<FilterColumn>,
}

fn parse_filter_criteria_xml(filter_column_span: &str) -> Option<FilterCriteria> {
    let mut iter = XmlIter::new(filter_column_span);
    let mut values = Vec::new();
    let mut date_groups = Vec::new();
    let mut is_blank = false;
    let mut custom_filters: Vec<(String, String)> = Vec::new();
    let mut custom_and = true;
    let mut top10: Option<(bool, bool, f64)> = None;

    while let Some(ev) = iter.next_ev() {
        let (Ev::Open(ref tag, ref attrs) | Ev::SelfClose(ref tag, ref attrs)) = ev else {
            continue;
        };
        match tag.split(':').next_back().unwrap_or(tag.as_str()) {
            "filters" => {
                if matches!(attr_get(attrs, "blank"), Some("1")) {
                    is_blank = true;
                }
            }
            "filter" => {
                if let Some(v) = attr_get(attrs, "val") {
                    values.push(v.to_string());
                }
            }
            "dateGroupItem" => {
                date_groups.push(DateGroupItem {
                    year: attr_get(attrs, "year").and_then(|v| v.parse().ok()),
                    month: attr_get(attrs, "month").and_then(|v| v.parse().ok()),
                    day: attr_get(attrs, "day").and_then(|v| v.parse().ok()),
                    hour: attr_get(attrs, "hour").and_then(|v| v.parse().ok()),
                    minute: attr_get(attrs, "minute").and_then(|v| v.parse().ok()),
                    second: attr_get(attrs, "second").and_then(|v| v.parse().ok()),
                    date_time_grouping: attr_get(attrs, "dateTimeGrouping")
                        .unwrap_or("")
                        .to_string(),
                });
            }
            "customFilters" => {
                custom_and = matches!(attr_get(attrs, "and"), Some("1"));
            }
            "customFilter" => {
                custom_filters.push((
                    attr_get(attrs, "operator").unwrap_or("equal").to_string(),
                    attr_get(attrs, "val").unwrap_or("").to_string(),
                ));
            }
            "top10" => {
                top10 = Some((
                    attr_get(attrs, "top").map(|v| v != "0").unwrap_or(true),
                    matches!(attr_get(attrs, "percent"), Some("1")),
                    attr_get(attrs, "val")
                        .and_then(|v| v.parse().ok())
                        .unwrap_or(0.0),
                ));
            }
            _ => {}
        }
    }
    if let Some((top, percent, val)) = top10 {
        return Some(FilterCriteria::Top10 { top, percent, val });
    }
    if !custom_filters.is_empty() {
        let mut it = custom_filters.into_iter();
        let (op1, val1) = it.next().unwrap();
        let (op2, val2) = match it.next() {
            Some((o, v)) => (Some(o), Some(v)),
            None => (None, None),
        };
        return Some(FilterCriteria::Custom {
            op1,
            val1,
            and: custom_and,
            op2,
            val2,
        });
    }
    if !date_groups.is_empty() {
        return Some(FilterCriteria::DateGroup(date_groups));
    }
    if is_blank {
        return Some(FilterCriteria::Blank);
    }
    if !values.is_empty() {
        return Some(FilterCriteria::Values(values));
    }
    None
}

fn parse_filter_column_xml(span: &str) -> Option<FilterColumn> {
    let (tag_start, tag_close_rel, full_name) = find_next_open_tag(span, 0)?;
    let name_end = tag_start + 1 + full_name.len();
    let raw_attrs = &span[name_end..name_end + tag_close_rel];
    let attrs_str = raw_attrs.trim_end().strip_suffix('/').unwrap_or(raw_attrs);
    let attrs = parse_attrs(attrs_str);
    let col_offset = attr_get(&attrs, "colId")?.parse().ok()?;
    Some(FilterColumn {
        col_offset,
        hidden_button: matches!(attr_get(&attrs, "hiddenButton"), Some("1")),
        show_button: attr_get(&attrs, "showButton")
            .map(|v| v != "0")
            .unwrap_or(true),
        criteria: parse_filter_criteria_xml(span)?,
        raw_span: Some(span.to_string()),
        dirty: false,
    })
}

/// Parses a sheet's standalone `<autoFilter>`, if any (0.16.0-B). Tolerant of any single
/// `<filterColumn>` that fails to parse (skipped, matching `xlsx_data_validations`'s own
/// per-record tolerance) -- a bare `<autoFilter ref="...">` with no children at all
/// parses to `Some` with an empty `columns` list, not `None`.
pub(crate) fn xlsx_autofilter(sheet_xml: &str) -> Option<AutoFilterDef> {
    let span = extract_raw_element(sheet_xml, "autoFilter")?;
    let ref_range = span_attr_str(&span, "ref").and_then(|s| parse_merge_ref(&s))?;
    let columns = extract_records(&span, "autoFilter", "filterColumn")
        .iter()
        .filter_map(|s| parse_filter_column_xml(s))
        .collect();
    Some(AutoFilterDef { ref_range, columns })
}

/// One `<dataValidation>` record (0.16.0-C), parsed from a raw span `extract_records`
/// already isolated. `raw_span` is the write-time source of truth: it preserves every
/// attribute/child this struct doesn't model (e.g. the real `xr:uid` extension GUID seen
/// on `fixture3`'s own data validation) byte-for-byte for anything not explicitly
/// re-patched. `dirty` marks a rule whose `sqref` has been shifted since `raw_span` was
/// captured (a structural edit) -- the writer only re-patches `raw_span` (via
/// `with_attr`, touching just the `sqref` attribute) for `dirty` rules, so an untouched
/// rule's bytes never move at all. Freshly-added rules (`add_data_validation_on_sheet`)
/// build `raw_span` directly from the given fields and are never `dirty` -- there is
/// nothing stale to reconcile.
#[derive(Clone, Debug, PartialEq)]
pub struct DataValidationRule {
    pub validation_type: String,
    pub operator: Option<String>,
    /// Raw `<formula1>` text, unparsed and unevaluated -- a `list` type's formula1 is
    /// often a literal comma-separated string (`"Yes,No,Maybe"`) rather than a cell
    /// reference; this project's own persist-only scope for 0.16.0 never needs to tell
    /// the two apart.
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub allow_blank: bool,
    pub show_input_message: bool,
    pub prompt_title: Option<String>,
    pub prompt: Option<String>,
    pub show_error_message: bool,
    pub error_style: Option<String>,
    pub error_title: Option<String>,
    pub error: Option<String>,
    /// Parsed `sqref` areas, 1-based inclusive rects -- `ST_Sqref` is a SPACE-delimited
    /// list of ranges (distinct from `<definedName>`'s comma-delimited multi-area
    /// grammar), each of which may be a single cell (`"E1"`, no colon, confirmed against
    /// `fixture3`'s real `sqref="E1"`) or a range (`"A1:C4"`).
    pub sqref: Vec<MergeRect>,
    pub dirty: bool,
    pub raw_span: String,
}

/// Fields for a NEW data-validation rule (`Vm::add_data_validation_on_sheet`), grouped
/// into one struct to keep that method's (and `build_data_validation_span`'s, `src/
/// lib.rs`) signature small -- mirrors `StyleAttrEdit`'s own grouping of `set_style`'s
/// many optional fields. `show_input_message`/`show_error_message` are deliberately NOT
/// separate fields here: the caller-facing `add_data_validation` derives them from
/// whether a prompt/error message was actually given (see that method's own doc
/// comment), so by the time a `DataValidationSpec` exists they're already resolved and
/// live on `DataValidationRule` directly, not duplicated here.
#[derive(Clone, Debug)]
pub struct DataValidationSpec {
    pub validation_type: String,
    pub operator: Option<String>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
    pub allow_blank: bool,
    pub show_input_message: bool,
    pub prompt_title: Option<String>,
    pub prompt: Option<String>,
    pub show_error_message: bool,
    pub error_style: Option<String>,
    pub error_title: Option<String>,
    pub error: Option<String>,
}

/// Splits and parses an `ST_Sqref` value ("A1:B2 D5") into its individual 1-based
/// inclusive ranges, tolerant of any unparseable token (skipped, not fatal) -- matches
/// this reader's existing "contributes nothing" tolerance elsewhere (e.g. an
/// unresolvable table part). Uses `elixcee_types::parse_range_addr`'s own single-cell-or-
/// range convention, not `parse_merge_ref`'s colon-required one -- a real `sqref` token
/// commonly has no colon at all.
fn parse_sqref(s: &str) -> Vec<MergeRect> {
    s.split_whitespace()
        .filter_map(elixcee_types::parse_range_addr)
        .collect()
}

fn parse_data_validation_xml(span: &str) -> Option<DataValidationRule> {
    let mut iter = XmlIter::new(span);
    let mut rule: Option<DataValidationRule> = None;
    let mut in_formula1 = false;
    let mut in_formula2 = false;
    let mut f1_text = String::new();
    let mut f2_text = String::new();

    let as_bool = |attrs: &[Attr], name: &str| {
        attr_get(attrs, name)
            .map(|v| matches!(v, "1" | "true" | "TRUE"))
            .unwrap_or(false)
    };

    while let Some(ev) = iter.next_ev() {
        match ev {
            Ev::Open(ref tag, ref attrs) | Ev::SelfClose(ref tag, ref attrs) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "dataValidation" => {
                        rule = Some(DataValidationRule {
                            validation_type: attr_get(attrs, "type").unwrap_or("none").to_string(),
                            operator: attr_get(attrs, "operator").map(|s| s.to_string()),
                            formula1: None,
                            formula2: None,
                            allow_blank: as_bool(attrs, "allowBlank"),
                            show_input_message: as_bool(attrs, "showInputMessage"),
                            prompt_title: attr_get(attrs, "promptTitle").map(|s| s.to_string()),
                            prompt: attr_get(attrs, "prompt").map(|s| s.to_string()),
                            show_error_message: as_bool(attrs, "showErrorMessage"),
                            error_style: attr_get(attrs, "errorStyle").map(|s| s.to_string()),
                            error_title: attr_get(attrs, "errorTitle").map(|s| s.to_string()),
                            error: attr_get(attrs, "error").map(|s| s.to_string()),
                            sqref: attr_get(attrs, "sqref")
                                .map(parse_sqref)
                                .unwrap_or_default(),
                            dirty: false,
                            raw_span: span.to_string(),
                        });
                    }
                    "formula1" if !matches!(ev, Ev::SelfClose(_, _)) => {
                        in_formula1 = true;
                        f1_text.clear();
                    }
                    "formula2" if !matches!(ev, Ev::SelfClose(_, _)) => {
                        in_formula2 = true;
                        f2_text.clear();
                    }
                    _ => {}
                }
            }
            Ev::Close(ref tag) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "formula1" if in_formula1 => {
                        if let Some(r) = rule.as_mut() {
                            r.formula1 = Some(f1_text.clone());
                        }
                        in_formula1 = false;
                    }
                    "formula2" if in_formula2 => {
                        if let Some(r) = rule.as_mut() {
                            r.formula2 = Some(f2_text.clone());
                        }
                        in_formula2 = false;
                    }
                    _ => {}
                }
            }
            Ev::Text(ref text) => {
                if in_formula1 {
                    f1_text.push_str(text);
                } else if in_formula2 {
                    f2_text.push_str(text);
                }
            }
        }
    }
    rule
}

/// Parses every `<dataValidation>` record inside `sheet_xml`'s `<dataValidations>`
/// container (0.16.0-C), tolerant of any single record that fails to parse (skipped, not
/// fatal). Empty when the sheet has no `<dataValidations>` at all.
pub(crate) fn xlsx_data_validations(sheet_xml: &str) -> Vec<DataValidationRule> {
    extract_records(sheet_xml, "dataValidations", "dataValidation")
        .iter()
        .filter_map(|span| parse_data_validation_xml(span))
        .collect()
}

/// One `<tableColumn>` entry inside a `<table>`'s `<tableColumns>` (0.16.0-A1).
#[derive(Clone, Debug, PartialEq)]
pub struct TableColumn {
    pub id: Option<String>,
    pub name: String,
    pub totals_row_function: Option<String>,
    pub totals_row_label: Option<String>,
    /// Raw `<calculatedColumnFormula>` text, unparsed and unevaluated -- structured
    /// references are out of scope for 0.16.0-A (see `TableDef`'s own doc comment).
    pub calculated_column_formula: Option<String>,
}

/// A parsed `xl/tables/tableN.xml` part (0.16.0-A1), read-only. `<fonts>`/`<fills>`/
/// `<borders>`-style deep interpretation is never needed here -- only the table's own
/// structural metadata is modeled.
#[derive(Clone, Debug, PartialEq)]
pub struct TableDef {
    pub name: String,
    /// Distinct from `name` in real Excel (the identifier structured references and
    /// the UI use, unique workbook-wide) -- defaults to `name` when the `displayName`
    /// attribute happens to be absent (real Excel always writes both, but nothing
    /// technically requires it).
    pub display_name: String,
    /// 1-based inclusive `(top-left, bottom-right)`, from `ref="A1:C4"`.
    pub ref_range: MergeRect,
    /// Defaults to 1 per ECMA-376 §18.5.1.2 when the `headerRowCount` attribute is
    /// absent (no fixture in this repo exercises a non-default value).
    pub header_row_count: u32,
    pub totals_row_count: u32,
    /// Defaults to `true` per the `CT_Table` XSD when the `totalsRowShown` attribute
    /// is absent (real Excel/`fixture3` always writes it explicitly either way).
    pub totals_row_shown: bool,
    pub columns: Vec<TableColumn>,
    /// `<tableStyleInfo name="...">`'s name, when present.
    pub style_name: Option<String>,
    /// The table's own nested `<autoFilter ref="...">`, when present -- structural
    /// only.
    pub auto_filter_ref: Option<MergeRect>,
    /// `auto_filter_ref`'s own `<filterColumn>` children (0.16.0-B2), parsed the same
    /// way as a standalone `AutoFilterDef.columns` -- reuses `FilterColumn`/
    /// `FilterCriteria` and the same `col_offset` convention (0-based, relative to
    /// `auto_filter_ref`'s own left edge, not `ref_range`'s). Empty when
    /// `auto_filter_ref` is `None` or has no active criteria yet.
    pub autofilter_columns: Vec<FilterColumn>,
    /// Normalized `xl/tables/tableN.xml` part path this table was parsed from
    /// (0.16.0-A2) -- lets the writer find the right `raw_entries` key to surgically
    /// patch. Empty for a table built programmatically (never happens today -- no
    /// table-creation API exists yet, 0.16.0-A3).
    pub source_part: String,
    /// Edits requested via `edit_table` since load, applied against `source_part`'s
    /// RAW bytes at save time (`apply_table_edits`) rather than reserializing this
    /// struct -- `TableDef`/`TableColumn` are lossy read projections (no `id`,
    /// `xr:uid`/`xr3:uid` extension GUIDs, or original attribute order), so a full
    /// reserialize would silently drop them. See
    /// `internal_docs/tables-0.16.0-a-design.md`'s A2 Addendum.
    pub(crate) pending_edits: Vec<TableEditOp>,
}

/// One requested change to a table, recorded at `edit_table` call time and applied,
/// in order, against the table's original raw XML bytes at save time (0.16.0-A2).
#[derive(Clone, Debug, PartialEq)]
pub(crate) enum TableEditOp {
    SetDisplayName(String),
    Resize(MergeRect),
    /// The nested `<autoFilter ref="...">`'s own `ref`, kept in sync with `Resize`
    /// whenever a structural edit shifts both together (`shift_tables_for_structural_edit`)
    /// -- closes a real gap 0.16.0-A1 left behind: the in-memory shift was already
    /// correct, but never reached the saved file. See the A2 Addendum.
    ResizeAutoFilter(MergeRect),
    SetStyle(Option<String>),
    SetTotalsRowShown(bool),
    AddColumn(String),
    RemoveColumn(String),
    /// Sets (replacing any existing entry for the same `col_offset`) one
    /// `<filterColumn>` on the table's own nested `<autoFilter>` (0.16.0-B2).
    SetFilterColumn(u32, FilterCriteria),
    /// Removes `col_offset`'s `<filterColumn>` entry, if any.
    ClearFilterColumn(u32),
}

fn col_letters(mut col: u32) -> String {
    let mut s = String::new();
    while col > 0 {
        let rem = ((col - 1) % 26) as u8;
        s.push((b'A' + rem) as char);
        col = (col - 1) / 26;
    }
    s.chars().rev().collect()
}

fn format_merge_ref(rect: &MergeRect) -> String {
    let ((r1, c1), (r2, c2)) = *rect;
    format!("{}{}:{}{}", col_letters(c1), r1, col_letters(c2), r2)
}

fn span_attr_str(span: &str, attr_name: &str) -> Option<String> {
    let (tag_start, tag_close_rel, full_name) = find_next_open_tag(span, 0)?;
    let name_end = tag_start + 1 + full_name.len();
    let raw_attrs = &span[name_end..name_end + tag_close_rel];
    let attrs_str = raw_attrs.trim_end().strip_suffix('/').unwrap_or(raw_attrs);
    attr_get(&parse_attrs(attrs_str), attr_name).map(|s| s.to_string())
}

/// Order of `<table>`'s own top-level children per ECMA-376 `CT_Table`'s sequence,
/// as confirmed against real bytes (`autoFilter` -> `tableColumns` -> `tableStyleInfo`;
/// `sortState` -- unseen in any fixture -- placed per spec sequence, not independently
/// verified). Shared by every `with_ordered_child` call in `apply_table_edits`.
const TABLE_CHILD_ORDER: &[&str] = &["autoFilter", "sortState", "tableColumns", "tableStyleInfo"];

/// Applies `edits` in order against `table_xml` (the table's ORIGINAL raw bytes),
/// returning the patched document. Surgical, not a reserialize: every op touches only
/// the specific attribute/child it changes via `with_attr`/`with_ordered_child`,
/// leaving `id`, `xr:uid`/`xr3:uid`, attribute order, and every untouched column's
/// raw span byte-identical. See `TableDef::pending_edits`.
pub(crate) fn apply_table_edits(table_xml: &str, edits: &[TableEditOp]) -> String {
    let mut xml = table_xml.to_string();
    for edit in edits {
        xml = match edit {
            TableEditOp::SetDisplayName(name) => with_attr(&xml, "displayName", name),
            TableEditOp::Resize(rect) => with_attr(&xml, "ref", &format_merge_ref(rect)),
            TableEditOp::ResizeAutoFilter(rect) => match extract_raw_element(&xml, "autoFilter") {
                Some(old) => {
                    let new_child = with_attr(&old, "ref", &format_merge_ref(rect));
                    with_ordered_child(&xml, "autoFilter", TABLE_CHILD_ORDER, Some(&new_child))
                }
                None => xml,
            },
            TableEditOp::SetStyle(Some(name)) => {
                let new_child = match extract_raw_element(&xml, "tableStyleInfo") {
                    Some(old) => with_attr(&old, "name", name),
                    None => format!("<tableStyleInfo name=\"{}\"/>", crate::xml_escape(name)),
                };
                with_ordered_child(&xml, "tableStyleInfo", TABLE_CHILD_ORDER, Some(&new_child))
            }
            TableEditOp::SetStyle(None) => {
                with_ordered_child(&xml, "tableStyleInfo", TABLE_CHILD_ORDER, None)
            }
            TableEditOp::SetTotalsRowShown(shown) => {
                with_attr(&xml, "totalsRowShown", if *shown { "1" } else { "0" })
            }
            TableEditOp::AddColumn(name) => {
                let mut spans = extract_records(&xml, "tableColumns", "tableColumn");
                let next_id = spans
                    .iter()
                    .filter_map(|s| span_attr_str(s, "id").and_then(|v| v.parse::<u32>().ok()))
                    .max()
                    .unwrap_or(0)
                    + 1;
                spans.push(format!(
                    "<tableColumn id=\"{next_id}\" name=\"{}\"/>",
                    crate::xml_escape(name)
                ));
                let new_child = format!(
                    "<tableColumns count=\"{}\">{}</tableColumns>",
                    spans.len(),
                    spans.concat()
                );
                with_ordered_child(&xml, "tableColumns", TABLE_CHILD_ORDER, Some(&new_child))
            }
            TableEditOp::RemoveColumn(name) => {
                let spans: Vec<String> = extract_records(&xml, "tableColumns", "tableColumn")
                    .into_iter()
                    .filter(|s| span_attr_str(s, "name").as_deref() != Some(name.as_str()))
                    .collect();
                let new_child = format!(
                    "<tableColumns count=\"{}\">{}</tableColumns>",
                    spans.len(),
                    spans.concat()
                );
                with_ordered_child(&xml, "tableColumns", TABLE_CHILD_ORDER, Some(&new_child))
            }
            TableEditOp::SetFilterColumn(col_offset, criteria) => {
                match extract_raw_element(&xml, "autoFilter") {
                    Some(af) => {
                        let ref_str = span_attr_str(&af, "ref").unwrap_or_default();
                        let mut cols: Vec<String> =
                            extract_records(&af, "autoFilter", "filterColumn")
                                .into_iter()
                                .filter(|c| {
                                    span_attr_str(c, "colId").and_then(|v| v.parse::<u32>().ok())
                                        != Some(*col_offset)
                                })
                                .collect();
                        cols.push(crate::build_filter_column_xml(&FilterColumn {
                            col_offset: *col_offset,
                            hidden_button: false,
                            show_button: true,
                            criteria: criteria.clone(),
                            raw_span: None,
                            dirty: false,
                        }));
                        let new_af = crate::rebuild_autofilter_container(
                            Some(&af),
                            &ref_str,
                            &cols.concat(),
                        );
                        with_ordered_child(&xml, "autoFilter", TABLE_CHILD_ORDER, Some(&new_af))
                    }
                    None => xml,
                }
            }
            TableEditOp::ClearFilterColumn(col_offset) => {
                match extract_raw_element(&xml, "autoFilter") {
                    Some(af) => {
                        let ref_str = span_attr_str(&af, "ref").unwrap_or_default();
                        let cols: Vec<String> = extract_records(&af, "autoFilter", "filterColumn")
                            .into_iter()
                            .filter(|c| {
                                span_attr_str(c, "colId").and_then(|v| v.parse::<u32>().ok())
                                    != Some(*col_offset)
                            })
                            .collect();
                        let new_af = crate::rebuild_autofilter_container(
                            Some(&af),
                            &ref_str,
                            &cols.concat(),
                        );
                        with_ordered_child(&xml, "autoFilter", TABLE_CHILD_ORDER, Some(&new_af))
                    }
                    None => xml,
                }
            }
        };
    }
    xml
}

/// Every `Id="rIdN"` value declared by a `.rels` document's `<Relationship>` elements, as
/// raw strings (0.16.0-A3) -- used to pick a fresh, non-colliding id when inserting one
/// more relationship into an existing worksheet `.rels` file. Unlike `workbook_rels_decls`
/// (which deliberately drops ids, since that caller always reassigns fresh ones for every
/// entry it carries over), a NEW relationship inserted alongside existing ones must avoid
/// colliding with whichever ids the source already used.
pub(crate) fn relationship_ids(xml: &str) -> Vec<String> {
    let mut ids = vec![];
    let mut iter = XmlIter::new(xml);
    while let Some(ev) = iter.next_ev() {
        if let Ev::Open(ref tag, ref attrs) | Ev::SelfClose(ref tag, ref attrs) = ev {
            let local = tag.split(':').next_back().unwrap_or(tag);
            if local == "Relationship"
                && let Some(id) = attr_get(attrs, "Id")
            {
                ids.push(id.to_string());
            }
        }
    }
    ids
}

/// Serializes a freshly-created `TableDef` (0.16.0-A3) into a full `<table>...</table>`
/// document from scratch. Safe only because a NEW table has no existing raw bytes to
/// preserve -- unlike `apply_table_edits`, which surgically patches an EXISTING table's
/// original bytes specifically to avoid dropping the `id`/`xr:uid`/`xr3:uid` extension
/// GUIDs `TableDef` doesn't store. Omits those GUIDs entirely -- confirmed safely
/// omittable by inspecting a real `openpyxl`-authored table directly, which ships to real
/// users without them. Includes an `<autoFilter ref="...">` matching the table's own
/// `ref` (also confirmed against the same real sample: openpyxl always includes one, even
/// with no filter criteria), embedding any `autofilter_columns` already set on this
/// `TableDef` before its first save (0.16.0-B2) -- always built fresh via
/// `build_filter_column_xml`, never a `raw_span` to preserve, since the table itself
/// never existed on disk before now.
pub(crate) fn render_table_xml(table: &TableDef, table_id: u32) -> String {
    let table_ref = format_merge_ref(&table.ref_range);
    let mut out = format!(
        concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
            "id=\"{}\" name=\"{}\" displayName=\"{}\" ref=\"{}\" headerRowCount=\"1\" ",
            "totalsRowShown=\"0\">",
        ),
        table_id,
        crate::xml_escape(&table.name),
        crate::xml_escape(&table.display_name),
        table_ref,
    );
    // 0.16.0-B2: a freshly created table can already have filter criteria set on it
    // (via set_table_*_filter, before the first save) -- embed them the same way
    // apply_table_edits' SetFilterColumn op does for an existing table, via
    // build_filter_column_xml, since a brand-new table's own filter columns are
    // ALWAYS fresh (never a raw_span to preserve).
    if table.autofilter_columns.is_empty() {
        out.push_str(&format!("<autoFilter ref=\"{table_ref}\"/>"));
    } else {
        let body: String = table
            .autofilter_columns
            .iter()
            .map(crate::build_filter_column_xml)
            .collect();
        out.push_str(&format!(
            "<autoFilter ref=\"{table_ref}\">{body}</autoFilter>"
        ));
    }
    out.push_str(&format!("<tableColumns count=\"{}\">", table.columns.len()));
    for (i, col) in table.columns.iter().enumerate() {
        out.push_str(&format!(
            "<tableColumn id=\"{}\" name=\"{}\"/>",
            i + 1,
            crate::xml_escape(&col.name)
        ));
    }
    out.push_str("</tableColumns>");
    if let Some(style) = &table.style_name {
        out.push_str(&format!(
            "<tableStyleInfo name=\"{}\" showFirstColumn=\"0\" showLastColumn=\"0\" \
             showRowStripes=\"1\" showColumnStripes=\"0\"/>",
            crate::xml_escape(style)
        ));
    }
    out.push_str("</table>\n");
    out
}

pub enum SheetCell {
    Integer(i64),
    Float(f64),
    Str(String),
    Bool(bool),
    Error(ExcelError),
}

/// Read a spreadsheet file into sheets. Supports .xlsx, .xlsm, .ods.
///
/// The extension is validated before opening the path. This keeps the path-based
/// API's format boundary explicit and avoids treating an XLSX payload as an
/// arbitrary input format. Extension matching is case-insensitive.
pub fn read_workbook(path: &str) -> Result<Vec<WorkbookSheet>, String> {
    let extension = std::path::Path::new(path)
        .extension()
        .and_then(|value| value.to_str());
    if extension.is_some_and(|value| value.eq_ignore_ascii_case("ods")) {
        read_ods(path)
    } else if extension.is_some_and(|value| {
        value.eq_ignore_ascii_case("xlsx") || value.eq_ignore_ascii_case("xlsm")
    }) {
        read_xlsx(path)
    } else {
        Err("unsupported input extension; use .xlsx, .xlsm, or .ods".to_string())
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

pub(crate) fn xml_unescape(s: &str) -> String {
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

/// Per-entry decompressed cap. Large worksheets need more room than the previous 64 MB
/// ceiling, but an explicit limit keeps malformed XML from turning one read into an
/// unbounded allocation.
const ZIP_ENTRY_MAX_BYTES: u64 = 256 * 1024 * 1024;
const ZIP_MAX_ENTRIES: usize = 10_000;
const ZIP_MAX_TOTAL_BYTES: u64 = 1024 * 1024 * 1024;
const ZIP_MAX_COMPRESSION_RATIO: u64 = 1_000;
const XML_MAX_ELEMENTS: usize = 1_000_000;
const XML_MAX_ATTRIBUTES: usize = 2_000_000;
const XML_MAX_ATTRIBUTE_VALUE_BYTES: usize = 16 * 1024 * 1024;
const XML_MAX_TEXT_NODE_BYTES: usize = 64 * 1024 * 1024;
const XML_MAX_DEPTH: usize = 1_024;
const WORKBOOK_MAX_SHEETS: usize = 4_096;
const SHEET_MAX_CELLS: usize = 5_000_000;
const SHEET_MAX_MERGES: usize = 1_000_000;
const SHARED_STRINGS_MAX_COUNT: usize = 1_000_000;
const SHARED_STRINGS_MAX_TOTAL_BYTES: usize = 256 * 1024 * 1024;
const DEFINED_NAMES_MAX_COUNT: usize = 100_000;

fn validate_workbook_model_count(sheet_count: usize) -> Result<(), String> {
    if sheet_count > WORKBOOK_MAX_SHEETS {
        return Err(format!(
            "workbook has too many sheets ({}; maximum is {})",
            sheet_count, WORKBOOK_MAX_SHEETS
        ));
    }
    Ok(())
}

fn validate_shared_strings(strings: &[String]) -> Result<(), String> {
    if strings.len() > SHARED_STRINGS_MAX_COUNT {
        return Err(format!(
            "shared strings table is too large ({}; maximum is {})",
            strings.len(),
            SHARED_STRINGS_MAX_COUNT
        ));
    }
    let total_bytes = strings
        .iter()
        .try_fold(0usize, |total, value| total.checked_add(value.len()));
    let Some(total_bytes) = total_bytes else {
        return Err("shared strings size overflows usize".to_string());
    };
    if total_bytes > SHARED_STRINGS_MAX_TOTAL_BYTES {
        return Err(format!(
            "shared strings table is too large ({} bytes; maximum is {} bytes)",
            total_bytes, SHARED_STRINGS_MAX_TOTAL_BYTES
        ));
    }
    Ok(())
}

fn validate_sheet_model(
    sheet_name: &str,
    cell_count: usize,
    merged_range_count: usize,
) -> Result<(), String> {
    if cell_count > SHEET_MAX_CELLS {
        return Err(format!(
            "sheet has too many cells: {sheet_name} ({}; maximum is {})",
            cell_count, SHEET_MAX_CELLS
        ));
    }
    if merged_range_count > SHEET_MAX_MERGES {
        return Err(format!(
            "sheet has too many merged ranges: {sheet_name} ({}; maximum is {})",
            merged_range_count, SHEET_MAX_MERGES
        ));
    }
    Ok(())
}

fn validate_zip_entry_metadata(
    name: &str,
    uncompressed: u64,
    compressed: u64,
    total_before: u64,
) -> Result<u64, String> {
    if name.starts_with('/')
        || name.starts_with('\\')
        || name.split('/').any(|part| part == "..")
        || name.contains('\0')
    {
        return Err(format!("ZIP entry has an unsafe path: {name}"));
    }
    if uncompressed > ZIP_ENTRY_MAX_BYTES {
        return Err(format!(
            "ZIP entry is too large: {name} ({} bytes; maximum is {})",
            uncompressed, ZIP_ENTRY_MAX_BYTES
        ));
    }
    let total_uncompressed = total_before
        .checked_add(uncompressed)
        .ok_or_else(|| "ZIP archive uncompressed size overflows u64".to_string())?;
    if total_uncompressed > ZIP_MAX_TOTAL_BYTES {
        return Err(format!(
            "ZIP archive expands beyond the maximum size ({} bytes; maximum is {})",
            total_uncompressed, ZIP_MAX_TOTAL_BYTES
        ));
    }
    if compressed > 0 && uncompressed / compressed > ZIP_MAX_COMPRESSION_RATIO {
        return Err(format!(
            "ZIP entry has an excessive compression ratio: {name} ({}:1; maximum is {}:1)",
            uncompressed / compressed,
            ZIP_MAX_COMPRESSION_RATIO
        ));
    }
    Ok(total_uncompressed)
}

fn validate_zip_archive<R: Read + Seek>(archive: &mut ZipArchive<R>) -> Result<(), String> {
    if archive.len() > ZIP_MAX_ENTRIES {
        return Err(format!(
            "ZIP archive has too many entries ({}; maximum is {})",
            archive.len(),
            ZIP_MAX_ENTRIES
        ));
    }
    let mut total_uncompressed = 0u64;
    for i in 0..archive.len() {
        let entry = archive.by_index(i).map_err(|e| e.to_string())?;
        total_uncompressed = validate_zip_entry_metadata(
            entry.name(),
            entry.size(),
            entry.compressed_size(),
            total_uncompressed,
        )?;
    }
    Ok(())
}

fn contains_ascii_case_insensitive(haystack: &[u8], needle: &[u8]) -> bool {
    haystack.windows(needle.len()).any(|window| {
        window
            .iter()
            .zip(needle)
            .all(|(actual, expected)| actual.to_ascii_lowercase() == *expected)
    })
}

fn validate_xml_budget(name: &str, xml: &str) -> Result<(), String> {
    let raw = xml.as_bytes();
    if contains_ascii_case_insensitive(raw, b"<!doctype")
        || contains_ascii_case_insensitive(raw, b"<!entity")
    {
        return Err(format!(
            "XML document uses a forbidden DTD or entity declaration: {name}"
        ));
    }

    let mut elements = 0usize;
    let mut attributes = 0usize;
    let mut depth = 0usize;
    let mut iter = XmlIter::new(xml);
    while let Some(event) = iter.next_ev() {
        match event {
            Ev::Open(_, attrs) => {
                elements = elements
                    .checked_add(1)
                    .ok_or_else(|| format!("XML document element count overflows: {name}"))?;
                if elements > XML_MAX_ELEMENTS {
                    return Err(format!(
                        "XML document has too many elements: {name} (maximum is {})",
                        XML_MAX_ELEMENTS
                    ));
                }
                depth += 1;
                if depth > XML_MAX_DEPTH {
                    return Err(format!(
                        "XML document is nested too deeply: {name} (maximum is {})",
                        XML_MAX_DEPTH
                    ));
                }
                for attr in attrs {
                    attributes = attributes
                        .checked_add(1)
                        .ok_or_else(|| format!("XML document attribute count overflows: {name}"))?;
                    if attributes > XML_MAX_ATTRIBUTES {
                        return Err(format!(
                            "XML document has too many attributes: {name} (maximum is {})",
                            XML_MAX_ATTRIBUTES
                        ));
                    }
                    if attr.value.len() > XML_MAX_ATTRIBUTE_VALUE_BYTES {
                        return Err(format!(
                            "XML attribute value is too long: {name} (maximum is {} bytes)",
                            XML_MAX_ATTRIBUTE_VALUE_BYTES
                        ));
                    }
                }
            }
            Ev::SelfClose(_, attrs) => {
                elements = elements
                    .checked_add(1)
                    .ok_or_else(|| format!("XML document element count overflows: {name}"))?;
                if elements > XML_MAX_ELEMENTS {
                    return Err(format!(
                        "XML document has too many elements: {name} (maximum is {})",
                        XML_MAX_ELEMENTS
                    ));
                }
                for attr in attrs {
                    attributes = attributes
                        .checked_add(1)
                        .ok_or_else(|| format!("XML document attribute count overflows: {name}"))?;
                    if attributes > XML_MAX_ATTRIBUTES {
                        return Err(format!(
                            "XML document has too many attributes: {name} (maximum is {})",
                            XML_MAX_ATTRIBUTES
                        ));
                    }
                    if attr.value.len() > XML_MAX_ATTRIBUTE_VALUE_BYTES {
                        return Err(format!(
                            "XML attribute value is too long: {name} (maximum is {} bytes)",
                            XML_MAX_ATTRIBUTE_VALUE_BYTES
                        ));
                    }
                }
            }
            Ev::Close(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| format!("XML document has an unmatched closing tag: {name}"))?;
            }
            Ev::Text(text) => {
                if text.len() > XML_MAX_TEXT_NODE_BYTES {
                    return Err(format!(
                        "XML text node is too long: {name} (maximum is {} bytes)",
                        XML_MAX_TEXT_NODE_BYTES
                    ));
                }
            }
        }
    }
    if depth != 0 {
        return Err(format!("XML document has unclosed elements: {name}"));
    }
    Ok(())
}

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
    validate_xml_budget(name, &s)?;
    Ok(s)
}

#[cfg(feature = "python")]
pub(crate) fn validate_zip_archive_for_stream<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
) -> Result<(), String> {
    validate_zip_archive(archive)
}
#[cfg(feature = "python")]
pub(crate) fn zip_read_text_for_stream<R: Read + Seek>(
    archive: &mut ZipArchive<R>,
    name: &str,
) -> Result<String, String> {
    zip_read_text(archive, name)
}
#[cfg(feature = "python")]
pub(crate) fn xlsx_workbook_sheets_for_stream(
    xml: &str,
) -> Vec<(String, String, Option<String>, Option<String>)> {
    xlsx_workbook_sheets(xml)
}
#[cfg(feature = "python")]
pub(crate) fn xlsx_rels_for_stream(xml: &str, suffix: &str) -> HashMap<String, String> {
    xlsx_rels(xml, suffix)
}
#[cfg(feature = "python")]
pub(crate) fn xlsx_sheet_cells_for_stream(xml: &str, shared: &[String]) -> XlsxSheetData {
    xlsx_sheet_cells(xml, shared, &[])
}
#[cfg(feature = "python")]
pub(crate) fn xlsx_shared_strings_for_stream(xml: &str) -> Vec<String> {
    xlsx_shared_strings(xml)
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
    validate_zip_archive(&mut archive)?;
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

/// The OOXML relationships namespace URI -- every hardcoded `r:id="..."` / `r:id=` this
/// writer emits assumes the `r:` prefix is bound to exactly this URI.
pub(crate) const OFFICE_REL_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Guarantees `attrs` (a captured root tag's raw attribute string, from
/// `extract_root_attrs`) binds the `r:` prefix to `OFFICE_REL_NS` before it's reused
/// verbatim as a regenerated root tag's own attributes.
///
/// XML namespace binding is about the URI, not the prefix spelling -- a real, fully valid
/// OOXML producer is free to write `xmlns:rel="<the same URI>"` and `<sheet rel:id="..">`
/// instead of the conventional `xmlns:r=".."`/`r:id=".."`. Carrying such a source's root
/// attrs through unchanged while this writer's own worksheet-emission code still
/// hardcodes the literal `r:` prefix (`build_xlsx_workbook`'s `<sheet r:id="..">`, and
/// `build_xlsx_sheet`'s future 0.10.0-D r:id-bearing elements) produces `r:id` with `r:`
/// bound to nothing at all -- an unbound-namespace-prefix XML error every strict consumer
/// (openpyxl/lxml, Excel itself) rejects outright. Found via a real report against the
/// released `0.10.0`, reproduced with a synthetic fixture built by renaming a genuine
/// openpyxl-authored file's `xmlns:r`/`r:id` to `xmlns:rel`/`rel:id` (still valid OOXML on
/// its own) and round-tripping it through elixcee.
///
/// Returns the (possibly appended-to) attrs string when `r:` can be made to resolve
/// correctly -- unchanged if already correct, with `xmlns:r="<OFFICE_REL_NS>"` appended
/// if the prefix was simply never declared (the common real-world case: no source ever
/// uses anything but `r:`, so this is almost always a no-op). Returns `None` only when
/// `xmlns:r` is already bound to some OTHER, different URI -- rebinding it in place would
/// require rewriting every other place in `attrs` that might rely on the original
/// binding, which isn't worth the risk for a shape no real producer has ever been seen to
/// generate; the caller falls back to the writer's own safe hardcoded root tag instead of
/// risking a subtly wrong rebind.
pub(crate) fn ensure_r_prefix_bound(attrs: &str) -> Option<String> {
    match parse_attrs(attrs).into_iter().find(|a| a.name == "xmlns:r") {
        Some(a) if a.value == OFFICE_REL_NS => Some(attrs.to_string()),
        Some(_) => None,
        None => Some(format!("{attrs} xmlns:r=\"{OFFICE_REL_NS}\"")),
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

/// Extracts every raw, byte-for-byte `<local_name ..>...</local_name>`/`<local_name ../>`
/// top-level element in `xml` matched by local name (namespace prefix ignored), in document
/// order — unlike `extract_raw_element` above (first occurrence only), this is for elements
/// `CT_Worksheet` allows to repeat (`maxOccurs="unbounded"`), e.g. `conditionalFormatting`:
/// a worksheet commonly has one `<conditionalFormatting sqref="...">` block per distinct
/// range/rule-set, unlike `sheetPr`/`sheetViews`/`dataValidations`/etc., which never repeat.
/// Same non-self-nesting assumption as `extract_raw_element` (true for `conditionalFormatting`
/// — its children are `cfRule`/`extLst`, never another `conditionalFormatting`). Stops the
/// scan (rather than misparsing) if an opening tag's matching close tag is missing.
pub(crate) fn extract_all_raw_elements(xml: &str, local_name: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut search_from = 0;
    while let Some((tag_start, tag_close_rel, full_name)) = find_next_open_tag(xml, search_from) {
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
            out.push(xml[tag_start..start_tag_end].to_string());
            search_from = start_tag_end;
            continue;
        }
        let close_tag = format!("</{}>", full_name);
        let Some(end_rel) = xml[start_tag_end..].find(&close_tag) else {
            break;
        };
        let end = start_tag_end + end_rel + close_tag.len();
        out.push(xml[tag_start..end].to_string());
        search_from = end;
    }
    out
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

/// Extracts each `<definedName ...>value</definedName>` child inside `xml`'s
/// `<definedNames>...</definedNames>` container, as `(whole_element, (text_start, text_end))`
/// pairs -- `text_start`/`text_end` are byte offsets INTO `whole_element` bounding just the
/// inner value, so a caller rewriting that value can splice it back in via
/// `format!("{}{}{}", &el[..text_start], new_value, &el[text_end..])` without touching the
/// element's own attributes (`name`, `localSheetId`, `hidden`, ...). Empty if `<definedNames>`
/// is absent or has no `<definedName>` children. Unlike `extract_hyperlinks`'s `<hyperlink>`
/// (always self-closing per the real `CT_Hyperlink` XSD), `<definedName>`'s content is a
/// required `ST_Formula` string (`CT_DefinedName`, never `minOccurs="0"`) -- a self-closing
/// `<definedName/>` is invalid but handled defensively anyway, with an empty text span at the
/// end of its own opening tag, matching this reader's general "don't crash on malformed input"
/// posture; a `<definedName>` missing its closing tag stops the scan rather than misparsing
/// whatever XML happens to follow.
pub(crate) fn extract_defined_name_elements(xml: &str) -> Vec<(String, (usize, usize))> {
    let Some(container) = extract_raw_element(xml, "definedNames") else {
        return Vec::new();
    };
    let mut out = Vec::new();
    let mut search_from = 0;
    while let Some((tag_start, tag_close_rel, full_name)) =
        find_next_open_tag(&container, search_from)
    {
        if full_name.rsplit(':').next().unwrap_or(&full_name) != "definedName" {
            search_from = tag_start + 1;
            continue;
        }
        let name_end = tag_start + 1 + full_name.len();
        let start_tag_end = name_end + tag_close_rel + 1;
        let self_closing = container[name_end..name_end + tag_close_rel]
            .trim_end()
            .ends_with('/');
        if self_closing {
            let el = container[tag_start..start_tag_end].to_string();
            let end = el.len();
            out.push((el, (end, end)));
            search_from = start_tag_end;
            continue;
        }
        let close_tag = format!("</{}>", full_name);
        let Some(end_rel) = container[start_tag_end..].find(&close_tag) else {
            break;
        };
        let text_end = start_tag_end + end_rel;
        let end = text_end + close_tag.len();
        out.push((
            container[tag_start..end].to_string(),
            (start_tag_end - tag_start, text_end - tag_start),
        ));
        search_from = end;
    }
    out
}

/// True iff `element_xml`'s own root tag (as returned by `extract_raw_element` -- the
/// element's opening `<` at byte 0) carries an `r:id` attribute. Used to gate restoring a
/// single element that COULD be relationship-backed but usually isn't (`<pageSetup>`,
/// unlike e.g. `<pageMargins>`, genuinely has an optional `r:id` per the real
/// `CT_PageSetup` XSD, referencing a `printerSettings` part): a plain `<pageSetup>` with
/// no `r:id` is always safe to restore verbatim (no relationship dependency at all); one
/// WITH `r:id` needs the same `rels_survived` gate 0.10.0-D's other relationship-backed
/// elements use, which isn't wired up here yet -- no real fixture has ever shown that
/// shape (see `fixtures/pristine/INVENTORY.md`'s "confirmed absent" list). `attr_get(&attrs,
/// "id")`, same precise (not string-match) technique as `extract_hyperlinks` above.
pub(crate) fn root_tag_has_rid(element_xml: &str) -> bool {
    let Some((tag_start, tag_close_rel, full_name)) = find_next_open_tag(element_xml, 0) else {
        return false;
    };
    let name_end = tag_start + 1 + full_name.len();
    let attrs = parse_attrs(&element_xml[name_end..name_end + tag_close_rel]);
    attr_get(&attrs, "id").is_some()
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
    fn ensure_r_prefix_bound_leaves_a_correct_binding_untouched() {
        let attrs = concat!(
            "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
        );
        assert_eq!(ensure_r_prefix_bound(attrs), Some(attrs.to_string()));
    }

    #[test]
    fn ensure_r_prefix_bound_appends_the_binding_when_absent() {
        // Real report against the released 0.10.0: a source that binds the
        // relationships namespace to a different prefix (`rel:`, `rel:id=".."`) is
        // fully valid OOXML on its own -- XML namespace binding is about the URI, not
        // the prefix spelling. This writer's own <sheet r:id="..."> always hardcodes
        // the literal `r:` prefix, so it must be added, not assumed present.
        let attrs = concat!(
            "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
            "xmlns:rel=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
        );
        let result = ensure_r_prefix_bound(attrs).unwrap();
        assert!(
            result.starts_with(attrs),
            "must not disturb the original attrs: {result}"
        );
        assert!(
            result.contains(
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
            ),
            "must append the r: binding: {result}"
        );
    }

    #[test]
    fn ensure_r_prefix_bound_appends_when_the_relationships_namespace_is_absent_entirely() {
        let attrs = "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
        let result = ensure_r_prefix_bound(attrs).unwrap();
        assert!(
            result.contains(
                "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
            ),
            "must append the r: binding even when no relationships namespace was declared \
             under any prefix: {result}"
        );
    }

    #[test]
    fn ensure_r_prefix_bound_refuses_to_reuse_attrs_when_r_is_bound_to_a_different_uri() {
        let attrs = concat!(
            "xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" ",
            "xmlns:r=\"http://example.com/totally-unrelated\"",
        );
        assert_eq!(
            ensure_r_prefix_bound(attrs),
            None,
            "must not silently rebind an r: prefix a source is already using for \
             something else -- caller falls back to the writer's own safe default instead"
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
    fn extract_all_raw_elements_returns_every_occurrence_in_document_order() {
        // fixture3's real shape has exactly one, but CT_Worksheet allows more than one --
        // a worksheet with two distinct conditional-formatting range/rule-sets.
        let xml = concat!(
            "<worksheet><sheetData/>",
            "<conditionalFormatting sqref=\"A1:A5\">",
            "<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"greaterThan\">",
            "<formula>10</formula></cfRule></conditionalFormatting>",
            "<conditionalFormatting sqref=\"B1:B5\">",
            "<cfRule type=\"cellIs\" dxfId=\"1\" priority=\"2\" operator=\"lessThan\">",
            "<formula>0</formula></cfRule></conditionalFormatting>",
            "</worksheet>",
        );
        let all = extract_all_raw_elements(xml, "conditionalFormatting");
        assert_eq!(all.len(), 2);
        assert!(all[0].starts_with("<conditionalFormatting sqref=\"A1:A5\">"));
        assert!(all[0].ends_with("</conditionalFormatting>"));
        assert!(all[1].starts_with("<conditionalFormatting sqref=\"B1:B5\">"));
    }

    #[test]
    fn extract_all_raw_elements_returns_empty_when_absent() {
        let xml = "<worksheet><sheetData/></worksheet>";
        assert_eq!(
            extract_all_raw_elements(xml, "conditionalFormatting"),
            Vec::<String>::new()
        );
    }

    #[test]
    fn extract_all_raw_elements_handles_a_single_self_closing_occurrence() {
        let xml = "<worksheet><conditionalFormatting sqref=\"A1\"/></worksheet>";
        assert_eq!(
            extract_all_raw_elements(xml, "conditionalFormatting"),
            vec!["<conditionalFormatting sqref=\"A1\"/>".to_string()]
        );
    }

    #[test]
    fn extract_all_raw_elements_does_not_match_a_differently_named_element() {
        let xml = "<worksheet><conditionalFormattingRule/></worksheet>";
        assert_eq!(
            extract_all_raw_elements(xml, "conditionalFormatting"),
            Vec::<String>::new()
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

    #[test]
    fn extract_defined_name_elements_returns_element_and_inner_text_span() {
        let xml = concat!(
            "<workbook><definedNames>",
            "<definedName name=\"MyRange\">Sheet1!$A$1:$A$3</definedName>",
            "<definedName name=\"Other\" localSheetId=\"0\">Sheet1!$B$1</definedName>",
            "</definedNames></workbook>",
        );
        let elements = extract_defined_name_elements(xml);
        assert_eq!(elements.len(), 2);
        let (el0, (s0, e0)) = &elements[0];
        assert_eq!(
            el0,
            "<definedName name=\"MyRange\">Sheet1!$A$1:$A$3</definedName>"
        );
        assert_eq!(&el0[*s0..*e0], "Sheet1!$A$1:$A$3");
        let (el1, (s1, e1)) = &elements[1];
        assert_eq!(
            el1,
            "<definedName name=\"Other\" localSheetId=\"0\">Sheet1!$B$1</definedName>"
        );
        assert_eq!(&el1[*s1..*e1], "Sheet1!$B$1");
    }

    #[test]
    fn extract_defined_name_elements_empty_when_container_absent() {
        let xml = "<workbook><sheets/></workbook>";
        assert!(extract_defined_name_elements(xml).is_empty());
    }

    #[test]
    fn extract_defined_name_elements_empty_when_no_children() {
        let xml = "<workbook><definedNames/></workbook>";
        assert!(extract_defined_name_elements(xml).is_empty());
    }

    #[test]
    fn extract_cell_xfs_returns_self_closing_and_child_bearing_spans_verbatim() {
        let xml = concat!(
            "<styleSheet><cellXfs count=\"3\">",
            "<xf/>",
            "<xf numFmtId=\"4\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"1\"/>",
            "<xf numFmtId=\"0\" fontId=\"1\"><alignment horizontal=\"center\"/></xf>",
            "</cellXfs></styleSheet>",
        );
        let xfs = extract_cell_xfs(xml);
        assert_eq!(
            xfs,
            vec![
                "<xf/>".to_string(),
                "<xf numFmtId=\"4\" fontId=\"0\" fillId=\"0\" borderId=\"0\" applyNumberFormat=\"1\"/>"
                    .to_string(),
                "<xf numFmtId=\"0\" fontId=\"1\"><alignment horizontal=\"center\"/></xf>".to_string(),
            ]
        );
    }

    #[test]
    fn extract_cell_xfs_empty_when_container_absent() {
        assert!(extract_cell_xfs("<styleSheet><fonts/></styleSheet>").is_empty());
    }

    #[test]
    fn with_num_fmt_id_inserts_into_a_bare_self_closing_xf() {
        assert_eq!(with_num_fmt_id("<xf/>", 4), "<xf numFmtId=\"4\"/>");
    }

    #[test]
    fn with_num_fmt_id_replaces_an_existing_numfmtid_preserving_other_attrs() {
        let xf = "<xf numFmtId=\"9\" fontId=\"2\" fillId=\"1\" applyNumberFormat=\"1\"/>";
        let out = with_num_fmt_id(xf, 4);
        // Order of the untouched attributes doesn't matter for correctness -- parse back
        // out and compare as a set, matching how this function's own dedup caller treats
        // equality (byte-identical strings from the same construction path).
        let attrs = parse_attrs(&out[3..out.len() - 2]);
        assert_eq!(attr_get(&attrs, "numFmtId"), Some("4"));
        assert_eq!(attr_get(&attrs, "fontId"), Some("2"));
        assert_eq!(attr_get(&attrs, "fillId"), Some("1"));
        assert_eq!(attr_get(&attrs, "applyNumberFormat"), Some("1"));
    }

    #[test]
    fn with_num_fmt_id_preserves_child_elements_on_a_non_self_closing_xf() {
        let xf = "<xf numFmtId=\"0\" fontId=\"1\"><alignment horizontal=\"center\"/></xf>";
        let out = with_num_fmt_id(xf, 14);
        assert!(out.ends_with("<alignment horizontal=\"center\"/></xf>"));
        assert!(out.contains("numFmtId=\"14\""));
        assert!(!out.contains("numFmtId=\"0\""));
    }

    #[test]
    fn with_num_fmt_id_is_idempotent_for_dedup_when_called_twice_on_the_same_input() {
        // The find-or-append caller relies on two independently constructed candidates
        // from the SAME source xf + same target id being byte-identical.
        let xf = "<xf fontId=\"3\" borderId=\"2\"/>";
        assert_eq!(with_num_fmt_id(xf, 7), with_num_fmt_id(xf, 7));
    }

    #[test]
    fn resolve_number_format_id_reuses_a_builtin() {
        let custom = HashMap::new();
        match resolve_number_format_id("#,##0.00", &custom) {
            ResolvedNumFmt::Existing(id) => assert_eq!(id, 4),
            ResolvedNumFmt::New(_) => panic!("expected an existing builtin id"),
        }
    }

    #[test]
    fn resolve_number_format_id_reuses_an_existing_custom_entry() {
        let mut custom = HashMap::new();
        custom.insert(164, "0.00\"kg\"".to_string());
        match resolve_number_format_id("0.00\"kg\"", &custom) {
            ResolvedNumFmt::Existing(id) => assert_eq!(id, 164),
            ResolvedNumFmt::New(_) => panic!("expected the existing custom id to be reused"),
        }
    }

    #[test]
    fn resolve_number_format_id_mints_164_when_no_custom_entries_exist() {
        let custom = HashMap::new();
        match resolve_number_format_id("0.00\"kg\"", &custom) {
            ResolvedNumFmt::New(id) => assert_eq!(id, 164),
            ResolvedNumFmt::Existing(_) => panic!("a genuinely custom format has no builtin match"),
        }
    }

    #[test]
    fn resolve_number_format_id_mints_one_past_the_highest_existing_custom_id() {
        let mut custom = HashMap::new();
        custom.insert(164, "0.00\"kg\"".to_string());
        custom.insert(170, "0.00\"lb\"".to_string());
        match resolve_number_format_id("[Red]0.00", &custom) {
            ResolvedNumFmt::New(id) => assert_eq!(id, 171),
            ResolvedNumFmt::Existing(_) => panic!("not a builtin or existing custom format"),
        }
    }

    // ── 0.15.0-B: extract_records / with_attr generalization ────────────────────

    #[test]
    fn extract_records_generalizes_to_fonts_and_fills() {
        let xml = concat!(
            "<styleSheet><fonts count=\"2\"><font/>",
            "<font><b val=\"1\"/><sz val=\"14\"/></font></fonts>",
            "<fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>",
            "</styleSheet>",
        );
        assert_eq!(
            extract_records(xml, "fonts", "font"),
            vec![
                "<font/>".to_string(),
                "<font><b val=\"1\"/><sz val=\"14\"/></font>".to_string(),
            ]
        );
        assert_eq!(
            extract_records(xml, "fills", "fill"),
            vec!["<fill><patternFill patternType=\"none\"/></fill>".to_string()]
        );
    }

    #[test]
    fn extract_records_matches_extract_cell_xfs_for_cellxfs() {
        let xml = "<styleSheet><cellXfs count=\"1\"><xf fontId=\"1\"/></cellXfs></styleSheet>";
        assert_eq!(extract_records(xml, "cellXfs", "xf"), extract_cell_xfs(xml));
    }

    #[test]
    fn with_attr_matches_with_num_fmt_id_for_numfmtid() {
        let xf = "<xf numFmtId=\"9\" fontId=\"2\"/>";
        assert_eq!(with_attr(xf, "numFmtId", "4"), with_num_fmt_id(xf, 4));
    }

    #[test]
    fn with_attr_sets_a_new_attribute_on_a_self_closing_span() {
        assert_eq!(
            with_attr("<xf fontId=\"0\"/>", "applyFont", "1"),
            "<xf fontId=\"0\" applyFont=\"1\"/>"
        );
    }

    #[test]
    fn with_attr_preserves_children_on_a_non_self_closing_span() {
        let out = with_attr(
            "<xf fillId=\"0\"><alignment vertical=\"center\"/></xf>",
            "fillId",
            "3",
        );
        assert!(out.contains("fillId=\"3\""));
        assert!(out.ends_with("<alignment vertical=\"center\"/></xf>"));
    }

    // ── 0.15.0-C1: named-style lookup ────────────────────────────────────────────

    #[test]
    fn named_style_xf_id_finds_a_real_fixture_shaped_entry() {
        let xml = concat!(
            "<styleSheet><cellStyles count=\"2\">",
            "<cellStyle name=\"ハイパーリンク\" xfId=\"1\" builtinId=\"8\"/>",
            "<cellStyle name=\"標準\" xfId=\"0\" builtinId=\"0\"/>",
            "</cellStyles></styleSheet>",
        );
        assert_eq!(named_style_xf_id(xml, "ハイパーリンク"), Some(1));
        assert_eq!(named_style_xf_id(xml, "標準"), Some(0));
    }

    #[test]
    fn named_style_xf_id_none_for_an_unknown_name() {
        let xml = "<styleSheet><cellStyles count=\"1\"><cellStyle name=\"標準\" xfId=\"0\"/></cellStyles></styleSheet>";
        assert_eq!(named_style_xf_id(xml, "Bad"), None);
    }

    #[test]
    fn named_style_xf_id_none_with_no_cellstyles_element() {
        assert_eq!(
            named_style_xf_id("<styleSheet></styleSheet>", "Normal"),
            None
        );
    }

    #[test]
    fn span_attr_u32_reads_an_existing_attribute() {
        assert_eq!(
            span_attr_u32("<xf fontId=\"7\" borderId=\"2\"/>", "fontId"),
            7
        );
    }

    #[test]
    fn span_attr_u32_defaults_to_zero_when_absent() {
        assert_eq!(span_attr_u32("<xf/>", "fontId"), 0);
    }

    // ── 0.15.0-B: with_child / with_ordered_child ────────────────────────────────

    #[test]
    fn with_child_inserts_into_a_bare_self_closing_parent() {
        assert_eq!(
            with_child("<font/>", "b", Some("<b val=\"1\"/>")),
            "<font><b val=\"1\"/></font>"
        );
    }

    #[test]
    fn with_child_appends_to_an_existing_non_self_closing_parent() {
        assert_eq!(
            with_child("<font><sz val=\"11\"/></font>", "b", Some("<b val=\"1\"/>")),
            "<font><sz val=\"11\"/><b val=\"1\"/></font>"
        );
    }

    #[test]
    fn with_child_replaces_an_existing_child_in_place() {
        assert_eq!(
            with_child(
                "<font><b val=\"1\"/><sz val=\"11\"/></font>",
                "sz",
                Some("<sz val=\"14\"/>")
            ),
            "<font><b val=\"1\"/><sz val=\"14\"/></font>"
        );
    }

    #[test]
    fn with_child_removes_an_existing_child_when_given_none() {
        assert_eq!(
            with_child("<font><b val=\"1\"/><sz val=\"11\"/></font>", "b", None),
            "<font><sz val=\"11\"/></font>"
        );
    }

    #[test]
    fn with_child_remove_is_a_noop_when_child_absent() {
        assert_eq!(
            with_child("<font><sz val=\"11\"/></font>", "b", None),
            "<font><sz val=\"11\"/></font>"
        );
    }

    #[test]
    fn with_ordered_child_inserts_at_correct_position_among_present_siblings() {
        // "top" (order index 2) inserted where only "left" (index 0) exists must land
        // AFTER left, matching real CT_Border sequence order, not appended blindly.
        let out = with_ordered_child(
            "<border><left style=\"thin\"/></border>",
            "top",
            &BORDER_SIDE_ORDER,
            Some("<top style=\"thick\"/>"),
        );
        assert_eq!(
            out,
            "<border><left style=\"thin\"/><top style=\"thick\"/></border>"
        );
    }

    #[test]
    fn with_ordered_child_inserts_before_a_later_sibling_when_earlier_is_added() {
        // "left" (index 0) inserted where only "top" (index 2) exists must land BEFORE
        // top, not after it.
        let out = with_ordered_child(
            "<border><top style=\"thick\"/></border>",
            "left",
            &BORDER_SIDE_ORDER,
            Some("<left style=\"thin\"/>"),
        );
        assert_eq!(
            out,
            "<border><left style=\"thin\"/><top style=\"thick\"/></border>"
        );
    }

    #[test]
    fn with_ordered_child_replaces_in_place_keeping_position() {
        let out = with_ordered_child(
            "<xf><alignment vertical=\"center\"/><protection locked=\"1\"/></xf>",
            "alignment",
            &XF_CHILD_ORDER,
            Some("<alignment horizontal=\"center\"/>"),
        );
        assert_eq!(
            out,
            "<xf><alignment horizontal=\"center\"/><protection locked=\"1\"/></xf>"
        );
    }

    // ── 0.15.0-B: font/border/alignment/protection merge primitives ─────────────

    #[test]
    fn with_font_edit_only_touches_requested_properties() {
        // fixture4/6's real, in-use hyperlink font shape: underlined, theme-colored, sized.
        let font = "<font><u/><sz val=\"12\"/><color theme=\"10\"/><name val=\"Calibri\"/></font>";
        let edit = FontEdit {
            bold: Some(true),
            ..Default::default()
        };
        let out = with_font_edit(font, &edit);
        assert!(out.contains("<b val=\"1\"/>"));
        assert!(out.contains("<u/>"));
        assert!(out.contains("<color theme=\"10\"/>"));
        assert!(out.contains("<sz val=\"12\"/>"));
        assert!(out.contains("<name val=\"Calibri\"/>"));
    }

    #[test]
    fn with_font_edit_replaces_an_existing_property() {
        let font = "<font><sz val=\"11\"/></font>";
        let edit = FontEdit {
            size: Some(14.0),
            ..Default::default()
        };
        let out = with_font_edit(font, &edit);
        assert!(out.contains("<sz val=\"14\"/>"));
        assert!(!out.contains("val=\"11\""));
    }

    #[test]
    fn with_font_edit_bold_false_writes_explicit_val_zero_not_a_removal() {
        let out = with_font_edit(
            "<font/>",
            &FontEdit {
                bold: Some(false),
                ..Default::default()
            },
        );
        assert_eq!(out, "<font><b val=\"0\"/></font>");
    }

    #[test]
    fn build_solid_fill_uses_fgcolor_and_indexed_64_bgcolor_sentinel() {
        assert_eq!(
            build_solid_fill("FF4472C4"),
            "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FF4472C4\"/><bgColor indexed=\"64\"/></patternFill></fill>"
        );
    }

    #[test]
    fn with_border_edit_touches_only_the_requested_side() {
        let border = "<border><left style=\"thin\"/><right/><top/><bottom/><diagonal/></border>";
        let edit = BorderEdit {
            top: Some(BorderSideEdit {
                style: Some("thick".to_string()),
                color_argb: Some("FF000000".to_string()),
            }),
            ..Default::default()
        };
        let out = with_border_edit(border, &edit);
        assert!(out.contains("<left style=\"thin\"/>"));
        assert!(out.contains("<top style=\"thick\"><color rgb=\"FF000000\"/></top>"));
        assert!(out.contains("<right/>"));
        assert!(out.contains("<bottom/>"));
        assert!(out.contains("<diagonal/>"));
    }

    #[test]
    fn merged_alignment_span_preserves_existing_vertical_when_setting_horizontal() {
        // Every real fixture's <xf> already carries vertical="center" -- setting
        // horizontal must not silently drop it.
        let xf = "<xf><alignment vertical=\"center\"/></xf>";
        let edit = AlignmentEdit {
            horizontal: Some("center".to_string()),
            ..Default::default()
        };
        let out = merged_alignment_span(xf, &edit);
        assert!(out.contains("vertical=\"center\""));
        assert!(out.contains("horizontal=\"center\""));
    }

    #[test]
    fn merged_alignment_span_creates_a_fresh_alignment_when_none_exists() {
        let out = merged_alignment_span(
            "<xf/>",
            &AlignmentEdit {
                wrap_text: Some(true),
                ..Default::default()
            },
        );
        assert_eq!(out, "<alignment wrapText=\"1\"/>");
    }

    #[test]
    fn merged_protection_span_merges_onto_existing() {
        let xf = "<xf><protection locked=\"1\"/></xf>";
        let out = merged_protection_span(
            xf,
            &ProtectionEdit {
                hidden: Some(true),
                ..Default::default()
            },
        );
        assert!(out.contains("locked=\"1\""));
        assert!(out.contains("hidden=\"1\""));
    }

    #[test]
    fn root_tag_has_rid_false_for_a_relationship_free_page_setup() {
        // fixture5_chart_image_freeze_print.xlsm's real shape.
        let xml = r#"<pageSetup paperSize="9" orientation="portrait" horizontalDpi="0" verticalDpi="0"/>"#;
        assert!(!root_tag_has_rid(xml));
    }

    #[test]
    fn root_tag_has_rid_true_for_an_rid_bearing_page_setup() {
        let xml = r#"<pageSetup paperSize="9" r:id="rId1"/>"#;
        assert!(root_tag_has_rid(xml));
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
    let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;
    validate_zip_archive(&mut archive)?;
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
    validate_zip_archive(&mut archive)?;
    let wb_xml = zip_read_text(&mut archive, "xl/workbook.xml")?;
    let sheet_refs = xlsx_workbook_sheets(&wb_xml);
    validate_workbook_model_count(sheet_refs.len())?;
    let date1904 = xlsx_workbook_date1904(&wb_xml);

    let rels_xml = zip_read_text(&mut archive, "xl/_rels/workbook.xml.rels")?;
    let rels = xlsx_rels(&rels_xml, "/worksheet");

    let shared: Vec<String> = match zip_read_text(&mut archive, "xl/sharedStrings.xml") {
        Ok(xml) => {
            let strings = xlsx_shared_strings(&xml);
            validate_shared_strings(&strings)?;
            strings
        }
        Err(_) => vec![],
    };

    let styles = match zip_read_text(&mut archive, "xl/styles.xml") {
        Ok(xml) => xlsx_styles(&xml),
        Err(_) => XlsxStyles::default(),
    };

    let mut sheets = vec![];
    for (name, rid, sheet_id, sheet_state) in sheet_refs {
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
        validate_sheet_model(
            &name,
            sheet_data.cells.len(),
            sheet_data.merged_ranges.len(),
        )?;
        // GitHub #4: resolved from the same style_ids BufferSheet::style_ids already
        // carries -- see WorkbookSheet::cell_number_formats' doc comment.
        let cell_number_formats: HashMap<(u32, u32), String> = sheet_data
            .style_ids
            .iter()
            .filter_map(|(&pos, &fmt_id)| {
                resolve_number_format(fmt_id, &styles.number_formats).map(|code| (pos, code))
            })
            .collect();
        // Tables (0.16.0-A1): resolved via the sheet's OWN `.rels` file, not the
        // workbook-level `rels` map above -- a `<tablePart r:id="...">` is scoped to
        // this one sheet, same OPC convention as hyperlinks/drawings. Tolerant of a
        // missing `.rels`/unresolvable target/unparseable table part -- each just
        // contributes nothing, matching this reader's convention elsewhere.
        let mut tables = Vec::new();
        let table_rids = xlsx_table_part_rids(&sheet_xml);
        if !table_rids.is_empty()
            && let Ok(sheet_rels_xml) =
                zip_read_text(&mut archive, &crate::part_rels_name(&zip_path))
        {
            let table_rels = xlsx_rels(&sheet_rels_xml, "/table");
            let base = crate::rels_target_dir(&crate::part_rels_name(&zip_path)).to_string();
            for rid in &table_rids {
                let Some(target) = table_rels.get(rid) else {
                    continue;
                };
                let resolved = crate::normalize_part_path(&format!("{base}{target}"));
                if let Ok(table_xml) = zip_read_text(&mut archive, &resolved)
                    && let Some(mut t) = parse_table_xml(&table_xml)
                {
                    t.source_part = resolved;
                    tables.push(t);
                }
            }
        }
        let data_validations = xlsx_data_validations(&sheet_xml);
        let autofilter = xlsx_autofilter(&sheet_xml);
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
                cell_number_formats,
                sheet_state,
                row_heights: sheet_data.row_heights,
                column_widths: sheet_data.column_widths,
                row_styles: sheet_data.row_styles,
                column_styles: sheet_data.column_styles,
                tables,
                data_validations,
                autofilter,
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

/// Returns `[(sheet_name, rId, sheetId, state)]` in document order. `state` is the
/// `<sheet state="...">` attribute's raw value (`None` when absent, the default
/// meaning visible) -- see `WorkbookSheet::sheet_state`'s doc comment for why this
/// stays a raw string here rather than a resolved type.
fn xlsx_workbook_sheets(xml: &str) -> Vec<(String, String, Option<String>, Option<String>)> {
    let mut iter = XmlIter::new(xml);
    let mut result = vec![];
    while let Some(ev) = iter.next_ev() {
        if let Ev::SelfClose(ref tag, ref attrs) = ev {
            let local = tag.split(':').next_back().unwrap_or(tag);
            if local == "sheet"
                && let (Some(name), Some(rid)) = (attr_get(attrs, "name"), attr_get(attrs, "id"))
            {
                let sheet_id = attr_get(attrs, "sheetId").map(|s| s.to_string());
                let state = attr_get(attrs, "state").map(|s| s.to_string());
                result.push((name.to_string(), rid.to_string(), sheet_id, state));
            }
        }
    }
    result
}

/// Returns `[(name, raw_text)]` in document order, from every
/// `<definedName name="...">TEXT</definedName>` inside `xl/workbook.xml`'s
/// `<definedNames>`. `raw_text` is the exact formula-text content (e.g.
/// `"Sheet1!$A$1:$A$3"`), unresolved -- see `Vm::defined_names`'s own doc
/// comment for why resolving it into a sheet+address isn't attempted.
/// `localSheetId`-scoped (sheet-local) and workbook-scoped names are not
/// distinguished here -- both are returned under their own `name` attribute
/// exactly as written; `Vm::defined_names` is what decides how to flatten
/// them into one map.
pub(crate) fn xlsx_defined_names(xml: &str) -> Result<Vec<(String, String)>, String> {
    let mut iter = XmlIter::new(xml);
    let mut result = vec![];
    let mut current_name: Option<String> = None;
    let mut current_text = String::new();
    while let Some(ev) = iter.next_ev() {
        match &ev {
            Ev::Open(tag, attrs) => {
                let local = tag.split(':').next_back().unwrap_or(tag);
                if local == "definedName" {
                    current_name = attr_get(attrs, "name").map(|s| s.to_string());
                    current_text.clear();
                }
            }
            Ev::Close(tag) => {
                let local = tag.split(':').next_back().unwrap_or(tag);
                if local == "definedName"
                    && let Some(name) = current_name.take()
                {
                    if result.len() >= DEFINED_NAMES_MAX_COUNT {
                        return Err(format!(
                            "defined-name table is too large (more than {}; maximum is {})",
                            DEFINED_NAMES_MAX_COUNT, DEFINED_NAMES_MAX_COUNT
                        ));
                    }
                    result.push((name, current_text.clone()));
                }
            }
            Ev::Text(text) => {
                if current_name.is_some() {
                    current_text.push_str(text);
                }
            }
            Ev::SelfClose(_, _) => {}
        }
    }
    Ok(result)
}

/// Returns `{rId → target_path}` for relationships whose `Type` ends with
/// `type_suffix` (e.g. `"/worksheet"`, `"/table"` — 0.16.0-A1 generalized this from a
/// worksheet-only filter once a second relationship type needed the identical parse).
fn xlsx_rels(xml: &str, type_suffix: &str) -> HashMap<String, String> {
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
                && ty.ends_with(type_suffix)
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

/// ECMA-376 Part 1 §18.8.30's built-in `numFmtId` -> format-code table (ids 0-49; 23-36
/// are reserved for legacy international formats the spec itself never assigns a code
/// to, and are omitted here the same way -- not a gap this reader introduced). A fixed,
/// published constant, unlike this project's usual "no writer code until a real fixture
/// shows the shape" rule for OOXML structural elements: there's nothing to discover here,
/// every real `.xlsx`/`.xlsm` producer uses these exact ids for these exact meanings.
const BUILTIN_NUMBER_FORMATS: &[(u32, &str)] = &[
    (0, "General"),
    (1, "0"),
    (2, "0.00"),
    (3, "#,##0"),
    (4, "#,##0.00"),
    (5, "$#,##0;($#,##0)"),
    (6, "$#,##0;[Red]($#,##0)"),
    (7, "$#,##0.00;($#,##0.00)"),
    (8, "$#,##0.00;[Red]($#,##0.00)"),
    (9, "0%"),
    (10, "0.00%"),
    (11, "0.00E+00"),
    (12, "# ?/?"),
    (13, "# ??/??"),
    (14, "m/d/yyyy"),
    (15, "d-mmm-yy"),
    (16, "d-mmm"),
    (17, "mmm-yy"),
    (18, "h:mm AM/PM"),
    (19, "h:mm:ss AM/PM"),
    (20, "h:mm"),
    (21, "h:mm:ss"),
    (22, "m/d/yyyy h:mm"),
    (37, "#,##0 ;(#,##0)"),
    (38, "#,##0 ;[Red](#,##0)"),
    (39, "#,##0.00;(#,##0.00)"),
    (40, "#,##0.00;[Red](#,##0.00)"),
    (41, "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)"),
    (42, "_($* #,##0_);_($* (#,##0);_($* \"-\"_);_(@_)"),
    (43, "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)"),
    (44, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"),
    (45, "mm:ss"),
    (46, "[h]:mm:ss"),
    (47, "mm:ss.0"),
    (48, "##0.0E+0"),
    (49, "@"),
];

/// Resolves a `numFmtId` to its format-code string -- a workbook-level custom
/// `<numFmt formatCode="...">` definition first (an id can only mean one thing in a
/// given file, but a custom entry is the file's own explicit statement of what an id
/// means, so it takes priority over the built-in table on the vanishingly rare chance
/// both exist for the same id), falling back to `BUILTIN_NUMBER_FORMATS`. `None` for
/// numFmtId 0 (General -- nothing to report) or any other id neither source defines.
fn resolve_number_format(num_fmt_id: u32, custom_formats: &HashMap<u32, String>) -> Option<String> {
    if let Some(code) = custom_formats.get(&num_fmt_id) {
        return Some(code.clone());
    }
    if num_fmt_id == 0 {
        return None;
    }
    BUILTIN_NUMBER_FORMATS
        .iter()
        .find(|(id, _)| *id == num_fmt_id)
        .map(|(_, code)| code.to_string())
}

/// `xl/styles.xml`'s custom `<numFmt numFmtId="N" formatCode="...">` definitions --
/// `xlsx_styles`'s own `number_formats` field, exposed narrowly (not the whole
/// read-only-shaped `XlsxStyles` struct) for `set_number_format`'s write-direction lookup
/// (`resolve_number_format_id` below).
pub(crate) fn custom_number_formats(xml: &str) -> HashMap<u32, String> {
    xlsx_styles(xml).number_formats
}

/// A target format-code string resolves to either an id some existing record (built-in or
/// this file's own custom `<numFmt>`) already means, or a brand-new custom id this file
/// has never used before.
pub(crate) enum ResolvedNumFmt {
    Existing(u32),
    New(u32),
}

/// Inverts `resolve_number_format`: given a target format-code string, finds the
/// `numFmtId` that already means it (checking the file's own custom definitions first,
/// then the built-in 0-49 table -- matching `resolve_number_format`'s own priority), or
/// mints the next free custom id (ECMA-376 §18.8.30: ids 0-163 are reserved/built-in:
/// starts at 164, or one past the highest custom id this file already defines).
pub(crate) fn resolve_number_format_id(
    format_code: &str,
    custom_formats: &HashMap<u32, String>,
) -> ResolvedNumFmt {
    if let Some((&id, _)) = custom_formats
        .iter()
        .find(|(_, code)| code.as_str() == format_code)
    {
        return ResolvedNumFmt::Existing(id);
    }
    if let Some((id, _)) = BUILTIN_NUMBER_FORMATS
        .iter()
        .find(|(_, code)| *code == format_code)
    {
        return ResolvedNumFmt::Existing(*id);
    }
    let next = custom_formats
        .keys()
        .copied()
        .max()
        .map_or(164, |m| m.max(163) + 1);
    ResolvedNumFmt::New(next)
}

/// Extracts the raw, byte-for-byte `<record .../>`/`<record>...</record>` spans directly
/// inside `xml`'s `<container>...</container>` element, in document order. Generalized
/// from 0.15.0-A's `set_number_format`-only `extract_cell_xfs` (kept below as a thin
/// wrapper for its own call sites/tests) once 0.15.0-B's font/fill/border needed the exact
/// same extraction shape for `<fonts>`/`<fills>`/`<borders>` -- real duplication removed,
/// not a speculative generalization. Handles both self-closing and child-bearing forms
/// (matching `extract_defined_name_elements`'s approach). Empty if `container` is absent
/// or has no `record` children.
pub(crate) fn extract_records(xml: &str, container: &str, record: &str) -> Vec<String> {
    let Some(container_span) = extract_raw_element(xml, container) else {
        return Vec::new();
    };
    let mut out = Vec::new();
    let mut search_from = 0;
    while let Some((tag_start, tag_close_rel, full_name)) =
        find_next_open_tag(&container_span, search_from)
    {
        if full_name.rsplit(':').next().unwrap_or(&full_name) != record {
            search_from = tag_start + 1;
            continue;
        }
        let name_end = tag_start + 1 + full_name.len();
        let start_tag_end = name_end + tag_close_rel + 1;
        let self_closing = container_span[name_end..name_end + tag_close_rel]
            .trim_end()
            .ends_with('/');
        if self_closing {
            out.push(container_span[tag_start..start_tag_end].to_string());
            search_from = start_tag_end;
            continue;
        }
        let close_tag = format!("</{}>", full_name);
        let Some(end_rel) = container_span[start_tag_end..].find(&close_tag) else {
            break;
        };
        let end = start_tag_end + end_rel + close_tag.len();
        out.push(container_span[tag_start..end].to_string());
        search_from = end;
    }
    out
}

/// `extract_records(xml, "cellXfs", "xf")` -- a cell's `s="N"` is a 0-based index into
/// this list. Named/kept separately from `extract_records` since every 0.15.0-A call site
/// already spells this name.
pub(crate) fn extract_cell_xfs(xml: &str) -> Vec<String> {
    extract_records(xml, "cellXfs", "xf")
}

/// Looks up `name`'s `xfId` from `xml`'s `<cellStyles>` element (0.15.0-C1 named-style
/// apply) -- `<cellStyle name="..." xfId="N" .../>` entries, matched by exact name.
/// `xfId` is a 0-based index into `<cellStyleXfs>`, a second style table parallel to (but
/// never confused with) `<cellXfs>`. `None` if `<cellStyles>` is absent or has no entry
/// with this name.
pub(crate) fn named_style_xf_id(xml: &str, name: &str) -> Option<u32> {
    extract_records(xml, "cellStyles", "cellStyle")
        .iter()
        .find_map(|span| {
            let (tag_start, tag_close_rel, full_name) = find_next_open_tag(span, 0)?;
            let name_end = tag_start + 1 + full_name.len();
            let raw_attrs = &span[name_end..name_end + tag_close_rel];
            let attrs_str = raw_attrs.trim_end().strip_suffix('/').unwrap_or(raw_attrs);
            let attrs = parse_attrs(attrs_str);
            if attr_get(&attrs, "name") != Some(name) {
                return None;
            }
            attr_get(&attrs, "xfId").and_then(|v| v.parse().ok())
        })
}

/// Clones `span` (any self-closing-or-not element span, e.g. from `extract_records`) with
/// `attr_name` set to `attr_value` -- added if absent, replaced if present -- and every
/// OTHER attribute and child element preserved verbatim. This is 0.15.0-A/B's shared core
/// safety primitive: it never needs to understand what any other attribute or child means,
/// only copy it into the new record unchanged. Generalized from 0.15.0-A's
/// `numFmtId`-only `with_num_fmt_id` (kept below as a thin wrapper) once 0.15.0-B needed
/// the identical shape for `fontId`/`fillId`/`borderId`/`applyFont`/`applyFill`/
/// `applyBorder`/`applyAlignment`/`applyProtection`/alignment's and protection's own
/// attributes.
pub(crate) fn with_attr(span: &str, attr_name: &str, attr_value: &str) -> String {
    let Some((tag_start, tag_close_rel, full_name)) = find_next_open_tag(span, 0) else {
        return span.to_string();
    };
    let name_end = tag_start + 1 + full_name.len();
    let tag_close_abs = name_end + tag_close_rel;
    let raw_attrs = &span[name_end..tag_close_abs];
    let self_closing = raw_attrs.trim_end().ends_with('/');
    let attrs_str = if self_closing {
        raw_attrs.trim_end().strip_suffix('/').unwrap_or(raw_attrs)
    } else {
        raw_attrs
    };
    let mut new_attrs = String::new();
    for a in parse_attrs(attrs_str) {
        if a.name.rsplit(':').next() == Some(attr_name) {
            continue;
        }
        new_attrs.push(' ');
        new_attrs.push_str(&a.name);
        new_attrs.push_str("=\"");
        new_attrs.push_str(&crate::xml_escape(&a.value));
        new_attrs.push('"');
    }
    new_attrs.push_str(&format!(
        " {attr_name}=\"{}\"",
        crate::xml_escape(attr_value)
    ));
    if self_closing {
        format!("<{full_name}{new_attrs}/>")
    } else {
        format!("<{full_name}{new_attrs}{}", &span[tag_close_abs..])
    }
}

/// Reads `attr_name` off `span`'s own opening tag as a `u32`, defaulting to 0 when absent
/// or unparseable -- matching this codebase's existing "index 0 is the implicit default"
/// convention for `fontId`/`fillId`/`borderId` (a bare `<xf/>` with no explicit pointer
/// means "the first/default record", same as a cell with no `s="N"` meaning style index 0).
pub(crate) fn span_attr_u32(span: &str, attr_name: &str) -> u32 {
    let Some((tag_start, tag_close_rel, full_name)) = find_next_open_tag(span, 0) else {
        return 0;
    };
    let name_end = tag_start + 1 + full_name.len();
    let raw_attrs = &span[name_end..name_end + tag_close_rel];
    let attrs_str = raw_attrs.trim_end().strip_suffix('/').unwrap_or(raw_attrs);
    attr_get(&parse_attrs(attrs_str), attr_name)
        .and_then(|v| v.parse().ok())
        .unwrap_or(0)
}

pub(crate) fn with_num_fmt_id(xf_span: &str, new_id: u32) -> String {
    with_attr(xf_span, "numFmtId", &new_id.to_string())
}

/// Clones `parent_span` with its `child_tag` child upserted: `Some(new_child)` replaces an
/// existing `child_tag` child (found via `extract_raw_element`, so `child_tag` must not
/// self-nest -- true for every 0.15.0-B target) or appends one at the very end if absent
/// (promoting a self-closing parent to an open/close pair when a child must be added where
/// none exists); `None` removes an existing `child_tag` child if present, a no-op
/// otherwise. Order among siblings is not schema-significant here -- CT_Font's
/// `EG_FontProperty` is an unordered choice group, confirmed empirically: a real
/// `openpyxl`-authored file emits two different fonts with two different child orders in
/// the same `styleSheet` and reopens fine either way. Use `with_ordered_child` instead for
/// a genuinely order-significant sequence (`<border>`'s sides, `<xf>`'s
/// alignment/protection).
pub(crate) fn with_child(parent_span: &str, child_tag: &str, new_child: Option<&str>) -> String {
    let existing = extract_raw_element(parent_span, child_tag);
    match (new_child, existing) {
        (Some(new_child), Some(old)) => parent_span.replacen(old.as_str(), new_child, 1),
        (Some(new_child), None) => insert_before_close(parent_span, new_child),
        (None, Some(old)) => parent_span.replacen(&old, "", 1),
        (None, None) => parent_span.to_string(),
    }
}

/// Same upsert/remove semantics as `with_child`, but a freshly-inserted child is placed at
/// its correct position among `order` (which must list every sibling this codebase ever
/// writes for this parent, in real schema sequence order) rather than always last --
/// `<border>`'s `left/right/top/bottom/diagonal` and `<xf>`'s `alignment`/`protection` are
/// both genuine ordered `xsd:sequence`s (unlike `<font>`'s children, see `with_child`).
/// Replacing an EXISTING child keeps its current position unchanged (already correct).
pub(crate) fn with_ordered_child(
    parent_span: &str,
    child_tag: &str,
    order: &[&str],
    new_child: Option<&str>,
) -> String {
    let existing = extract_raw_element(parent_span, child_tag);
    let Some(new_child) = new_child else {
        return match existing {
            Some(old) => parent_span.replacen(&old, "", 1),
            None => parent_span.to_string(),
        };
    };
    if let Some(old) = &existing {
        return parent_span.replacen(old.as_str(), new_child, 1);
    }
    if let Some(pos) = order.iter().position(|s| *s == child_tag) {
        for later in &order[pos + 1..] {
            if let Some(later_span) = extract_raw_element(parent_span, later) {
                return parent_span.replacen(
                    later_span.as_str(),
                    &format!("{new_child}{later_span}"),
                    1,
                );
            }
        }
    }
    insert_before_close(parent_span, new_child)
}

/// Inserts `new_child` as the last child of `parent_span`, promoting a self-closing
/// element (`<font/>`) to an open/close pair (`<font>{new_child}</font>`) when it has no
/// children yet, or splicing just before the existing closing tag otherwise. Shared tail
/// of `with_child`/`with_ordered_child`'s insert-when-absent case. Also reused directly by
/// `save_xlsx_impl` (0.16.0-A3) to insert a new `<Relationship>` into an existing worksheet
/// `.rels` document before its `</Relationships>` close -- `find_next_open_tag` already
/// skips the leading `<?xml ...?>` declaration, so passing a whole `.rels` document (not
/// just an inner element span) works unchanged.
pub(crate) fn insert_before_close(parent_span: &str, new_child: &str) -> String {
    let Some((tag_start, tag_close_rel, full_name)) = find_next_open_tag(parent_span, 0) else {
        return parent_span.to_string();
    };
    let name_end = tag_start + 1 + full_name.len();
    let tag_close_abs = name_end + tag_close_rel;
    let self_closing = parent_span[name_end..tag_close_abs]
        .trim_end()
        .ends_with('/');
    if self_closing {
        let attrs = parent_span[name_end..tag_close_abs].trim_end();
        let attrs = attrs.strip_suffix('/').unwrap_or(attrs);
        format!("<{full_name}{attrs}>{new_child}</{full_name}>")
    } else {
        let close_tag = format!("</{full_name}>");
        match parent_span.rfind(&close_tag) {
            Some(pos) => format!(
                "{}{}{}",
                &parent_span[..pos],
                new_child,
                &parent_span[pos..]
            ),
            None => parent_span.to_string(),
        }
    }
}

/// CT_Border's real, order-significant side sequence -- see `with_ordered_child`.
pub(crate) const BORDER_SIDE_ORDER: [&str; 5] = ["left", "right", "top", "bottom", "diagonal"];
/// CT_Xf's real, order-significant `(alignment?, protection?, extLst?)` sequence (this
/// codebase never writes `extLst`) -- see `with_ordered_child`.
pub(crate) const XF_CHILD_ORDER: [&str; 2] = ["alignment", "protection"];

/// 0.15.0-B `font={...}` request -- `None` on any field means "leave this property
/// exactly as the cell's current font already has it," matching `set_number_format`'s own
/// "only touch what was asked" contract. Booleans are written as an explicit
/// `val="1"`/`val="0"` on their own child element rather than relying on
/// presence-means-true/absence-means-false -- confirmed against a real `openpyxl`-authored
/// file that a genuine Excel-ecosystem writer already does this (`<b val="1"/>`, not a bare
/// `<b/>`), and `CT_BooleanProperty`'s own `val` attribute supports both explicit forms, so
/// this is not a deviation from real-world convention, just the simpler of two valid ones
/// (no separate "remove to mean false" logic needed).
#[derive(Debug, Clone, Default)]
pub struct FontEdit {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub underline: Option<bool>,
    pub strike: Option<bool>,
    pub size: Option<f64>,
    /// Already-normalized 8-hex-digit ARGB (`"FFRRGGBB"`) -- normalization from a
    /// caller-supplied 6-digit RGB happens at the Python-API boundary (`src/lib.rs`), not
    /// here; this module only ever emits exactly what it's given.
    pub color_argb: Option<String>,
    pub name: Option<String>,
}

impl FontEdit {
    /// Merges `other`'s explicitly-set (`Some`) fields onto `self`, leaving whatever
    /// `other` left `None` untouched -- used when a cell already has a pending
    /// `set_style(font=...)` edit and a later call, before the same save, sets more font
    /// properties, so the earlier ones aren't lost (`Vm::pending_style_attrs`'s own
    /// merge-not-overwrite contract).
    pub(crate) fn merge_from(&mut self, other: &FontEdit) {
        if other.bold.is_some() {
            self.bold = other.bold;
        }
        if other.italic.is_some() {
            self.italic = other.italic;
        }
        if other.underline.is_some() {
            self.underline = other.underline;
        }
        if other.strike.is_some() {
            self.strike = other.strike;
        }
        if other.size.is_some() {
            self.size = other.size;
        }
        if other.color_argb.is_some() {
            self.color_argb = other.color_argb.clone();
        }
        if other.name.is_some() {
            self.name = other.name.clone();
        }
    }
}

/// CT_Font's real child sequence, restricted to the properties `FontEdit` supports (order
/// among these is not schema-significant, see `with_child`, but a consistent order is
/// still used for freshly-appended children rather than an arbitrary one).
const FONT_CHILD_ORDER: [&str; 7] = ["name", "b", "i", "strike", "color", "sz", "u"];

/// Clones `font_span` (from `extract_records(xml, "fonts", "font")`) with `edit` applied --
/// every property `edit` leaves as `None` is preserved on the cloned record exactly as it
/// was (0.15.0's "preserve unknown style attributes" safety requirement: cloning
/// `fixture4`'s real in-use hyperlink font while only setting `bold` must keep its
/// existing `<u/>`/`<color theme="10"/>`/`<name>`/`<sz>` untouched).
pub(crate) fn with_font_edit(font_span: &str, edit: &FontEdit) -> String {
    let mut out = font_span.to_string();
    if let Some(b) = edit.bold {
        out = with_child(
            &out,
            "b",
            Some(&format!("<b val=\"{}\"/>", if b { 1 } else { 0 })),
        );
    }
    if let Some(i) = edit.italic {
        out = with_child(
            &out,
            "i",
            Some(&format!("<i val=\"{}\"/>", if i { 1 } else { 0 })),
        );
    }
    if let Some(u) = edit.underline {
        out = with_child(
            &out,
            "u",
            Some(if u {
                "<u val=\"single\"/>"
            } else {
                "<u val=\"none\"/>"
            }),
        );
    }
    if let Some(s) = edit.strike {
        out = with_child(
            &out,
            "strike",
            Some(&format!("<strike val=\"{}\"/>", if s { 1 } else { 0 })),
        );
    }
    if let Some(sz) = edit.size {
        out = with_child(&out, "sz", Some(&format!("<sz val=\"{sz}\"/>")));
    }
    if let Some(color) = &edit.color_argb {
        out = with_child(
            &out,
            "color",
            Some(&format!("<color rgb=\"{}\"/>", crate::xml_escape(color))),
        );
    }
    if let Some(name) = &edit.name {
        out = with_child(
            &out,
            "name",
            Some(&format!("<name val=\"{}\"/>", crate::xml_escape(name))),
        );
    }
    let _ = &FONT_CHILD_ORDER; // documents the intended order; with_child always appends last (order-insignificant here).
    out
}

/// Builds a brand-new `<fill>...</fill>` for a solid color -- 0.15.0-B deliberately
/// REPLACES the whole record rather than merging (unlike font/border), since "set this
/// cell's fill to solid color X" is a wholesale pattern-type change in real Excel's own
/// UI, not a patch onto whatever pattern/gradient existed before. `bgColor
/// indexed="64"` is the real, well-known Excel sentinel for "no second color" on a solid
/// fill -- `fgColor` is the one visible color; a caller-supplied single color naturally
/// maps to `fgColor` alone, matching how Excel's own single-color fill picker writes this
/// shape. `color_argb` is already-normalized 8-hex-digit ARGB, same convention as
/// `FontEdit::color_argb`.
pub(crate) fn build_solid_fill(color_argb: &str) -> String {
    format!(
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"{}\"/><bgColor indexed=\"64\"/></patternFill></fill>",
        crate::xml_escape(color_argb)
    )
}

/// One `<border>` side (`<left>`/`<right>`/`<top>`/`<bottom>`/`<diagonal>`) -- `None`
/// fields preserve the side exactly as it already was, same "only touch what was asked"
/// contract as `FontEdit`.
#[derive(Debug, Clone, Default)]
pub struct BorderSideEdit {
    pub style: Option<String>,
    pub color_argb: Option<String>,
}

/// 0.15.0-B `border={...}` request -- each side is independently optional; a side left
/// `None` is preserved on the cloned `<border>` record exactly as it was.
#[derive(Debug, Clone, Default)]
pub struct BorderEdit {
    pub left: Option<BorderSideEdit>,
    pub right: Option<BorderSideEdit>,
    pub top: Option<BorderSideEdit>,
    pub bottom: Option<BorderSideEdit>,
    pub diagonal: Option<BorderSideEdit>,
}

impl BorderEdit {
    /// Merges `other`'s explicitly-set sides onto `self` -- a whole side (style + color
    /// together) is the merge unit, matching `with_border_edit`'s own per-side
    /// granularity. See `FontEdit::merge_from` for why this exists.
    pub(crate) fn merge_from(&mut self, other: &BorderEdit) {
        if other.left.is_some() {
            self.left = other.left.clone();
        }
        if other.right.is_some() {
            self.right = other.right.clone();
        }
        if other.top.is_some() {
            self.top = other.top.clone();
        }
        if other.bottom.is_some() {
            self.bottom = other.bottom.clone();
        }
        if other.diagonal.is_some() {
            self.diagonal = other.diagonal.clone();
        }
    }
}

fn build_border_side_span(tag: &str, side: &BorderSideEdit) -> String {
    let style_attr = side
        .style
        .as_ref()
        .map(|s| format!(" style=\"{}\"", crate::xml_escape(s)))
        .unwrap_or_default();
    match &side.color_argb {
        Some(color) => format!(
            "<{tag}{style_attr}><color rgb=\"{}\"/></{tag}>",
            crate::xml_escape(color)
        ),
        None if style_attr.is_empty() => format!("<{tag}/>"),
        None => format!("<{tag}{style_attr}/>"),
    }
}

/// Clones `border_span` (from `extract_records(xml, "borders", "border")`) with `edit`
/// applied -- each requested side replaces (or inserts, at its correct schema position via
/// `with_ordered_child`/`BORDER_SIDE_ORDER`) that side's span; every other side, present or
/// absent, is preserved exactly as it was.
pub(crate) fn with_border_edit(border_span: &str, edit: &BorderEdit) -> String {
    let mut out = border_span.to_string();
    if let Some(side) = &edit.left {
        out = with_ordered_child(
            &out,
            "left",
            &BORDER_SIDE_ORDER,
            Some(&build_border_side_span("left", side)),
        );
    }
    if let Some(side) = &edit.right {
        out = with_ordered_child(
            &out,
            "right",
            &BORDER_SIDE_ORDER,
            Some(&build_border_side_span("right", side)),
        );
    }
    if let Some(side) = &edit.top {
        out = with_ordered_child(
            &out,
            "top",
            &BORDER_SIDE_ORDER,
            Some(&build_border_side_span("top", side)),
        );
    }
    if let Some(side) = &edit.bottom {
        out = with_ordered_child(
            &out,
            "bottom",
            &BORDER_SIDE_ORDER,
            Some(&build_border_side_span("bottom", side)),
        );
    }
    if let Some(side) = &edit.diagonal {
        out = with_ordered_child(
            &out,
            "diagonal",
            &BORDER_SIDE_ORDER,
            Some(&build_border_side_span("diagonal", side)),
        );
    }
    out
}

/// 0.15.0-B `alignment={...}` request -- `None` fields preserve the cell's existing
/// `<alignment>` attributes exactly (every real fixture in this project already carries
/// `vertical="center"` on every `<xf>`; replacing wholesale instead of merging would
/// silently drop it, a real "preserve unknown attributes" violation).
#[derive(Debug, Clone, Default)]
pub struct AlignmentEdit {
    pub horizontal: Option<String>,
    pub vertical: Option<String>,
    pub wrap_text: Option<bool>,
    pub indent: Option<u32>,
}

impl AlignmentEdit {
    /// See `FontEdit::merge_from`.
    pub(crate) fn merge_from(&mut self, other: &AlignmentEdit) {
        if other.horizontal.is_some() {
            self.horizontal = other.horizontal.clone();
        }
        if other.vertical.is_some() {
            self.vertical = other.vertical.clone();
        }
        if other.wrap_text.is_some() {
            self.wrap_text = other.wrap_text;
        }
        if other.indent.is_some() {
            self.indent = other.indent;
        }
    }
}

/// 0.15.0-B `protection={...}` request -- same merge-not-replace contract as
/// `AlignmentEdit`.
#[derive(Debug, Clone, Default)]
pub struct ProtectionEdit {
    pub locked: Option<bool>,
    pub hidden: Option<bool>,
}

impl ProtectionEdit {
    /// See `FontEdit::merge_from`.
    pub(crate) fn merge_from(&mut self, other: &ProtectionEdit) {
        if other.locked.is_some() {
            self.locked = other.locked;
        }
        if other.hidden.is_some() {
            self.hidden = other.hidden;
        }
    }
}

/// Merges `edit` onto `xf_span`'s EXISTING `<alignment>` child (or a fresh `<alignment/>`
/// if none exists yet) -- attribute-level merge via `with_attr`, not a wholesale replace.
/// Returns just the new `<alignment .../>` span; the caller still has to splice it back
/// into `xf_span` via `with_ordered_child`/`XF_CHILD_ORDER`.
pub(crate) fn merged_alignment_span(xf_span: &str, edit: &AlignmentEdit) -> String {
    let mut out =
        extract_raw_element(xf_span, "alignment").unwrap_or_else(|| "<alignment/>".to_string());
    if let Some(h) = &edit.horizontal {
        out = with_attr(&out, "horizontal", h);
    }
    if let Some(v) = &edit.vertical {
        out = with_attr(&out, "vertical", v);
    }
    if let Some(w) = edit.wrap_text {
        out = with_attr(&out, "wrapText", if w { "1" } else { "0" });
    }
    if let Some(indent) = edit.indent {
        out = with_attr(&out, "indent", &indent.to_string());
    }
    out
}

/// Same merge-onto-existing-child contract as `merged_alignment_span`, for `<protection>`.
pub(crate) fn merged_protection_span(xf_span: &str, edit: &ProtectionEdit) -> String {
    let mut out =
        extract_raw_element(xf_span, "protection").unwrap_or_else(|| "<protection/>".to_string());
    if let Some(l) = edit.locked {
        out = with_attr(&out, "locked", if l { "1" } else { "0" });
    }
    if let Some(h) = edit.hidden {
        out = with_attr(&out, "hidden", if h { "1" } else { "0" });
    }
    out
}

/// Parses a single worksheet XML into a 1-based (row, col) → SheetCell map,
/// plus any `<mergeCells><mergeCell ref="..."/></mergeCells>` ranges
/// (Milestone B6c2) and hidden row/column metadata (Milestone B7b).
/// A small return struct, not a growing bare tuple — B6c2 hit a
/// `clippy::type_complexity` error the first time this function's return
/// type grew, so this sidesteps a repeat of that churn.
pub(crate) struct XlsxSheetData {
    pub(crate) cells: HashMap<(u32, u32), SheetCell>,
    /// The first worksheet row encountered. This is kept separately from `cells`
    /// because a valid `<row r="N"/>` has no cell entries at all.
    #[cfg_attr(not(feature = "python"), allow(dead_code))]
    pub(crate) first_row: Option<u32>,
    merged_ranges: Vec<MergeRect>,
    /// Hidden row intervals, 1-based inclusive `(start, end)` — coalesced
    /// from consecutive `<row r=".." hidden="1">` tags (Milestone B7b).
    hidden_rows: Vec<(u32, u32)>,
    /// Hidden column intervals, 1-based inclusive `(start, end)` — read
    /// directly from `<col min=".." max=".." hidden="1">` (Milestone
    /// B7b), already interval-shaped in the XML, no coalescing needed.
    hidden_columns: Vec<(u32, u32)>,
    /// Per-row explicit height in points (P2), from `<row r=".." ht=".."
    /// customHeight="1">` — `customHeight="1"` is required for `ht` to
    /// actually apply in real Excel, so a bare `ht` without it is not
    /// recorded (matches real Excel's own behavior, confirmed via
    /// ECMA-376's `CT_Row` semantics). Sparse: only rows with an explicit
    /// height get an entry, matching `cell_number_formats`'s sparsity.
    row_heights: HashMap<u32, f64>,
    /// Column width ranges in "characters" (P2), 1-based inclusive
    /// `(min, max, width)` — from `<col min=".." max=".." width=".."
    /// customWidth="1"/>`, same `customWidth="1"`-required caveat as
    /// `row_heights`' `customHeight`. Already range-shaped in the XML, no
    /// coalescing needed, same as `hidden_columns`.
    column_widths: Vec<(u32, u32, f64)>,
    /// Per-row default style index (0.15.0-C2) — see `WorkbookSheet::row_styles`.
    row_styles: HashMap<u32, u32>,
    /// Column default-style ranges (0.15.0-C2) — see `WorkbookSheet::column_styles`.
    column_styles: Vec<(u32, u32, u32)>,
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
    let mut row_heights: HashMap<u32, f64> = HashMap::new();
    let mut column_widths: Vec<(u32, u32, f64)> = Vec::new();
    let mut row_styles: HashMap<u32, u32> = HashMap::new();
    let mut column_styles: Vec<(u32, u32, u32)> = Vec::new();
    let mut formulas: HashMap<(u32, u32), String> = HashMap::new();
    let mut dimension: Option<MergeRect> = None;
    let mut style_ids: HashMap<(u32, u32), u32> = HashMap::new();
    let mut raw_style_indices: HashMap<(u32, u32), u32> = HashMap::new();
    let mut first_row: Option<u32> = None;
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
                            if first_row.is_none() && cur_row != 0 {
                                first_row = Some(cur_row);
                            }
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
                        if attr_is_true(attrs, "customHeight")
                            && let Some(ht) = attr_get(attrs, "ht").and_then(|s| s.parse().ok())
                        {
                            row_heights.insert(cur_row, ht);
                        }
                        // `customFormat="1"` is required for `s` to mean "this row's own
                        // default style" (0.15.0-C2), same required-flag convention as
                        // `customHeight`/`ht` above.
                        if attr_is_true(attrs, "customFormat")
                            && let Some(s) = attr_get(attrs, "s").and_then(|s| s.parse().ok())
                        {
                            row_styles.insert(cur_row, s);
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
                        if attr_is_true(attrs, "customWidth") {
                            let min = attr_get(attrs, "min").and_then(|s| s.parse().ok());
                            let max = attr_get(attrs, "max").and_then(|s| s.parse().ok());
                            let width = attr_get(attrs, "width").and_then(|s| s.parse().ok());
                            if let (Some(min), Some(max), Some(width)) = (min, max, width) {
                                column_widths.push((min, max, width));
                            }
                        }
                        // `style="N"` (0.15.0-C2) needs no required-flag guard the way
                        // `hidden`/`customWidth` do -- its own presence IS the signal,
                        // matching real Excel's own `<col style>` convention.
                        if let Some(style) = attr_get(attrs, "style").and_then(|s| s.parse().ok()) {
                            let min = attr_get(attrs, "min").and_then(|s| s.parse().ok());
                            let max = attr_get(attrs, "max").and_then(|s| s.parse().ok());
                            if let (Some(min), Some(max)) = (min, max) {
                                column_styles.push((min, max, style));
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
        first_row,
        merged_ranges,
        hidden_rows,
        hidden_columns,
        row_heights,
        column_widths,
        row_styles,
        column_styles,
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
        "str" => Some(SheetCell::Str(v.to_string())),
        // One of the 7 classic error strings -> a real error-typed cell; anything else
        // (a newer dynamic-array error like #SPILL!, or malformed input) falls back to a
        // plain string rather than guessing -- see ExcelError::from_str's doc comment.
        "e" => Some(match ExcelError::from_str(v) {
            Ok(err) => SheetCell::Error(err),
            Err(()) => SheetCell::Str(v.to_string()),
        }),
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

/// Extracts each `<tablePart r:id="...">`'s relationship id from a worksheet's
/// `<tableParts>` element (0.16.0-A1), in document order. Empty if the sheet has no
/// `<tableParts>` at all.
fn xlsx_table_part_rids(sheet_xml: &str) -> Vec<String> {
    let Some(tp) = extract_raw_element(sheet_xml, "tableParts") else {
        return Vec::new();
    };
    let mut iter = XmlIter::new(&tp);
    let mut rids = Vec::new();
    while let Some(ev) = iter.next_ev() {
        if let Ev::SelfClose(ref tag, ref attrs) = ev
            && tag.split(':').next_back() == Some("tablePart")
            && let Some(rid) = attr_get(attrs, "id")
        {
            rids.push(rid.to_string());
        }
    }
    rids
}

/// Parses one `xl/tables/tableN.xml` part's real content into a `TableDef` (0.16.0-A1).
/// `None` if the document has no `<table ref="...">` with a parseable `ref` -- a
/// malformed table part is simply dropped rather than surfacing an error, matching this
/// reader's existing tolerant convention for every other optional/opaque fragment.
fn parse_table_xml(xml: &str) -> Option<TableDef> {
    let mut iter = XmlIter::new(xml);
    let mut table: Option<TableDef> = None;
    let mut cur_column: Option<TableColumn> = None;
    let mut in_calc_formula = false;
    let mut calc_formula_text = String::new();

    while let Some(ev) = iter.next_ev() {
        match ev {
            Ev::Open(ref tag, ref attrs) | Ev::SelfClose(ref tag, ref attrs) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "table" => {
                        let name = attr_get(attrs, "name").unwrap_or("").to_string();
                        let display_name =
                            attr_get(attrs, "displayName").unwrap_or(&name).to_string();
                        if let Some(ref_range) = attr_get(attrs, "ref").and_then(parse_merge_ref) {
                            table = Some(TableDef {
                                name,
                                display_name,
                                ref_range,
                                header_row_count: attr_get(attrs, "headerRowCount")
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(1),
                                totals_row_count: attr_get(attrs, "totalsRowCount")
                                    .and_then(|s| s.parse().ok())
                                    .unwrap_or(0),
                                totals_row_shown: attr_get(attrs, "totalsRowShown")
                                    .map(|v| matches!(v, "1" | "true" | "TRUE"))
                                    .unwrap_or(true),
                                columns: Vec::new(),
                                style_name: None,
                                auto_filter_ref: None,
                                autofilter_columns: Vec::new(),
                                source_part: String::new(),
                                pending_edits: Vec::new(),
                            });
                        }
                    }
                    "tableColumn" => {
                        cur_column = Some(TableColumn {
                            id: attr_get(attrs, "id").map(|s| s.to_string()),
                            name: attr_get(attrs, "name").unwrap_or("").to_string(),
                            totals_row_function: attr_get(attrs, "totalsRowFunction")
                                .map(|s| s.to_string()),
                            totals_row_label: attr_get(attrs, "totalsRowLabel")
                                .map(|s| s.to_string()),
                            calculated_column_formula: None,
                        });
                        if matches!(ev, Ev::SelfClose(_, _))
                            && let (Some(t), Some(c)) = (table.as_mut(), cur_column.take())
                        {
                            t.columns.push(c);
                        }
                    }
                    "calculatedColumnFormula" if !matches!(ev, Ev::SelfClose(_, _)) => {
                        in_calc_formula = true;
                        calc_formula_text.clear();
                    }
                    "tableStyleInfo" => {
                        if let Some(t) = table.as_mut() {
                            t.style_name = attr_get(attrs, "name").map(|s| s.to_string());
                        }
                    }
                    "autoFilter" => {
                        if let Some(t) = table.as_mut() {
                            t.auto_filter_ref = attr_get(attrs, "ref").and_then(parse_merge_ref);
                        }
                    }
                    _ => {}
                }
            }
            Ev::Close(ref tag) => {
                let local = tag.split(':').next_back().unwrap_or(tag.as_str());
                match local {
                    "calculatedColumnFormula" if in_calc_formula => {
                        if let Some(c) = cur_column.as_mut() {
                            c.calculated_column_formula = Some(calc_formula_text.clone());
                        }
                        in_calc_formula = false;
                    }
                    "tableColumn" => {
                        if let (Some(t), Some(c)) = (table.as_mut(), cur_column.take()) {
                            t.columns.push(c);
                        }
                    }
                    _ => {}
                }
            }
            Ev::Text(ref text) => {
                if in_calc_formula {
                    calc_formula_text.push_str(text);
                }
            }
        }
    }
    // 0.16.0-B2: the streaming loop above only ever reads `<autoFilter>`'s own `ref`
    // attribute -- its `<filterColumn>` children are parsed the same way as a
    // standalone `AutoFilterDef`'s (`xlsx_autofilter`), via a second, targeted
    // extraction against the whole document rather than threading state through the
    // streaming iterator above.
    if let Some(t) = table.as_mut()
        && t.auto_filter_ref.is_some()
        && let Some(af_span) = extract_raw_element(xml, "autoFilter")
    {
        t.autofilter_columns = extract_records(&af_span, "autoFilter", "filterColumn")
            .iter()
            .filter_map(|s| parse_filter_column_xml(s))
            .collect();
    }
    table
}

#[cfg(test)]
mod table_parsing_tests {
    use super::*;

    const REAL_TABLE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1"
  name="Table1" displayName="Table1" ref="A1:C4" totalsRowShown="0">
  <autoFilter ref="A1:C4"/>
  <tableColumns count="3">
    <tableColumn id="1" name="Name"/>
    <tableColumn id="2" name="Qty"/>
    <tableColumn id="3" name="Status"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0"
    showRowStripes="1" showColumnStripes="0"/>
</table>"#;

    // Mirrors a real fixture's shape more closely than REAL_TABLE_XML: the Microsoft
    // xr:uid/xr3:uid extension GUIDs every real `<table>`/`<tableColumn>` carries
    // (0.16.0-A2's whole reason for a surgical patch instead of reserialize-from-struct).
    const TABLE_XML_WITH_GUIDS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1"
  xr:uid="{00000000-0001-0000-0000-000000000001}" name="Table1" displayName="Table1"
  ref="A1:C4" totalsRowShown="0">
  <autoFilter ref="A1:C4"/>
  <tableColumns count="3">
    <tableColumn id="1" xr3:uid="{00000000-0001-0000-0000-000000000002}" name="Name"/>
    <tableColumn id="2" xr3:uid="{00000000-0001-0000-0000-000000000003}" name="Qty"/>
    <tableColumn id="3" xr3:uid="{00000000-0001-0000-0000-000000000004}" name="Status"/>
  </tableColumns>
  <tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0"
    showRowStripes="1" showColumnStripes="0"/>
</table>"#;

    #[test]
    fn apply_table_edits_rename_preserves_id_and_guid_untouched() {
        let out = apply_table_edits(
            TABLE_XML_WITH_GUIDS,
            &[TableEditOp::SetDisplayName("Renamed".to_string())],
        );
        assert!(out.contains(r#"displayName="Renamed""#));
        assert!(out.contains(r#"id="1""#));
        assert!(out.contains(r#"xr:uid="{00000000-0001-0000-0000-000000000001}""#));
        // Untouched columns keep their own xr3:uid GUIDs verbatim.
        assert!(out.contains(r#"xr3:uid="{00000000-0001-0000-0000-000000000002}""#));
        assert!(out.contains(r#"xr3:uid="{00000000-0001-0000-0000-000000000003}""#));
        assert!(out.contains(r#"xr3:uid="{00000000-0001-0000-0000-000000000004}""#));
    }

    #[test]
    fn apply_table_edits_resize_updates_only_ref() {
        let out = apply_table_edits(REAL_TABLE_XML, &[TableEditOp::Resize(((1, 1), (5, 3)))]);
        let t = parse_table_xml(&out).unwrap();
        assert_eq!(t.ref_range, ((1, 1), (5, 3)));
        // The nested autoFilter's own ref is a SEPARATE op -- resizing the table alone
        // must not accidentally touch it.
        assert_eq!(t.auto_filter_ref, Some(((1, 1), (4, 3))));
    }

    #[test]
    fn apply_table_edits_resize_auto_filter_updates_only_the_nested_ref() {
        let out = apply_table_edits(
            REAL_TABLE_XML,
            &[TableEditOp::ResizeAutoFilter(((1, 1), (5, 3)))],
        );
        let t = parse_table_xml(&out).unwrap();
        assert_eq!(t.ref_range, ((1, 1), (4, 3)));
        assert_eq!(t.auto_filter_ref, Some(((1, 1), (5, 3))));
    }

    #[test]
    fn apply_table_edits_set_filter_column_adds_a_fresh_filter_column() {
        let out = apply_table_edits(
            REAL_TABLE_XML,
            &[TableEditOp::SetFilterColumn(
                1,
                FilterCriteria::Values(vec!["Yes".to_string()]),
            )],
        );
        assert!(out.contains(r#"colId="1""#));
        assert!(out.contains(r#"<filter val="Yes"/>"#));
        // The table's own ref/other siblings are untouched.
        assert!(out.contains(r#"ref="A1:C4""#));
    }

    #[test]
    fn apply_table_edits_set_filter_column_replaces_an_existing_entry_for_the_same_col_offset() {
        let with_one = apply_table_edits(
            REAL_TABLE_XML,
            &[TableEditOp::SetFilterColumn(1, FilterCriteria::Blank)],
        );
        let out = apply_table_edits(
            &with_one,
            &[TableEditOp::SetFilterColumn(
                1,
                FilterCriteria::Values(vec!["A".to_string()]),
            )],
        );
        assert_eq!(out.matches("filterColumn").count(), 2); // one open + one close tag
        assert!(!out.contains("blank"));
        assert!(out.contains(r#"<filter val="A"/>"#));
    }

    #[test]
    fn apply_table_edits_set_filter_column_preserves_an_unrelated_columns_raw_bytes() {
        let with_two = apply_table_edits(
            REAL_TABLE_XML,
            &[
                TableEditOp::SetFilterColumn(0, FilterCriteria::Blank),
                TableEditOp::SetFilterColumn(1, FilterCriteria::Values(vec!["A".to_string()])),
            ],
        );
        let out = apply_table_edits(
            &with_two,
            &[TableEditOp::SetFilterColumn(
                1,
                FilterCriteria::Values(vec!["B".to_string()]),
            )],
        );
        // Column 0's own criteria (untouched by the second call) survives verbatim.
        assert!(out.contains(r#"colId="0""#));
        assert!(out.contains(r#"<filters blank="1"/>"#));
        assert!(out.contains(r#"<filter val="B"/>"#));
        assert!(!out.contains(r#"<filter val="A"/>"#));
    }

    #[test]
    fn apply_table_edits_clear_filter_column_removes_only_the_targeted_entry() {
        let with_two = apply_table_edits(
            REAL_TABLE_XML,
            &[
                TableEditOp::SetFilterColumn(0, FilterCriteria::Blank),
                TableEditOp::SetFilterColumn(1, FilterCriteria::Values(vec!["A".to_string()])),
            ],
        );
        let out = apply_table_edits(&with_two, &[TableEditOp::ClearFilterColumn(0)]);
        assert!(!out.contains(r#"colId="0""#));
        assert!(out.contains(r#"colId="1""#));
        let t = parse_table_xml(&out).unwrap();
        assert_eq!(t.autofilter_columns.len(), 1);
        assert_eq!(t.autofilter_columns[0].col_offset, 1);
    }

    #[test]
    fn apply_table_edits_clear_filter_column_leaves_a_bare_self_closing_autofilter_when_empty() {
        let with_one = apply_table_edits(
            REAL_TABLE_XML,
            &[TableEditOp::SetFilterColumn(0, FilterCriteria::Blank)],
        );
        let out = apply_table_edits(&with_one, &[TableEditOp::ClearFilterColumn(0)]);
        assert!(out.contains(r#"<autoFilter ref="A1:C4"/>"#));
        let t = parse_table_xml(&out).unwrap();
        assert!(t.autofilter_columns.is_empty());
    }

    #[test]
    fn parse_table_xml_reads_a_real_nested_filter_column() {
        let xml = r#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          id="1" name="Table1" displayName="Table1" ref="A1:C4">
          <autoFilter ref="A1:C4">
            <filterColumn colId="0"><filters><filter val="X"/></filters></filterColumn>
          </autoFilter>
          <tableColumns count="3">
            <tableColumn id="1" name="Name"/><tableColumn id="2" name="Qty"/>
            <tableColumn id="3" name="Status"/>
          </tableColumns>
        </table>"#;
        let t = parse_table_xml(xml).unwrap();
        assert_eq!(t.autofilter_columns.len(), 1);
        assert_eq!(t.autofilter_columns[0].col_offset, 0);
        assert_eq!(
            t.autofilter_columns[0].criteria,
            FilterCriteria::Values(vec!["X".to_string()])
        );
    }

    #[test]
    fn apply_table_edits_set_style_replaces_the_whole_element() {
        let out = apply_table_edits(
            REAL_TABLE_XML,
            &[TableEditOp::SetStyle(Some("TableStyleLight1".to_string()))],
        );
        let t = parse_table_xml(&out).unwrap();
        assert_eq!(t.style_name.as_deref(), Some("TableStyleLight1"));
    }

    #[test]
    fn apply_table_edits_set_totals_row_shown() {
        let out = apply_table_edits(REAL_TABLE_XML, &[TableEditOp::SetTotalsRowShown(true)]);
        let t = parse_table_xml(&out).unwrap();
        assert!(t.totals_row_shown);
    }

    #[test]
    fn apply_table_edits_add_column_assigns_a_fresh_id_and_preserves_existing_guids() {
        let out = apply_table_edits(
            TABLE_XML_WITH_GUIDS,
            &[TableEditOp::AddColumn("Total".to_string())],
        );
        let t = parse_table_xml(&out).unwrap();
        assert_eq!(t.columns.len(), 4);
        assert_eq!(t.columns[3].name, "Total");
        assert_eq!(t.columns[3].id.as_deref(), Some("4")); // max existing id (3) + 1
        assert!(out.contains(r#"xr3:uid="{00000000-0001-0000-0000-000000000002}""#));
        assert!(out.contains(r#"count="4""#));
    }

    #[test]
    fn apply_table_edits_remove_column_drops_only_the_named_one() {
        let out = apply_table_edits(
            TABLE_XML_WITH_GUIDS,
            &[TableEditOp::RemoveColumn("Qty".to_string())],
        );
        let t = parse_table_xml(&out).unwrap();
        assert_eq!(t.columns.len(), 2);
        assert_eq!(t.columns[0].name, "Name");
        assert_eq!(t.columns[1].name, "Status");
        // The surviving columns' own GUIDs are untouched -- Qty's own span is gone,
        // Name's and Status's are byte-identical, not renumbered.
        assert!(out.contains(r#"xr3:uid="{00000000-0001-0000-0000-000000000002}""#));
        assert!(out.contains(r#"xr3:uid="{00000000-0001-0000-0000-000000000004}""#));
        assert!(!out.contains(r#"xr3:uid="{00000000-0001-0000-0000-000000000003}""#));
        assert!(out.contains(r#"count="2""#));
    }

    #[test]
    fn parse_table_xml_extracts_every_field_from_a_real_shape() {
        let t = parse_table_xml(REAL_TABLE_XML).unwrap();
        assert_eq!(t.name, "Table1");
        assert_eq!(t.display_name, "Table1");
        assert_eq!(t.ref_range, ((1, 1), (4, 3)));
        assert_eq!(t.header_row_count, 1); // absent -> spec default
        assert_eq!(t.totals_row_count, 0);
        assert!(!t.totals_row_shown); // explicit totalsRowShown="0"
        assert_eq!(t.style_name.as_deref(), Some("TableStyleMedium2"));
        assert_eq!(t.auto_filter_ref, Some(((1, 1), (4, 3))));
        assert_eq!(t.columns.len(), 3);
        assert_eq!(t.columns[0].name, "Name");
        assert_eq!(t.columns[1].name, "Qty");
        assert_eq!(t.columns[2].name, "Status");
        assert!(
            t.columns
                .iter()
                .all(|c| c.calculated_column_formula.is_none())
        );
    }

    #[test]
    fn parse_table_xml_defaults_totals_row_shown_to_true_when_absent() {
        let xml = r#"<table name="T" displayName="T" ref="A1:B2">
          <tableColumns count="1"><tableColumn id="1" name="X"/></tableColumns>
        </table>"#;
        let t = parse_table_xml(xml).unwrap();
        assert!(t.totals_row_shown);
        assert!(t.auto_filter_ref.is_none());
        assert!(t.style_name.is_none());
    }

    #[test]
    fn parse_table_xml_captures_a_calculated_column_formula_as_raw_text() {
        let xml = r#"<table name="T" displayName="T" ref="A1:B2">
          <tableColumns count="1">
            <tableColumn id="1" name="Total">
              <calculatedColumnFormula>[@Qty]*[@Price]</calculatedColumnFormula>
            </tableColumn>
          </tableColumns>
        </table>"#;
        let t = parse_table_xml(xml).unwrap();
        assert_eq!(
            t.columns[0].calculated_column_formula.as_deref(),
            Some("[@Qty]*[@Price]")
        );
    }

    #[test]
    fn parse_table_xml_returns_none_when_ref_is_missing_or_unparseable() {
        assert!(parse_table_xml(r#"<table name="T" displayName="T"/>"#).is_none());
        assert!(parse_table_xml(r#"<table name="T" displayName="T" ref="A1"/>"#).is_none());
    }

    #[test]
    fn xlsx_table_part_rids_extracts_every_id_in_document_order() {
        let sheet_xml = r#"<worksheet><tableParts count="2">
          <tablePart r:id="rId3"/><tablePart r:id="rId7"/>
        </tableParts></worksheet>"#;
        assert_eq!(xlsx_table_part_rids(sheet_xml), vec!["rId3", "rId7"]);
    }

    #[test]
    fn xlsx_table_part_rids_is_empty_without_a_tableparts_element() {
        assert!(xlsx_table_part_rids("<worksheet></worksheet>").is_empty());
    }

    #[test]
    fn relationship_ids_extracts_every_id_ignoring_type_and_target() {
        let xml = concat!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
            r#"<Relationship Id="rId1" Type="a" Target="b"/>"#,
            r#"<Relationship Id="rId2" Type="c" Target="d"/>"#,
            r#"</Relationships>"#,
        );
        assert_eq!(relationship_ids(xml), vec!["rId1", "rId2"]);
    }

    #[test]
    fn relationship_ids_is_empty_for_a_relationships_document_with_no_entries() {
        let xml = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
        assert!(relationship_ids(xml).is_empty());
    }

    #[test]
    fn insert_before_close_inserts_a_relationship_into_an_existing_rels_document() {
        let xml = concat!(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
            r#"<Relationship Id="rId1" Type="a" Target="b"/>"#,
            r#"</Relationships>"#,
        );
        let out = insert_before_close(xml, r#"<Relationship Id="rId2" Type="c" Target="d"/>"#);
        assert_eq!(relationship_ids(&out), vec!["rId1", "rId2"]);
        // The existing relationship's own bytes survive untouched.
        assert!(out.contains(r#"<Relationship Id="rId1" Type="a" Target="b"/>"#));
    }

    fn sample_table_def(ref_range: MergeRect) -> TableDef {
        TableDef {
            name: "Table1".to_string(),
            display_name: "Table1".to_string(),
            ref_range,
            header_row_count: 1,
            totals_row_count: 0,
            totals_row_shown: false,
            columns: vec![
                TableColumn {
                    id: None,
                    name: "Name".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                    calculated_column_formula: None,
                },
                TableColumn {
                    id: None,
                    name: "Qty".to_string(),
                    totals_row_function: None,
                    totals_row_label: None,
                    calculated_column_formula: None,
                },
            ],
            style_name: None,
            auto_filter_ref: Some(ref_range),
            autofilter_columns: Vec::new(),
            source_part: String::new(),
            pending_edits: Vec::new(),
        }
    }

    #[test]
    fn render_table_xml_produces_output_that_parse_table_xml_reads_back_correctly() {
        let table = sample_table_def(((1, 1), (3, 2)));
        let xml = render_table_xml(&table, 7);
        let parsed = parse_table_xml(&xml).expect("must parse back");
        assert_eq!(parsed.name, "Table1");
        assert_eq!(parsed.display_name, "Table1");
        assert_eq!(parsed.ref_range, ((1, 1), (3, 2)));
        assert_eq!(parsed.header_row_count, 1);
        assert!(!parsed.totals_row_shown);
        assert_eq!(
            parsed
                .columns
                .iter()
                .map(|c| c.name.as_str())
                .collect::<Vec<_>>(),
            vec!["Name", "Qty"]
        );
        assert_eq!(parsed.auto_filter_ref, Some(((1, 1), (3, 2))));
        assert!(parsed.style_name.is_none());
    }

    #[test]
    fn render_table_xml_includes_a_style_when_one_is_set() {
        let mut table = sample_table_def(((1, 1), (1, 1)));
        table.style_name = Some("TableStyleMedium2".to_string());
        let xml = render_table_xml(&table, 1);
        assert!(xml.contains(r#"<tableStyleInfo name="TableStyleMedium2""#));
        let parsed = parse_table_xml(&xml).expect("must parse back");
        assert_eq!(parsed.style_name.as_deref(), Some("TableStyleMedium2"));
    }

    #[test]
    fn render_table_xml_embeds_filter_criteria_already_set_before_the_first_save() {
        // Regression: a table created via create_table can already have set_table_*_filter
        // called on it before ANY save -- render_table_xml (the from-scratch write path,
        // distinct from apply_table_edits' patch-existing-bytes path) must embed
        // `autofilter_columns` too, or the criteria silently never reaches disk.
        let mut table = sample_table_def(((1, 1), (1, 1)));
        table.autofilter_columns = vec![FilterColumn {
            col_offset: 0,
            hidden_button: false,
            show_button: true,
            criteria: FilterCriteria::Blank,
            raw_span: None,
            dirty: false,
        }];
        let xml = render_table_xml(&table, 1);
        assert!(xml.contains(r#"colId="0""#));
        assert!(xml.contains(r#"<filters blank="1"/>"#));
        let parsed = parse_table_xml(&xml).expect("must parse back");
        assert_eq!(parsed.autofilter_columns.len(), 1);
        assert_eq!(parsed.autofilter_columns[0].criteria, FilterCriteria::Blank);
    }

    #[test]
    fn render_table_xml_omits_xr_uid_extension_guids() {
        // Confirmed safe against a real openpyxl-authored table (0.16.0-A3 design doc's
        // A3 Addendum) -- real Excel-authored tables always carry xr:uid/xr3:uid, but
        // they're Excel's own added-on-first-touch metadata, not required for a valid,
        // openable file.
        let xml = render_table_xml(&sample_table_def(((1, 1), (1, 1))), 1);
        assert!(!xml.contains("xr:uid"));
        assert!(!xml.contains("xr3:uid"));
    }
}

#[cfg(test)]
mod input_extension_tests {
    use super::read_workbook;

    #[test]
    fn read_workbook_rejects_unsupported_extension_before_opening() {
        let error = match read_workbook("/definitely/not/there.xlsb") {
            Ok(_) => panic!("unsupported extension should be rejected"),
            Err(error) => error,
        };
        assert_eq!(
            error,
            "unsupported input extension; use .xlsx, .xlsm, or .ods"
        );
    }

    #[test]
    fn read_workbook_rejects_missing_extension_before_opening() {
        let error = match read_workbook("/definitely/not/there") {
            Ok(_) => panic!("missing extension should be rejected"),
            Err(error) => error,
        };
        assert_eq!(
            error,
            "unsupported input extension; use .xlsx, .xlsm, or .ods"
        );
    }
}

#[cfg(test)]
mod data_validation_parsing_tests {
    use super::*;

    const REAL_LIST_DV: &str = r#"<dataValidations count="1"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E1" xr:uid="{BF4C2CDE-5B18-5247-880B-6E29EFBEE104}"><formula1>"Yes,No,Maybe"</formula1></dataValidation></dataValidations>"#;

    #[test]
    fn xlsx_data_validations_extracts_every_field_from_a_real_shape() {
        let sheet_xml = format!("<worksheet>{REAL_LIST_DV}</worksheet>");
        let rules = xlsx_data_validations(&sheet_xml);
        assert_eq!(rules.len(), 1);
        let r = &rules[0];
        assert_eq!(r.validation_type, "list");
        assert_eq!(r.operator, None);
        assert_eq!(r.formula1.as_deref(), Some(r#""Yes,No,Maybe""#));
        assert_eq!(r.formula2, None);
        assert!(r.allow_blank);
        assert!(r.show_input_message);
        assert!(r.show_error_message);
        assert_eq!(r.sqref, vec![((1, 5), (1, 5))]);
        assert!(!r.dirty);
        assert!(r.raw_span.contains("xr:uid"));
    }

    #[test]
    fn xlsx_data_validations_is_empty_without_a_datavalidations_element() {
        assert!(xlsx_data_validations("<worksheet></worksheet>").is_empty());
    }

    #[test]
    fn xlsx_data_validations_reads_an_operator_and_two_formulas() {
        let sheet_xml = r#"<worksheet><dataValidations count="1"><dataValidation type="whole" operator="between" allowBlank="0" sqref="A1:A5"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations></worksheet>"#;
        let rules = xlsx_data_validations(sheet_xml);
        assert_eq!(rules.len(), 1);
        assert_eq!(rules[0].operator.as_deref(), Some("between"));
        assert_eq!(rules[0].formula1.as_deref(), Some("1"));
        assert_eq!(rules[0].formula2.as_deref(), Some("10"));
        assert!(!rules[0].allow_blank);
        assert_eq!(rules[0].sqref, vec![((1, 1), (5, 1))]);
    }

    #[test]
    fn xlsx_data_validations_reads_multiple_records_in_document_order() {
        let sheet_xml = r#"<worksheet><dataValidations count="2"><dataValidation type="list" sqref="A1"><formula1>"X,Y"</formula1></dataValidation><dataValidation type="custom" sqref="B1"><formula1>ISNUMBER(B1)</formula1></dataValidation></dataValidations></worksheet>"#;
        let rules = xlsx_data_validations(sheet_xml);
        assert_eq!(rules.len(), 2);
        assert_eq!(rules[0].validation_type, "list");
        assert_eq!(rules[1].validation_type, "custom");
    }

    #[test]
    fn xlsx_autofilter_is_none_without_an_autofilter_element() {
        assert!(xlsx_autofilter("<worksheet></worksheet>").is_none());
    }

    #[test]
    fn xlsx_autofilter_parses_a_bare_ref_with_no_columns() {
        let af = xlsx_autofilter(r#"<worksheet><autoFilter ref="A1:C21"/></worksheet>"#).unwrap();
        assert_eq!(af.ref_range, ((1, 1), (21, 3)));
        assert!(af.columns.is_empty());
    }

    #[test]
    fn xlsx_autofilter_parses_a_values_filter_column() {
        let xml = r#"<worksheet><autoFilter ref="A1:E21"><filterColumn colId="0" hiddenButton="0" showButton="1"><filters><filter val="Item1"/><filter val="Item2"/></filters></filterColumn></autoFilter></worksheet>"#;
        let af = xlsx_autofilter(xml).unwrap();
        assert_eq!(af.columns.len(), 1);
        let col = &af.columns[0];
        assert_eq!(col.col_offset, 0);
        assert!(!col.hidden_button);
        assert!(col.show_button);
        assert_eq!(
            col.criteria,
            FilterCriteria::Values(vec!["Item1".to_string(), "Item2".to_string()])
        );
    }

    #[test]
    fn xlsx_autofilter_parses_a_one_condition_custom_filter() {
        let xml = r#"<worksheet><autoFilter ref="A1:B5"><filterColumn colId="1" hiddenButton="0" showButton="1"><customFilters and="0"><customFilter val="5" operator="greaterThan"/></customFilters></filterColumn></autoFilter></worksheet>"#;
        let af = xlsx_autofilter(xml).unwrap();
        assert_eq!(
            af.columns[0].criteria,
            FilterCriteria::Custom {
                op1: "greaterThan".to_string(),
                val1: "5".to_string(),
                and: false,
                op2: None,
                val2: None,
            }
        );
    }

    #[test]
    fn xlsx_autofilter_parses_a_two_condition_and_custom_filter() {
        let xml = r#"<worksheet><autoFilter ref="A1:B5"><filterColumn colId="0" hiddenButton="0" showButton="1"><customFilters and="1"><customFilter val="10" operator="greaterThanOrEqual"/><customFilter val="20" operator="lessThanOrEqual"/></customFilters></filterColumn></autoFilter></worksheet>"#;
        let af = xlsx_autofilter(xml).unwrap();
        assert_eq!(
            af.columns[0].criteria,
            FilterCriteria::Custom {
                op1: "greaterThanOrEqual".to_string(),
                val1: "10".to_string(),
                and: true,
                op2: Some("lessThanOrEqual".to_string()),
                val2: Some("20".to_string()),
            }
        );
    }

    #[test]
    fn xlsx_autofilter_parses_a_blank_filter() {
        let xml = r#"<worksheet><autoFilter ref="A1:D5"><filterColumn colId="3" hiddenButton="0" showButton="1"><filters blank="1"/></filterColumn></autoFilter></worksheet>"#;
        let af = xlsx_autofilter(xml).unwrap();
        assert_eq!(af.columns[0].criteria, FilterCriteria::Blank);
    }

    #[test]
    fn xlsx_autofilter_parses_a_top10_filter() {
        let xml = r#"<worksheet><autoFilter ref="A1:B21"><filterColumn colId="1" hiddenButton="0" showButton="1"><top10 top="1" percent="0" val="5"/></filterColumn></autoFilter></worksheet>"#;
        let af = xlsx_autofilter(xml).unwrap();
        assert_eq!(
            af.columns[0].criteria,
            FilterCriteria::Top10 {
                top: true,
                percent: false,
                val: 5.0
            }
        );
    }

    #[test]
    fn xlsx_autofilter_parses_a_date_group_filter() {
        let xml = r#"<worksheet><autoFilter ref="A1:B5"><filterColumn colId="1" hiddenButton="0" showButton="1"><filters calendarType="gregorian"><dateGroupItem year="2024" month="1" dateTimeGrouping="month"/></filters></filterColumn></autoFilter></worksheet>"#;
        let af = xlsx_autofilter(xml).unwrap();
        assert_eq!(
            af.columns[0].criteria,
            FilterCriteria::DateGroup(vec![DateGroupItem {
                year: Some(2024),
                month: Some(1),
                day: None,
                hour: None,
                minute: None,
                second: None,
                date_time_grouping: "month".to_string(),
            }])
        );
    }

    #[test]
    fn xlsx_autofilter_reads_multiple_filter_columns_in_document_order() {
        let xml = r#"<worksheet><autoFilter ref="A1:C21"><filterColumn colId="0" hiddenButton="0" showButton="1"><filters blank="1"/></filterColumn><filterColumn colId="2" hiddenButton="0" showButton="1"><filters blank="1"/></filterColumn></autoFilter></worksheet>"#;
        let af = xlsx_autofilter(xml).unwrap();
        assert_eq!(af.columns.len(), 2);
        assert_eq!(af.columns[0].col_offset, 0);
        assert_eq!(af.columns[1].col_offset, 2);
    }

    #[test]
    fn parse_sqref_handles_single_cell_and_multi_area() {
        assert_eq!(parse_sqref("E1"), vec![((1, 5), (1, 5))]);
        assert_eq!(
            parse_sqref("A1:A5 C1:C5"),
            vec![((1, 1), (5, 1)), ((1, 3), (5, 3))]
        );
    }

    #[test]
    fn parse_sqref_tolerates_an_unparseable_token() {
        // Real-world tolerance, matching this reader's convention elsewhere: a bad
        // token contributes nothing rather than failing the whole parse.
        assert_eq!(
            parse_sqref("A1 !!! C1"),
            vec![((1, 1), (1, 1)), ((1, 3), (1, 3))]
        );
    }
}

// ── ODS reader ────────────────────────────────────────────────────────────────

fn read_ods(path: &str) -> Result<Vec<WorkbookSheet>, String> {
    let file = std::fs::File::open(path).map_err(|e| e.to_string())?;
    let mut archive = ZipArchive::new(file).map_err(|e| e.to_string())?;
    validate_zip_archive(&mut archive)?;
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
                            cell_number_formats: HashMap::new(),
                            sheet_state: None,
                            row_heights: HashMap::new(),
                            column_widths: Vec::new(),
                            row_styles: HashMap::new(),
                            column_styles: Vec::new(),
                            tables: Vec::new(),
                            data_validations: Vec::new(),
                            autofilter: None,
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
                    Some("1".to_string()),
                    None
                ),
                (
                    "Sheet2".to_string(),
                    "rId2".to_string(),
                    Some("5".to_string()),
                    None
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
            vec![("Sheet1".to_string(), "rId1".to_string(), None, None)]
        );
    }

    #[test]
    fn xlsx_workbook_sheets_captures_the_state_attribute() {
        let xml = r#"<sheets>
<sheet name="Sheet1" sheetId="1" r:id="rId1"/>
<sheet name="Sheet2" sheetId="2" r:id="rId2" state="hidden"/>
<sheet name="Sheet3" sheetId="3" r:id="rId3" state="veryHidden"/>
</sheets>"#;
        let result = xlsx_workbook_sheets(xml);
        let states: Vec<Option<String>> = result.into_iter().map(|(_, _, _, s)| s).collect();
        assert_eq!(
            states,
            vec![
                None,
                Some("hidden".to_string()),
                Some("veryHidden".to_string())
            ]
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

#[cfg(test)]
mod defined_names_tests {
    use super::*;

    #[test]
    fn xlsx_defined_names_captures_name_and_raw_text() {
        let xml = r#"<workbook><definedNames>
<definedName name="MyRange">Sheet1!$A$1:$A$3</definedName>
<definedName name="Other" localSheetId="0">Sheet1!$B$1</definedName>
</definedNames></workbook>"#;
        assert_eq!(
            xlsx_defined_names(xml).unwrap(),
            vec![
                ("MyRange".to_string(), "Sheet1!$A$1:$A$3".to_string()),
                ("Other".to_string(), "Sheet1!$B$1".to_string()),
            ]
        );
    }

    #[test]
    fn xlsx_defined_names_is_empty_when_absent() {
        let xml = r#"<workbook><sheets><sheet name="Sheet1" r:id="rId1"/></sheets></workbook>"#;
        assert_eq!(
            xlsx_defined_names(xml).unwrap(),
            Vec::<(String, String)>::new()
        );
    }

    #[test]
    fn xlsx_defined_names_xml_unescapes_the_text_content() {
        let xml = r#"<definedNames><definedName name="X">Sheet1!$A$1 &amp; "text"</definedName></definedNames>"#;
        assert_eq!(
            xlsx_defined_names(xml).unwrap(),
            vec![("X".to_string(), "Sheet1!$A$1 & \"text\"".to_string())]
        );
    }

    #[test]
    fn xlsx_defined_names_rejects_a_table_over_the_count_limit() {
        let mut xml = String::from("<workbook><definedNames>");
        for index in 0..=DEFINED_NAMES_MAX_COUNT {
            xml.push_str(&format!(
                "<definedName name=\"Name{index}\">Sheet1!$A$1</definedName>"
            ));
        }
        xml.push_str("</definedNames></workbook>");

        let error = xlsx_defined_names(&xml).unwrap_err();
        assert_eq!(
            error,
            "defined-name table is too large (more than 100000; maximum is 100000)"
        );
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

    // ── P2: row height / column width parsing ───────────────────────────────

    #[test]
    fn xlsx_sheet_cells_reads_a_custom_row_height() {
        let xml = r#"<worksheet><sheetData>
<row r="5" ht="30.5" customHeight="1"><c r="A5"><v>1</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.row_heights.get(&5), Some(&30.5));
    }

    #[test]
    fn xlsx_sheet_cells_ignores_ht_without_custom_height() {
        // Real Excel ignores a bare `ht` without `customHeight="1"` -- some
        // producers emit `ht` alongside an auto-fit row without ever setting
        // the flag, and that must not be recorded as an explicit height.
        let xml = r#"<worksheet><sheetData>
<row r="5" ht="30.5"><c r="A5"><v>1</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert!(data.row_heights.is_empty());
    }

    #[test]
    fn xlsx_sheet_cells_reads_a_custom_column_width_range() {
        let xml = r#"<worksheet><cols>
<col min="2" max="4" width="12.5" customWidth="1"/>
</cols><sheetData></sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.column_widths, vec![(2, 4, 12.5)]);
    }

    #[test]
    fn xlsx_sheet_cells_ignores_width_without_custom_width() {
        let xml = r#"<worksheet><cols>
<col min="2" max="4" width="12.5"/>
</cols><sheetData></sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert!(data.column_widths.is_empty());
    }

    #[test]
    fn xlsx_sheet_cells_row_height_and_hidden_are_independent() {
        // A row can be both explicitly-heighted and hidden at once -- confirm
        // one attribute's parsing doesn't clobber the other's.
        let xml = r#"<worksheet><sheetData>
<row r="5" ht="20" customHeight="1" hidden="1"/>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        assert_eq!(data.row_heights.get(&5), Some(&20.0));
        assert_eq!(data.hidden_rows, vec![(5, 5)]);
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
    fn xlsx_sheet_cells_reads_a_t_e_cell_as_a_real_error_not_a_string() {
        // Real shape confirmed live from fixture5_chart_image_freeze_print.xlsm's D8 (see
        // ROADMAP.md Known gaps item 14): `t="e"` used to be treated identically to
        // `t="str"`, silently dropping the error type.
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1" t="e"><v>#VALUE!</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        match data.cells.get(&(1, 1)) {
            Some(SheetCell::Error(e)) => assert_eq!(e.as_str(), "#VALUE!"),
            other => panic!(
                "expected Error(\"#VALUE!\") at A1, got {:?}",
                other.is_some()
            ),
        }
    }

    #[test]
    fn xlsx_sheet_cells_falls_back_to_a_plain_string_for_an_unrecognized_t_e_value() {
        // A newer dynamic-array error (#SPILL!) or malformed input -- reader.rs only
        // recognizes the 7 classic error strings (no fixture evidence of any other), so
        // this must not panic or silently invent a wrong error code.
        let xml = r#"<worksheet><sheetData>
<row r="1"><c r="A1" t="e"><v>#SPILL!</v></c></row>
</sheetData></worksheet>"#;
        let data = xlsx_sheet_cells(xml, &[], &[]);
        match data.cells.get(&(1, 1)) {
            Some(SheetCell::Str(s)) => assert_eq!(s, "#SPILL!"),
            other => panic!("expected Str(\"#SPILL!\") at A1, got {:?}", other.is_some()),
        }
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

    // ── GitHub #4: resolve_number_format ────────────────────────────────────

    #[test]
    fn resolve_number_format_finds_a_builtin_date_format() {
        assert_eq!(
            resolve_number_format(14, &HashMap::new()).as_deref(),
            Some("m/d/yyyy")
        );
    }

    #[test]
    fn resolve_number_format_general_is_none() {
        assert_eq!(resolve_number_format(0, &HashMap::new()), None);
    }

    #[test]
    fn resolve_number_format_an_unknown_id_with_no_custom_definition_is_none() {
        assert_eq!(resolve_number_format(9999, &HashMap::new()), None);
    }

    #[test]
    fn resolve_number_format_prefers_a_custom_definition_over_the_builtin_table() {
        let mut custom = HashMap::new();
        custom.insert(14, "yyyy-mm-dd".to_string());
        assert_eq!(
            resolve_number_format(14, &custom).as_deref(),
            Some("yyyy-mm-dd")
        );
    }

    #[test]
    fn resolve_number_format_finds_a_custom_format_above_id_163() {
        let mut custom = HashMap::new();
        custom.insert(164, "0.00\"kg\"".to_string());
        assert_eq!(
            resolve_number_format(164, &custom).as_deref(),
            Some("0.00\"kg\"")
        );
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
    use std::io::Write;
    use zip::write::SimpleFileOptions;

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
            (SheetCell::Error(x), Some(SheetCell::Error(y))) => x == y,
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

    #[test]
    fn zip_validation_rejects_path_traversal_entries() {
        let mut zip = zip::ZipWriter::new(Cursor::new(Vec::new()));
        zip.start_file("../outside.xml", SimpleFileOptions::default())
            .unwrap();
        zip.write_all(b"not valid workbook data").unwrap();
        let bytes = zip.finish().unwrap().into_inner();
        let mut archive = ZipArchive::new(Cursor::new(bytes)).unwrap();
        let error = validate_zip_archive(&mut archive).unwrap_err();
        assert!(error.contains("unsafe path"));
    }

    #[test]
    fn zip_entry_metadata_rejects_each_resource_limit() {
        let error =
            validate_zip_entry_metadata("large.xml", ZIP_ENTRY_MAX_BYTES + 1, 1, 0).unwrap_err();
        assert!(error.contains("too large"));

        let error =
            validate_zip_entry_metadata("combined.xml", 1, 1, ZIP_MAX_TOTAL_BYTES).unwrap_err();
        assert!(error.contains("maximum size"));

        let error = validate_zip_entry_metadata("bomb.xml", ZIP_MAX_COMPRESSION_RATIO + 1, 1, 0)
            .unwrap_err();
        assert!(error.contains("compression ratio"));

        let error = validate_zip_entry_metadata("/absolute.xml", 1, 1, 0).unwrap_err();
        assert!(error.contains("unsafe path"));
    }

    #[test]
    fn zip_entry_metadata_accepts_limits_without_exceeding_them() {
        assert_eq!(
            validate_zip_entry_metadata(
                "ok.xml",
                ZIP_ENTRY_MAX_BYTES,
                ZIP_ENTRY_MAX_BYTES / ZIP_MAX_COMPRESSION_RATIO,
                0
            )
            .unwrap(),
            ZIP_ENTRY_MAX_BYTES
        );
        assert_eq!(
            validate_zip_entry_metadata("ok.xml", 1, 1, ZIP_MAX_TOTAL_BYTES - 1).unwrap(),
            ZIP_MAX_TOTAL_BYTES
        );
        assert_eq!(
            validate_zip_entry_metadata("ok.xml", ZIP_MAX_COMPRESSION_RATIO, 1, 0).unwrap(),
            ZIP_MAX_COMPRESSION_RATIO
        );
    }

    #[test]
    fn xml_budget_rejects_external_entity_declarations() {
        let error = validate_xml_budget(
            "workbook.xml",
            "<!DOCTYPE workbook [<!ENTITY x SYSTEM 'file:///secret'>]><workbook/>",
        )
        .unwrap_err();
        assert!(error.contains("DTD or entity"));
    }

    #[test]
    fn xml_budget_rejects_unclosed_documents() {
        let error = validate_xml_budget("sheet.xml", "<worksheet><sheetData/>").unwrap_err();
        assert!(error.contains("unclosed"));
    }

    #[test]
    fn xml_budget_accepts_a_normal_document_and_xsd_boolean_literals() {
        let xml =
            r#"<?xml version="1.0"?><worksheet><row hidden="true"><c r="A1"/></row></worksheet>"#;
        validate_xml_budget("sheet.xml", xml).expect("normal XML should stay within the budget");
    }

    #[test]
    fn workbook_model_limits_reject_excessive_shape_counts() {
        let error = validate_workbook_model_count(WORKBOOK_MAX_SHEETS + 1).unwrap_err();
        assert!(error.contains("too many sheets"));

        let error = validate_sheet_model("Sheet1", SHEET_MAX_CELLS + 1, 0).unwrap_err();
        assert!(error.contains("too many cells"));

        let error = validate_sheet_model("Sheet1", 0, SHEET_MAX_MERGES + 1).unwrap_err();
        assert!(error.contains("too many merged ranges"));
    }

    #[test]
    fn shared_string_limit_rejects_excessive_count() {
        let strings = vec![String::new(); SHARED_STRINGS_MAX_COUNT + 1];
        let error = validate_shared_strings(&strings).unwrap_err();
        assert!(error.contains("shared strings table is too large"));
    }
}
