//! Shared value types for elixcee, extracted so a future crate (e.g. an
//! XLSX-only or WASM build) can depend on them without pulling in the full
//! VBA parser/VM. std-only: no dependency on any other elixcee crate.

/// Excel worksheet error values (#DIV/0!, #N/A, etc.)
#[derive(Debug, Clone, PartialEq)]
pub enum ExcelError {
    DivZero, // #DIV/0!
    NA,      // #N/A
    Value,   // #VALUE!
    Ref,     // #REF!
    Name,    // #NAME?
    Num,     // #NUM!
    Null,    // #NULL!
}

impl ExcelError {
    pub fn as_str(&self) -> &'static str {
        match self {
            ExcelError::DivZero => "#DIV/0!",
            ExcelError::NA      => "#N/A",
            ExcelError::Value   => "#VALUE!",
            ExcelError::Ref     => "#REF!",
            ExcelError::Name    => "#NAME?",
            ExcelError::Num     => "#NUM!",
            ExcelError::Null    => "#NULL!",
        }
    }
}

impl std::fmt::Display for ExcelError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.as_str())
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Variant {
    Integer(i64),
    Float(f64),
    Str(String),
    Boolean(bool),
    Date(i64),           // Excel serial date — displays as "YYYY-MM-DD"
    Error(ExcelError),   // Excel error value (#DIV/0!, #N/A, …)
    Empty,
    Array(Vec<Variant>),                  // 0-indexed 1D array
    Record(std::collections::HashMap<String, Variant>), // UDT instance (p.x, p.y, …)
}

pub fn serial_to_display(s: i64) -> String {
    let (y, m, d) = serial_to_ymd(s);
    format!("{:04}-{:02}-{:02}", y, m, d)
}

impl std::fmt::Display for Variant {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Variant::Integer(n) => write!(f, "{}", n),
            Variant::Float(v)   => write!(f, "{}", v),
            Variant::Str(s)     => write!(f, "{}", s),
            Variant::Boolean(b) => write!(f, "{}", if *b { "True" } else { "False" }),
            Variant::Date(s)    => write!(f, "{}", serial_to_display(*s)),
            Variant::Error(e)   => write!(f, "{}", e),
            Variant::Empty      => write!(f, ""),
            Variant::Array(a)   => write!(f, "[{}]", a.iter().map(|v| v.to_string()).collect::<Vec<_>>().join(", ")),
            Variant::Record(m)  => {
                let mut pairs: Vec<String> = m.iter().map(|(k, v)| format!("{}: {}", k, v)).collect();
                pairs.sort();
                write!(f, "{{{}}}", pairs.join(", "))
            }
        }
    }
}

#[derive(Debug, Clone)]
pub struct CellContent {
    pub formula: Option<String>,
    pub value: Variant,
}

// ── Date-serial math ────────────────────────────────────────────────────────
// Moved from formula/eval.rs together with serial_to_ymd, which depends on
// both — previously private to that file, now pub since a different crate
// needs them; bodies unchanged.

pub fn is_leap(y: i32) -> bool { (y % 4 == 0 && y % 100 != 0) || y % 400 == 0 }

pub fn days_in_month(y: i32, m: u32) -> u32 {
    match m {
        1|3|5|7|8|10|12 => 31,
        4|6|9|11        => 30,
        2               => if is_leap(y) { 29 } else { 28 },
        _               => 0,
    }
}

pub fn serial_to_ymd(mut s: i64) -> (i32, u32, u32) {
    // Undo the Excel leap-year bug offset for dates after serial 60
    if s > 60 { s -= 1; }
    // s is now days since Jan 1 1900 (1-based)
    let mut y = 1900i32;
    loop {
        let days = if is_leap(y) { 366i64 } else { 365 };
        if s <= days { break; }
        s -= days;
        y += 1;
    }
    let mut m = 1u32;
    loop {
        let dim = days_in_month(y, m) as i64;
        if s <= dim { break; }
        s -= dim;
        m += 1;
    }
    (y, m, s as u32)
}

// ── Range address helpers ─────────────────────────────────────────────────────
// Moved from vm/mod.rs verbatim — pure string↔coordinate parsing, no
// further crate:: dependencies of its own.

fn col_letters_to_num_vm(s: &str) -> u32 {
    s.chars().fold(0u32, |acc, c| acc * 26 + (c.to_ascii_uppercase() as u32 - b'A' as u32 + 1))
}

pub fn parse_cell_addr(addr: &str) -> Option<(u32, u32)> {
    let addr = addr.trim().to_uppercase();
    let alpha_end = addr.find(|c: char| c.is_ascii_digit())?;
    if alpha_end == 0 { return None; }
    let col = col_letters_to_num_vm(&addr[..alpha_end]);
    let row: u32 = addr[alpha_end..].parse().ok()?;
    Some((row, col))
}

pub fn parse_range_addr(addr: &str) -> Option<((u32, u32), (u32, u32))> {
    let addr = addr.trim();
    if let Some(i) = addr.find(':') {
        Some((parse_cell_addr(&addr[..i])?, parse_cell_addr(&addr[i+1..])?))
    } else {
        let c = parse_cell_addr(addr)?;
        Some((c, c))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── ExcelError ───────────────────────────────────────────────────────

    #[test]
    fn excel_error_as_str_covers_every_variant() {
        assert_eq!(ExcelError::DivZero.as_str(), "#DIV/0!");
        assert_eq!(ExcelError::NA.as_str(), "#N/A");
        assert_eq!(ExcelError::Value.as_str(), "#VALUE!");
        assert_eq!(ExcelError::Ref.as_str(), "#REF!");
        assert_eq!(ExcelError::Name.as_str(), "#NAME?");
        assert_eq!(ExcelError::Num.as_str(), "#NUM!");
        assert_eq!(ExcelError::Null.as_str(), "#NULL!");
    }

    #[test]
    fn excel_error_display_matches_as_str() {
        assert_eq!(ExcelError::DivZero.to_string(), "#DIV/0!");
    }

    // ── Variant ──────────────────────────────────────────────────────────

    #[test]
    fn variant_display_formats_every_kind() {
        assert_eq!(Variant::Integer(42).to_string(), "42");
        assert_eq!(Variant::Float(3.5).to_string(), "3.5");
        assert_eq!(Variant::Str("hi".into()).to_string(), "hi");
        assert_eq!(Variant::Boolean(true).to_string(), "True");
        assert_eq!(Variant::Boolean(false).to_string(), "False");
        assert_eq!(Variant::Empty.to_string(), "");
        assert_eq!(Variant::Error(ExcelError::NA).to_string(), "#N/A");
        assert_eq!(
            Variant::Array(vec![Variant::Integer(1), Variant::Integer(2)]).to_string(),
            "[1, 2]"
        );
    }

    #[test]
    fn variant_display_formats_a_date_via_serial_to_display() {
        assert_eq!(Variant::Date(45000).to_string(), serial_to_display(45000));
        assert_eq!(Variant::Date(1).to_string(), "1900-01-01");
    }

    #[test]
    fn variant_record_display_sorts_keys_for_determinism() {
        let mut m = std::collections::HashMap::new();
        m.insert("b".to_string(), Variant::Integer(2));
        m.insert("a".to_string(), Variant::Integer(1));
        assert_eq!(Variant::Record(m).to_string(), "{a: 1, b: 2}");
    }

    // ── CellContent ──────────────────────────────────────────────────────

    #[test]
    fn cell_content_holds_formula_and_value() {
        let c = CellContent { formula: Some("=A1+1".into()), value: Variant::Integer(2) };
        assert_eq!(c.formula.as_deref(), Some("=A1+1"));
        assert_eq!(c.value, Variant::Integer(2));
    }

    // ── serial_to_ymd / serial_to_display ───────────────────────────────

    #[test]
    fn serial_to_ymd_handles_the_epoch() {
        // Excel serial 1 = 1900-01-01.
        assert_eq!(serial_to_ymd(1), (1900, 1, 1));
    }

    #[test]
    fn serial_to_ymd_reproduces_the_excel_1900_leap_year_bug() {
        // Verified against this exact (pre-move) algorithm, not assumed:
        // serials 60 and 61 both decode to 1900-03-01 — this
        // implementation's `s > 60 { s -= 1 }` offset collapses the
        // fictitious "1900-02-29" into the following real day rather than
        // reconstructing it as a distinct date. Not this commit's concern
        // to change (a pure move preserves it exactly); this test exists
        // so a future change to this function trips a test, not just a
        // silent behavior drift.
        assert_eq!(serial_to_ymd(59), (1900, 2, 28));
        assert_eq!(serial_to_ymd(60), (1900, 3, 1));
        assert_eq!(serial_to_ymd(61), (1900, 3, 1));
        assert_eq!(serial_to_ymd(62), (1900, 3, 2));
    }

    #[test]
    fn serial_to_display_formats_as_iso_date() {
        assert_eq!(serial_to_display(1), "1900-01-01");
    }

    // ── parse_cell_addr / parse_range_addr ──────────────────────────────

    #[test]
    fn parse_cell_addr_handles_single_and_double_letter_columns() {
        assert_eq!(parse_cell_addr("A1"), Some((1, 1)));
        assert_eq!(parse_cell_addr("Z1"), Some((1, 26)));
        assert_eq!(parse_cell_addr("AA1"), Some((1, 27)));
    }

    #[test]
    fn parse_cell_addr_is_case_insensitive_and_trims_whitespace() {
        assert_eq!(parse_cell_addr(" a1 "), Some((1, 1)));
    }

    #[test]
    fn parse_cell_addr_rejects_malformed_input() {
        assert_eq!(parse_cell_addr("1A"), None);
        assert_eq!(parse_cell_addr(""), None);
        assert_eq!(parse_cell_addr("A"), None);
    }

    #[test]
    fn parse_range_addr_handles_a_single_cell_as_a_1x1_range() {
        assert_eq!(parse_range_addr("B2"), Some(((2, 2), (2, 2))));
    }

    #[test]
    fn parse_range_addr_handles_a_two_cell_range() {
        assert_eq!(parse_range_addr("A1:C3"), Some(((1, 1), (3, 3))));
    }

    #[test]
    fn parse_range_addr_rejects_a_malformed_piece() {
        assert_eq!(parse_range_addr("A1:"), None);
        assert_eq!(parse_range_addr("!!"), None);
    }
}
