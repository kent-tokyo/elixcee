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
            ExcelError::NA => "#N/A",
            ExcelError::Value => "#VALUE!",
            ExcelError::Ref => "#REF!",
            ExcelError::Name => "#NAME?",
            ExcelError::Num => "#NUM!",
            ExcelError::Null => "#NULL!",
        }
    }

    /// The classic BIFF numeric error code SheetJS (and every serious XLSX consumer's
    /// in-memory model) uses for `t:"e"` cells instead of the display string -- confirmed
    /// live against the real `xlsx` oracle's own `BErr` table
    /// (`compat/node_modules/xlsx/xlsx.js`): reading a real Excel-authored `t="e"` cell
    /// through `XLSX.read()` always comes back as `{t:"e", v:<this code>, w:<as_str()>}`,
    /// never the display string in `v`. Used by `crates/elixcee-wasm` so `@elixcee/xlsx`'s
    /// `read()` matches that shape exactly.
    pub fn biff_code(&self) -> u8 {
        match self {
            ExcelError::Null => 0x00,
            ExcelError::DivZero => 0x07,
            ExcelError::Value => 0x0F,
            ExcelError::Ref => 0x17,
            ExcelError::Name => 0x1D,
            ExcelError::Num => 0x24,
            ExcelError::NA => 0x2A,
        }
    }
}

impl std::fmt::Display for ExcelError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.as_str())
    }
}

/// The `as_str` inverse: parses one of the 7 classic error strings (e.g. from a `t="e"`
/// cell's `<v>` text). An unrecognized string (a newer dynamic-array error like `#SPILL!`,
/// or plain malformed input) is `Err(())`, not a panic -- callers fall back to treating the
/// value as an opaque string rather than guessing at an error code with no fixture evidence.
impl std::str::FromStr for ExcelError {
    type Err = ();

    fn from_str(s: &str) -> Result<Self, Self::Err> {
        match s {
            "#DIV/0!" => Ok(ExcelError::DivZero),
            "#N/A" => Ok(ExcelError::NA),
            "#VALUE!" => Ok(ExcelError::Value),
            "#REF!" => Ok(ExcelError::Ref),
            "#NAME?" => Ok(ExcelError::Name),
            "#NUM!" => Ok(ExcelError::Num),
            "#NULL!" => Ok(ExcelError::Null),
            _ => Err(()),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Variant {
    Integer(i64),
    Float(f64),
    Str(String),
    Boolean(bool),
    Date(i64),         // Excel serial date — displays as "YYYY-MM-DD"
    Error(ExcelError), // Excel error value (#DIV/0!, #N/A, …)
    Empty,
    /// VBA's `Null` — "no valid data" (as from a database NULL), a
    /// genuinely different concept from `Empty` (an uninitialized Variant).
    /// `IsNull(Null)` is True and `IsEmpty(Null)` is False, and vice versa;
    /// arithmetic and comparison propagate Null where they treat Empty as
    /// 0/"". Kept as its own variant precisely so those rules are
    /// expressible — folding it into `Empty` is what made every documented
    /// Null-propagation rule unimplementable before.
    Null,
    /// 0-indexed, always-1D — spreadsheet Range-value multi-cell reads
    /// (`Range("A1:B3").Value`), `formula::eval`'s array-formula/SUMPRODUCT/
    /// array-constant results, and `DimArrayRecord`/`ArrayRecordSet`'s
    /// record-array storage. None of these are VBA `Dim`-declared arrays —
    /// see `VbaArray` below for those.
    Array(Vec<Variant>),
    /// A real (possibly multi-dimensional) VBA array: `Dim arr(3, 2)`,
    /// `Dim arr(1 To 3, -2 To 2)`, `ReDim`, `Array(...)`, `Split(...)`.
    /// Deliberately a separate variant from `Array` above rather than a
    /// wholesale replacement — `Array`'s other three callers (Range-value
    /// reads, formula-engine array results, record arrays) aren't VBA
    /// `Dim`-declared arrays and have no per-dimension bounds to track.
    VbaArray(VbaArray),
    Record(std::collections::HashMap<String, Variant>), // UDT instance (p.x, p.y, …)
}

/// The bounds of one dimension of a VBA-declared array: an explicit `lo To
/// hi`, or the implicit `Option Base .. To <upper>` form collapsed to the
/// same shape. `upper < lower` is legal VBA (as when `Option Base 1` meets a
/// bare `Dim arr(0)`) and just means zero elements along that dimension.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct ArrayBound {
    pub lower: i64,
    pub upper: i64,
}

impl ArrayBound {
    pub fn len(&self) -> i64 {
        (self.upper - self.lower + 1).max(0)
    }

    pub fn is_empty(&self) -> bool {
        self.len() == 0
    }
}

// ponytail: fixed ceiling rather than a configurable one — raise if a real
// macro ever legitimately needs more; this exists to fail fast (as "Out of
// memory", VBA's own Error 7) on a fuzzer-shaped `Dim arr(2000000000,
// 2000000000)` instead of trying to allocate it.
pub const MAX_ARRAY_ELEMENTS: usize = 10_000_000;

/// A real, possibly multi-dimensional VBA array. `elements` is stored flat,
/// row-major (the first dimension varies slowest) — an implementation
/// detail invisible to VBA code, since every access goes through
/// `linear_index`/`get`/`set`.
#[derive(Debug, Clone, PartialEq)]
pub struct VbaArray {
    pub bounds: Vec<ArrayBound>,
    pub elements: Vec<Variant>,
}

impl VbaArray {
    /// Builds a zero-filled (`Variant::Empty`) array with the given
    /// per-dimension bounds. `Err("Out of memory")` (VBA's real Error 7) if
    /// the element-count product overflows or exceeds `MAX_ARRAY_ELEMENTS`.
    pub fn new_zeroed(bounds: Vec<ArrayBound>) -> Result<Self, String> {
        if bounds.is_empty() {
            return Err("array must have at least one dimension".to_string());
        }
        let mut total: i64 = 1;
        for b in &bounds {
            total = total
                .checked_mul(b.len())
                .filter(|&t| t >= 0 && (t as u64) <= MAX_ARRAY_ELEMENTS as u64)
                .ok_or_else(|| "Out of memory".to_string())?;
        }
        Ok(VbaArray {
            bounds,
            elements: vec![Variant::Empty; total as usize],
        })
    }

    /// A 1-D, 0-based array built directly from already-computed elements —
    /// `Array(...)`/`Split(...)`'s shape.
    pub fn from_vec(elements: Vec<Variant>) -> Self {
        let upper = elements.len() as i64 - 1;
        VbaArray {
            bounds: vec![ArrayBound { lower: 0, upper }],
            elements,
        }
    }

    pub fn rank(&self) -> usize {
        self.bounds.len()
    }

    /// `dimension` is 1-based, matching VBA's own `LBound(arr, n)`.
    pub fn lbound(&self, dimension: usize) -> Result<i64, String> {
        self.bound(dimension).map(|b| b.lower)
    }

    /// `dimension` is 1-based, matching VBA's own `UBound(arr, n)`.
    pub fn ubound(&self, dimension: usize) -> Result<i64, String> {
        self.bound(dimension).map(|b| b.upper)
    }

    fn bound(&self, dimension: usize) -> Result<&ArrayBound, String> {
        if dimension == 0 || dimension > self.bounds.len() {
            return Err("Subscript out of range".to_string());
        }
        Ok(&self.bounds[dimension - 1])
    }

    /// The flat index for `indices` (one per dimension, in `arr(i, j, ...)`
    /// order). `Err("Subscript out of range")` if the count doesn't match
    /// `rank()`, or any index falls outside its dimension's bounds — VBA's
    /// real Error 9 either way.
    pub fn linear_index(&self, indices: &[i64]) -> Result<usize, String> {
        Self::linear_index_for(&self.bounds, indices)
    }

    /// Same computation as `linear_index`, taking `bounds` directly rather
    /// than a whole `VbaArray` — lets a caller bounds-check a write against
    /// a cheap `Vec<ArrayBound>` clone instead of cloning `elements` too.
    pub fn linear_index_for(bounds: &[ArrayBound], indices: &[i64]) -> Result<usize, String> {
        if indices.len() != bounds.len() {
            return Err("Subscript out of range".to_string());
        }
        let mut idx: i64 = 0;
        for (&sub, bound) in indices.iter().zip(bounds.iter()) {
            if sub < bound.lower || sub > bound.upper {
                return Err("Subscript out of range".to_string());
            }
            idx = idx * bound.len() + (sub - bound.lower);
        }
        Ok(idx as usize)
    }

    pub fn get(&self, indices: &[i64]) -> Result<&Variant, String> {
        self.linear_index(indices).map(|i| &self.elements[i])
    }

    pub fn set(&mut self, indices: &[i64], value: Variant) -> Result<(), String> {
        let i = self.linear_index(indices)?;
        self.elements[i] = value;
        Ok(())
    }
}

pub fn serial_to_display(s: i64) -> String {
    let (y, m, d) = serial_to_ymd(s);
    format!("{:04}-{:02}-{:02}", y, m, d)
}

impl std::fmt::Display for Variant {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Variant::Integer(n) => write!(f, "{}", n),
            Variant::Float(v) => write!(f, "{}", v),
            Variant::Str(s) => write!(f, "{}", s),
            Variant::Boolean(b) => write!(f, "{}", if *b { "True" } else { "False" }),
            Variant::Date(s) => write!(f, "{}", serial_to_display(*s)),
            Variant::Error(e) => write!(f, "{}", e),
            Variant::Empty => write!(f, ""),
            // Matches what real VBA's own `Debug.Print Null` prints. Never
            // reached by `&` concatenation, which applies the documented
            // Null rules (both-Null -> Null, one-Null -> "") before ever
            // formatting an operand.
            Variant::Null => write!(f, "Null"),
            Variant::Array(a) => write!(
                f,
                "[{}]",
                a.iter()
                    .map(|v| v.to_string())
                    .collect::<Vec<_>>()
                    .join(", ")
            ),
            Variant::VbaArray(a) => write!(
                f,
                "[{}]",
                a.elements
                    .iter()
                    .map(|v| v.to_string())
                    .collect::<Vec<_>>()
                    .join(", ")
            ),
            Variant::Record(m) => {
                let mut pairs: Vec<String> =
                    m.iter().map(|(k, v)| format!("{}: {}", k, v)).collect();
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

pub fn is_leap(y: i32) -> bool {
    (y % 4 == 0 && y % 100 != 0) || y % 400 == 0
}

pub fn days_in_month(y: i32, m: u32) -> u32 {
    match m {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 => {
            if is_leap(y) {
                29
            } else {
                28
            }
        }
        _ => 0,
    }
}

pub fn serial_to_ymd(mut s: i64) -> (i32, u32, u32) {
    // Undo the Excel leap-year bug offset for dates after serial 60
    if s > 60 {
        s -= 1;
    }
    // s is now days since Jan 1 1900 (1-based)
    let mut y = 1900i32;
    loop {
        let days = if is_leap(y) { 366i64 } else { 365 };
        if s <= days {
            break;
        }
        s -= days;
        y += 1;
    }
    let mut m = 1u32;
    loop {
        let dim = days_in_month(y, m) as i64;
        if s <= dim {
            break;
        }
        s -= dim;
        m += 1;
    }
    (y, m, s as u32)
}

// ── Range address helpers ─────────────────────────────────────────────────────
// Moved from vm/mod.rs verbatim — pure string↔coordinate parsing, no
// further crate:: dependencies of its own.

fn col_letters_to_num_vm(s: &str) -> u32 {
    s.chars().fold(0u32, |acc, c| {
        acc * 26 + (c.to_ascii_uppercase() as u32 - b'A' as u32 + 1)
    })
}

pub fn parse_cell_addr(addr: &str) -> Option<(u32, u32)> {
    let addr = addr.trim().to_uppercase();
    let alpha_end = addr.find(|c: char| c.is_ascii_digit())?;
    if alpha_end == 0 {
        return None;
    }
    let col = col_letters_to_num_vm(&addr[..alpha_end]);
    let row: u32 = addr[alpha_end..].parse().ok()?;
    Some((row, col))
}

pub fn parse_range_addr(addr: &str) -> Option<((u32, u32), (u32, u32))> {
    let addr = addr.trim();
    if let Some(i) = addr.find(':') {
        Some((
            parse_cell_addr(&addr[..i])?,
            parse_cell_addr(&addr[i + 1..])?,
        ))
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

    #[test]
    fn excel_error_from_str_is_the_as_str_inverse_for_every_variant() {
        for e in [
            ExcelError::DivZero,
            ExcelError::NA,
            ExcelError::Value,
            ExcelError::Ref,
            ExcelError::Name,
            ExcelError::Num,
            ExcelError::Null,
        ] {
            assert_eq!(e.as_str().parse::<ExcelError>().as_ref(), Ok(&e));
        }
    }

    #[test]
    fn excel_error_from_str_rejects_an_unrecognized_string() {
        assert_eq!("#SPILL!".parse::<ExcelError>(), Err(()));
        assert_eq!("not an error at all".parse::<ExcelError>(), Err(()));
    }

    #[test]
    fn excel_error_biff_code_matches_the_real_oracles_own_berr_table() {
        // Confirmed live against compat/node_modules/xlsx/xlsx.js's BErr table (see
        // ExcelError::biff_code's doc comment) -- these codes are not invented here.
        assert_eq!(ExcelError::Null.biff_code(), 0x00);
        assert_eq!(ExcelError::DivZero.biff_code(), 0x07);
        assert_eq!(ExcelError::Value.biff_code(), 0x0F);
        assert_eq!(ExcelError::Ref.biff_code(), 0x17);
        assert_eq!(ExcelError::Name.biff_code(), 0x1D);
        assert_eq!(ExcelError::Num.biff_code(), 0x24);
        assert_eq!(ExcelError::NA.biff_code(), 0x2A);
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

    // ── VbaArray ─────────────────────────────────────────────────────────

    fn dim2(lo0: i64, hi0: i64, lo1: i64, hi1: i64) -> VbaArray {
        VbaArray::new_zeroed(vec![
            ArrayBound {
                lower: lo0,
                upper: hi0,
            },
            ArrayBound {
                lower: lo1,
                upper: hi1,
            },
        ])
        .unwrap()
    }

    #[test]
    fn two_distinct_second_dimension_indices_write_distinct_elements() {
        let mut a = dim2(0, 3, 0, 2);
        a.set(&[2, 0], Variant::Integer(111)).unwrap();
        a.set(&[2, 1], Variant::Integer(222)).unwrap();
        assert_eq!(*a.get(&[2, 0]).unwrap(), Variant::Integer(111));
        assert_eq!(*a.get(&[2, 1]).unwrap(), Variant::Integer(222));
    }

    #[test]
    fn lbound_ubound_report_each_dimension_independently() {
        let a = dim2(1, 3, -2, 2);
        assert_eq!(a.lbound(1), Ok(1));
        assert_eq!(a.ubound(1), Ok(3));
        assert_eq!(a.lbound(2), Ok(-2));
        assert_eq!(a.ubound(2), Ok(2));
    }

    #[test]
    fn dimension_zero_is_subscript_out_of_range() {
        let a = dim2(0, 3, 0, 2);
        assert_eq!(a.ubound(0), Err("Subscript out of range".to_string()));
        assert_eq!(a.lbound(0), Err("Subscript out of range".to_string()));
    }

    #[test]
    fn a_dimension_beyond_rank_is_subscript_out_of_range() {
        let a = dim2(0, 3, 0, 2);
        assert_eq!(a.ubound(3), Err("Subscript out of range".to_string()));
    }

    #[test]
    fn too_few_or_too_many_subscripts_are_rejected_not_silently_accepted() {
        let a = dim2(0, 3, 0, 2);
        assert!(a.get(&[1]).is_err());
        assert!(a.get(&[1, 1, 1]).is_err());
    }

    #[test]
    fn an_index_outside_its_dimensions_bounds_is_subscript_out_of_range() {
        let a = dim2(0, 3, 0, 2);
        assert_eq!(a.get(&[4, 0]), Err("Subscript out of range".to_string()));
        assert_eq!(a.get(&[0, -1]), Err("Subscript out of range".to_string()));
    }

    #[test]
    fn element_count_overflow_is_rejected_as_out_of_memory() {
        let huge = vec![
            ArrayBound {
                lower: 0,
                upper: i64::MAX / 2,
            },
            ArrayBound {
                lower: 0,
                upper: i64::MAX / 2,
            },
        ];
        assert_eq!(VbaArray::new_zeroed(huge), Err("Out of memory".to_string()));
    }

    #[test]
    fn a_practically_huge_but_non_overflowing_array_hits_the_element_cap() {
        let too_big = vec![ArrayBound {
            lower: 0,
            upper: (MAX_ARRAY_ELEMENTS as i64) + 1,
        }];
        assert_eq!(
            VbaArray::new_zeroed(too_big),
            Err("Out of memory".to_string())
        );
    }

    #[test]
    fn from_vec_is_1d_zero_based_matching_array_and_splits_shape() {
        let a = VbaArray::from_vec(vec![
            Variant::Integer(1),
            Variant::Integer(2),
            Variant::Integer(3),
        ]);
        assert_eq!(a.rank(), 1);
        assert_eq!(a.lbound(1), Ok(0));
        assert_eq!(a.ubound(1), Ok(2));
        assert_eq!(*a.get(&[0]).unwrap(), Variant::Integer(1));
        assert_eq!(*a.get(&[2]).unwrap(), Variant::Integer(3));
    }

    #[test]
    fn an_inverted_bound_is_a_legal_zero_length_dimension() {
        // `Option Base 1` meeting a bare `Dim arr(0)` — real VBA allows this
        // and treats it as a zero-length dimension, not an error.
        let a = VbaArray::new_zeroed(vec![ArrayBound { lower: 1, upper: 0 }]).unwrap();
        assert_eq!(a.elements.len(), 0);
        assert!(a.get(&[1]).is_err());
    }

    // ── CellContent ──────────────────────────────────────────────────────

    #[test]
    fn cell_content_holds_formula_and_value() {
        let c = CellContent {
            formula: Some("=A1+1".into()),
            value: Variant::Integer(2),
        };
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
