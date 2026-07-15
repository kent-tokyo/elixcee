//! Property-based workbook test runner (Milestone B5a): the `test-workbook`
//! subcommand runs a macro against a starting `.xlsx`/`.ods` workbook many
//! times with generated boundary-value inputs, checking each run for
//! panics, runtime errors, timeouts, and Excel error values in a result
//! range. Every case starts from a completely fresh `Vm` and a fresh read
//! of the workbook file — no cell, variable, MsgBox-log, or deadline state
//! survives from one case to the next.
//!
//! **Known limitation, not fixed in this phase**: `RANDARRAY`/`Rnd`'s PRNG
//! (`src/formula/eval.rs`) is a *thread-local*, not a `Vm` field, so a fresh
//! `Vm` per case does not reset it — draws continue across cases on the
//! same thread. `--seed`/`--case` replay is only guaranteed to reproduce
//! identical *input generation* (which boundary value gets written where),
//! not VBA-visible randomness for a macro that calls `RANDARRAY`/`Rnd`.
//! Neither `boundary_numeric` nor `boundary_string` (the only strategies in
//! this phase) invoke any VBA-side randomness, so this doesn't bite v1.
//!
//! The TOML fixture format is parsed by a hand-rolled, deliberately
//! minimal subset parser (`parse_fixture` below) rather than a real TOML
//! dependency — `toml` is a `[dev-dependencies]`-only crate (added for
//! `tests/blackbox.rs`), and pulling it into the release binary would
//! reverse the project's zero-new-runtime-dependency principle (the same
//! one Milestone B2 invoked to reject a TOML project manifest). This
//! parser only supports what the fixture schema below needs: flat
//! `key = value` lines and `[[inputs]]`/`[[assertions]]` array-of-tables —
//! same scope-limiting philosophy as `reader.rs`'s minimal XML parser. Any
//! construct outside that subset (inline tables, multi-line strings,
//! dotted keys, trailing junk after a value) is a hard parse error, not a
//! silent skip — a silently-misparsed fixture that still "runs" would
//! produce a confusing green result instead of a clear failure.

use crate::diagnostics::{json_string, variant_to_json};
use crate::parser::ast::Program;
use crate::vm::{CellContent, Variant, Vm, parse_sheet_range_addr};
use std::collections::HashMap;
use std::time::{Duration, Instant};

// ── Fixture schema ────────────────────────────────────────────────────────────

pub struct Fixture {
    pub name: String,
    pub workbook: String,
    pub vba_files: Vec<String>,
    pub macro_name: String,
    pub cases: u64,
    pub seed: u64,
    pub timeout_secs: u64,
    pub inputs: Vec<InputSpec>,
    pub assertions: Vec<AssertionSpec>,
}

pub struct InputSpec {
    pub range: String,
    pub strategy: String,
}

pub struct AssertionSpec {
    pub range: String,
    pub rule: String,
}

// ── Minimal TOML-subset parser ────────────────────────────────────────────────

enum TomlValue {
    Str(String),
    Int(i64),
    StrArray(Vec<String>),
}

enum Section {
    None,
    Inputs,
    Assertions,
}

pub fn parse_fixture(text: &str) -> Result<Fixture, String> {
    let mut top: HashMap<String, TomlValue> = HashMap::new();
    let mut inputs: Vec<HashMap<String, TomlValue>> = Vec::new();
    let mut assertions: Vec<HashMap<String, TomlValue>> = Vec::new();
    let mut section = Section::None;

    for (i, raw_line) in text.lines().enumerate() {
        let line_no = i + 1;
        let line = raw_line.trim();
        if line.is_empty() || line.starts_with('#') {
            continue;
        }

        if let Some(name) = line.strip_prefix("[[").and_then(|r| r.strip_suffix("]]")) {
            match name.trim() {
                "inputs" => {
                    inputs.push(HashMap::new());
                    section = Section::Inputs;
                }
                "assertions" => {
                    assertions.push(HashMap::new());
                    section = Section::Assertions;
                }
                other => return Err(format!("line {}: unknown section '[[{}]]'", line_no, other)),
            }
            continue;
        }
        if line.starts_with('[') {
            return Err(format!(
                "line {}: unsupported TOML construct (only [[inputs]]/[[assertions]] sections are supported): {}",
                line_no, line
            ));
        }

        let Some(eq_pos) = line.find('=') else {
            return Err(format!(
                "line {}: expected 'key = value', got: {}",
                line_no, line
            ));
        };
        let key = line[..eq_pos].trim();
        if key.is_empty() || key.contains('.') || key.contains(char::is_whitespace) {
            return Err(format!(
                "line {}: unsupported key syntax: '{}'",
                line_no, key
            ));
        }
        let value = parse_toml_value(line[eq_pos + 1..].trim(), line_no)?;

        match section {
            Section::None => {
                top.insert(key.to_string(), value);
            }
            Section::Inputs => {
                inputs
                    .last_mut()
                    .expect("section entered via [[inputs]]")
                    .insert(key.to_string(), value);
            }
            Section::Assertions => {
                assertions
                    .last_mut()
                    .expect("section entered via [[assertions]]")
                    .insert(key.to_string(), value);
            }
        }
    }

    let name = require_str(&top, "name", "fixture")?;
    let workbook = require_str(&top, "workbook", "fixture")?;
    let vba_files = require_str_array(&top, "vba_files", "fixture")?;
    let macro_name = require_str(&top, "macro", "fixture")?;
    let cases = require_int(&top, "cases", "fixture")?;
    let seed = require_int(&top, "seed", "fixture")?;
    let timeout_secs = optional_int(&top, "timeout_secs", "fixture", 10)?;

    if inputs.is_empty() {
        return Err("fixture: at least one [[inputs]] entry is required".to_string());
    }
    if assertions.is_empty() {
        return Err("fixture: at least one [[assertions]] entry is required".to_string());
    }

    let inputs = inputs
        .iter()
        .map(|m| {
            Ok(InputSpec {
                range: require_str(m, "range", "[[inputs]]")?,
                strategy: require_str(m, "strategy", "[[inputs]]")?,
            })
        })
        .collect::<Result<Vec<_>, String>>()?;
    let assertions = assertions
        .iter()
        .map(|m| {
            Ok(AssertionSpec {
                range: require_str(m, "range", "[[assertions]]")?,
                rule: require_str(m, "rule", "[[assertions]]")?,
            })
        })
        .collect::<Result<Vec<_>, String>>()?;

    Ok(Fixture {
        name,
        workbook,
        vba_files,
        macro_name,
        cases,
        seed,
        timeout_secs,
        inputs,
        assertions,
    })
}

fn parse_toml_value(s: &str, line_no: usize) -> Result<TomlValue, String> {
    if let Some(inner) = s.strip_prefix('"').and_then(|r| r.strip_suffix('"')) {
        return Ok(TomlValue::Str(unescape_toml_string(inner, line_no)?));
    }
    if let Some(inner) = s.strip_prefix('[').and_then(|r| r.strip_suffix(']')) {
        let mut items = Vec::new();
        for part in split_top_level_commas(inner) {
            let part = part.trim();
            if part.is_empty() {
                continue;
            }
            match parse_toml_value(part, line_no)? {
                TomlValue::Str(s) => items.push(s),
                _ => return Err(format!("line {}: array elements must be strings", line_no)),
            }
        }
        return Ok(TomlValue::StrArray(items));
    }
    if let Ok(n) = s.parse::<i64>() {
        return Ok(TomlValue::Int(n));
    }
    Err(format!(
        "line {}: unsupported value syntax: '{}'",
        line_no, s
    ))
}

fn split_top_level_commas(s: &str) -> Vec<&str> {
    let mut parts = Vec::new();
    let mut in_quote = false;
    let mut start = 0;
    for (i, c) in s.char_indices() {
        match c {
            '"' => in_quote = !in_quote,
            ',' if !in_quote => {
                parts.push(&s[start..i]);
                start = i + 1;
            }
            _ => {}
        }
    }
    parts.push(&s[start..]);
    parts
}

fn unescape_toml_string(s: &str, line_no: usize) -> Result<String, String> {
    let mut out = String::with_capacity(s.len());
    let mut chars = s.chars();
    while let Some(c) = chars.next() {
        if c != '\\' {
            out.push(c);
            continue;
        }
        match chars.next() {
            Some('"') => out.push('"'),
            Some('\\') => out.push('\\'),
            Some('n') => out.push('\n'),
            Some('t') => out.push('\t'),
            Some('r') => out.push('\r'),
            other => {
                return Err(format!(
                    "line {}: unsupported escape sequence '\\{:?}'",
                    line_no, other
                ));
            }
        }
    }
    Ok(out)
}

fn require_str(
    map: &HashMap<String, TomlValue>,
    key: &str,
    context: &str,
) -> Result<String, String> {
    match map.get(key) {
        Some(TomlValue::Str(s)) => Ok(s.clone()),
        Some(_) => Err(format!("{}: '{}' must be a string", context, key)),
        None => Err(format!("{}: missing required field '{}'", context, key)),
    }
}

fn require_str_array(
    map: &HashMap<String, TomlValue>,
    key: &str,
    context: &str,
) -> Result<Vec<String>, String> {
    match map.get(key) {
        Some(TomlValue::StrArray(v)) => Ok(v.clone()),
        Some(_) => Err(format!(
            "{}: '{}' must be an array of strings",
            context, key
        )),
        None => Err(format!("{}: missing required field '{}'", context, key)),
    }
}

fn require_int(map: &HashMap<String, TomlValue>, key: &str, context: &str) -> Result<u64, String> {
    match map.get(key) {
        Some(TomlValue::Int(n)) if *n >= 0 => Ok(*n as u64),
        Some(TomlValue::Int(_)) => Err(format!("{}: '{}' must not be negative", context, key)),
        Some(_) => Err(format!("{}: '{}' must be an integer", context, key)),
        None => Err(format!("{}: missing required field '{}'", context, key)),
    }
}

fn optional_int(
    map: &HashMap<String, TomlValue>,
    key: &str,
    context: &str,
    default: u64,
) -> Result<u64, String> {
    match map.get(key) {
        Some(_) => require_int(map, key, context),
        None => Ok(default),
    }
}

// ── Seeded case generator ─────────────────────────────────────────────────────

/// Independent from `formula/eval.rs`'s thread-local xorshift64 (same core
/// algorithm, copied) — this one is explicitly seeded so case generation is
/// fully deterministic, unlike the thread-local RNG behind `RANDARRAY`.
struct CaseRng {
    state: u64,
}

impl CaseRng {
    fn new(seed: u64) -> Self {
        CaseRng {
            state: if seed == 0 {
                0x9E37_79B9_7F4A_7C15
            } else {
                seed
            },
        }
    }

    fn next_u64(&mut self) -> u64 {
        let mut s = self.state;
        s ^= s << 13;
        s ^= s >> 7;
        s ^= s << 17;
        self.state = s;
        s
    }
}

/// Derives a per-case seed so `--case N` always reproduces the same draws
/// whether run inside the full `cases` loop or standalone.
fn case_seed(base_seed: u64, case_index: u64) -> u64 {
    base_seed
        .wrapping_add(case_index)
        .wrapping_mul(0x9E37_79B9_7F4A_7C15)
}

// ── Strategies (v1: two, splitting the roadmap's 8 boundary values along the
// numeric/string divide implied by "boundary_numeric" itself) ────────────────

fn boundary_numeric_pool() -> Vec<Variant> {
    vec![
        Variant::Empty,
        Variant::Integer(0),
        Variant::Integer(1),
        Variant::Integer(-1),
        // Chosen over i64::MAX/MIN: these sit just past VBA's classic
        // Integer/Long overflow boundaries, which is where realistic
        // spreadsheet-macro bugs actually show up.
        Variant::Integer(999_999_999),
        Variant::Integer(-999_999_999),
    ]
}

fn boundary_string_pool() -> Vec<Variant> {
    vec![
        Variant::Str(String::new()),
        Variant::Str("test".to_string()),
        Variant::Str("a".repeat(1000)),
    ]
}

fn resolve_strategy(name: &str) -> Result<Vec<Variant>, String> {
    match name {
        "boundary_numeric" => Ok(boundary_numeric_pool()),
        "boundary_string" => Ok(boundary_string_pool()),
        other => Err(format!("unknown strategy '{}'", other)),
    }
}

fn col_to_letters(mut col: u32) -> String {
    let mut bytes = Vec::new();
    while col > 0 {
        col -= 1;
        bytes.push(b'A' + (col % 26) as u8);
        col /= 26;
    }
    bytes.reverse();
    String::from_utf8(bytes).unwrap()
}

// ── Execution ─────────────────────────────────────────────────────────────────

#[derive(Clone)]
pub struct InputUsed {
    pub address: String,
    pub value: Variant,
}

pub struct FailureDetail {
    pub rule: String,
    pub address: Option<String>,
    pub actual: Option<String>,
    pub message: Option<String>,
}

pub enum FixtureResult {
    Passed {
        seed: u64,
        cases_run: u64,
    },
    Failed {
        seed: u64,
        case_index: u64,
        inputs_used: Vec<InputUsed>,
        failure: FailureDetail,
    },
}

/// Runs `fixture` against `programs` (already parsed once, outside this
/// function — the AST is immutable and safe to reuse across cases). Each
/// case: derives its seed, builds a fresh `Vm`, reloads `workbook_path`
/// from scratch, writes generated input values, runs the macro under a
/// wall-clock deadline and `catch_unwind`, then checks assertions — fully
/// independent of every other case. Stops at the first failing case
/// (fail-fast, matching `proptest`'s own convention and keeping this
/// subcommand's "exactly one JSON object per invocation" contract).
pub fn run_fixture(
    fixture: &Fixture,
    programs: &[(String, Program)],
    workbook_path: &str,
    seed_override: Option<u64>,
    case_override: Option<u64>,
) -> Result<FixtureResult, String> {
    let base_seed = seed_override.unwrap_or(fixture.seed);
    let input_pools: Vec<Vec<Variant>> = fixture
        .inputs
        .iter()
        .map(|i| resolve_strategy(&i.strategy))
        .collect::<Result<_, _>>()?;

    let case_indices: Vec<u64> = match case_override {
        Some(n) => vec![n],
        None => (0..fixture.cases).collect(),
    };
    let cases_run = case_indices.len() as u64;

    for case_index in case_indices {
        let seed = case_seed(base_seed, case_index);
        let mut rng = CaseRng::new(seed);

        let mut vm = Vm::new();
        vm.load_workbook_file(workbook_path)
            .map_err(|e| format!("failed to load workbook: {}", e))?;

        let mut inputs_used = Vec::new();
        // [[inputs]] in TOML declaration order, cells row-major within each
        // range — draw order must be pinned for --case replay to reproduce
        // the exact same values every time.
        for (spec, pool) in fixture.inputs.iter().zip(input_pools.iter()) {
            let (sheet, (r1, c1), (r2, c2)) = parse_sheet_range_addr(&spec.range, &vm.active_sheet)
                .ok_or_else(|| format!("invalid range '{}'", spec.range))?;
            vm.ensure_sheet(&sheet);
            let prev_active = vm.active_sheet.clone();
            vm.active_sheet = sheet.clone();
            for r in r1..=r2 {
                for c in c1..=c2 {
                    let value = pool[(rng.next_u64() as usize) % pool.len()].clone();
                    let address = format!("{}!{}{}", sheet, col_to_letters(c), r);
                    vm.cells_mut().insert(
                        (r, c),
                        CellContent {
                            formula: None,
                            value: value.clone(),
                        },
                    );
                    inputs_used.push(InputUsed { address, value });
                }
            }
            vm.active_sheet = prev_active;
        }

        vm.deadline = Some(Instant::now() + Duration::from_secs(fixture.timeout_secs));

        let run_result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            if programs.len() == 1 {
                vm.run_sub(&programs[0].1, &fixture.macro_name)
            } else {
                vm.run_sub_multi(programs, &fixture.macro_name)
            }
        }));

        match run_result {
            Err(_panic) => {
                return Ok(FixtureResult::Failed {
                    seed: base_seed,
                    case_index,
                    inputs_used,
                    failure: FailureDetail {
                        rule: "no_panic".to_string(),
                        address: None,
                        actual: None,
                        message: Some("macro execution panicked".to_string()),
                    },
                });
            }
            Ok(Err(e)) => {
                let rule = if e.starts_with("TIMEOUT:") {
                    "no_timeout"
                } else {
                    "no_runtime_error"
                };
                return Ok(FixtureResult::Failed {
                    seed: base_seed,
                    case_index,
                    inputs_used,
                    failure: FailureDetail {
                        rule: rule.to_string(),
                        address: None,
                        actual: None,
                        message: Some(e),
                    },
                });
            }
            Ok(Ok(())) => {}
        }

        for spec in &fixture.assertions {
            match spec.rule.as_str() {
                "no_excel_errors" => {
                    let (sheet, (r1, c1), (r2, c2)) =
                        parse_sheet_range_addr(&spec.range, &vm.active_sheet)
                            .ok_or_else(|| format!("invalid range '{}'", spec.range))?;
                    // A missing sheet is a fixture/config problem (typo, or
                    // the macro genuinely didn't produce the expected
                    // sheet) — hard error rather than silently treating
                    // "sheet doesn't exist" as "no errors found".
                    let cells = vm.get_sheet_cells(&sheet).ok_or_else(|| {
                        format!(
                            "assertion range '{}': sheet '{}' does not exist",
                            spec.range, sheet
                        )
                    })?;
                    for r in r1..=r2 {
                        for c in c1..=c2 {
                            if let Some(content) = cells.get(&(r, c))
                                && let Variant::Error(e) = &content.value
                            {
                                return Ok(FixtureResult::Failed {
                                    seed: base_seed,
                                    case_index,
                                    inputs_used,
                                    failure: FailureDetail {
                                        rule: "no_excel_errors".to_string(),
                                        address: Some(format!(
                                            "{}!{}{}",
                                            sheet,
                                            col_to_letters(c),
                                            r
                                        )),
                                        actual: Some(e.as_str().to_string()),
                                        message: None,
                                    },
                                });
                            }
                        }
                    }
                }
                other => return Err(format!("unknown assertion rule '{}'", other)),
            }
        }
    }

    Ok(FixtureResult::Passed {
        seed: base_seed,
        cases_run,
    })
}

// ── Output ────────────────────────────────────────────────────────────────────

pub fn to_json(result: &FixtureResult) -> String {
    match result {
        FixtureResult::Passed { seed, cases_run } => {
            format!(
                "{{\"schema_version\":1,\"ok\":true,\"seed\":{},\"cases_run\":{}}}",
                seed, cases_run
            )
        }
        FixtureResult::Failed {
            seed,
            case_index,
            inputs_used,
            failure,
        } => {
            let inputs_json: Vec<String> = inputs_used
                .iter()
                .map(|iu| {
                    format!(
                        "{{\"address\":{},\"value\":{}}}",
                        json_string(&iu.address),
                        variant_to_json(&iu.value)
                    )
                })
                .collect();
            let mut fields = vec![format!("\"rule\":{}", json_string(&failure.rule))];
            if let Some(a) = &failure.address {
                fields.push(format!("\"address\":{}", json_string(a)));
            }
            if let Some(a) = &failure.actual {
                fields.push(format!("\"actual\":{}", json_string(a)));
            }
            if let Some(m) = &failure.message {
                fields.push(format!("\"message\":{}", json_string(m)));
            }
            format!(
                "{{\"schema_version\":1,\"ok\":false,\"seed\":{},\"case_index\":{},\"inputs\":[{}],\"failure\":{{{}}}}}",
                seed,
                case_index,
                inputs_json.join(","),
                fields.join(",")
            )
        }
    }
}

fn display_variant(v: &Variant) -> String {
    match v {
        Variant::Empty => "(empty)".to_string(),
        other => other.to_string(),
    }
}

pub fn to_plain_text(result: &FixtureResult) -> String {
    match result {
        FixtureResult::Passed { seed, cases_run } => {
            format!("ok: {} case(s) passed (seed {})", cases_run, seed)
        }
        FixtureResult::Failed {
            seed,
            case_index,
            inputs_used,
            failure,
        } => {
            let mut line = format!(
                "FAIL: case {} (seed {}) - {}",
                case_index, seed, failure.rule
            );
            if let Some(a) = &failure.address {
                line.push_str(&format!(" at {}", a));
            }
            if let Some(a) = &failure.actual {
                line.push_str(&format!(": {}", a));
            }
            if let Some(m) = &failure.message {
                line.push_str(&format!(": {}", m));
            }
            for iu in inputs_used {
                line.push_str(&format!(
                    "\n  {} = {}",
                    iu.address,
                    display_variant(&iu.value)
                ));
            }
            line
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser;

    const EXAMPLE_FIXTURE: &str = r#"
name = "order calculation"
workbook = "fixtures/orders.xlsx"
vba_files = ["Main.bas"]
macro = "Main.Process"
cases = 100
seed = 42

[[inputs]]
range = "Input!B2:B10"
strategy = "boundary_numeric"

[[assertions]]
range = "Result!A1:F100"
rule = "no_excel_errors"
"#;

    #[test]
    fn parse_fixture_round_trips_the_example_schema() {
        let f = parse_fixture(EXAMPLE_FIXTURE).unwrap();
        assert_eq!(f.name, "order calculation");
        assert_eq!(f.workbook, "fixtures/orders.xlsx");
        assert_eq!(f.vba_files, vec!["Main.bas".to_string()]);
        assert_eq!(f.macro_name, "Main.Process");
        assert_eq!(f.cases, 100);
        assert_eq!(f.seed, 42);
        assert_eq!(f.timeout_secs, 10); // default, not specified in the example
        assert_eq!(f.inputs.len(), 1);
        assert_eq!(f.inputs[0].range, "Input!B2:B10");
        assert_eq!(f.inputs[0].strategy, "boundary_numeric");
        assert_eq!(f.assertions.len(), 1);
        assert_eq!(f.assertions[0].range, "Result!A1:F100");
        assert_eq!(f.assertions[0].rule, "no_excel_errors");
    }

    #[test]
    fn parse_fixture_honors_an_explicit_timeout_secs() {
        // Inserted before any [[section]] — a key = value line after a
        // [[inputs]]/[[assertions]] header belongs to that table entry,
        // not to the top-level fixture (matches real TOML semantics).
        let text = EXAMPLE_FIXTURE.replacen("seed = 42\n", "seed = 42\ntimeout_secs = 5\n", 1);
        assert_eq!(parse_fixture(&text).unwrap().timeout_secs, 5);
    }

    #[test]
    fn parse_fixture_rejects_a_dotted_key() {
        let text = "name = \"x\"\nworkbook = \"w.xlsx\"\nvba_files = [\"a.bas\"]\nmacro = \"Main\"\ncases = 1\nseed = 1\na.b = 1\n[[inputs]]\nrange = \"A1\"\nstrategy = \"boundary_numeric\"\n[[assertions]]\nrange = \"A1\"\nrule = \"no_excel_errors\"\n";
        assert!(parse_fixture(text).is_err());
    }

    #[test]
    fn parse_fixture_rejects_an_inline_table() {
        let text = "x = { a = 1 }\n";
        assert!(parse_fixture(text).is_err());
    }

    #[test]
    fn parse_fixture_rejects_an_unknown_section() {
        let text = "[[bogus]]\nfoo = \"bar\"\n";
        assert!(parse_fixture(text).is_err());
    }

    #[test]
    fn parse_fixture_requires_at_least_one_input_and_assertion() {
        let no_inputs = "name = \"x\"\nworkbook = \"w.xlsx\"\nvba_files = [\"a.bas\"]\nmacro = \"Main\"\ncases = 1\nseed = 1\n[[assertions]]\nrange = \"A1\"\nrule = \"no_excel_errors\"\n";
        assert!(parse_fixture(no_inputs).is_err());
    }

    #[test]
    fn case_rng_is_deterministic_for_the_same_seed() {
        let mut a = CaseRng::new(42);
        let mut b = CaseRng::new(42);
        let seq_a: Vec<u64> = (0..10).map(|_| a.next_u64()).collect();
        let seq_b: Vec<u64> = (0..10).map(|_| b.next_u64()).collect();
        assert_eq!(seq_a, seq_b);
    }

    #[test]
    fn case_rng_differs_for_different_seeds() {
        let mut a = CaseRng::new(1);
        let mut b = CaseRng::new(2);
        assert_ne!(a.next_u64(), b.next_u64());
    }

    #[test]
    fn case_seed_is_deterministic_and_case_specific() {
        assert_eq!(case_seed(42, 17), case_seed(42, 17));
        assert_ne!(case_seed(42, 17), case_seed(42, 18));
    }

    #[test]
    fn boundary_numeric_pool_has_the_documented_values() {
        let pool = boundary_numeric_pool();
        assert!(pool.contains(&Variant::Empty));
        assert!(pool.contains(&Variant::Integer(0)));
        assert!(pool.contains(&Variant::Integer(1)));
        assert!(pool.contains(&Variant::Integer(-1)));
        assert!(pool.contains(&Variant::Integer(999_999_999)));
        assert!(pool.contains(&Variant::Integer(-999_999_999)));
    }

    #[test]
    fn boundary_string_pool_has_the_documented_values() {
        let pool = boundary_string_pool();
        assert!(pool.contains(&Variant::Str(String::new())));
        assert!(pool.contains(&Variant::Str("test".to_string())));
        assert!(
            pool.iter()
                .any(|v| matches!(v, Variant::Str(s) if s.len() == 1000))
        );
    }

    #[test]
    fn resolve_strategy_rejects_an_unknown_name() {
        assert!(resolve_strategy("bogus").is_err());
    }

    fn build_workbook_fixture(path: &str) {
        let vm = Vm::new();
        crate::save_workbook(&vm, path).unwrap();
    }

    /// `Main` writes `=100/B2` as a real formula (not raw VBA division,
    /// which raises a hard VBA runtime error rather than producing a
    /// stored Excel error value) into A1 — a deterministic way to make
    /// `no_excel_errors` fire exactly when the drawn input is `0` (one of
    /// `boundary_numeric`'s pool values), so the test doesn't depend on
    /// guessing which case number draws which value. Both cells live on
    /// the same (default) sheet to avoid needing cross-sheet `Range()`
    /// addressing inside the VBA macro itself — the fixture's own
    /// `Sheet!Range` parsing (exercised via `Input!B2`/`Result!A1` below)
    /// is a separate, already-covered code path.
    const DIVIDE_MACRO: &str = "Sub Main()\n    Range(\"A1\").Formula = \"=100/B2\"\nEnd Sub\n";

    #[test]
    fn run_fixture_reports_no_excel_errors_with_case_index_and_replays_identically() {
        let path = std::env::temp_dir().join("elixcee_testworkbook_divide.xlsx");
        build_workbook_fixture(path.to_str().unwrap());
        let program = parser::parse(DIVIDE_MACRO).unwrap();
        let programs = vec![("main".to_string(), program)];

        let fixture = Fixture {
            name: "divide".to_string(),
            workbook: path.to_str().unwrap().to_string(),
            vba_files: vec![],
            macro_name: "Main".to_string(),
            cases: 50,
            seed: 7,
            timeout_secs: 5,
            inputs: vec![InputSpec {
                range: "Sheet1!B2".to_string(),
                strategy: "boundary_numeric".to_string(),
            }],
            assertions: vec![AssertionSpec {
                range: "Sheet1!A1".to_string(),
                rule: "no_excel_errors".to_string(),
            }],
        };

        let result = run_fixture(&fixture, &programs, &fixture.workbook, None, None).unwrap();
        let (seed, case_index, inputs_used) = match &result {
            FixtureResult::Failed {
                seed,
                case_index,
                inputs_used,
                failure,
            } => {
                assert_eq!(failure.rule, "no_excel_errors");
                assert_eq!(failure.address.as_deref(), Some("sheet1!A1"));
                assert_eq!(failure.actual.as_deref(), Some("#DIV/0!"));
                (*seed, *case_index, inputs_used.clone())
            }
            FixtureResult::Passed { .. } => {
                panic!("expected a division-by-zero failure across 50 cases")
            }
        };

        // Replaying the exact failing case in isolation must reproduce the
        // identical failure and identical drawn input.
        let replay = run_fixture(
            &fixture,
            &programs,
            &fixture.workbook,
            Some(seed),
            Some(case_index),
        )
        .unwrap();
        match replay {
            FixtureResult::Failed {
                case_index: replay_case,
                inputs_used: replay_inputs,
                failure,
                ..
            } => {
                assert_eq!(replay_case, case_index);
                assert_eq!(failure.actual.as_deref(), Some("#DIV/0!"));
                assert_eq!(replay_inputs[0].value, inputs_used[0].value);
                assert_eq!(replay_inputs[0].address, inputs_used[0].address);
            }
            FixtureResult::Passed { .. } => {
                panic!("replay of a failing case must fail identically")
            }
        }
    }

    #[test]
    fn run_fixture_passes_when_the_macro_never_divides_by_a_drawn_zero() {
        let path = std::env::temp_dir().join("elixcee_testworkbook_noop.xlsx");
        build_workbook_fixture(path.to_str().unwrap());
        let program = parser::parse("Sub Main()\n    Cells(1, 1).Value = 1\nEnd Sub\n").unwrap();
        let programs = vec![("main".to_string(), program)];

        let fixture = Fixture {
            name: "noop".to_string(),
            workbook: path.to_str().unwrap().to_string(),
            vba_files: vec![],
            macro_name: "Main".to_string(),
            cases: 20,
            seed: 1,
            timeout_secs: 5,
            inputs: vec![InputSpec {
                range: "Sheet1!B2".to_string(),
                strategy: "boundary_numeric".to_string(),
            }],
            assertions: vec![AssertionSpec {
                range: "Sheet1!A1".to_string(),
                rule: "no_excel_errors".to_string(),
            }],
        };

        let result = run_fixture(&fixture, &programs, &fixture.workbook, None, None).unwrap();
        match result {
            FixtureResult::Passed { cases_run, .. } => assert_eq!(cases_run, 20),
            FixtureResult::Failed { .. } => {
                panic!("a macro that never divides should never fail no_excel_errors")
            }
        }
    }

    #[test]
    fn to_json_success_shape() {
        let json = to_json(&FixtureResult::Passed {
            seed: 42,
            cases_run: 100,
        });
        assert!(json.contains("\"ok\":true"));
        assert!(json.contains("\"seed\":42"));
        assert!(json.contains("\"cases_run\":100"));
    }

    #[test]
    fn to_json_failure_shape_matches_the_documented_contract() {
        let result = FixtureResult::Failed {
            seed: 42,
            case_index: 17,
            inputs_used: vec![InputUsed {
                address: "Input!B2".to_string(),
                value: Variant::Integer(-1),
            }],
            failure: FailureDetail {
                rule: "no_excel_errors".to_string(),
                address: Some("Result!C8".to_string()),
                actual: Some("#DIV/0!".to_string()),
                message: None,
            },
        };
        let json = to_json(&result);
        assert!(json.contains("\"ok\":false"));
        assert!(json.contains("\"case_index\":17"));
        assert!(json.contains("\"address\":\"Input!B2\""));
        assert!(json.contains("\"value\":-1"));
        assert!(json.contains("\"rule\":\"no_excel_errors\""));
        assert!(json.contains("\"address\":\"Result!C8\""));
        assert!(json.contains("\"actual\":\"#DIV/0!\""));
    }
}
