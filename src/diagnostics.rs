//! Structured error classification and hand-rolled JSON output for the CLI's
//! `--json` agent contract. No `serde` dependency — the output shapes are
//! flat and fixed, so a small escaper is all that's needed.
//!
//! Classification happens here, at the CLI boundary, by pattern-matching the
//! existing `Result<_, String>` error text produced by the parser/VM/reader.
//! This keeps every other module's error type untouched.

use crate::parser::ast::SourceSpan;
use crate::vm::Variant;

/// Where a `SourceSpan` (char offset) lands in a source file — 1-based line
/// and column, matching editor conventions.
pub struct SourceLocation {
    pub file: String,
    pub line: u32,
    pub column: u32,
}

/// Convert a char-offset span into a 1-based line/column by scanning
/// `source` once. A CLI run reports at most one error, so a single O(n)
/// scan is simplest and plenty fast — a precomputed line-offset index would
/// only pay off for batch lookups across many diagnostics at once (e.g. a
/// future `check` command that reports every issue in a file at once).
pub fn locate(source: &str, file: &str, span: SourceSpan) -> SourceLocation {
    let mut line = 1u32;
    let mut column = 1u32;
    for c in source.chars().take(span.start as usize) {
        if c == '\n' {
            line += 1;
            column = 1;
        } else {
            column += 1;
        }
    }
    SourceLocation {
        file: file.to_string(),
        line,
        column,
    }
}

/// Escape a string for embedding inside a JSON string literal (without the
/// surrounding quotes).
pub fn json_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out
}

/// A quoted, escaped JSON string literal.
pub fn json_string(s: &str) -> String {
    format!("\"{}\"", json_escape(s))
}

/// Render a `Variant` as a JSON value (number/bool/string/null — matches the
/// display conventions the plain-text CLI already uses for Array/Record).
pub fn variant_to_json(v: &Variant) -> String {
    match v {
        Variant::Integer(n) => n.to_string(),
        Variant::Float(f) if f.is_finite() => f.to_string(),
        // NaN/Infinity aren't valid JSON number literals — fall back to a
        // quoted label rather than emitting invalid JSON.
        Variant::Float(f) if f.is_nan() => json_string("NaN"),
        Variant::Float(f) if *f > 0.0 => json_string("Infinity"),
        Variant::Float(_) => json_string("-Infinity"),
        Variant::Str(s) => json_string(s),
        Variant::Boolean(b) => {
            if *b {
                "true".into()
            } else {
                "false".into()
            }
        }
        Variant::Date(s) => json_string(&crate::vm::serial_to_display(*s)),
        Variant::Error(e) => json_string(e.as_str()),
        Variant::Empty => "null".into(),
        Variant::Array(_) => json_string("[array]"),
        Variant::Record(_) => json_string("[record]"),
    }
}

/// A structured, machine-readable error for the `--json` CLI contract.
///
/// `message` is always the raw underlying error text (no CLI-added prefix
/// like "parse error: ") — `code`/`kind` already convey the category.
pub struct ElixceeError {
    pub code: &'static str,
    pub kind: &'static str,
    pub message: String,
    /// Where in the source this happened — `None` for failures that occur
    /// before/outside macro execution (io errors, `--sheet` setup errors,
    /// or a runtime error that somehow occurs before any statement runs).
    pub location: Option<SourceLocation>,
}

impl ElixceeError {
    /// File read/write failures (reading the VBA source, `--file`, `--output`).
    pub fn io_error(message: String) -> Self {
        ElixceeError {
            code: "E3001",
            kind: "io_error",
            message,
            location: None,
        }
    }

    /// `parser::parse` failures.
    pub fn parse_error(message: String) -> Self {
        ElixceeError {
            code: "E2001",
            kind: "parse_error",
            message,
            location: None,
        }
    }

    /// Pre-execution `--sheet` resolution failures (distinct from a
    /// `Sheets("X")` reference failing *during* macro execution).
    pub fn sheet_setup_error(message: String) -> Self {
        ElixceeError {
            code: "E3002",
            kind: "sheet_setup_error",
            message,
            location: None,
        }
    }

    /// `Vm::run_sub` failures — sub-classified since one call can fail for
    /// several distinct reasons.
    pub fn runtime_error(message: String) -> Self {
        let (code, kind) = classify_runtime_error(&message);
        ElixceeError {
            code,
            kind,
            message,
            location: None,
        }
    }

    /// Attach a source location (or clear it — pass `None` if the caller
    /// couldn't resolve one, e.g. no statement had executed yet).
    pub fn with_location(mut self, location: Option<SourceLocation>) -> Self {
        self.location = location;
        self
    }

    /// `messages` carries any MsgBox text recorded before this failure (e.g.
    /// a macro that shows progress via MsgBox and then hits a runtime
    /// error) — pass `&[]` for failures that happen before the macro starts
    /// running (nothing could have fired yet).
    pub fn to_json(&self, messages: &[String]) -> String {
        let messages_json = format!(
            "[{}]",
            messages
                .iter()
                .map(|m| json_string(m))
                .collect::<Vec<_>>()
                .join(",")
        );
        let location_json = match &self.location {
            Some(loc) => format!(
                "{{\"file\":{},\"line\":{},\"column\":{}}}",
                json_string(&loc.file),
                loc.line,
                loc.column,
            ),
            None => "null".to_string(),
        };
        format!(
            "{{\"schema_version\":1,\"ok\":false,\"error\":{{\"code\":\"{}\",\"kind\":\"{}\",\"message\":{},\"location\":{}}},\"messages\":{}}}",
            self.code,
            self.kind,
            json_string(&self.message),
            location_json,
            messages_json,
        )
    }
}

fn classify_runtime_error(msg: &str) -> (&'static str, &'static str) {
    if msg.starts_with("Undefined variable: '") {
        return ("E1001", "undefined_variable");
    }
    if msg.starts_with("Sub/Function '")
        || msg.starts_with("Unknown VBA function: '")
        || (msg.starts_with("Sub '") && msg.ends_with("' not found"))
    {
        return ("E1002", "undefined_sub_or_function");
    }
    if msg.starts_with("Sheet '") && msg.ends_with("' not found") {
        return ("E1003", "sheet_not_found");
    }
    if msg.starts_with("MsgBox: ") {
        return ("E1004", "msgbox_blocked");
    }
    ("E1099", "runtime_error")
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser;
    use crate::vm::Vm;

    fn run_err(src: &str, sub: &str) -> String {
        let prog = parser::parse(src).expect("should parse");
        let mut vm = Vm::new();
        vm.run_sub(&prog, sub).expect_err("should fail at runtime")
    }

    #[test]
    fn classifies_undefined_variable() {
        let msg = run_err("Sub Main()\n    x = totla + 1\nEnd Sub\n", "Main");
        let err = ElixceeError::runtime_error(msg);
        assert_eq!(err.code, "E1001");
        assert_eq!(err.kind, "undefined_variable");
    }

    #[test]
    fn classifies_missing_entrypoint_sub() {
        let msg = run_err("Sub Main()\n    x = 1\nEnd Sub\n", "DoesNotExist");
        let err = ElixceeError::runtime_error(msg);
        assert_eq!(err.code, "E1002");
        assert_eq!(err.kind, "undefined_sub_or_function");
    }

    #[test]
    fn classifies_unknown_function_call() {
        let msg = run_err(
            "Sub Main()\n    x = TotallyUnknownFunc(1)\nEnd Sub\n",
            "Main",
        );
        let err = ElixceeError::runtime_error(msg);
        assert_eq!(err.code, "E1002");
        assert_eq!(err.kind, "undefined_sub_or_function");
    }

    #[test]
    fn classifies_msgbox_blocked() {
        let prog = parser::parse("Sub Main()\n    MsgBox \"hi\"\nEnd Sub\n").unwrap();
        let mut vm = Vm::new();
        vm.error_on_msgbox = true;
        let msg = vm.run_sub(&prog, "Main").expect_err("should fail");
        let err = ElixceeError::runtime_error(msg);
        assert_eq!(err.code, "E1004");
        assert_eq!(err.kind, "msgbox_blocked");
    }

    #[test]
    fn classifies_sheet_not_found_by_construction() {
        // Not reachable via run_sub today (Sheets("X") auto-creates), but
        // set_active_sheet (used by main.rs for --sheet) produces this text.
        let (code, kind) = classify_runtime_error("Sheet 'Ghost' not found");
        assert_eq!(code, "E1003");
        assert_eq!(kind, "sheet_not_found");
    }

    #[test]
    fn unclassified_error_falls_back() {
        let (code, kind) = classify_runtime_error("something else entirely");
        assert_eq!(code, "E1099");
        assert_eq!(kind, "runtime_error");
    }

    #[test]
    fn locate_finds_line_and_column() {
        let src = "Sub Main()\n    x = totla\nEnd Sub\n";
        // "    x = totla" — 8 chars ("    x = ") before "totla" -> column 9.
        let offset = src.find("totla").unwrap() as u32;
        let loc = locate(
            src,
            "Main.bas",
            SourceSpan {
                start: offset,
                end: offset + 5,
            },
        );
        assert_eq!(loc.file, "Main.bas");
        assert_eq!(loc.line, 2);
        assert_eq!(loc.column, 9);
    }

    #[test]
    fn locate_handles_crlf_without_double_counting_lines() {
        let src = "Sub Main()\r\n    x = totla\r\nEnd Sub\r\n";
        let offset = src.find("totla").unwrap() as u32;
        let loc = locate(
            src,
            "Main.bas",
            SourceSpan {
                start: offset,
                end: offset + 5,
            },
        );
        assert_eq!(loc.line, 2);
    }

    #[test]
    fn locate_handles_first_line() {
        let src = "x = totla\n";
        let offset = src.find("totla").unwrap() as u32;
        let loc = locate(
            src,
            "Main.bas",
            SourceSpan {
                start: offset,
                end: offset + 5,
            },
        );
        assert_eq!(loc.line, 1);
        assert_eq!(loc.column, 5);
    }

    #[test]
    fn escapes_quotes_and_backslashes() {
        // Windows paths are backslash-heavy — must round-trip safely.
        assert_eq!(
            json_escape(r#"C:\data\"file".xlsx"#),
            r#"C:\\data\\\"file\".xlsx"#
        );
    }

    #[test]
    fn escapes_control_characters() {
        assert_eq!(json_escape("a\nb\tc"), "a\\nb\\tc");
    }

    #[test]
    fn variant_to_json_guards_non_finite_floats() {
        // NaN/Infinity aren't valid JSON number literals. No currently-known
        // VBA/formula path leaves a non-finite float in a cell (func_sqrt and
        // func_norm_inv both guard their inputs and return an Excel error
        // value instead), but this is cheap insurance against a future
        // formula function that misses that guard.
        assert_eq!(variant_to_json(&Variant::Float(f64::NAN)), "\"NaN\"");
        assert_eq!(
            variant_to_json(&Variant::Float(f64::INFINITY)),
            "\"Infinity\""
        );
        assert_eq!(
            variant_to_json(&Variant::Float(f64::NEG_INFINITY)),
            "\"-Infinity\""
        );
        assert_eq!(variant_to_json(&Variant::Float(1.5)), "1.5");
    }
}
