//! Structured error classification and hand-rolled JSON output for the CLI's
//! `--json` agent contract. No `serde` dependency — the output shapes are
//! flat and fixed, so a small escaper is all that's needed.
//!
//! Classification happens here, at the CLI boundary, by pattern-matching the
//! existing `Result<_, String>` error text produced by the parser/VM/reader.
//! This keeps every other module's error type untouched.

use crate::parser::ast::SourceSpan;
use crate::vm::{TraceEvent, Variant};

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
        // VBA's `Null` serializes as JSON null, same as `Empty` — neither
        // has a representable cell value, and adding a new `--json` shape
        // for one of them would be a contract change (see
        // docs/agent-contract.md). The Empty-vs-Null distinction is a
        // VBA-language one, observable via `IsNull`/`TypeName`, not a
        // wire-format one.
        Variant::Empty | Variant::Null => "null".into(),
        Variant::Array(_) => json_string("[array]"),
        Variant::VbaArray(_) => json_string("[array]"),
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
    /// VBA `Err` properties captured for an uncaught runtime failure.
    pub error_evidence: Option<crate::vm::ErrorEvidence>,
    /// Optional redaction-safe execution events. Empty unless tracing was
    /// explicitly enabled for the CLI run.
    pub trace: Vec<TraceEvent>,
}

impl ElixceeError {
    /// File read/write failures (reading the VBA source, `--file`, `--output`).
    pub fn io_error(message: String) -> Self {
        ElixceeError {
            code: "E3001",
            kind: "io_error",
            message,
            location: None,
            error_evidence: None,
            trace: vec![],
        }
    }

    /// `parser::parse` failures.
    pub fn parse_error(message: String) -> Self {
        ElixceeError {
            code: "E2001",
            kind: "parse_error",
            message,
            location: None,
            error_evidence: None,
            trace: vec![],
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
            error_evidence: None,
            trace: vec![],
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
            error_evidence: None,
            trace: vec![],
        }
    }

    /// Build a runtime diagnostic from the VM's structured failure category.
    /// The string classifier remains a fallback for failures raised before the
    /// execution boundary or by older callers.
    pub fn runtime_error_with_kind(
        message: String,
        failure: Option<crate::vm::RuntimeFailureKind>,
    ) -> Self {
        let (code, kind) = failure
            .map(runtime_failure_fields)
            .unwrap_or_else(|| classify_runtime_error(&message));
        ElixceeError {
            code,
            kind,
            message,
            location: None,
            error_evidence: None,
            trace: vec![],
        }
    }

    /// Attach a source location (or clear it — pass `None` if the caller
    /// couldn't resolve one, e.g. no statement had executed yet).
    pub fn with_location(mut self, location: Option<SourceLocation>) -> Self {
        self.location = location;
        self
    }

    /// Attach structured VBA error properties without changing the stable
    /// human-readable message or the legacy JSON shape when absent.
    pub fn with_error_evidence(mut self, error_evidence: Option<crate::vm::ErrorEvidence>) -> Self {
        self.error_evidence = error_evidence;
        self
    }

    pub fn with_trace(mut self, trace: Vec<TraceEvent>) -> Self {
        self.trace = trace;
        self
    }

    /// Return the stable terminal category used by the run-mode JSON contract.
    /// The detailed `error.kind` remains available for diagnostics; this field
    /// lets batch consumers classify completion without parsing messages.
    pub fn termination_class(&self) -> &'static str {
        if self.message.starts_with("TIMEOUT:") {
            "timeout"
        } else if self.message.starts_with("CANCELED:") {
            "canceled"
        } else if matches!(
            self.kind,
            "security_blocked_external_effect" | "msgbox_blocked"
        ) {
            "policy_blocked"
        } else if self.kind == "parse_error" {
            "parse_error"
        } else if self.kind == "io_error" {
            "io_error"
        } else if self.kind == "sheet_setup_error" {
            "setup_error"
        } else {
            "runtime_error"
        }
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
        let evidence_field = self
            .error_evidence
            .as_ref()
            .map(|evidence| {
                format!(
                    ",\"error_evidence\":{{\"number\":{},\"description\":{},\"source\":{},\"help_file\":{},\"help_context\":{}}}",
                    evidence.number,
                    json_string(&evidence.description),
                    json_string(&evidence.source),
                    json_string(&evidence.help_file),
                    evidence.help_context,
                )
            })
            .unwrap_or_default();
        let trace_field = if self.trace.is_empty() {
            String::new()
        } else {
            format!(
                ",\"trace\":[{}]",
                self.trace
                    .iter()
                    .map(trace_event_json)
                    .collect::<Vec<_>>()
                    .join(",")
            )
        };
        format!(
            "{{\"schema_version\":1,\"ok\":false,\"termination_class\":\"{}\",\"error\":{{\"code\":\"{}\",\"kind\":\"{}\",\"message\":{},\"location\":{}}}{},\"messages\":{}{}}}",
            self.termination_class(),
            self.code,
            self.kind,
            json_string(&self.message),
            location_json,
            evidence_field,
            messages_json,
            trace_field,
        )
    }
}

fn trace_event_json(event: &TraceEvent) -> String {
    let span = event
        .span
        .map(|span| format!("{{\"start\":{},\"end\":{}}}", span.start, span.end))
        .unwrap_or_else(|| "null".to_string());
    format!(
        "{{\"sequence\":{},\"execution_id\":{},\"source_hash\":{},\"kind\":{},\"procedure\":{},\"span\":{},\"detail\":{}}}",
        event.sequence,
        json_string(&event.execution_id),
        json_string(&event.source_hash),
        json_string(&event.kind),
        event
            .procedure
            .as_deref()
            .map(json_string)
            .unwrap_or_else(|| "null".to_string()),
        span,
        json_string(&event.detail),
    )
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
    // Real VBA's error 91, raised by the VM through the single
    // `vm::OBJECT_NOT_SET` constant — an exact match, not a prefix, because
    // that constant is the whole message and nothing else produces it.
    if msg == crate::vm::OBJECT_NOT_SET {
        return ("E1007", "object_variable_not_set");
    }
    ("E1099", "runtime_error")
}

fn runtime_failure_fields(failure: crate::vm::RuntimeFailureKind) -> (&'static str, &'static str) {
    match failure {
        crate::vm::RuntimeFailureKind::UndefinedVariable => ("E1001", "undefined_variable"),
        crate::vm::RuntimeFailureKind::UndefinedSubOrFunction => {
            ("E1002", "undefined_sub_or_function")
        }
        crate::vm::RuntimeFailureKind::SheetNotFound => ("E1003", "sheet_not_found"),
        crate::vm::RuntimeFailureKind::MsgBoxBlocked => ("E1004", "msgbox_blocked"),
        crate::vm::RuntimeFailureKind::ObjectVariableNotSet => ("E1007", "object_variable_not_set"),
        crate::vm::RuntimeFailureKind::SecurityBlockedExternalEffect => {
            ("E1011", "security_blocked_external_effect")
        }
        crate::vm::RuntimeFailureKind::Generic => ("E1099", "runtime_error"),
    }
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
    fn structured_runtime_kind_is_used_without_message_reclassification() {
        let err = ElixceeError::runtime_error_with_kind(
            "implementation-specific text".to_string(),
            Some(crate::vm::RuntimeFailureKind::SecurityBlockedExternalEffect),
        );
        assert_eq!(err.code, "E1011");
        assert_eq!(err.kind, "security_blocked_external_effect");
    }

    #[test]
    fn json_includes_error_evidence_only_when_attached() {
        let plain = ElixceeError::runtime_error("boom".to_string());
        assert!(!plain.to_json(&[]).contains("error_evidence"));
        let enriched = plain.with_error_evidence(Some(crate::vm::ErrorEvidence {
            number: 513,
            description: "boom".to_string(),
            source: "Module1".to_string(),
            help_file: "help.chm".to_string(),
            help_context: 42,
        }));
        assert!(enriched.to_json(&[]).contains(
            "\"error_evidence\":{\"number\":513,\"description\":\"boom\",\"source\":\"Module1\",\"help_file\":\"help.chm\",\"help_context\":42}"
        ));
    }

    #[test]
    fn json_includes_trace_on_runtime_errors_only_when_enabled() {
        let mut error = ElixceeError::runtime_error("boom".to_string());
        assert!(!error.to_json(&[]).contains("\"trace\""));
        error.trace = vec![TraceEvent {
            sequence: 0,
            execution_id: "job-7".to_string(),
            source_hash: "fnv1a64:abc".to_string(),
            kind: "failure".to_string(),
            procedure: Some("Main".to_string()),
            span: None,
            detail: "runtime failure".to_string(),
        }];
        let json = error.to_json(&[]);
        assert!(json.contains("\"trace\":[{\"sequence\":0,\"execution_id\":\"job-7\""));
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

    #[test]
    fn termination_class_separates_timeout_cancel_and_policy_failures() {
        assert_eq!(
            ElixceeError::runtime_error("TIMEOUT: deadline".to_string()).termination_class(),
            "timeout"
        );
        assert_eq!(
            ElixceeError::runtime_error("CANCELED: host request".to_string()).termination_class(),
            "canceled"
        );
        assert_eq!(
            ElixceeError::runtime_error_with_kind(
                "blocked".to_string(),
                Some(crate::vm::RuntimeFailureKind::SecurityBlockedExternalEffect),
            )
            .termination_class(),
            "policy_blocked"
        );
    }
}
