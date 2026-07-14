/// Integration tests for the `--json` CLI contract: runs the built `elixcee`
/// binary directly and validates its stdout by parsing it with a real JSON
/// parser (serde_json is a dev-only dependency — it does not affect the
/// release binary, which still emits JSON via the hand-rolled writer in
/// `src/diagnostics.rs`).
use serde_json::Value;
use std::fs;
use std::process::Command;

/// Run the built binary with `--json` plus any extra args, asserting stdout
/// is exactly one line and that it parses as JSON. Returns (exit_success, parsed_json).
fn run_json_with_args(args: &[&std::ffi::OsStr]) -> (bool, Value) {
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(args)
        .arg("--json")
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).expect("stdout is utf8");
    let stderr = String::from_utf8(output.stderr).expect("stderr is utf8");

    assert_eq!(
        stdout.lines().count(),
        1,
        "--json must emit exactly one line of stdout, got: {:?} (stderr: {:?})",
        stdout,
        stderr,
    );
    let parsed: Value = serde_json::from_str(stdout.trim())
        .unwrap_or_else(|e| panic!("stdout was not valid JSON ({}): {:?}", e, stdout));
    (output.status.success(), parsed)
}

fn write_vba(vba: &str, tag: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!("elixcee_cli_json_{}.bas", tag));
    fs::write(&path, vba).expect("write temp vba file");
    path
}

fn run_json(vba: &str, macro_name: &str, tag: &str) -> (bool, Value) {
    let path = write_vba(vba, tag);
    run_json_with_args(&[path.as_os_str(), std::ffi::OsStr::new(macro_name)])
}

// ── Success ─────────────────────────────────────────────────────────────────

#[test]
fn json_success_is_single_clean_object() {
    let (ok, v) = run_json(
        "Sub Main()\n    Cells(1, 1).Value = 42\nEnd Sub\n",
        "Main",
        "success",
    );
    assert!(ok, "expected exit 0: {:?}", v);
    assert_eq!(v["schema_version"], 1);
    assert_eq!(v["ok"], true);
    assert_eq!(v["entrypoint"], "Main");
    assert!(v["duration_ms"].as_f64().unwrap() >= 0.0);
    assert_eq!(v["cells"][0]["sheet"], "sheet1");
    assert_eq!(v["cells"][0]["address"], "A1");
    assert_eq!(v["cells"][0]["value"], 42);
    assert_eq!(v["messages"], serde_json::json!([]));
}

#[test]
fn json_msgbox_does_not_corrupt_stdout() {
    let (ok, v) = run_json(
        "Sub Main()\n    MsgBox \"hi there\"\n    Cells(1, 1).Value = 1\nEnd Sub\n",
        "Main",
        "msgbox",
    );
    assert!(ok, "{:?}", v);
    assert_eq!(v["messages"], serde_json::json!(["hi there"]));
}

// ── Multi-module (Milestone B2) ──────────────────────────────────────────────

#[test]
fn json_multi_file_run_resolves_qualified_entrypoint() {
    // Attribute VB_Name names the module explicitly, independent of the
    // temp file's actual filename (which the file-stem fallback would use).
    let helper_path = write_vba(
        "Attribute VB_Name = \"Helpers\"\nSub Helper()\n    y = 1\nEnd Sub\n",
        "multi_helper",
    );
    let main_path = write_vba(
        "Attribute VB_Name = \"MainMod\"\nSub Main()\n    Call Helper()\n    Cells(1, 1).Value = 42\nEnd Sub\n",
        "multi_main",
    );
    let (ok, v) = run_json_with_args(&[
        helper_path.as_os_str(),
        main_path.as_os_str(),
        std::ffi::OsStr::new("MainMod.Main"),
    ]);
    assert!(ok, "{:?}", v);
    assert_eq!(v["ok"], true);
    assert_eq!(v["cells"][0]["value"], 42);
}

#[test]
fn json_multi_file_run_rejects_a_genuine_sub_collision() {
    let a = write_vba("Sub Main()\n    x = 1\nEnd Sub\n", "collide_a");
    let b = write_vba("Sub Main()\n    x = 2\nEnd Sub\n", "collide_b");
    let (ok, v) = run_json_with_args(&[a.as_os_str(), b.as_os_str(), std::ffi::OsStr::new("Main")]);
    assert!(!ok, "{:?}", v);
    assert_eq!(v["ok"], false);
    let msg = v["error"]["message"].as_str().unwrap_or_default();
    assert!(msg.contains("duplicate Sub 'main'"), "{:?}", v);
}

// ── One failure per CLI stage ────────────────────────────────────────────────

#[test]
fn json_missing_file_is_io_error() {
    let missing = std::env::temp_dir().join("elixcee_cli_json_does_not_exist.bas");
    let _ = fs::remove_file(&missing);
    let (ok, v) = run_json_with_args(&[missing.as_os_str(), std::ffi::OsStr::new("Main")]);
    assert!(!ok, "{:?}", v);
    assert_eq!(v["ok"], false);
    assert_eq!(v["error"]["code"], "E3001");
    assert_eq!(v["error"]["kind"], "io_error");
    // Happens before the macro file is even read — nothing to locate.
    assert_eq!(v["error"]["location"], serde_json::Value::Null);
}

#[test]
fn json_syntax_error_is_parse_error() {
    let (ok, v) = run_json("Sub Main(\n    x = 1\n", "Main", "parse_error");
    assert!(!ok, "{:?}", v);
    assert_eq!(v["error"]["code"], "E2001");
    assert_eq!(v["error"]["kind"], "parse_error");
    assert!(v["error"]["location"]["line"].as_u64().is_some(), "{:?}", v);
    assert!(
        v["error"]["location"]["column"].as_u64().is_some(),
        "{:?}",
        v
    );
}

#[test]
fn json_undefined_variable_is_runtime_error() {
    let (ok, v) = run_json(
        "Sub Main()\n    x = totla + 1\nEnd Sub\n",
        "Main",
        "undefined_var",
    );
    assert!(!ok, "{:?}", v);
    assert_eq!(v["error"]["code"], "E1001");
    assert_eq!(v["error"]["kind"], "undefined_variable");
    // Statement-level granularity (documented MVP scope): location points at
    // the start of the whole "x = totla + 1" statement on line 2 (column 5,
    // where "x" is), not the "totla" sub-expression specifically.
    assert_eq!(v["error"]["location"]["line"], 2);
    assert_eq!(v["error"]["location"]["column"], 5);
    assert!(
        v["error"]["location"]["file"]
            .as_str()
            .unwrap()
            .ends_with("undefined_var.bas"),
        "{:?}",
        v
    );
}

#[test]
fn json_error_still_carries_messages_shown_before_the_failure() {
    let (ok, v) = run_json(
        "Sub Main()\n    MsgBox \"seen before failure\"\n    x = totla + 1\nEnd Sub\n",
        "Main",
        "msgbox_then_error",
    );
    assert!(!ok, "{:?}", v);
    assert_eq!(v["error"]["code"], "E1001");
    assert_eq!(v["messages"], serde_json::json!(["seen before failure"]));
}

#[test]
fn json_missing_entrypoint_is_runtime_error() {
    let (ok, v) = run_json(
        "Sub Main()\n    x = 1\nEnd Sub\n",
        "NoSuchMacro",
        "missing_entrypoint",
    );
    assert!(!ok, "{:?}", v);
    assert_eq!(v["error"]["code"], "E1002");
    // Fails before any statement ever executes — nothing to locate.
    assert_eq!(v["error"]["location"], serde_json::Value::Null);
}

#[test]
fn json_output_and_sheet_setup_stages() {
    let vba_path = write_vba(
        "Sub Main()\n    Cells(1, 1).Value = 99\nEnd Sub\n",
        "output_src",
    );
    let out_path = std::env::temp_dir().join("elixcee_cli_json_output.xlsx");
    let _ = fs::remove_file(&out_path);

    let (ok, v) = run_json_with_args(&[
        vba_path.as_os_str(),
        std::ffi::OsStr::new("Main"),
        std::ffi::OsStr::new("--output"),
        out_path.as_os_str(),
    ]);
    assert!(ok, "{:?}", v);
    assert_eq!(v["ok"], true);
    assert!(out_path.exists(), "--output file was not written");

    // Feed the written workbook back in with a bad --sheet: must still be a
    // single clean JSON error object (E3002), not two objects / mixed text.
    let (ok2, v2) = run_json_with_args(&[
        vba_path.as_os_str(),
        std::ffi::OsStr::new("Main"),
        std::ffi::OsStr::new("--file"),
        out_path.as_os_str(),
        std::ffi::OsStr::new("--sheet"),
        std::ffi::OsStr::new("DoesNotExist"),
    ]);
    assert!(!ok2, "{:?}", v2);
    assert_eq!(v2["error"]["code"], "E3002");
    assert_eq!(v2["error"]["kind"], "sheet_setup_error");
}

// ── String content: quotes, backslashes, control chars, Unicode, edge sizes ──

#[test]
fn json_cell_values_with_special_characters_round_trip() {
    let long_string = "x".repeat(5000);
    let vba = format!(
        "Sub Main()\n    \
         Cells(1, 1).Value = \"quote\"\"and\\backslash\"\n    \
         Cells(2, 1).Value = \"line1\" & Chr(10) & \"line2\" & Chr(9) & \"tabbed\"\n    \
         Cells(3, 1).Value = \"日本語 emoji test\"\n    \
         Cells(4, 1).Value = \"\"\n    \
         Cells(5, 1).Value = \"{long}\"\n\
         End Sub\n",
        long = long_string,
    );
    let (ok, v) = run_json(&vba, "Main", "special_chars");
    assert!(ok, "{:?}", v);
    let cells = v["cells"].as_array().unwrap();
    let value_at = |addr: &str| {
        cells
            .iter()
            .find(|c| c["address"] == addr)
            .unwrap_or_else(|| panic!("no cell {}", addr))["value"]
            .clone()
    };
    assert_eq!(value_at("A1"), "quote\"and\\backslash");
    assert_eq!(value_at("A2"), "line1\nline2\ttabbed");
    assert_eq!(value_at("A3"), "日本語 emoji test");
    assert_eq!(value_at("A4"), "");
    assert_eq!(value_at("A5"), long_string);
}

#[test]
fn json_io_error_message_with_backslash_in_path_round_trips() {
    // Backslash is just a regular filename character on macOS/Linux — this
    // exercises real backslash-escaping in the JSON message without needing
    // a Windows path.
    let missing = std::env::temp_dir().join("elixcee_cli_json_back\\slash_missing.bas");
    let _ = fs::remove_file(&missing);
    let (ok, v) = run_json_with_args(&[missing.as_os_str(), std::ffi::OsStr::new("Main")]);
    assert!(!ok, "{:?}", v);
    assert_eq!(v["error"]["code"], "E3001");
    let msg = v["error"]["message"].as_str().unwrap();
    assert!(
        msg.contains('\\'),
        "expected literal backslash in message: {}",
        msg
    );
}

#[test]
fn json_excel_error_value_round_trips() {
    let (ok, v) = run_json(
        "Sub Main()\n    Range(\"A1\").Formula = \"=1/0\"\nEnd Sub\n",
        "Main",
        "excel_error_value",
    );
    assert!(ok, "{:?}", v);
    let cells = v["cells"].as_array().unwrap();
    let a1 = cells.iter().find(|c| c["address"] == "A1").unwrap();
    assert_eq!(a1["value"], "#DIV/0!");
}

// ── Non-JSON mode stays byte-for-byte the same ───────────────────────────────

#[test]
fn non_json_output_is_unchanged() {
    let path = write_vba(
        "Sub Main()\n    MsgBox \"hello\"\n    Cells(1, 1).Value = 42\n    Cells(2, 1).Value = \"hi\"\nEnd Sub\n",
        "non_json_unchanged",
    );
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .arg(&path)
        .arg("Main")
        .output()
        .expect("run elixcee binary");
    assert!(output.status.success());
    let stdout = String::from_utf8(output.stdout).unwrap();
    // MsgBox prints immediately during execution, before the final cell dump —
    // same order as before the msgbox_log refactor.
    assert_eq!(stdout, "hello\nA1\t42\nA2\thi\n");
}
