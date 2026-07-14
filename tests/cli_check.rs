/// Integration tests for the `elixcee check` subcommand: spawns the built
/// binary directly and validates its stdout, mirroring the pattern in
/// `tests/cli_json.rs` (serde_json is a dev-only dependency).
use serde_json::Value;
use std::fs;
use std::process::Command;

fn write_vba(vba: &str, tag: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!("elixcee_cli_check_{}.bas", tag));
    fs::write(&path, vba).expect("write temp vba file");
    path
}

/// Run `elixcee check <file> [--entry macro] --json`, asserting stdout is
/// exactly one line of valid JSON. Returns (exit_success, parsed_json).
fn run_check_json(vba: &str, macro_name: Option<&str>, tag: &str) -> (bool, Value) {
    let path = write_vba(vba, tag);
    let mut args = vec!["check".to_string(), path.to_string_lossy().to_string()];
    if let Some(m) = macro_name {
        args.push("--entry".to_string());
        args.push(m.to_string());
    }
    args.push("--json".to_string());

    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(&args)
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).expect("stdout is utf8");
    let stderr = String::from_utf8(output.stderr).expect("stderr is utf8");

    assert_eq!(
        stdout.lines().count(),
        1,
        "check --json must emit exactly one line of stdout, got: {:?} (stderr: {:?})",
        stdout,
        stderr,
    );
    let parsed: Value = serde_json::from_str(stdout.trim())
        .unwrap_or_else(|e| panic!("stdout was not valid JSON ({}): {:?}", e, stdout));
    (output.status.success(), parsed)
}

#[test]
fn clean_file_has_no_diagnostics() {
    let (ok, v) = run_check_json(
        "Sub Main()\n    Cells(1, 1).Value = 42\nEnd Sub\n",
        Some("Main"),
        "clean",
    );
    assert!(ok, "{:?}", v);
    assert_eq!(v["schema_version"], 1);
    assert_eq!(v["ok"], true);
    assert_eq!(v["diagnostics"], serde_json::json!([]));
}

#[test]
fn missing_file_is_io_error() {
    let missing = std::env::temp_dir().join("elixcee_cli_check_does_not_exist.bas");
    let _ = fs::remove_file(&missing);
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(["check", missing.to_str().unwrap(), "--json"])
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).expect("stdout is utf8");
    assert_eq!(stdout.lines().count(), 1, "{:?}", stdout);
    let v: Value = serde_json::from_str(stdout.trim()).expect("valid json");
    assert!(!output.status.success());
    assert_eq!(v["ok"], false);
    let diags = v["diagnostics"].as_array().unwrap();
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0]["severity"], "error");
    assert_eq!(diags[0]["code"], "E3001");
    assert_eq!(diags[0]["kind"], "io_error");
    assert_eq!(diags[0]["location"], serde_json::Value::Null);
}

#[test]
fn parse_error_is_reported_with_location() {
    let (ok, v) = run_check_json("Sub Main(\n    x = 1\n", Some("Main"), "parse_error");
    assert!(!ok, "{:?}", v);
    assert_eq!(v["ok"], false);
    let diags = v["diagnostics"].as_array().unwrap();
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0]["severity"], "error");
    assert_eq!(diags[0]["code"], "E2001");
    assert!(diags[0]["location"]["line"].as_u64().is_some(), "{:?}", v);
}

#[test]
fn missing_entrypoint_is_reported() {
    let (ok, v) = run_check_json(
        "Sub Main()\n    x = 1\nEnd Sub\n",
        Some("NoSuchMacro"),
        "missing_entrypoint",
    );
    assert!(!ok, "{:?}", v);
    let diags = v["diagnostics"].as_array().unwrap();
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0]["code"], "E1002");
    assert_eq!(diags[0]["location"], serde_json::Value::Null);
}

#[test]
fn msgbox_is_an_info_diagnostic_not_an_error() {
    let (ok, v) = run_check_json(
        "Sub Main()\n    MsgBox \"hi\"\nEnd Sub\n",
        Some("Main"),
        "msgbox",
    );
    assert!(ok, "MsgBox alone should not fail check: {:?}", v);
    assert_eq!(v["ok"], true);
    let diags = v["diagnostics"].as_array().unwrap();
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0]["severity"], "info");
    assert_eq!(diags[0]["code"], "I1001");
    assert_eq!(diags[0]["location"]["line"], 2);
}

#[test]
fn undefined_function_call_is_reported() {
    let (ok, v) = run_check_json(
        "Sub Main()\n    x = Bogus(1)\nEnd Sub\n",
        Some("Main"),
        "undefined_call",
    );
    assert!(!ok, "{:?}", v);
    assert_eq!(v["ok"], false);
    let diags = v["diagnostics"].as_array().unwrap();
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0]["severity"], "error");
    assert_eq!(diags[0]["code"], "E1002");
    assert!(diags[0]["location"]["line"].as_u64().is_some(), "{:?}", v);
}

#[test]
fn unsupported_construct_is_reported_as_info() {
    let (ok, v) = run_check_json(
        "Sub Main()\n    Range(\"A1\").NumberFormat = \"0.00\"\nEnd Sub\n",
        Some("Main"),
        "unsupported_construct",
    );
    assert!(ok, "an info-only diagnostic must not fail check: {:?}", v);
    let diags = v["diagnostics"].as_array().unwrap();
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0]["severity"], "info");
    assert_eq!(diags[0]["code"], "I1002");
    assert!(diags[0]["location"]["line"].as_u64().is_some(), "{:?}", v);
}

#[test]
fn module_level_unsupported_construct_is_reported_as_info() {
    let (ok, v) = run_check_json(
        "Public Const MAX_RETRIES = 5\nSub Main()\n    x = 1\nEnd Sub\n",
        Some("Main"),
        "module_level_unsupported",
    );
    assert!(ok, "an info-only diagnostic must not fail check: {:?}", v);
    let diags = v["diagnostics"].as_array().unwrap();
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0]["severity"], "info");
    assert_eq!(diags[0]["code"], "I1002");
    assert_eq!(diags[0]["location"]["line"], 1);
}

#[test]
fn macro_name_is_optional() {
    let (ok, v) = run_check_json("Sub Main()\n    x = 1\nEnd Sub\n", None, "no_macro_name");
    assert!(ok, "{:?}", v);
    assert_eq!(v["diagnostics"], serde_json::json!([]));
}

#[test]
fn check_never_executes_the_macro() {
    // A macro whose only observable effect is a cell write. `check` has no
    // --output/--file surface at all, so the only way to prove it didn't
    // run is that no execution-only side channel (MsgBox capture, cell
    // dump) appears anywhere in the output — the diagnostics array is the
    // only thing check ever emits.
    let (ok, v) = run_check_json(
        "Sub Main()\n    MsgBox \"side effect\"\n    Cells(1, 1).Value = 999\nEnd Sub\n",
        Some("Main"),
        "no_execution",
    );
    assert!(ok, "{:?}", v);
    assert!(
        v.get("cells").is_none(),
        "check must not run the macro: {:?}",
        v
    );
    assert!(
        v.get("messages").is_none(),
        "check must not run the macro: {:?}",
        v
    );
    let diags = v["diagnostics"].as_array().unwrap();
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0]["code"], "I1001");
}

#[test]
fn non_json_mode_prints_ok_and_exits_zero_when_clean() {
    let path = write_vba("Sub Main()\n    x = 1\nEnd Sub\n", "plain_clean");
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(["check", path.to_str().unwrap(), "--entry", "Main"])
        .output()
        .expect("run elixcee binary");
    assert!(output.status.success());
    assert_eq!(String::from_utf8(output.stdout).unwrap(), "ok\n");
}

#[test]
fn non_json_mode_prints_diagnostic_lines_and_exits_nonzero_on_error() {
    let path = write_vba("Sub Main()\n    x = 1\nEnd Sub\n", "plain_error");
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(["check", path.to_str().unwrap(), "--entry", "Nope"])
        .output()
        .expect("run elixcee binary");
    assert!(!output.status.success());
    let stdout = String::from_utf8(output.stdout).unwrap();
    assert!(stdout.contains("E1002"), "{:?}", stdout);
    assert!(stdout.contains("error"), "{:?}", stdout);
}

// ── Multi-module (Milestone B2) ──────────────────────────────────────────────

#[test]
fn multi_file_check_with_entry_resolves_qualified_entrypoint_and_cross_module_call() {
    let helper = write_vba(
        "Attribute VB_Name = \"Helpers\"\nSub Helper()\n    y = 1\nEnd Sub\n",
        "multi_helper",
    );
    let main = write_vba(
        "Attribute VB_Name = \"MainMod\"\nSub Main()\n    Call Helper()\nEnd Sub\n",
        "multi_main",
    );
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            "check",
            helper.to_str().unwrap(),
            main.to_str().unwrap(),
            "--entry",
            "MainMod.Main",
            "--json",
        ])
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).unwrap();
    let v: Value = serde_json::from_str(stdout.trim()).expect("valid json");
    assert!(output.status.success(), "{:?}", v);
    assert_eq!(v["ok"], true);
    assert_eq!(v["diagnostics"], serde_json::json!([]));
}

#[test]
fn multi_file_check_without_entry_checks_every_module_with_no_entrypoint_assertion() {
    // The natural "check every module in the project" invocation — no
    // --entry, any number of files, no positional macro name at all.
    let a = write_vba("Sub Foo()\n    MsgBox \"hi\"\nEnd Sub\n", "noentry_a");
    let b = write_vba("Sub Bar()\n    x = 1\nEnd Sub\n", "noentry_b");
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(["check", a.to_str().unwrap(), b.to_str().unwrap(), "--json"])
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).unwrap();
    let v: Value = serde_json::from_str(stdout.trim()).expect("valid json");
    assert!(output.status.success(), "{:?}", v);
    let diags = v["diagnostics"].as_array().unwrap();
    assert_eq!(diags.len(), 1);
    assert_eq!(diags[0]["code"], "I1001"); // the MsgBox in module a
}

#[test]
fn multi_file_check_reports_a_cross_module_sub_collision() {
    let a = write_vba("Sub Main()\n    x = 1\nEnd Sub\n", "collide_a");
    let b = write_vba("Sub Main()\n    x = 2\nEnd Sub\n", "collide_b");
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(["check", a.to_str().unwrap(), b.to_str().unwrap(), "--json"])
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).unwrap();
    let v: Value = serde_json::from_str(stdout.trim()).expect("valid json");
    assert!(!output.status.success(), "{:?}", v);
    assert_eq!(v["ok"], false);
    let diags = v["diagnostics"].as_array().unwrap();
    assert!(
        diags.iter().any(|d| d["code"] == "E1005"),
        "expected an E1005 diagnostic: {:?}",
        v
    );
}
