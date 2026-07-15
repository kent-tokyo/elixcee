/// Integration tests for the `elixcee test-workbook` subcommand: runs the
/// built binary directly, mirroring the pattern in `tests/cli_snapshot.rs`
/// (serde_json is a dev-only dependency for parsing `--json` output — it
/// does not affect the release binary, which uses the hand-rolled TOML
/// parser in `src/testworkbook.rs` for the fixture file itself).
///
/// Every expected value here was captured by actually running the built
/// binary during development, not hand-guessed — e.g. real VBA division
/// (`x = 100 / y`) raises a hard VBA runtime error rather than producing a
/// stored Excel error value, so `no_excel_errors` fixtures here use
/// `Range(...).Formula = "=100/B2"` (a real formula) instead.
///
/// `no_panic` is exercised only at the unit level (`src/testworkbook.rs`'s
/// `catch_unwind` wrapping is structurally verified there); no naturally
/// occurring Rust panic trigger was found via a real VBA construct in this
/// codebase during development (array/index access is already
/// bounds-checked and returns `Result::Err`, not a panic — a side effect of
/// this codebase's existing fuzzing history), so no end-to-end "confirm a
/// panicking macro fails with no_panic" case is included here.
use serde_json::Value;
use std::fs;
use std::path::PathBuf;
use std::process::Command;

/// Builds a fixture directory containing `orders.xlsx` (a fresh, single
/// default-sheet workbook), the given `.bas` source, and a `fixture.toml`
/// referencing both by relative path — proving path resolution is relative
/// to the fixture file's own directory, not the process's CWD.
fn build_fixture_dir(tag: &str, vba: &str, fixture_toml: &str) -> PathBuf {
    let dir = std::env::temp_dir().join(format!("elixcee_cli_test_workbook_{}", tag));
    let _ = fs::remove_dir_all(&dir);
    fs::create_dir_all(&dir).unwrap();

    let blank_bas = dir.join("blank.bas");
    fs::write(&blank_bas, "Sub Main()\n    x = 1\nEnd Sub\n").unwrap();
    let workbook_path = dir.join("orders.xlsx");
    let out = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            blank_bas.to_str().unwrap(),
            "Main",
            "--output",
            workbook_path.to_str().unwrap(),
        ])
        .output()
        .expect("build blank workbook fixture");
    assert!(
        out.status.success(),
        "{:?}",
        String::from_utf8_lossy(&out.stderr)
    );

    fs::write(dir.join("Main.bas"), vba).unwrap();
    fs::write(dir.join("fixture.toml"), fixture_toml).unwrap();
    dir
}

fn run_json(fixture_path: &std::path::Path, extra_args: &[&str]) -> (bool, Value) {
    let mut args = vec![fixture_path.to_str().unwrap(), "--json"];
    args.extend_from_slice(extra_args);
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .arg("test-workbook")
        .args(&args)
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).expect("stdout is utf8");
    let stderr = String::from_utf8(output.stderr).expect("stderr is utf8");
    assert_eq!(
        stdout.lines().count(),
        1,
        "test-workbook --json must emit exactly one line, got: {:?} (stderr: {:?})",
        stdout,
        stderr,
    );
    let parsed: Value = serde_json::from_str(stdout.trim())
        .unwrap_or_else(|e| panic!("stdout was not valid JSON ({}): {:?}", e, stdout));
    (output.status.success(), parsed)
}

const DIVIDE_FIXTURE_TOML: &str = r#"
name = "divide test"
workbook = "orders.xlsx"
vba_files = ["Main.bas"]
macro = "Main"
cases = 50
seed = 42

[[inputs]]
range = "Sheet1!B2"
strategy = "boundary_numeric"

[[assertions]]
range = "Sheet1!A1"
rule = "no_excel_errors"
"#;

const DIVIDE_MACRO: &str = "Sub Main()\n    Range(\"A1\").Formula = \"=100/B2\"\nEnd Sub\n";

#[test]
fn json_failing_fixture_reports_seed_case_index_inputs_and_failure() {
    let dir = build_fixture_dir("failing", DIVIDE_MACRO, DIVIDE_FIXTURE_TOML);
    let (ok, v) = run_json(&dir.join("fixture.toml"), &[]);
    assert!(!ok, "{:?}", v);
    assert_eq!(v["schema_version"], 1);
    assert_eq!(v["ok"], false);
    assert_eq!(v["seed"], 42);
    assert_eq!(v["case_index"], 0);
    assert_eq!(v["inputs"][0]["address"], "sheet1!B2");
    assert_eq!(v["failure"]["rule"], "no_excel_errors");
    assert_eq!(v["failure"]["address"], "sheet1!A1");
    assert_eq!(v["failure"]["actual"], "#DIV/0!");
}

#[test]
fn seed_and_case_replay_reproduces_the_identical_failure() {
    let dir = build_fixture_dir("replay", DIVIDE_MACRO, DIVIDE_FIXTURE_TOML);
    let (_, first) = run_json(&dir.join("fixture.toml"), &[]);
    let case_index = first["case_index"].as_i64().unwrap();

    let (ok, replay) = run_json(
        &dir.join("fixture.toml"),
        &["--seed", "42", "--case", &case_index.to_string()],
    );
    assert!(!ok, "{:?}", replay);
    assert_eq!(replay["case_index"], case_index);
    assert_eq!(replay["inputs"], first["inputs"]);
    assert_eq!(replay["failure"], first["failure"]);
}

#[test]
fn json_passing_fixture_reports_ok_true_and_cases_run() {
    let dir = build_fixture_dir(
        "passing",
        "Sub Main()\n    Cells(1, 1).Value = 1\nEnd Sub\n",
        r#"
name = "pass test"
workbook = "orders.xlsx"
vba_files = ["Main.bas"]
macro = "Main"
cases = 15
seed = 1

[[inputs]]
range = "Sheet1!B2"
strategy = "boundary_string"

[[assertions]]
range = "Sheet1!A1"
rule = "no_excel_errors"
"#,
    );
    let (ok, v) = run_json(&dir.join("fixture.toml"), &[]);
    assert!(ok, "{:?}", v);
    assert_eq!(v["ok"], true);
    assert_eq!(v["seed"], 1);
    assert_eq!(v["cases_run"], 15);
}

#[test]
fn timeout_fires_close_to_the_configured_deadline_not_never() {
    let dir = build_fixture_dir(
        "timeout",
        "Sub Main()\n    Do While True\n        x = 1\n    Loop\nEnd Sub\n",
        r#"
name = "timeout test"
workbook = "orders.xlsx"
vba_files = ["Main.bas"]
macro = "Main"
cases = 1
seed = 1
timeout_secs = 1

[[inputs]]
range = "Sheet1!B2"
strategy = "boundary_numeric"

[[assertions]]
range = "Sheet1!A1"
rule = "no_excel_errors"
"#,
    );
    let start = std::time::Instant::now();
    let (ok, v) = run_json(&dir.join("fixture.toml"), &[]);
    let elapsed = start.elapsed();
    assert!(!ok, "{:?}", v);
    assert_eq!(v["failure"]["rule"], "no_timeout");
    assert!(
        v["failure"]["message"]
            .as_str()
            .unwrap()
            .starts_with("TIMEOUT:")
    );
    assert!(
        elapsed.as_secs_f64() < 10.0,
        "timeout overshot far past its 1s deadline: {:?}",
        elapsed
    );
}

#[test]
fn non_json_mode_prints_a_plain_text_summary() {
    let dir = build_fixture_dir(
        "plaintext",
        "Sub Main()\n    Cells(1, 1).Value = 1\nEnd Sub\n",
        r#"
name = "pass test"
workbook = "orders.xlsx"
vba_files = ["Main.bas"]
macro = "Main"
cases = 5
seed = 1

[[inputs]]
range = "Sheet1!B2"
strategy = "boundary_string"

[[assertions]]
range = "Sheet1!A1"
rule = "no_excel_errors"
"#,
    );
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(["test-workbook", dir.join("fixture.toml").to_str().unwrap()])
        .output()
        .expect("run elixcee binary");
    assert!(output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("ok:"));
    assert!(stdout.contains("5"));
}

#[test]
fn missing_fixture_file_is_an_io_error_in_json_mode() {
    let path = std::env::temp_dir().join("elixcee_cli_test_workbook_does_not_exist.toml");
    let _ = fs::remove_file(&path);
    let (ok, v) = run_json(&path, &[]);
    assert!(!ok, "{:?}", v);
    assert_eq!(v["error"]["code"], "E3001");
    assert_eq!(v["error"]["kind"], "io_error");
}

#[test]
fn malformed_fixture_toml_is_a_clear_error_not_a_silent_empty_run() {
    let dir = std::env::temp_dir().join("elixcee_cli_test_workbook_malformed");
    let _ = fs::remove_dir_all(&dir);
    fs::create_dir_all(&dir).unwrap();
    // An inline table is outside this hand-rolled parser's supported subset.
    fs::write(dir.join("fixture.toml"), "name = { bad = 1 }\n").unwrap();
    let (ok, v) = run_json(&dir.join("fixture.toml"), &[]);
    assert!(!ok, "{:?}", v);
    assert_eq!(v["error"]["code"], "E3001");
}
