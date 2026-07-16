/// Integration tests for the `elixcee diagnose-workbook` subcommand
/// (Milestone B6d): runs the built binary directly, mirroring
/// `tests/cli_test_workbook.rs`'s own conventions (same fixture format,
/// same `--json` single-line contract) plus the new `root_causes` field and
/// `--cases` override.
///
/// The smoke fixture uses `ArrayIndexOutOfBounds` — an *input-dependent*
/// root cause (a drawn boundary value can flip an array index in or out of
/// bounds across cases) — not a structural one (merge/shape/protection),
/// since those fire identically regardless of which case runs and wouldn't
/// demonstrate this command's actual value over a single plain `diagnose`
/// call. Every expected value here was captured by actually running the
/// built binary during development, not hand-guessed.
use serde_json::Value;
use std::fs;
use std::path::PathBuf;
use std::process::Command;

fn build_fixture_dir(tag: &str, vba: &str, fixture_toml: &str) -> PathBuf {
    let dir = std::env::temp_dir().join(format!("elixcee_cli_diagnose_workbook_{}", tag));
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
        .arg("diagnose-workbook")
        .args(&args)
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).expect("stdout is utf8");
    let stderr = String::from_utf8(output.stderr).expect("stderr is utf8");
    assert_eq!(
        stdout.lines().count(),
        1,
        "diagnose-workbook --json must emit exactly one line, got: {:?} (stderr: {:?})",
        stdout,
        stderr,
    );
    let parsed: Value = serde_json::from_str(stdout.trim())
        .unwrap_or_else(|e| panic!("stdout was not valid JSON ({}): {:?}", e, stdout));
    (output.status.success(), parsed)
}

const ARRAY_OOB_MACRO: &str =
    "Sub Main()\n    Dim arr(5)\n    arr(Range(\"B2\").Value) = 1\nEnd Sub\n";

const ARRAY_OOB_FIXTURE_TOML: &str = r#"
name = "array oob"
workbook = "orders.xlsx"
vba_files = ["Main.bas"]
macro = "Main"
cases = 20
seed = 42

[[inputs]]
range = "Sheet1!B2"
strategy = "boundary_numeric"

[[assertions]]
range = "Sheet1!A1"
rule = "no_excel_errors"
"#;

#[test]
fn json_failing_fixture_reports_a_classified_array_index_root_cause() {
    let dir = build_fixture_dir("oob", ARRAY_OOB_MACRO, ARRAY_OOB_FIXTURE_TOML);
    let (ok, v) = run_json(&dir.join("fixture.toml"), &[]);
    assert!(!ok, "{:?}", v);
    assert_eq!(v["schema_version"], 1);
    assert_eq!(v["ok"], false);
    assert_eq!(v["seed"], 42);
    assert_eq!(v["root_causes"][0]["code"], "ARRAY_INDEX_OUT_OF_BOUNDS");
    assert_eq!(v["root_causes"][0]["name"], "arr");
    assert_eq!(v["root_causes"][0]["lower"], 0);
    assert_eq!(v["root_causes"][0]["upper"], 5);
}

#[test]
fn seed_and_case_replay_reproduces_the_identical_classified_failure() {
    let dir = build_fixture_dir("oob-replay", ARRAY_OOB_MACRO, ARRAY_OOB_FIXTURE_TOML);
    let (_, first) = run_json(&dir.join("fixture.toml"), &[]);
    let case_index = first["case_index"].as_i64().unwrap();

    let (ok, replay) = run_json(
        &dir.join("fixture.toml"),
        &["--seed", "42", "--case", &case_index.to_string()],
    );
    assert!(!ok, "{:?}", replay);
    assert_eq!(replay["case_index"], case_index);
    assert_eq!(replay["inputs"], first["inputs"]);
    assert_eq!(replay["root_causes"], first["root_causes"]);
}

#[test]
fn cases_override_runs_fewer_cases_than_the_fixture_declares_on_a_passing_run() {
    let dir = build_fixture_dir(
        "cases-override",
        "Sub Main()\n    Cells(1, 1).Value = 1\nEnd Sub\n",
        r#"
name = "pass test"
workbook = "orders.xlsx"
vba_files = ["Main.bas"]
macro = "Main"
cases = 100
seed = 1

[[inputs]]
range = "Sheet1!B2"
strategy = "boundary_string"

[[assertions]]
range = "Sheet1!A1"
rule = "no_excel_errors"
"#,
    );
    let (ok, v) = run_json(&dir.join("fixture.toml"), &["--cases", "5"]);
    assert!(ok, "{:?}", v);
    assert_eq!(v["cases_run"], 5);
}

#[test]
fn json_passing_fixture_reports_ok_true_with_no_root_causes_field_needed() {
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
    assert_eq!(v["cases_run"], 15);
}

#[test]
fn non_json_mode_reports_a_root_cause_line_for_a_classified_failure() {
    let dir = build_fixture_dir("plaintext", ARRAY_OOB_MACRO, ARRAY_OOB_FIXTURE_TOML);
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            "diagnose-workbook",
            dir.join("fixture.toml").to_str().unwrap(),
        ])
        .output()
        .expect("run elixcee binary");
    assert!(!output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("root cause: ARRAY_INDEX_OUT_OF_BOUNDS"));
}
