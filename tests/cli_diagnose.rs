/// Integration tests for the `elixcee diagnose` subcommand (Milestone
/// B6a): runs the built binary directly, mirroring the pattern in
/// `tests/cli_test_workbook.rs` (serde_json is a dev-only dependency for
/// parsing `--json` output — the release binary itself has no JSON crate
/// dependency; `src/diagnose.rs`'s JSON is hand-rolled, same rationale as
/// every other subcommand's output).
use serde_json::Value;
use std::fs;
use std::path::PathBuf;
use std::process::Command;

/// Builds a scratch directory with `report.xlsx` (containing one cell
/// written in each of `sheets`, so those sheets actually exist) and a
/// `main.bas` holding `vba`. Returns `(dir, workbook_path, macro_bas_path)`.
fn build_dir(tag: &str, sheets: &[&str], vba: &str) -> (PathBuf, PathBuf, PathBuf) {
    let dir = std::env::temp_dir().join(format!("elixcee_cli_diagnose_{}", tag));
    let _ = fs::remove_dir_all(&dir);
    fs::create_dir_all(&dir).unwrap();

    let setup_bas = dir.join("setup.bas");
    let setup_src: String = sheets
        .iter()
        .map(|s| format!("    Worksheets(\"{}\").Cells(1,1).Value = 1\n", s))
        .collect();
    fs::write(&setup_bas, format!("Sub Setup()\n{}End Sub\n", setup_src)).unwrap();
    let workbook_path = dir.join("report.xlsx");
    let out = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            setup_bas.to_str().unwrap(),
            "Setup",
            "--output",
            workbook_path.to_str().unwrap(),
        ])
        .output()
        .expect("build workbook fixture");
    assert!(
        out.status.success(),
        "{:?}",
        String::from_utf8_lossy(&out.stderr)
    );

    let macro_bas = dir.join("main.bas");
    fs::write(&macro_bas, vba).unwrap();
    (dir, workbook_path, macro_bas)
}

fn run_json(
    macro_bas: &std::path::Path,
    workbook: &std::path::Path,
    entrypoint: &str,
) -> (bool, Value) {
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .arg("diagnose")
        .arg(macro_bas.to_str().unwrap())
        .args([
            "--file",
            workbook.to_str().unwrap(),
            "--entrypoint",
            entrypoint,
            "--json",
        ])
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).expect("stdout is utf8");
    let stderr = String::from_utf8(output.stderr).expect("stderr is utf8");
    assert_eq!(
        stdout.lines().count(),
        1,
        "diagnose --json must emit exactly one line, got: {:?} (stderr: {:?})",
        stdout,
        stderr,
    );
    let parsed: Value = serde_json::from_str(stdout.trim())
        .unwrap_or_else(|e| panic!("stdout was not valid JSON ({}): {:?}", e, stdout));
    (output.status.success(), parsed)
}

#[test]
fn missing_worksheet_by_name_reports_a_did_you_mean_suggestion() {
    // The exact scenario from the user's own B6a request: a typo'd sheet
    // year, with a close-name candidate one edit away.
    let (_, workbook, macro_bas) = build_dir(
        "missing_sheet",
        &["入力", "売上2026", "集計"],
        "Sub Run()\n    Worksheets(\"売上2025\").Range(\"A1\").Value = 1\nEnd Sub\n",
    );
    let (ok, v) = run_json(&macro_bas, &workbook, "Run");
    assert!(!ok, "{:?}", v);
    assert_eq!(v["ok"], false);
    assert_eq!(v["root_causes"][0]["code"], "WORKSHEET_NOT_FOUND");
    assert_eq!(v["root_causes"][0]["requested"], "売上2025");
    assert_eq!(v["root_causes"][0]["suggested"], "売上2026");
    assert!(
        v["location"]["file"]
            .as_str()
            .unwrap()
            .ends_with("main.bas")
    );
}

#[test]
fn missing_workbook_reports_a_workbook_not_found_root_cause() {
    let (_, workbook, macro_bas) = build_dir(
        "missing_workbook",
        &[],
        "Sub Run()\n    Workbooks(\"data.xlsx\").Worksheets(1).Cells(1,1).Value = 1\nEnd Sub\n",
    );
    let (ok, v) = run_json(&macro_bas, &workbook, "Run");
    assert!(!ok, "{:?}", v);
    assert_eq!(v["root_causes"][0]["code"], "WORKBOOK_NOT_FOUND");
    assert_eq!(v["root_causes"][0]["requested"], "data.xlsx");
}

#[test]
fn array_out_of_bounds_reports_zero_based_bounds_evidence() {
    let (_, workbook, macro_bas) = build_dir(
        "array_oob",
        &[],
        "Sub Run()\n    Dim values(5)\n    values(9) = 1\nEnd Sub\n",
    );
    let (ok, v) = run_json(&macro_bas, &workbook, "Run");
    assert!(!ok, "{:?}", v);
    assert_eq!(v["root_causes"][0]["code"], "ARRAY_INDEX_OUT_OF_BOUNDS");
    assert_eq!(v["root_causes"][0]["name"], "values");
    assert_eq!(v["root_causes"][0]["index"], 9);
    assert_eq!(v["root_causes"][0]["lower"], 0);
    assert_eq!(v["root_causes"][0]["upper"], 5);
}

#[test]
fn numeric_sheet_index_out_of_range_is_diagnosed() {
    let (_, workbook, macro_bas) = build_dir(
        "numeric_index",
        &[],
        "Sub Run()\n    x = Worksheets(9).Cells(1,1).Value\nEnd Sub\n",
    );
    let (ok, v) = run_json(&macro_bas, &workbook, "Run");
    assert!(!ok, "{:?}", v);
    assert_eq!(v["root_causes"][0]["code"], "WORKSHEET_NOT_FOUND");
    assert_eq!(v["root_causes"][0]["requested"], "9");
}

#[test]
fn passing_macro_reports_ok_true() {
    let (_, workbook, macro_bas) = build_dir(
        "passing",
        &[],
        "Sub Run()\n    Cells(1,1).Value = 1\nEnd Sub\n",
    );
    let (ok, v) = run_json(&macro_bas, &workbook, "Run");
    assert!(ok, "{:?}", v);
    assert_eq!(v["ok"], true);
}

#[test]
fn on_error_resume_next_does_not_hide_the_failure() {
    let (_, workbook, macro_bas) = build_dir(
        "on_error",
        &[],
        "Sub Run()\n    On Error Resume Next\n    Worksheets(\"Missing\").Range(\"A1\").Value = 1\n    x = 2\nEnd Sub\n",
    );
    let (ok, v) = run_json(&macro_bas, &workbook, "Run");
    assert!(
        !ok,
        "diagnose must not let On Error Resume Next hide the resolution failure: {:?}",
        v
    );
    assert_eq!(v["root_causes"][0]["code"], "WORKSHEET_NOT_FOUND");
}

#[test]
fn non_json_mode_prints_a_plain_text_summary() {
    let (_, workbook, macro_bas) = build_dir(
        "plaintext",
        &["Input"],
        "Sub Run()\n    Worksheets(\"Wrong\").Range(\"A1\").Value = 1\nEnd Sub\n",
    );
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .arg("diagnose")
        .arg(macro_bas.to_str().unwrap())
        .args(["--file", workbook.to_str().unwrap(), "--entrypoint", "Run"])
        .output()
        .expect("run elixcee binary");
    assert!(!output.status.success());
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("WORKSHEET_NOT_FOUND"));
}

#[test]
fn missing_workbook_file_is_an_io_error_in_json_mode() {
    let dir = std::env::temp_dir().join("elixcee_cli_diagnose_missing_file");
    let _ = fs::remove_dir_all(&dir);
    fs::create_dir_all(&dir).unwrap();
    let macro_bas = dir.join("main.bas");
    fs::write(&macro_bas, "Sub Run()\nEnd Sub\n").unwrap();
    let missing_workbook = dir.join("does_not_exist.xlsx");
    let (ok, v) = run_json(&macro_bas, &missing_workbook, "Run");
    assert!(!ok, "{:?}", v);
    assert_eq!(v["error"]["code"], "E3001");
    assert_eq!(v["error"]["kind"], "io_error");
}

/// Confirms the opt-in design end to end: the exact same macro construct
/// (writing to a not-yet-existing sheet by name) succeeds under plain `run`
/// mode via its existing auto-vivify convenience, while `diagnose` treats
/// it as a hard resolution failure — proving `strict_resolution` really is
/// isolated to `diagnose` and doesn't regress `run`.
#[test]
fn run_mode_auto_vivify_is_unaffected_by_diagnose_strict_mode() {
    let (_, workbook, macro_bas) = build_dir(
        "run_unaffected",
        &[],
        "Sub Run()\n    Worksheets(\"BrandNew\").Cells(1,1).Value = 42\nEnd Sub\n",
    );

    let run_output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            macro_bas.to_str().unwrap(),
            "Run",
            "--file",
            workbook.to_str().unwrap(),
            "--json",
        ])
        .output()
        .expect("run elixcee binary");
    assert!(
        run_output.status.success(),
        "{:?}",
        String::from_utf8_lossy(&run_output.stderr)
    );
    let run_stdout = String::from_utf8_lossy(&run_output.stdout);
    assert!(run_stdout.contains("\"ok\":true"));

    let (diag_ok, diag_v) = run_json(&macro_bas, &workbook, "Run");
    assert!(!diag_ok, "{:?}", diag_v);
    assert_eq!(diag_v["root_causes"][0]["code"], "WORKSHEET_NOT_FOUND");
}
