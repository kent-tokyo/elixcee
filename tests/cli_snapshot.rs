/// Integration tests for the `elixcee snapshot` subcommand: runs the built
/// binary directly, mirroring the pattern in `tests/cli_json.rs`/
/// `tests/cli_check.rs` (serde_json is a dev-only dependency for parsing
/// `--json` output — it does not affect the release binary).
///
/// Every expected value here was captured by actually running the built
/// binary during development, not hand-guessed (a real `.xlsx` round-trips
/// through this repo's own writer/reader, which lowercases sheet names and
/// picks an Integer-vs-Float cell representation neither is safe to assume
/// without checking).
use serde_json::Value;
use std::fs;
use std::process::Command;

fn write_vba(vba: &str, tag: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!("elixcee_cli_snapshot_{}.bas", tag));
    fs::write(&path, vba).expect("write temp vba file");
    path
}

/// Runs `<vba>`'s `Main` and saves the result to a fresh workbook file
/// (`.xlsx` or `.ods`, picked by `ext`), returning its path — the same
/// `--output` round-trip `cli_json.rs` already uses.
fn build_workbook_fixture(vba: &str, tag: &str, ext: &str) -> std::path::PathBuf {
    let vba_path = write_vba(vba, tag);
    let out_path = std::env::temp_dir().join(format!("elixcee_cli_snapshot_{}.{}", tag, ext));
    let _ = fs::remove_file(&out_path);

    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            vba_path.as_os_str(),
            std::ffi::OsStr::new("Main"),
            std::ffi::OsStr::new("--output"),
            out_path.as_os_str(),
        ])
        .output()
        .expect("run elixcee binary to build fixture");
    assert!(
        output.status.success(),
        "failed to build workbook fixture: {:?}",
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(out_path.exists(), "--output file was not written");
    out_path
}

fn run_snapshot_json(path: &std::path::Path) -> (bool, Value) {
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(["snapshot", path.to_str().unwrap(), "--json"])
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).expect("stdout is utf8");
    let stderr = String::from_utf8(output.stderr).expect("stderr is utf8");
    assert_eq!(
        stdout.lines().count(),
        1,
        "snapshot --json must emit exactly one line, got: {:?} (stderr: {:?})",
        stdout,
        stderr,
    );
    let parsed: Value = serde_json::from_str(stdout.trim())
        .unwrap_or_else(|e| panic!("stdout was not valid JSON ({}): {:?}", e, stdout));
    (output.status.success(), parsed)
}

fn run_snapshot_markdown(path: &std::path::Path) -> (bool, String) {
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(["snapshot", path.to_str().unwrap()])
        .output()
        .expect("run elixcee binary");
    let stdout = String::from_utf8(output.stdout).expect("stdout is utf8");
    (output.status.success(), stdout)
}

#[test]
fn json_snapshot_of_a_single_sheet_workbook() {
    let path = build_workbook_fixture(
        "Sub Main()\n    Cells(1, 1).Value = 42\nEnd Sub\n",
        "single",
        "xlsx",
    );
    let (ok, v) = run_snapshot_json(&path);
    assert!(ok, "{:?}", v);
    assert_eq!(v["schema_version"], 1);
    assert_eq!(v["ok"], true);
    // The CLI-provided path is echoed verbatim; it's a temp path, so check
    // loosely rather than hard-coding it.
    assert!(v["file"].as_str().unwrap().ends_with(".xlsx"));

    let sheets = v["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 1);
    // The writer lowercases sheet names — confirmed by actually running it.
    assert_eq!(sheets[0]["name"], "sheet1");
    assert_eq!(sheets[0]["sheet_id"], "1");
    assert_eq!(sheets[0]["stable_id"], "sheet1");
    assert_eq!(sheets[0]["cells"][0]["address"], "A1");
    assert_eq!(sheets[0]["cells"][0]["value"], 42);
}

#[test]
fn json_snapshot_of_a_multi_sheet_workbook_has_distinct_stable_ids() {
    let path = build_workbook_fixture(
        "Sub Main()\n    \
            Cells(1, 1).Value = 1\n    \
            Sheets.Add\n    \
            Sheets(\"sheet2\").Cells(1, 1).Value = 2\n\
         End Sub\n",
        "multi",
        "xlsx",
    );
    let (ok, v) = run_snapshot_json(&path);
    assert!(ok, "{:?}", v);

    let sheets = v["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 2);
    let stable_ids: Vec<&str> = sheets
        .iter()
        .map(|s| s["stable_id"].as_str().unwrap())
        .collect();
    // Both sheets must get distinct stable_ids — the exact name<->id mapping
    // depends on the writer's sheet enumeration order (alphabetical by
    // display name, confirmed by direct code read), which this test
    // deliberately does not hard-code beyond "still distinct".
    assert_ne!(stable_ids[0], stable_ids[1]);
    assert_eq!(
        sheets
            .iter()
            .map(|s| s["cells"][0]["value"].as_i64().unwrap())
            .sum::<i64>(),
        3
    );
}

#[test]
fn markdown_snapshot_is_the_default_and_contains_expected_content() {
    let path = build_workbook_fixture(
        "Sub Main()\n    Cells(1, 1).Value = 42\nEnd Sub\n",
        "markdown",
        "xlsx",
    );
    let (ok, stdout) = run_snapshot_markdown(&path);
    assert!(ok, "{:?}", stdout);
    assert!(stdout.contains("# Workbook Snapshot:"));
    assert!(stdout.contains("| sheet1 | sheet1 | 1 |"));
    assert!(stdout.contains("| A1 | 42 |"));
}

#[test]
fn json_snapshot_of_an_ods_workbook_has_a_null_sheet_id_and_synthetic_stable_id() {
    // .ods has no attribute equivalent to XLSX's sheetId, so sheet_id must
    // always be null and stable_id always the synthetic positional form —
    // exercised end-to-end here, not just at the unit level, since this is
    // the one behavior this feature exists to distinguish from the .xlsx path.
    let path = build_workbook_fixture(
        "Sub Main()\n    Cells(1, 1).Value = 42\nEnd Sub\n",
        "ods",
        "ods",
    );
    let (ok, v) = run_snapshot_json(&path);
    assert!(ok, "{:?}", v);
    let sheets = v["sheets"].as_array().unwrap();
    assert_eq!(sheets.len(), 1);
    assert_eq!(sheets[0]["sheet_id"], Value::Null);
    assert_eq!(sheets[0]["stable_id"], "sheet1");
    assert_eq!(sheets[0]["cells"][0]["value"], 42);
}

#[test]
fn json_snapshot_of_a_missing_file_is_an_io_error() {
    let path = std::env::temp_dir().join("elixcee_cli_snapshot_does_not_exist.xlsx");
    let _ = fs::remove_file(&path);
    let (ok, v) = run_snapshot_json(&path);
    assert!(!ok, "{:?}", v);
    assert_eq!(v["ok"], false);
    assert_eq!(v["error"]["code"], "E3001");
    assert_eq!(v["error"]["kind"], "io_error");
}

#[test]
fn non_json_snapshot_of_a_missing_file_prints_to_stderr_and_exits_nonzero() {
    let path = std::env::temp_dir().join("elixcee_cli_snapshot_does_not_exist2.xlsx");
    let _ = fs::remove_file(&path);
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(["snapshot", path.to_str().unwrap()])
        .output()
        .expect("run elixcee binary");
    assert!(!output.status.success());
    let stderr = String::from_utf8_lossy(&output.stderr);
    assert!(stderr.starts_with("error: "));
}

#[test]
fn snapshot_of_an_unsupported_extension_is_an_io_error() {
    let path = std::env::temp_dir().join("elixcee_cli_snapshot_wrong_ext.txt");
    fs::write(&path, "not a workbook").unwrap();
    let (ok, v) = run_snapshot_json(&path);
    assert!(!ok, "{:?}", v);
    assert_eq!(v["error"]["code"], "E3001");
}
