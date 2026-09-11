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
use std::io::{BufRead, Cursor, Write};
use std::process::{Command, Stdio};
use std::thread;
use std::time::Duration;
use zip::write::{SimpleFileOptions, ZipWriter};

#[cfg(unix)]
unsafe extern "C" {
    fn kill(pid: i32, signal: i32) -> i32;
}

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

#[test]
fn snapshot_applies_the_reader_work_budget() {
    let path = build_workbook_fixture(
        "Sub Main()\n    Cells(1, 1).Value = 42\nEnd Sub\n",
        "budget",
        "xlsx",
    );
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            "snapshot",
            path.to_str().unwrap(),
            "--json",
            "--max-work-units",
            "1",
        ])
        .output()
        .expect("run elixcee binary");
    assert!(!output.status.success());
    let value: Value = serde_json::from_slice(&output.stdout).expect("JSON error output");
    assert_eq!(value["ok"], false);
    assert!(
        value["error"]["message"]
            .as_str()
            .unwrap()
            .contains("READER_WORK_BUDGET")
    );
}

#[test]
fn snapshot_honors_a_preexisting_cancel_file() {
    let path = build_workbook_fixture(
        "Sub Main()\n    Cells(1, 1).Value = 42\nEnd Sub\n",
        "cancel",
        "xlsx",
    );
    let cancel_file = std::env::temp_dir().join("elixcee_cli_snapshot_cancel.flag");
    fs::write(&cancel_file, "cancel").unwrap();
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            "snapshot",
            path.to_str().unwrap(),
            "--json",
            "--cancel-file",
            cancel_file.to_str().unwrap(),
        ])
        .output()
        .expect("run elixcee binary");
    let _ = fs::remove_file(&cancel_file);
    assert!(!output.status.success());
    let value: Value = serde_json::from_slice(&output.stdout).expect("JSON error output");
    assert_eq!(value["ok"], false);
    assert!(
        value["error"]["message"]
            .as_str()
            .unwrap()
            .contains("READER_CANCELED")
    );
}

#[cfg(unix)]
#[test]
fn snapshot_exits_with_cancellation_after_sigint() {
    let path = std::env::temp_dir().join("elixcee_cli_snapshot_sigint.xlsx");
    let mut sheet = String::from("<worksheet><sheetData>");
    for row in 1..=1_000_000 {
        sheet.push_str(&format!(
            "<row r=\"{row}\"><c r=\"A{row}\"><v>1</v></c></row>"
        ));
    }
    sheet.push_str("</sheetData></worksheet>");
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    zip.start_file("xl/workbook.xml", options).unwrap();
    zip.write_all(
        br#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#,
    )
    .unwrap();
    zip.start_file("xl/_rels/workbook.xml.rels", options)
        .unwrap();
    zip.write_all(br#"<Relationships><Relationship Id="rId1" Type="/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#).unwrap();
    zip.start_file("xl/worksheets/sheet1.xml", options).unwrap();
    zip.write_all(sheet.as_bytes()).unwrap();
    let bytes = zip.finish().unwrap().into_inner();
    fs::write(&path, bytes).unwrap();

    let mut child = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args(["snapshot", path.to_str().unwrap(), "--json"])
        .env("ELIXCEE_TEST_SIGNAL_READY", "1")
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .expect("spawn elixcee binary");
    let stderr = child.stderr.take().expect("capture child stderr");
    let (ready_sender, ready_receiver) = std::sync::mpsc::channel();
    thread::spawn(move || {
        let mut lines = std::io::BufReader::new(stderr).lines();
        while let Some(Ok(line)) = lines.next() {
            if line == "ELIXCEE_SIGNAL_READY" {
                let _ = ready_sender.send(());
                break;
            }
        }
    });
    if ready_receiver.recv_timeout(Duration::from_secs(2)).is_err() {
        let _ = child.kill();
        panic!("CLI did not install its signal handler before the deadline");
    }
    assert_eq!(unsafe { kill(child.id() as i32, 2) }, 0);
    let output = child.wait_with_output().expect("wait for SIGINT child");
    let _ = fs::remove_file(&path);
    eprintln!(
        "SIGINT status={:?} stdout={} stderr={}",
        output.status,
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    assert!(!output.status.success());
    let value: Value = serde_json::from_slice(&output.stdout).expect("JSON cancellation output");
    assert_eq!(value["ok"], false);
    assert!(
        value["error"]["message"]
            .as_str()
            .unwrap()
            .contains("READER_CANCELED")
    );
}
