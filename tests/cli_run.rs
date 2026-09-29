//! End-to-end coverage for the headless CLI path: VBA execution, recalculation,
//! JSON projection, and atomic workbook output must describe the same result.

use serde_json::Value;
use std::fs;
use std::process::Command;

#[test]
fn run_recalculates_formula_and_writes_the_requested_workbook() {
    let suffix = std::process::id();
    let directory = std::env::temp_dir().join(format!("elixcee-cli-run-{suffix}"));
    fs::create_dir_all(&directory).unwrap();
    let source = directory.join("main.bas");
    let output_path = directory.join("result.xlsx");
    fs::write(
        &source,
        "Sub Main()\n    Cells(1,1).Value = 10\n    Range(\"B1\").Formula = \"=A1*2\"\nEnd Sub\n",
    )
    .unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            source.to_str().unwrap(),
            "Main",
            "--output",
            output_path.to_str().unwrap(),
            "--json",
        ])
        .output()
        .unwrap();
    let stdout = String::from_utf8(output.stdout).unwrap();
    let stderr = String::from_utf8(output.stderr).unwrap();
    assert!(
        output.status.success(),
        "stdout={stdout:?} stderr={stderr:?}"
    );
    let json: Value = serde_json::from_str(stdout.trim()).unwrap();
    assert_eq!(json["ok"], true);
    assert_eq!(json["termination_class"], "success");
    assert_eq!(json["cells"][0]["address"], "A1");
    assert_eq!(json["cells"][0]["value"], 10);
    assert_eq!(json["cells"][1]["address"], "B1");
    assert_eq!(json["cells"][1]["value"], 20);
    assert!(output_path.is_file(), "CLI did not publish output workbook");

    let mut workbook = calamine::open_workbook::<calamine::Xlsx<_>, _>(&output_path).unwrap();
    let sheet = calamine::Reader::worksheet_range(&mut workbook, "sheet1").unwrap();
    assert_eq!(sheet.get((0, 0)).unwrap().to_string(), "10");
    assert_eq!(sheet.get((0, 1)).unwrap().to_string(), "20");
    let _ = fs::remove_dir_all(directory);
}

#[test]
fn trace_includes_recalculation_and_verified_save_boundaries() {
    let suffix = std::process::id();
    let directory = std::env::temp_dir().join(format!("elixcee-cli-trace-boundary-{suffix}"));
    fs::create_dir_all(&directory).unwrap();
    let source = directory.join("main.bas");
    let output_path = directory.join("result.xlsx");
    fs::write(&source, "Sub Main()\n    Cells(1,1).Value = 10\nEnd Sub\n").unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            source.to_str().unwrap(),
            "Main",
            "--output",
            output_path.to_str().unwrap(),
            "--trace",
            "boundary-test",
            "--json",
        ])
        .output()
        .unwrap();
    let stdout = String::from_utf8(output.stdout).unwrap();
    let stderr = String::from_utf8(output.stderr).unwrap();
    assert!(
        output.status.success(),
        "stdout={stdout:?} stderr={stderr:?}"
    );
    let json: Value = serde_json::from_str(stdout.trim()).unwrap();
    let kinds: Vec<_> = json["trace"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|event| event["kind"].as_str())
        .collect();
    let position = |kind: &str| kinds.iter().position(|value| *value == kind).unwrap();
    assert!(position("recalculate_start") < position("recalculate_end"));
    assert!(position("recalculate_end") < position("save_start"));
    assert!(position("save_start") < position("save_end"));
    assert!(output_path.is_file());
    let _ = fs::remove_dir_all(directory);
}

#[test]
fn failed_macro_does_not_replace_an_existing_output_workbook() {
    let suffix = std::process::id();
    let directory = std::env::temp_dir().join(format!("elixcee-cli-run-failure-{suffix}"));
    fs::create_dir_all(&directory).unwrap();
    let source = directory.join("failure.bas");
    let output_path = directory.join("result.xlsx");
    fs::write(
        &source,
        "Sub Main()\n    Cells(1,1).Value = 99\n    Err.Raise 5\nEnd Sub\n",
    )
    .unwrap();
    fs::write(&output_path, b"existing output").unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            source.to_str().unwrap(),
            "Main",
            "--output",
            output_path.to_str().unwrap(),
            "--json",
        ])
        .output()
        .unwrap();
    let stdout = String::from_utf8(output.stdout).unwrap();
    let stderr = String::from_utf8(output.stderr).unwrap();
    assert!(
        !output.status.success(),
        "stdout={stdout:?} stderr={stderr:?}"
    );
    let json: Value = serde_json::from_str(stdout.trim()).unwrap();
    assert_eq!(json["ok"], false);
    assert_eq!(json["termination_class"], "runtime_error");
    assert_eq!(json["error"]["kind"], "runtime_error");
    assert_eq!(fs::read(&output_path).unwrap(), b"existing output");
    let _ = fs::remove_dir_all(directory);
}

#[test]
fn save_failure_does_not_replace_an_existing_output_directory() {
    let suffix = std::process::id();
    let directory = std::env::temp_dir().join(format!("elixcee-cli-save-failure-{suffix}"));
    fs::create_dir_all(&directory).unwrap();
    let source = directory.join("main.bas");
    let output_path = directory.join("result.xlsx");
    fs::write(&source, "Sub Main()\n    Cells(1,1).Value = 99\nEnd Sub\n").unwrap();
    fs::create_dir(&output_path).unwrap();
    let sentinel = output_path.join("sentinel");
    fs::write(&sentinel, b"existing directory").unwrap();

    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            source.to_str().unwrap(),
            "Main",
            "--output",
            output_path.to_str().unwrap(),
            "--json",
        ])
        .output()
        .unwrap();
    let stdout = String::from_utf8(output.stdout).unwrap();
    let stderr = String::from_utf8(output.stderr).unwrap();
    assert!(
        !output.status.success(),
        "stdout={stdout:?} stderr={stderr:?}"
    );
    let json: Value = serde_json::from_str(stdout.trim()).unwrap();
    assert_eq!(json["ok"], false);
    assert_eq!(json["error"]["kind"], "io_error");
    assert_eq!(fs::read(&sentinel).unwrap(), b"existing directory");
    let _ = fs::remove_dir_all(directory);
}

#[test]
fn failure_injection_matrix_preserves_outputs_and_next_job_is_clean() {
    const CASES: usize = 100;
    let suffix = std::process::id();
    let directory = std::env::temp_dir().join(format!("elixcee-cli-failure-matrix-{suffix}"));
    fs::create_dir_all(&directory).unwrap();
    let source = directory.join("failure.bas");
    fs::write(
        &source,
        "Sub Main()\n    Cells(1,1).Value = 100\n    Err.Raise 5\nEnd Sub\n",
    )
    .unwrap();

    for case in 0..CASES {
        let output_path = directory.join(format!("result-{case}.xlsx"));
        let sentinel = format!("existing output {case}");
        fs::write(&output_path, sentinel.as_bytes()).unwrap();
        let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
            .args([
                source.to_str().unwrap(),
                "Main",
                "--output",
                output_path.to_str().unwrap(),
                "--json",
            ])
            .output()
            .unwrap();
        let stdout = String::from_utf8(output.stdout).unwrap();
        let stderr = String::from_utf8(output.stderr).unwrap();
        assert!(
            !output.status.success(),
            "case={case} stdout={stdout:?} stderr={stderr:?}"
        );
        let json: Value = serde_json::from_str(stdout.trim()).unwrap();
        assert_eq!(json["ok"], false, "case={case} json={json}");
        assert_eq!(
            json["error"]["kind"], "runtime_error",
            "case={case} json={json}"
        );
        assert_eq!(
            fs::read(&output_path).unwrap(),
            sentinel.as_bytes(),
            "case={case}"
        );
    }

    let success_source = directory.join("success.bas");
    let success_output = directory.join("success.xlsx");
    fs::write(
        &success_source,
        "Sub Main()\n    Cells(1,1).Value = 7\nEnd Sub\n",
    )
    .unwrap();
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            success_source.to_str().unwrap(),
            "Main",
            "--output",
            success_output.to_str().unwrap(),
            "--json",
        ])
        .output()
        .unwrap();
    let stdout = String::from_utf8(output.stdout).unwrap();
    let stderr = String::from_utf8(output.stderr).unwrap();
    assert!(
        output.status.success(),
        "stdout={stdout:?} stderr={stderr:?}"
    );
    let json: Value = serde_json::from_str(stdout.trim()).unwrap();
    assert_eq!(json["ok"], true);
    assert_eq!(json["cells"][0]["value"], 7);
    assert!(success_output.is_file());
    let _ = fs::remove_dir_all(directory);
}

#[test]
fn preexisting_cancel_file_stops_run_without_publishing_output() {
    const CASES: usize = 100;
    let suffix = std::process::id();
    let directory = std::env::temp_dir().join(format!("elixcee-cli-cancel-matrix-{suffix}"));
    fs::create_dir_all(&directory).unwrap();
    let source = directory.join("main.bas");
    let cancel_file = directory.join("cancel.flag");
    fs::write(&source, "Sub Main()\n    Cells(1,1).Value = 7\nEnd Sub\n").unwrap();
    fs::write(&cancel_file, "cancel").unwrap();

    for case in 0..CASES {
        let output_path = directory.join(format!("result-{case}.xlsx"));
        let sentinel = format!("existing output {case}");
        fs::write(&output_path, sentinel.as_bytes()).unwrap();
        let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
            .args([
                source.to_str().unwrap(),
                "Main",
                "--output",
                output_path.to_str().unwrap(),
                "--cancel-file",
                cancel_file.to_str().unwrap(),
                "--json",
            ])
            .output()
            .unwrap();
        let stdout = String::from_utf8(output.stdout).unwrap();
        let stderr = String::from_utf8(output.stderr).unwrap();
        assert!(
            !output.status.success(),
            "case={case} stdout={stdout:?} stderr={stderr:?}"
        );
        let json: Value = serde_json::from_str(stdout.trim()).unwrap();
        assert_eq!(json["ok"], false, "case={case} json={json}");
        assert!(
            json["error"]["message"]
                .as_str()
                .unwrap_or_default()
                .starts_with("CANCELED:"),
            "case={case} json={json}"
        );
        assert_eq!(
            fs::read(&output_path).unwrap(),
            sentinel.as_bytes(),
            "case={case}"
        );
    }
    let _ = fs::remove_dir_all(directory);
}

#[test]
fn headless_workflow_fixture_runs_against_its_declared_workbook() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"));
    let source = root.join("compat/vba-workflows/fixtures/Transfer.bas");
    let workbook = root.join("compat/corpus/workbooks/numeric_grid.xlsx");
    let suffix = std::process::id();
    let output_path = std::env::temp_dir().join(format!("elixcee-workflow-{suffix}.xlsx"));

    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .args([
            source.to_str().unwrap(),
            "TransferAndRecalculate",
            "--file",
            workbook.to_str().unwrap(),
            "--output",
            output_path.to_str().unwrap(),
            "--json",
        ])
        .output()
        .unwrap();
    let stdout = String::from_utf8(output.stdout).unwrap();
    let stderr = String::from_utf8(output.stderr).unwrap();
    assert!(
        output.status.success(),
        "stdout={stdout:?} stderr={stderr:?}"
    );
    let json: Value = serde_json::from_str(stdout.trim()).unwrap();
    let cells = json["cells"].as_array().unwrap();
    let cell_value = |address: &str| {
        cells
            .iter()
            .find(|cell| cell["address"] == address)
            .and_then(|cell| cell["value"].as_i64())
    };
    assert_eq!(cell_value("B2"), Some(30));
    assert_eq!(cell_value("B3"), Some(60));
    assert!(output_path.is_file());
    let _ = fs::remove_file(output_path);
}
