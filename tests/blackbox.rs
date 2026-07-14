/// Deterministic black-box tests (Milestone B3): a generic harness that
/// reads declarative `.toml` fixtures from `tests/fixtures/blackbox/` and
/// asserts the built `elixcee` binary's real `--json` stdout equals a fixed
/// expected JSON blob. This is additive to (not a replacement for)
/// `tests/cli_json.rs`/`tests/cli_check.rs`'s hand-written, per-scenario
/// assertions — this suite is for full-contract regression snapshots that
/// are cheap to add (drop a `.toml` file, no Rust required).
///
/// Fixture schema:
/// ```toml
/// name = "unique_fixture_name"
/// command = "run"          # or "check"
/// files = ["<VBA source 1>", "<VBA source 2>", ...]
/// args = ["Main"]           # appended verbatim after the file paths, before --json
/// expect_exit_code = 0
/// expect_json = """<the exact JSON stdout the binary should print>"""
/// ```
///
/// Two known-nondeterministic bits are normalized out before comparison, so
/// fixtures must follow these conventions:
/// - `duration_ms` is stripped entirely (see `normalize`) — omit it from
///   `expect_json`.
/// - Every temp file path the harness itself wrote (from `files`) is
///   replaced with the literal placeholder `<FILE>` in the raw stdout text
///   *before* JSON-parsing (see `scrub_paths`) — this covers not just
///   `error.location.file`/a `check` diagnostic's `location.file`, but also
///   the rarer case of a path embedded directly in a `message` string (e.g.
///   `E1006`'s duplicate-module-name message names the offending file).
///   Write `<FILE>` literally in `expect_json` wherever a real path would
///   otherwise appear.
///
/// Hard rule: no fixture may assert the relative order of *multiple*
/// cross-module Sub/Function collisions in one run. `find_cross_module_sub_
/// collisions`/`_func_collisions` (`src/parser/mod.rs`) end in a `HashMap`
/// iteration, so with two or more distinct collisions in a single
/// invocation, both which one run-mode's `.first()` reports and the order
/// of `check`'s per-collision `E1005` diagnostics are process-seed-dependent.
/// A fixture with exactly one collision (the only kind in this suite) is
/// unaffected.
use serde_json::Value;
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;

struct Fixture {
    name: String,
    command: String,
    files: Vec<String>,
    args: Vec<String>,
    expect_exit_code: i64,
    expect_json: String,
}

fn load_fixture(path: &Path) -> Fixture {
    let text = fs::read_to_string(path)
        .unwrap_or_else(|e| panic!("read fixture {}: {}", path.display(), e));
    let doc: toml::Value = text
        .parse()
        .unwrap_or_else(|e| panic!("parse fixture {}: {}", path.display(), e));

    let field_str = |key: &str| -> String {
        doc.get(key)
            .and_then(|v| v.as_str())
            .unwrap_or_else(|| panic!("{}: missing/non-string '{}'", path.display(), key))
            .to_string()
    };
    let string_array = |key: &str| -> Vec<String> {
        doc.get(key)
            .and_then(|v| v.as_array())
            .unwrap_or_else(|| panic!("{}: missing/non-array '{}'", path.display(), key))
            .iter()
            .map(|v| {
                v.as_str()
                    .unwrap_or_else(|| {
                        panic!("{}: '{}' entries must be strings", path.display(), key)
                    })
                    .to_string()
            })
            .collect()
    };

    Fixture {
        name: field_str("name"),
        command: field_str("command"),
        files: string_array("files"),
        args: doc
            .get("args")
            .map(|_| string_array("args"))
            .unwrap_or_default(),
        expect_exit_code: doc
            .get("expect_exit_code")
            .and_then(|v| v.as_integer())
            .unwrap_or_else(|| {
                panic!("{}: missing/non-integer 'expect_exit_code'", path.display())
            }),
        expect_json: field_str("expect_json"),
    }
}

fn write_files(name: &str, sources: &[String]) -> Vec<PathBuf> {
    sources
        .iter()
        .enumerate()
        .map(|(i, src)| {
            let path = std::env::temp_dir().join(format!("elixcee_blackbox_{}_{}.bas", name, i));
            fs::write(&path, src)
                .unwrap_or_else(|e| panic!("write temp fixture file {}: {}", path.display(), e));
            path
        })
        .collect()
}

/// Strips `duration_ms` keys — see the module doc comment for why.
fn normalize(v: &mut Value) {
    match v {
        Value::Object(map) => {
            map.remove("duration_ms");
            for val in map.values_mut() {
                normalize(val);
            }
        }
        Value::Array(arr) => arr.iter_mut().for_each(normalize),
        _ => {}
    }
}

/// Replaces every occurrence of a written temp file's path with `<FILE>` in
/// raw text, before JSON-parsing — see the module doc comment for why this
/// is done at the text level rather than by JSON key name.
fn scrub_paths(text: &str, paths: &[PathBuf]) -> String {
    let mut out = text.to_string();
    for p in paths {
        out = out.replace(p.to_string_lossy().as_ref(), "<FILE>");
    }
    out
}

fn run_one(fixture_path: &Path) -> Result<(), String> {
    let fixture = load_fixture(fixture_path);
    let file_paths = write_files(&fixture.name, &fixture.files);

    let mut cmd = Command::new(env!("CARGO_BIN_EXE_elixcee"));
    match fixture.command.as_str() {
        "check" => {
            cmd.arg("check");
        }
        "run" => {}
        other => return Err(format!("{}: unknown command '{}'", fixture.name, other)),
    }
    cmd.args(file_paths.iter().map(|p| p.as_os_str()));
    cmd.args(&fixture.args);
    cmd.arg("--json");

    let output = cmd
        .output()
        .map_err(|e| format!("{}: failed to run binary: {}", fixture.name, e))?;
    let stdout = String::from_utf8_lossy(&output.stdout).to_string();
    let stderr = String::from_utf8_lossy(&output.stderr).to_string();

    let exit_code = output.status.code().unwrap_or(-1) as i64;
    if exit_code != fixture.expect_exit_code {
        return Err(format!(
            "{}: expected exit code {}, got {} (stdout: {:?}, stderr: {:?})",
            fixture.name, fixture.expect_exit_code, exit_code, stdout, stderr
        ));
    }
    if stdout.lines().count() != 1 {
        return Err(format!(
            "{}: expected exactly one line of stdout, got: {:?} (stderr: {:?})",
            fixture.name, stdout, stderr
        ));
    }

    let scrubbed_stdout = scrub_paths(stdout.trim(), &file_paths);
    let mut actual: Value = serde_json::from_str(&scrubbed_stdout).map_err(|e| {
        format!(
            "{}: stdout was not valid JSON ({}): {:?}",
            fixture.name, e, stdout
        )
    })?;
    let mut expected: Value = serde_json::from_str(fixture.expect_json.trim()).map_err(|e| {
        format!(
            "{}: expect_json was not valid JSON ({}): {:?}",
            fixture.name, e, fixture.expect_json
        )
    })?;
    normalize(&mut actual);
    normalize(&mut expected);

    if actual != expected {
        return Err(format!(
            "{}: JSON mismatch\n  expected: {}\n  actual:   {}",
            fixture.name, expected, actual
        ));
    }
    Ok(())
}

#[test]
fn all_blackbox_fixtures_match_expected_json() {
    let dir = Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/blackbox");
    let mut fixture_paths: Vec<PathBuf> = fs::read_dir(&dir)
        .unwrap_or_else(|e| panic!("read fixtures dir {}: {}", dir.display(), e))
        .filter_map(|entry| entry.ok().map(|e| e.path()))
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("toml"))
        .collect();
    fixture_paths.sort();

    assert!(
        !fixture_paths.is_empty(),
        "no fixtures found under {}",
        dir.display()
    );

    let failures: Vec<String> = fixture_paths
        .iter()
        .filter_map(|p| run_one(p).err())
        .collect();

    assert!(
        failures.is_empty(),
        "{} of {} blackbox fixtures failed:\n\n{}",
        failures.len(),
        fixture_paths.len(),
        failures.join("\n\n")
    );
}
