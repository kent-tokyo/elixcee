/// Integration test for `elixcee --version`/`-V`, mirroring the
/// `Command::new(env!("CARGO_BIN_EXE_elixcee"))` pattern already used in
/// `tests/cli_snapshot.rs`/`tests/cli_check.rs`.
use std::process::Command;

fn run(flag: &str) -> String {
    let output = Command::new(env!("CARGO_BIN_EXE_elixcee"))
        .arg(flag)
        .output()
        .expect("run elixcee binary");
    assert!(
        output.status.success(),
        "elixcee {flag} exited non-zero: {:?}",
        output
    );
    String::from_utf8(output.stdout).expect("utf8 stdout")
}

#[test]
fn version_flag_prints_the_crate_version() {
    let stdout = run("--version");
    assert_eq!(
        stdout.trim(),
        format!("elixcee {}", env!("CARGO_PKG_VERSION"))
    );
}

#[test]
fn short_version_flag_matches_the_long_form() {
    assert_eq!(run("-V"), run("--version"));
}
