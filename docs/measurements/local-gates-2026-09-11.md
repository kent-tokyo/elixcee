# Self-contained local gates — 2026-09-11

The current `release/0.21.0` checkout passed
`bash scripts/check-local-gates.sh` on macOS arm64. The gate covered version
and measurement-boundary checks, formula dispatch documentation (90 functions
and 534 names), the OOXML feature matrix, the stream-measurement validator,
Rust formatting, 1,745 workspace/all-target tests, strict workspace clippy,
private Rustdoc, and offline Cargo audit.

The four nightly fuzz smoke targets also completed without a panic under the
configured 5-second / 1 GiB RSS limit. The run completed the package TypeScript
checks, operation-plan tests, package audit, WASM smoke, packed-consumer smoke,
and real Chrome browser smoke.

This is current-host evidence for self-contained gates. It does not establish
Linux/Windows behavior, Microsoft Excel oracle agreement, Excel Chart reopen
compatibility, or external-library comparison results. The fuzz build may
rewrite `fuzz/Cargo.lock`; that generated change was discarded after the run.
