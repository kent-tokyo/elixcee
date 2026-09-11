#!/usr/bin/env bash
# Run the evidence-producing gates that are self-contained on the current host.
# Deliberately excludes Excel/LibreOffice oracles, competitor comparisons, other OS
# installs, network fetches, GitHub Actions, and registry publication.
set -euo pipefail

ROOT=$(cd "$(dirname "$0")/.." && pwd)
cd "$ROOT"

run() {
  echo "[local-gates] $*"
  "$@"
}

run bash scripts/check-versions.sh
run bash scripts/check-measurement-boundary.sh
run env PYTHONDONTWRITEBYTECODE=1 python3 scripts/check-formula-dispatch.py --check-contracts --check-docs
run env PYTHONDONTWRITEBYTECODE=1 python3 scripts/check-ooxml-feature-matrix.py
run env TMPDIR=/private/tmp PYTHONDONTWRITEBYTECODE=1 python3 scripts/check-stream-writer-measurements.py --self-test
run git diff --check
run cargo fmt --check
run cargo test --offline --workspace --all-targets --quiet
run cargo clippy --offline --workspace --all-targets -- -D warnings
run cargo doc --offline --workspace --no-deps --features python --document-private-items
run cargo audit --no-fetch --stale

if command -v cargo-fuzz >/dev/null 2>&1 && command -v rustup >/dev/null 2>&1 \
  && rustup toolchain list | grep -q '^nightly'; then
  run cargo +nightly fuzz run fuzz_formula_parser -- -max_total_time=5 -rss_limit_mb=1024
  run cargo +nightly fuzz run fuzz_formula_eval -- -max_total_time=5 -rss_limit_mb=1024
  run cargo +nightly fuzz run fuzz_vba_parser -- -max_total_time=5 -rss_limit_mb=1024
  run cargo +nightly fuzz run fuzz_xlsx_reader -- -max_total_time=5 -rss_limit_mb=1024
else
  echo "[local-gates] fuzz smoke skipped: nightly cargo-fuzz toolchain unavailable" >&2
  exit 2
fi

pushd packages/xlsx >/dev/null
run npm run typecheck
run npm run typecheck:no-dom
run node test/operation-plan.test.mjs
run npm run audit:pack
run npm run wasm:smoke
run npm run pack:consumer
run npm run browser:smoke
popd >/dev/null

echo "[local-gates] all self-contained gates passed"
