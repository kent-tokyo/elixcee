#!/usr/bin/env bash
# Fails (non-zero exit) if the root Cargo.toml package version and
# pyproject.toml project version disagree — the two are versioned
# independently (maturin builds the Python wheel from pyproject.toml,
# cargo/crates.io from Cargo.toml), so nothing else catches this drift.
set -euo pipefail

cd "$(dirname "$0")/.."

cargo_version=$(grep -m1 '^version = "' Cargo.toml | sed -E 's/^version = "(.*)"$/\1/')
pyproject_version=$(grep -m1 '^version = "' pyproject.toml | sed -E 's/^version = "(.*)"$/\1/')

if [[ -z "$cargo_version" ]]; then
  echo "check-versions: could not find a version = \"...\" line in Cargo.toml" >&2
  exit 1
fi
if [[ -z "$pyproject_version" ]]; then
  echo "check-versions: could not find a version = \"...\" line in pyproject.toml" >&2
  exit 1
fi

if [[ "$cargo_version" != "$pyproject_version" ]]; then
  echo "check-versions: version mismatch — Cargo.toml version=$cargo_version, pyproject.toml version=$pyproject_version" >&2
  exit 1
fi

echo "check-versions: OK ($cargo_version)"
