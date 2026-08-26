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

# Root Cargo.toml pins an exact elixcee-types version (cargo publish's own
# verification build resolves it from the registry, not the local workspace path —
# see ROADMAP.md's "Packaging note" for the release this caught). Catches the pin
# drifting from the workspace member's actual version; doesn't catch the member's
# code changing without its version being bumped at all.
types_toml="crates/elixcee-types/Cargo.toml"
if [[ -f "$types_toml" ]]; then
  types_version=$(grep -m1 '^version = "' "$types_toml" | sed -E 's/^version = "(.*)"$/\1/')
  pinned_types_version=$(grep -m1 'elixcee-types = ' Cargo.toml | sed -E 's/.*version = "([^"]*)".*/\1/')

  if [[ -n "$pinned_types_version" && "$types_version" != "$pinned_types_version" ]]; then
    echo "check-versions: elixcee-types version mismatch — crates/elixcee-types/Cargo.toml version=$types_version, root Cargo.toml pins version=$pinned_types_version" >&2
    exit 1
  fi
fi

# @elixcee/xlsx versions independently of the root crate (see ROADMAP.md), so this
# doesn't cross-check its version against Cargo.toml/pyproject.toml — it only guards
# the one concrete drift this project has actually hit: "private" flipped to false
# (publish-ready) while "version" is still the 0.0.0-development placeholder nobody
# meant to actually publish.
xlsx_pkg="packages/xlsx/package.json"
if [[ -f "$xlsx_pkg" ]]; then
  xlsx_version=$(grep -m1 '"version":' "$xlsx_pkg" | sed -E 's/.*"version": *"([^"]*)".*/\1/')
  xlsx_private=$(grep -m1 '"private":' "$xlsx_pkg" | sed -E 's/.*"private": *([a-z]+).*/\1/')

  if [[ "$xlsx_private" == "false" && "$xlsx_version" == "0.0.0-development" ]]; then
    echo "check-versions: packages/xlsx/package.json has \"private\": false but \"version\" is still the 0.0.0-development placeholder — set a real version before publishing" >&2
    exit 1
  fi
fi

echo "check-versions: OK ($cargo_version)"
