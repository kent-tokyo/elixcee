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

# The pin/member check above only catches the two numbers disagreeing with each
# other — it can't catch elixcee-types' own source drifting away from what's
# actually live on crates.io at that shared version number, which is the gap
# that broke `cargo publish -p elixcee` during 0.11.0/0.12.0's release prep (see
# ROADMAP.md's "Packaging note"). Guard it with a committed source hash, the
# same pattern as crates/elixcee-wasm/wasm-size-baseline.json: fails if
# crates/elixcee-types/src/lib.rs has changed since it was last actually
# published, until a human/agent deliberately re-versions and regenerates the
# hash — exactly the explicit-remembering ROADMAP.md says this needs.
types_hash_file="crates/elixcee-types/PUBLISHED_HASH"
types_lib="crates/elixcee-types/src/lib.rs"
if [[ -f "$types_hash_file" && -f "$types_lib" ]]; then
  if command -v sha256sum >/dev/null 2>&1; then
    types_current_hash=$(sha256sum "$types_lib" | awk '{print $1}')
  else
    types_current_hash=$(shasum -a 256 "$types_lib" | awk '{print $1}')
  fi
  types_recorded_hash=$(grep -v '^#' "$types_hash_file" | tr -d '[:space:]')

  if [[ "$types_current_hash" != "$types_recorded_hash" ]]; then
    echo "check-versions: $types_lib has changed since it was last published at elixcee-types $types_version (recorded in $types_hash_file) — bump crates/elixcee-types/Cargo.toml's version (and the root Cargo.toml pin) before publishing elixcee off this source, then regenerate $types_hash_file once elixcee-types is actually republished" >&2
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
