#!/usr/bin/env bash
set -euo pipefail

manifest="Cargo.toml"
for path in \
  src/bin/measure_reader_inprocess.rs \
  src/bin/measure_reader_vm_load.rs \
  src/bin/measure_reader_write_inprocess.rs; do
  if ! rg -Fq "\"$path\"" "$manifest"; then
    echo "measurement boundary: missing package exclusion for $path" >&2
    exit 1
  fi
  if [[ ! -f "$path" ]]; then
    echo "measurement boundary: expected local helper is missing: $path" >&2
    exit 1
  fi
done

echo "measurement boundary: OK (local helpers excluded from package)"
