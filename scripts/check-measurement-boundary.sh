#!/usr/bin/env bash
set -euo pipefail

manifest="Cargo.toml"
if ! grep -Fq '"compat/benchmarks/"' "$manifest"; then
  echo "measurement boundary: benchmark dependencies must be excluded from package" >&2
  exit 1
fi
for path in \
  examples/bench_workbook.rs \
  src/bin/measure_formula_dirty.rs \
  src/bin/measure_reader_inprocess.rs \
  src/bin/measure_reader_vm_load.rs \
  src/bin/measure_reader_write_inprocess.rs; do
  if ! grep -Fq "\"$path\"" "$manifest"; then
    echo "measurement boundary: missing package exclusion for $path" >&2
    exit 1
  fi
  if [[ ! -f "$path" ]]; then
    echo "measurement boundary: expected local helper is missing: $path" >&2
    exit 1
  fi
done

echo "measurement boundary: OK (local helpers excluded from package)"
