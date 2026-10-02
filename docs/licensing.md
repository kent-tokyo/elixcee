# Licensing

## elixcee's license

elixcee (this repository, including any future `@elixcee/xlsx` package) is **MIT**
licensed as declared in [Cargo.toml](../Cargo.toml) and [pyproject.toml](../pyproject.toml).
The JS package carries the [MIT license text](../packages/xlsx/LICENSE).

## The compat target's license chain

The fixed `hyperformula@3.4.0` package is a development-only oracle. The probe
uses its `gpl-v3` license-key mode and does not bundle, import, or expose it from
elixcee's runtime. Do not copy HyperFormula code into the MIT-licensed product.

The private compatibility harness installs the official SheetJS 0.20.3 tarball
from `cdn.sheetjs.com`. The resolved package is **Apache-2.0** and declares no
runtime dependencies:

| Package | Version resolved | License |
|---|---|---|
| `xlsx` | 0.20.3 | Apache-2.0 |

This is a development-only dependency record, not legal advice. Recheck the
resolved package contents and notices before distribution.

## Obligation if code or text is ever ported

Apache-2.0 is permissive but not license-free: Apache License 2.0 §4(b) requires that any
modified files carry a prominent notice stating they were changed, and §4(d) requires
preserving a readable copy of any NOTICE-file attribution content in redistributed works.
If a future phase ever copies actual logic, algorithms, or text from `xlsx` or any of its
dependencies (as opposed to independently reimplementing equivalent behavior from
observed input/output pairs), that specific code must retain its Apache-2.0 license and
attribution — it cannot simply become MIT by virtue of living in this repository.

## Current status

**No SheetJS code has been vendored or ported into `@elixcee/xlsx`'s own source.**
[`compat/oracle`](../compat/oracle) installs and runs the real `xlsx` package from
SheetJS's official CDN as a `devDependency` for introspection and differential
testing. It is not included in elixcee runtime packages.

**As of Phase 1B-2B, `packages/xlsx` takes its first real *runtime* dependency**:
`ssf@0.11.2` (Apache-2.0, transitively pulling in `frac@1.1.2`, also Apache-2.0), used to
back `format_cell`/`sheet_to_csv`/`sheet_to_txt`'s number-format rendering — see
[package README](../packages/xlsx/README.md) for the current runtime boundary.
This is an ordinary npm `dependencies` declaration — `ssf`/`frac`'s
own source is never copied into or embedded in this repository's files — but it does mean
consumers of a future published package would transitively receive `ssf`/`frac`'s actual
package contents in their own `node_modules`, which is the normal npm dependency pattern,
not a §4(b)/(d) "vendored redistribution" case. Neither `ssf` nor `frac` ships an upstream
`NOTICE` file (checked directly, not assumed — only a `LICENSE`), so the §4(d) NOTICE-
propagation obligation is not actually triggered; as a conservative practice regardless,
their license text and package identity are recorded in
[`THIRD_PARTY_NOTICES.md`](../THIRD_PARTY_NOTICES.md) at the repository root.

If a later phase ever vendors a tarball or CDN copy for install-time redistribution
(rather than an ordinary `npm install`/`dependencies` declaration), that **does** trigger
full redistribution obligations and this document must be expanded — with a complete
per-package NOTICE inventory — before doing so.
