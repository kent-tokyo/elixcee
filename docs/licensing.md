# Licensing

## elixcee's license

elixcee (this repository, including any future `@elixcee/xlsx` package) is **MIT**
licensed (see repository root `LICENSE` / `Cargo.toml` / `pyproject.toml`).

## The compat target's license chain

`xlsx@0.18.5` (SheetJS) and every one of its 7 runtime dependencies are **Apache-2.0**,
confirmed via `npm view <pkg> license` against the live public registry:

| Package | Version resolved | License |
|---|---|---|
| `xlsx` | 0.18.5 | Apache-2.0 |
| `adler-32` | 1.3.1 | Apache-2.0 |
| `cfb` | 1.2.2 | Apache-2.0 |
| `codepage` | 1.15.0 | Apache-2.0 |
| `crc-32` | 1.2.2 | Apache-2.0 |
| `ssf` | 0.11.2 | Apache-2.0 |
| `wmf` | 1.0.2 | Apache-2.0 |
| `word` | 0.4.0 | Apache-2.0 |

All 8 packages resolve from the public npm registry today (verified live during Phase 0
planning) — no vendored-tarball fallback is needed for `compat/oracle` to install and run
the real package as a `devDependency`.

## Obligation if code or text is ever ported

Apache-2.0 is permissive but not license-free: Apache License 2.0 §4(b) requires that any
modified files carry a prominent notice stating they were changed, and §4(d) requires
preserving a readable copy of any NOTICE-file attribution content in redistributed works.
If a future phase ever copies actual logic, algorithms, or text from `xlsx` or any of its
dependencies (as opposed to independently reimplementing equivalent behavior from
observed input/output pairs), that specific code must retain its Apache-2.0 license and
attribution — it cannot simply become MIT by virtue of living in this repository.

## Current status

**No SheetJS code has been vendored or ported as of Phase 0.**
[`compat/oracle`](../compat/oracle) installs and runs the real `xlsx` package as an
ordinary `devDependency` for introspection and differential testing. That is consumption
(using the package as intended, via its public API), not redistribution, and carries no
NOTICE obligation by itself.

If a later phase ever vendors a tarball or CDN copy for install-time redistribution
(rather than an ordinary `npm install` at build/test time), that **does** trigger
redistribution obligations and this document must be expanded — with a full per-package
NOTICE inventory — before doing so.

A complete, per-package license/NOTICE inventory (beyond the license identifiers above)
is deferred to whichever phase first considers porting or vendoring actual SheetJS logic.
