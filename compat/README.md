# compat/

A Node.js project living inside an otherwise-Rust repository, for one reason: the
`@elixcee/xlsx` compatibility initiative needs to run the real `xlsx@0.18.5` npm package
(the "oracle") to know what it actually does, rather than guessing from documentation.
See [`docs/xlsx-compatibility-goal.md`](../docs/xlsx-compatibility-goal.md) and
[`docs/xlsx-architecture.md`](../docs/xlsx-architecture.md) for the why; this file is
just the how.

## Setup

```sh
cd compat
npm install
```

## `oracle/` — what does xlsx@0.18.5 actually expose?

```sh
npm run oracle:manifest
```

Regenerates `oracle/api-manifest.json`: a machine-derived record of the oracle's public
API surface (top-level exports, `.utils`, `.stream`, for both its CJS and ESM
entrypoints). Commit the regenerated file whenever the pinned `xlsx` version changes.

## `differential/` — comparing elixcee's output against the oracle

`differential/classify.mjs` defines the six-value verdict
(`MATCH`/`INTENTIONAL_SECURITY_DIVERGENCE`/`UNSUPPORTED`/`BUG`/`ORACLE_AMBIGUITY`/
`NONDETERMINISTIC`) every comparison must resolve to — see its module doc comment for
definitions and the comparison-normalization rule.

```sh
npm run differential:demo
```

Runs the Phase 0 plumbing check: proves an oracle-vs-oracle comparison yields `MATCH`
and an oracle-vs-placeholder comparison yields `UNSUPPORTED`. This is not real
compatibility coverage — see [`docs/xlsx-security-model.md`](../docs/xlsx-security-model.md)
for what real coverage will need to check.
