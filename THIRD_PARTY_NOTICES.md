# Third-party notices

`elixcee` (including `@elixcee/xlsx`) is MIT licensed — see [`docs/licensing.md`](docs/licensing.md)
for the full policy. This file records the third-party packages `@elixcee/xlsx`
(`packages/xlsx`) takes as **runtime** `dependencies`, as a conservative practice even
though neither package below ships an upstream `NOTICE` file (checked directly against
each package's published contents, not assumed) — meaning the Apache License 2.0 §4(d)
NOTICE-propagation obligation is not strictly triggered. Ordinary `devDependencies` used
only by `compat/` (the differential-testing harness, never shipped to `@elixcee/xlsx`
consumers) are not listed here; see `docs/licensing.md` for that broader inventory.

---

## ssf

- **Version used:** 0.11.2 (exact-pinned)
- **Package:** https://www.npmjs.com/package/ssf
- **License:** Apache License 2.0
- **Copyright:** Copyright (C) 2013-present SheetJS LLC
- **Used for:** the number-format ("Format Codes") evaluation engine backing
  `format_cell`, `sheet_to_csv`, and `sheet_to_txt`. Isolated behind
  `packages/xlsx/src/internal/ssf-adapter.cjs`, the only file that `require`s it — see
  `docs/xlsx-architecture.md`'s "SSF backend" decision for why this dependency exists and
  its intended-transitional status.

## frac

- **Version used:** 1.1.2 (resolved transitively via `ssf`'s own `~1.1.2` dependency)
- **Package:** https://www.npmjs.com/package/frac
- **License:** Apache License 2.0
- **Copyright:** Copyright (C) 2012-present SheetJS
- **Used for:** rational/fraction approximation, used internally by `ssf`'s `# ?/?`-style
  fraction format codes. Not `require`d directly by any file in this repository — it is
  `ssf`'s own dependency.

---

Full license texts are available in each package's published npm tarball
(`node_modules/ssf/LICENSE`, `node_modules/frac/LICENSE` after `npm install`) and at
<https://www.apache.org/licenses/LICENSE-2.0>.
