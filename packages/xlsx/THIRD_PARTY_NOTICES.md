# Third-party notices

`@elixcee/xlsx` itself is MIT licensed (see `LICENSE` in this package). This file records
the third-party packages it takes as **runtime** `dependencies`, as a conservative
practice even though neither package below ships an upstream `NOTICE` file (checked
directly against each package's published contents, not assumed) — meaning the Apache
License 2.0 §4(d) NOTICE-propagation obligation is not strictly triggered.

---

## ssf

- **Version used:** 0.11.2 (exact-pinned)
- **Package:** https://www.npmjs.com/package/ssf
- **License:** Apache License 2.0
- **Copyright:** Copyright (C) 2013-present SheetJS LLC
- **Used for:** the number-format ("Format Codes") evaluation engine backing
  `format_cell`, `sheet_to_csv`, and `sheet_to_txt`. Isolated behind
  `src/internal/ssf-adapter.cjs`, the only file in this package that `require`s it.

## frac

- **Version used:** 1.1.2 (resolved transitively via `ssf`'s own `~1.1.2` dependency)
- **Package:** https://www.npmjs.com/package/frac
- **License:** Apache License 2.0
- **Copyright:** Copyright (C) 2012-present SheetJS
- **Used for:** rational/fraction approximation, used internally by `ssf`'s `# ?/?`-style
  fraction format codes. Not `require`d directly by any file in this package — it is
  `ssf`'s own dependency.

---

Full license texts are available in each package's published npm tarball
(`node_modules/ssf/LICENSE`, `node_modules/frac/LICENSE` after `npm install`) and at
<https://www.apache.org/licenses/LICENSE-2.0>.

This file is scoped to `@elixcee/xlsx`'s own runtime dependencies. For the full
elixcee project's licensing policy (including the `xlsx`/SheetJS packages used only by
the development-time differential-testing harness, never shipped here), see
<https://github.com/kent-tokyo/elixcee/blob/main/docs/licensing.md>.
