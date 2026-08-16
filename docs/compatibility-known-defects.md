# Compatibility-known defects

A running log of oracle (`xlsx@0.18.5`) behaviors that look like bugs but are
deliberately reproduced anyway, because compatibility with the real, currently-shipping
package takes priority over "fixing" something on its behalf. Each entry is a compat
decision on record, not something to silently normalize or reject later without updating
this file and its differential test coverage.

Contrast with [`docs/xlsx-security-model.md`](xlsx-security-model.md)'s intentional
*divergences* — those are cases where elixcee deliberately does NOT match the oracle
(because matching would mean replicating a DoS/injection vector). The entries below are
the opposite: elixcee DOES match the oracle, even though the oracle's behavior is itself
questionable, because there's no security reason not to.

---

```yaml
compatibility-known-defect:
  api: book_append_sheet
  case: colon in sheet name
  oracle_behavior: accepted
  excel_validity: invalid or application-dependent
  elixcee_behavior: reproduced for compatibility
```

`check_ws_name`'s thrown error message reads `"Sheet name cannot contain : \ / ? * [ ]"`,
listing `:` among the forbidden characters — but the actual character check
(`badchars = "][*?/\\".split("")`, `xlsx.js`) never includes `:`. A sheet named
`"Sheet:1"` is accepted without error. This is a genuine mismatch between the oracle's
error message and its real behavior, not a documentation choice. `packages/xlsx`
reproduces both: the same accept-with-colon behavior AND the same (technically
inaccurate) error message text, since real-world code may already depend on either. See
`compat/differential/xlsx-utils.test.mjs`'s `book_append_sheet` scenarios for the
differential coverage (`"Sheet:1"` is one of the tested special-character names).

**Applies to future write-path work too**: when `packages/xlsx` eventually implements
XLSX writing, do not add validation or normalization for colon-containing sheet names
beyond what the oracle itself does — differential-test whatever the oracle actually
writes to the ZIP for such a name, rather than assuming Excel's own (stricter, and
UI-context-dependent) rules apply.
