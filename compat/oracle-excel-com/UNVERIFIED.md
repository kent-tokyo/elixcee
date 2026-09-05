# Excel scenario adapter: unverified boundary

The corpus-wide Windows/Excel COM runner remains unverified scaffolding.
[RunScenario.ps1](RunScenario.ps1) is not a validated production runner.
This status is specific to **scenario execution**, not every Excel-related test:
the separate [0.9.0-A workbook record](results/0.9.0-A_summary.md) covers selected
Excel-for-Mac authored/reopened fixtures, but not post-save macro execution.

## Evidence still required

1. Execute the [scenario corpus](../corpus/scenarios.json) on a real, licensed
   Windows/Excel installation, recording Excel version/build and host details.
2. Validate the adapter's setup, macro loading, workbook isolation, timeout,
   exception handling, and cleanup before trusting its output.
3. Resolve representation differences such as elixcee's array/record JSON
   placeholders. Do not invent a MATCH when the outputs are not comparable.
4. Feed validated output into the existing normalizer/classifier and report
   MATCH, divergence, unsupported, nondeterministic, and unavailable separately.
5. Keep results tied to the exact source/binary and input corpus. A rerun is needed
   when implementation or fixtures change.
6. Verify post-save macro execution independently if making preservation-plus-execution claims.

Older milestone reports mention parser gaps that may now be implemented.
Use [FUNCTIONS](../../FUNCTIONS.md) for current coverage, not that historical list.

## What other evidence does not establish

- LibreOffice results do not establish agreement with Microsoft Excel.
  ORACLE_UNAVAILABLE is a missing measurement, not a pass or a semantic mismatch.
- Synthetic expected outcomes and independently written semantic references do not
  replace a live Excel run.
- Successful reopening does not prove that all workbook features survived or that
  VBA still executes correctly after save.
- One successful run cannot settle every Excel version, workbook feature, or
  nondeterministic case.

Follow the [Windows execution plan](WINDOWS_EXECUTION.md) and
[adapter contract](CONTRACT.md) when a suitable environment is available.
This documentation cleanup did not execute the Windows adapter or add Excel evidence.
