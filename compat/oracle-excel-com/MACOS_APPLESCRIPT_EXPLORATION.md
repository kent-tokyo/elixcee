# Mac Excel + AppleScript exploration — real findings, inconclusive result

**Status: exploratory spike, paused, not a working adapter.** This is a different automation
surface from `CONTRACT.md`'s Windows-COM contract — Microsoft Excel for Mac (confirmed
present and licensed on this machine: version 16.108) has no COM, so a future macOS-side
adapter would need a different mechanism entirely. This document records what was actually
tried, live, against a real running Excel — not a design proposal.

## Why this exists

`CONTRACT.md`/`UNVERIFIED.md`/`WINDOWS_EXECUTION.md` assume a Windows+Excel session that has
never been available. Partway through this project's life, Mac Excel *did* become available
on the development machine, prompting a same-day live spike into whether it could serve as a
real-Excel oracle for either `compat/corpus`'s value-correctness scenarios or the "safe
round-trip" (`0.9.0`, see `ROADMAP.md`) file-preservation validation. The spike found a real,
working mechanism, then hit reliability problems it didn't have time to root-cause. Recorded
here so a future session doesn't have to rediscover any of this from scratch.

## What's confirmed live (not assumed)

1. **AppleScript can control Mac Excel** for ordinary operations — create/open/close/save
   workbooks (including `save workbook as ... file format macro enabled XML file format` for
   `.xlsm`), read/write cell values via `Range`/`Cells`, check the active workbook, list open
   workbooks and windows. All confirmed via direct `osascript` calls against a live instance.
2. **AppleScript's own dictionary has NO VBA-project manipulation surface at all** — read
   Excel's actual `.sdef` directly (`/Applications/Microsoft Excel.app/Contents/Resources/
   Excel.sdef`, ~8000 lines; the `sdef` CLI tool itself requires full Xcode and isn't
   available, but the resource file can be read directly). The *only* VBA-related item in the
   entire dictionary is a single read-only boolean, `has vb project`. No `VBComponents`, no
   `CodeModule`, nothing resembling Windows COM's `Workbook.VBProject.VBComponents.Add` +
   `CodeModule.AddFromString`. This is a hard, confirmed constraint, not a guess.
3. **`run VB macro "<name>" arg1 <value>` can call a VBA `Sub` with a String parameter**,
   passing an argument through from AppleScript into VBA. Confirmed via a minimal `PingTest()`
   Sub with no parameters, and via a parameterized `Bootstrap(vbaSource As String)` Sub.
4. **VBA itself (not AppleScript) can modify its own project at runtime** —
   `ThisWorkbook.VBProject.VBComponents.Add(1)` (add a standard module) and
   `.CodeModule.AddFromString(...)` are both callable from ordinary VBA code, cross-platform
   (this is a VBA-level API, not a Windows-only COM extension), gated behind the same "Trust
   access to the VBA project object model" setting Windows needs (`Excel > Settings >
   セキュリティ > マクロ セキュリティ > 開発者向けのマクロ設定`, confirmed already enabled on
   this machine, screenshot-verified).
5. **Combining 3+4 worked, twice, end to end**: a `Bootstrap(vbaSource As String)` Sub —
   already present in a saved `.xlsm`, called once via `run VB macro "Bootstrap" arg1
   <sourceText>` — added a new module containing a `Sub Scenario()` built from the passed-in
   source text, ran it via an internal `Application.Run "Scenario"` call, and the result was
   visible in the workbook's own cells, read back via ordinary AppleScript `Range` properties.
   Confirmed live, cell values matched exactly what the injected code set (`A1 = 42`, then
   `A1 = 100` / `B1 = "hello"` on a second, independent call).

## What broke, and wasn't root-caused before pausing

1. **A VBA runtime error inside the dynamically-injected `Scenario` code did not reach the
   Bootstrap Sub's own `On Error GoTo` handler** — instead of a clean caught error, Excel
   broke into an interactive VBE debug session, which hangs any further AppleScript call
   against the app (a `run VB macro` call blocked for the full 2-minute command timeout).
   Recovering required a human to manually intervene in the Excel UI (dismiss the break
   state) — not something `osascript` could do from the outside. Root cause not confirmed;
   plausible candidates, none verified: `Application.Run`'s error-propagation semantics for a
   VBA-internal call may differ from what COM-automation callers get on Windows (where an
   external automation client is normally treated as "top of stack" for error-trapping
   purposes, letting `On Error` in the immediate caller work as expected); VBE's error-trap
   mode (Windows has a Tools > Options > General > Error Trapping setting with three modes —
   no equivalent was found in Mac VBE's menus, itself worth confirming properly rather than
   assumed absent).
2. **After that hang-and-recover cycle, previously-working calls started failing** with a
   generic `パラメータのエラーです。(-50)` (parameter error) from Excel — reproducible even
   after fully quitting and relaunching Excel (ruling out "stuck process" as the sole cause;
   something in the saved `.xlsm` file's own state, or a deeper automation-bridge issue, is
   the more likely explanation, neither confirmed).
3. **Splitting inject-and-run into two separate external `run VB macro` calls made things
   worse, not better**: a module added via one external call (confirmed present and correct
   by opening the VBE and reading the code visually) could not then be invoked by a *second*,
   separate external `run VB macro "<name>"` call — it silently no-opped (no error, exit 0,
   no cell changes) rather than running. The working pattern was always inject-then-run
   *inside the same VBA-to-VBA call chain* (Bootstrap calling `Application.Run` itself), never
   two independent AppleScript-driven calls. This suggests externally-invoked `run VB macro`
   resolves against some compiled/registered macro table that a same-session `AddFromString`
   addition doesn't immediately join, while an internal `Application.Run` from already-running
   VBA code can see it. Not confirmed against Microsoft documentation — inferred from
   behavior only.
4. Reverting to the exact original working design, in the same already-exercised workbook,
   did **not** reproduce the original success — it hit the `-50` error described in (2)
   instead, without hanging this time. Whatever broke did not un-break on revert.

## What this means for `0.9.0`

The core mechanism this would need — get arbitrary VBA source into a real Excel process and
run it, driven by an external script — **is proven possible** on Mac, via VBA's own
self-modification API triggered through `run VB macro` with a string argument. That's a real,
useful finding distinct from "Mac Excel automation doesn't work" (it does, for the happy
path). But repeated-use reliability, and specifically the behavior needed for the corpus's
own `EXPECTED_RUNTIME_ERROR` scenarios (intentionally-failing VBA, ~8 of 581 today) to come
back as a clean, capturable result rather than a hang, is unresolved.

**Not recommended to resume immediately in the same environment/workbook state this spike
left behind** — a genuinely fresh attempt (brand-new `.xlsm`, the exact minimal
inject-and-run-internally design from item 5 above, nothing else) in a dedicated session,
without layering the failed alternate designs on top first, is the more likely path to a
clear result. Whether the error-hang problem is fixable at all from pure AppleScript
automation (vs. needing, say, a different Mac automation technology, or accepting that error
scenarios need a different capture strategy than success scenarios) is still an open
question.

## Artifacts left behind

`~/Desktop/roundtrip_oracle.xlsm` (real file, on the user's own Desktop, not in this repo) —
contains the working `Bootstrap`/`PingTest` Subs plus several empty leftover modules
(`Module2`–`Module5`) from the failed splitting experiments. Left as-is; not cleaned up as
part of this pause. A future session picking this back up should decide whether to reuse it
(after manually clearing the stray modules) or start over.
