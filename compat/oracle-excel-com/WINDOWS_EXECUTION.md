# Running the corpus on a future Windows + Excel environment

Step-by-step instructions for the identical ~580-scenario corpus through the (to-be-
implemented, currently `UNVERIFIED`) Excel COM adapter. None of these steps have been
carried out — this is a plan for a future session with actual Windows+Excel access, not a
record of something already done.

1. **Copy this repo's `compat/` directory** (or clone the repo) onto a Windows machine
   with a licensed Microsoft Excel install. `compat/corpus/scenarios.json` and
   `compat/corpus/workbooks/*.xlsx` travel as-is — they're plain JSON/XLSX, no
   platform-specific content.

2. **Enable VBA project object model access in Excel**: File > Options > Trust Center >
   Trust Center Settings > Macro Settings > check "Trust access to the VBA project object
   model". Required for `RunScenario.ps1`'s `VBComponents.Add`/`AddFromString` calls
   (see `CONTRACT.md`). This is a one-time, per-machine setting.

3. **Fix up `RunScenario.ps1`.** It is explicitly `UNVERIFIED` scaffolding (see its own
   header) — expect it to need real debugging against a live Excel instance before it
   runs end to end. Treat this step as "implement the adapter," not "run a working
   script."

4. **Run it** (once fixed) from PowerShell, from the repo's `compat/oracle-excel-com/`
   directory:

   ```powershell
   .\RunScenario.ps1 `
     -ScenariosJsonPath ..\corpus\scenarios.json `
     -WorkbooksDir ..\corpus\workbooks `
     -OutputJsonPath ..\corpus\results\microsoft-excel-results.json
   ```

   Record the actual `Application.Version`/`Application.Build` this ran against — see
   `CONTRACT.md`'s schema — since `UNVERIFIED.md` item 6 flags that this affects which
   worksheet functions are even expected to exist.

5. **Feed the output into the existing classifier.** `compat/corpus/run-classify.mjs`
   currently globs `results/libreoffice-results*.json` for oracle results; add a second
   glob for `results/microsoft-excel-results*.json` (or rename the pattern to be
   oracle-agnostic) — a one-line change, not a redesign, per `CONTRACT.md`'s framing.
   Run:

   ```sh
   cd compat/corpus
   node run-classify.mjs
   ```

6. **Report Excel's numbers on their own table**, tagged `oracle: "microsoft_excel"`.
   Per this project's explicit rule (see `../corpus/README.md` and `CONTRACT.md`'s
   closing note): never merge Excel's classify-results with LibreOffice's into one
   blended percentage. Two separate tables, two separate `oracle` tags, always.

7. **Cross-reference `UNVERIFIED.md`'s itemized list** and confirm each item — that
   document is written to be checked off by this exact run, not superseded by a new one.
