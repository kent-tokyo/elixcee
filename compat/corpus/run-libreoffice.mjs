// Drives scenarios.json against LibreOffice via its UNO Basic scripting interface,
// automated through `soffice --headless "vnd.sun.star.script:...")` CLI macro invocation
// (see compat/corpus/README.md's "Why this invocation shape" for why pyuno/unohttpd
// were tried first and abandoned in this environment). LibreOffice is a secondary
// reference implementation here, NOT an "Excel oracle" — see ../oracle-excel-com/ for the
// (unimplemented) real-Excel adapter contract. Every result record carries
// `"oracle": "libreoffice"` per this milestone's explicit requirement.
//
// KNOWN, REPRODUCIBLE LIMITATION (see README.md for the full isolation trail): invoking a
// document macro that touches the VBA object model (Range/Cells, i.e. ActiveSheet
// resolution) via the nested `getScriptProvider().getScript(uri).invoke()` path used here
// hangs indefinitely in this sandboxed headless environment — confirmed with `x = 1 + 1`
// (completes) vs. `v = Range("A1").Value` (hangs >90s) as the only variable, isolated
// across Hidden true/false, in-memory doc vs. saved-xlsm-and-reopened doc, and read vs.
// write. Root cause not identified (a `sample` capture during the hang was not taken —
// see README.md's open item). Consequently every scenario whose VBA touches Range/Cells
// — i.e. nearly this entire corpus, since that's what it's built to exercise — is
// expected to time out here. That is a real, honestly-measured negative result, not a
// simulated one: see results/libreoffice-results.json's actual counts after a run.
//
// Run: `node run-libreoffice.mjs` from compat/corpus/. Requires `soffice` on PATH
// (developed against LibreOffice 26.2.5.2 — see tests/fixtures/e2e/README.md for this
// repo's existing LibreOffice-version convention).
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const DIR = path.dirname(fileURLToPath(import.meta.url));
const PER_SCENARIO_TIMEOUT_MS = 8_000;

function xmlEscape(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Basic string-literal escaping: double any embedded quote (Basic's own escaping rule).
// Used for values baked into the generated Module1.xba as literal Basic string constants.
function basicLiteral(s) {
  return '"' + String(s).replace(/"/g, '""') + '"';
}

function bootstrapProfile(profileDir) {
  if (fs.existsSync(profileDir)) return;
  spawnSync('soffice', ['--headless', `-env:UserInstallation=file://${profileDir}`, '--convert-to', 'txt', '/dev/null'], {
    timeout: 30_000,
  });
}

function buildModule(scenario, workbookPath, outPath) {
  // NOTE: `&` MUST be written as the XML entity &amp; anywhere in this template — the
  // .xba file is XML, and a raw `&` produces a malformed document that LibreOffice loads
  // as empty (no compile error surfaced anywhere; the macro URI just silently resolves to
  // nothing). This bit the corpus build once already — see README.md's investigation log.
  const loadUrl = workbookPath ? `file://${workbookPath}` : 'private:factory/scalc';
  const vbaLiteral = basicLiteralFromMultiline(scenario.vbaSource);

  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">${xmlEscape(`REM  *****  BASIC  *****

Sub RunScenario
    Dim sOutPath As String
    sOutPath = ${basicLiteral(outPath)}
    Dim iFile As Integer
    iFile = FreeFile
    Open sOutPath For Output As #iFile
    Print #iFile, "STATUS=STARTED"

    Dim oDesktop As Object
    oDesktop = createUnoService("com.sun.star.frame.Desktop")
    Dim oArgs(0) As New com.sun.star.beans.PropertyValue
    oArgs(0).Name = "Hidden"
    oArgs(0).Value = True
    Dim oDoc As Object
    oDoc = oDesktop.loadComponentFromURL(${basicLiteral(loadUrl)}, "_blank", 0, oArgs())
    Print #iFile, "STATUS=DOC_LOADED"

    Dim oLibs As Object
    oLibs = oDoc.BasicLibraries
    If Not oLibs.hasByName("VBAProject") Then
        oLibs.createLibrary("VBAProject")
    End If
    oLibs.VBACompatibilityMode = True
    Dim oLib As Object
    oLib = oLibs.getByName("VBAProject")

    Dim sCode As String
    sCode = ${vbaLiteral}
    oLib.insertByName("Module1", sCode)
    Print #iFile, "STATUS=MODULE_INSERTED"
    Close #iFile

    oDoc.getScriptProvider().getScript( _
        ${basicLiteral(`vnd.sun.star.script:VBAProject.Module1.${scenario.entrypoint}?language=Basic&location=document`)}).invoke(Array(), Array(), Array())

    ' Reopen the log for append-style writes below (StarBasic has no append-open mode
    ' shorthand here, so re-derive iFile and write the rest fresh with everything
    ' collected first instead).
    iFile = FreeFile
    Open sOutPath For Output As #iFile
    Print #iFile, "STATUS=MACRO_INVOKED"

    Dim oSheet As Object
    oSheet = oDoc.CurrentController.ActiveSheet
    If IsNull(oSheet) Then
        oSheet = oDoc.Sheets.getByIndex(0)
    End If

    Dim oCursor As Object
    oCursor = oSheet.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    Dim nLastRow As Integer, nLastCol As Integer
    nLastRow = oCursor.RangeAddress.EndRow
    nLastCol = oCursor.RangeAddress.EndColumn

    Dim r As Integer, c As Integer
    For r = 0 To nLastRow
        For c = 0 To nLastCol
            Dim oCell As Object
            oCell = oSheet.getCellByPosition(c, r)
            Dim nType As Integer
            nType = oCell.getType()
            If nType = com.sun.star.table.CellContentType.VALUE Then
                Print #iFile, "CELL" + Chr(9) + oCell.AbsoluteName + Chr(9) + "number" + Chr(9) + Str(oCell.getValue())
            ElseIf nType = com.sun.star.table.CellContentType.TEXT Then
                Print #iFile, "CELL" + Chr(9) + oCell.AbsoluteName + Chr(9) + "string" + Chr(9) + oCell.getString()
            ElseIf nType = com.sun.star.table.CellContentType.FORMULA Then
                Print #iFile, "CELL" + Chr(9) + oCell.AbsoluteName + Chr(9) + "formula_number" + Chr(9) + Str(oCell.getValue())
            End If
        Next c
    Next r

    Print #iFile, "STATUS=DONE"
    Close #iFile

    oDoc.close(False)
    oDesktop.terminate()
End Sub
`)}</script:module>
`;
}

// Builds a StarBasic string-literal expression (possibly multiple `+`-joined literals,
// one per source line) from a multi-line string, using Chr(10) between lines. Avoids `&`
// in the GENERATED Basic (this harness's own concatenation operator) — not because `&`
// itself is broken (it isn't; see README.md), but to keep this function's own output
// simple to XML-escape in one pass via the caller's xmlEscape() over the whole template.
function basicLiteralFromMultiline(text) {
  const lines = text.split('\n');
  return lines.map((line) => basicLiteral(line)).join(' + Chr(10) + ');
}

function parseResult(outPath) {
  if (!fs.existsSync(outPath)) return { status: 'NO_OUTPUT', cells: [] };
  const text = fs.readFileSync(outPath, 'utf8');
  const lines = text.split('\n').filter(Boolean);
  let status = 'UNKNOWN';
  const cells = [];
  for (const line of lines) {
    if (line.startsWith('STATUS=')) status = line.slice('STATUS='.length);
    else if (line.startsWith('CELL\t')) {
      const [, address, type, rawValue] = line.split('\t');
      const value = type === 'string' ? rawValue : Number(rawValue);
      cells.push({ address, type, value });
    }
  }
  return { status, cells };
}

const scenarios = JSON.parse(fs.readFileSync(path.join(DIR, 'scenarios.json'), 'utf8'));
const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'elixcee-lo-corpus-'));
const profileDir = path.join(workDir, 'profile');
bootstrapProfile(profileDir);

const moduleTarget = path.join(profileDir, 'user', 'basic', 'Standard', 'Module1.xba');
const lockFile = path.join(profileDir, '.lock');

const results = [];
let completed = 0;
let timedOut = 0;

// Usage: node run-libreoffice.mjs [count] [startIndex] [outSuffix]
// `startIndex`/`outSuffix` let multiple instances of this script run concurrently over
// disjoint slices of scenarios.json (see README.md's "parallel shards" note) — each gets
// its own tmp workDir/profile (mkdtempSync below) and its own results file, so there's no
// shared mutable state between shards.
const limitArg = process.argv[2] ? parseInt(process.argv[2], 10) : scenarios.length;
const startArg = process.argv[3] ? parseInt(process.argv[3], 10) : 0;
const outSuffix = process.argv[4] || '';
const toRun = scenarios.slice(startArg, startArg + limitArg);

for (const scenario of toRun) {
  const workbookPath = scenario.workbook ? path.join(DIR, 'workbooks', `${scenario.workbook}.xlsx`) : null;
  const outPath = path.join(workDir, `${scenario.id}.out.txt`);

  fs.writeFileSync(moduleTarget, buildModule(scenario, workbookPath, outPath));
  fs.rmSync(lockFile, { force: true });

  const proc = spawnSync(
    'soffice',
    [
      '--headless',
      '--invisible',
      '--nologo',
      '--nofirststartwizard',
      `-env:UserInstallation=file://${profileDir}`,
      `vnd.sun.star.script:Standard.Module1.RunScenario?language=Basic&location=application`,
    ],
    { timeout: PER_SCENARIO_TIMEOUT_MS, killSignal: 'SIGKILL' }
  );

  const timedOutFlag = proc.signal === 'SIGKILL' || proc.signal === 'SIGTERM' || proc.error?.code === 'ETIMEDOUT';
  const parsed = parseResult(outPath);

  if (timedOutFlag && parsed.status !== 'DONE') {
    timedOut++;
    results.push({ id: scenario.id, category: scenario.category, oracle: 'libreoffice', ok: false, status: 'TIMEOUT', cells: [] });
    // A hung soffice.bin may survive SIGTERM to the wrapper; make sure nothing lingers
    // holding the shared profile before the next scenario.
    spawnSync('pkill', ['-9', '-f', `UserInstallation=file://${profileDir}`]);
  } else if (parsed.status === 'DONE') {
    completed++;
    results.push({ id: scenario.id, category: scenario.category, oracle: 'libreoffice', ok: true, status: 'DONE', cells: parsed.cells });
  } else {
    results.push({ id: scenario.id, category: scenario.category, oracle: 'libreoffice', ok: false, status: parsed.status, cells: parsed.cells });
  }
}

fs.rmSync(workDir, { recursive: true, force: true });

const outDir = path.join(DIR, 'results');
fs.mkdirSync(outDir, { recursive: true });
const outFile = path.join(outDir, `libreoffice-results${outSuffix}.json`);
fs.writeFileSync(outFile, JSON.stringify(results, null, 2) + '\n');

console.log(`ran ${results.length} scenarios against LibreOffice: ${completed} completed, ${timedOut} timed out`);
console.log(`wrote ${path.relative(DIR, outFile)}`);
