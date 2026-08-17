<#
UNVERIFIED — untested scaffolding, not a working runner.

This script has NEVER been run. It was written on macOS with no Windows and no Excel
available to test against (see UNVERIFIED.md). Treat every line as a best-effort sketch
of the COM automation calls described in CONTRACT.md, not as validated code. Before this
can be called "done," a real Windows+Excel session needs to run it, fix whatever is
wrong, and remove this header.

Intended usage (untested): for each scenario object decoded from ../corpus/scenarios.json,
call Invoke-Scenario, then serialize the returned result objects to
../corpus/results/microsoft-excel-results.json in the shape CONTRACT.md specifies.

Requires (all unverified in this environment):
  - Windows with a licensed Microsoft Excel install.
  - "Trust access to the VBA project object model" enabled in Excel's Trust Center
    (File > Options > Trust Center > Trust Center Settings > Macro Settings) — required
    for VBComponents.Add/CodeModule.AddFromString below; without it those calls raise a
    COM exception (untested which HRESULT).
  - PowerShell able to load the Excel.Application COM ProgID (standard on a machine with
    Excel installed; untested on this specific corpus).
#>

param(
    [Parameter(Mandatory = $true)][string]$ScenariosJsonPath,
    [Parameter(Mandatory = $true)][string]$WorkbooksDir,
    [Parameter(Mandatory = $true)][string]$OutputJsonPath
)

function Invoke-Scenario {
    param($Scenario, [string]$WorkbooksDir)

    # UNVERIFIED: COM automation calls below are sketched from documented Excel object
    # model behavior, not confirmed against a running Excel instance.
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $result = [ordered]@{
        id            = $Scenario.id
        category      = $Scenario.category
        oracle        = 'microsoft_excel'
        excel_version = "$($excel.Version) / $($excel.Build)"
        ok            = $false
        status        = 'ERROR'
        cells         = @()
        error         = $null
    }

    try {
        if ($Scenario.workbook) {
            $wbPath = Join-Path $WorkbooksDir "$($Scenario.workbook).xlsx"
            $workbook = $excel.Workbooks.Open($wbPath)
        }
        else {
            $workbook = $excel.Workbooks.Add()
        }

        # UNVERIFIED: requires "Trust access to the VBA project object model".
        $component = $workbook.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
        $component.CodeModule.AddFromString($Scenario.vbaSource)

        $excel.Run("$($component.Name).$($Scenario.entrypoint)")

        $sheet = $workbook.ActiveSheet
        $used = $sheet.UsedRange
        $cells = @()
        foreach ($cell in $used.Cells) {
            if ($null -eq $cell.Value2) { continue }
            # UNVERIFIED: type branching sketch — mirrors run-libreoffice.mjs's
            # getType()-based approach (never trust a bare numeric read for a text cell;
            # see ../corpus/normalize.mjs's doc comment for why that silently
            # manufactures false matches). HasFormula / VarType(cell.Value2) is the
            # likely real check but has not been exercised.
            $type = if ($cell.HasFormula) { 'formula_number' }
            elseif ($cell.Value2 -is [double]) { 'number' }
            else { 'string' }
            $cells += [ordered]@{ address = $cell.Address($false, $false); type = $type; value = $cell.Value2 }
        }

        $result.ok = $true
        $result.status = 'DONE'
        $result.cells = $cells

        $workbook.Close($false)
    }
    catch {
        $result.status = 'ERROR'
        $result.error = [ordered]@{ message = $_.Exception.Message; source = 'RunScenario.ps1 (UNVERIFIED)' }
    }
    finally {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    return $result
}

# UNVERIFIED end-to-end driver loop.
$scenarios = Get-Content $ScenariosJsonPath -Raw | ConvertFrom-Json
$results = @()
foreach ($scenario in $scenarios) {
    $results += Invoke-Scenario -Scenario $scenario -WorkbooksDir $WorkbooksDir
}
$results | ConvertTo-Json -Depth 6 | Set-Content $OutputJsonPath
Write-Host "UNVERIFIED script finished — wrote $OutputJsonPath. Confirm output before trusting it."
