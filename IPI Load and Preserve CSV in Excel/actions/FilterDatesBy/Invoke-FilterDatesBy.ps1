param (
    [Parameter(Mandatory)]
    [ValidateNotNull()]
    [object]$Workbook,

    [Parameter(Mandatory)]
    [scriptblock]$Log
)

# Workbook-ish Checks (safe across COM)
try {
    $null = $Workbook.Worksheets.Count
    $null = $Workbook.Name
} catch {
    & $Log "Parameter -Workbook must be an Excel Workbook COM object (e.g. from `$Excel.Workbooks.Add() or .Open())." "DarkRed" -Bold:$true
}

Import-Module "$PSScriptRoot\config\Config.psm1" -Force
$Config = Get-ScriptPaths

Import-Module $Config.Engine -Force

& $Log "Adding Dates dataset to filter by..." "Green" -Italic:$true
try {
    Add-FilterByDates -Workbook $Workbook -Config $Config -Log $Log
} catch {
    & $Log "Error: $($_.Exception.Message)" "DarkRed" -Bold:$true
}

& $Log "Adding Dates Filter Formula..." "Green" -Italic:$true
try {
    Add-DateFilterFormula -Workbook $Workbook -Config $Config -Log $Log
} catch {
    & $Log "Error: $($_.Exception.Message)" "DarkRed" -Bold:$true
}

& $Log "Completed Tasks" "Purple" -Bold:$true

# For Excel, do not call Quit(), just release the reference
if ($null -ne $Workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook.Parent) | Out-Null }
Remove-Variable Workbook

# Optional: force garbage collection
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
[GC]::Collect()