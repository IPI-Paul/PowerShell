function Get-LiteralPath {
    param ([string]$path)

    (Get-Item -LiteralPath "$PSScriptRoot\$path")
}

function Get-ScriptPaths {
    $shared = "..\..\..\shared"

    [PSCustomObject]@{
        ExcelHelper         = (Get-LiteralPath "$shared\cs\Excel 2D Data Helper.cs")
        FilterByDates       = (Get-LiteralPath "$shared\csv\Filter By Dates.csv")
        DateFilterFormula   = (Get-LiteralPath "..\md\Date Filter Formula.md")
        Engine              = (Get-LiteralPath "..\psm\FilterDatesBy.psm1")
    }
}

Export-ModuleMember -Function Get-ScriptPaths