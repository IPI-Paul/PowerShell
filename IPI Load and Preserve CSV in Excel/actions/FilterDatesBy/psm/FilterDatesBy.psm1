function Add-FilterByDates {
    param (
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$Workbook,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Config,
        [scriptblock]$Log
    )

    # Prevent multiple Add-Type calls
    if (-not ('ExcelHelper' -as [type])) {
        Add-Type -TypeDefinition (Get-Content $Config.ExcelHelper -Raw) -ReferencedAssemblies "Microsoft.Office.Interop.Excel"
    }

    $DateRecords              = Get-Content `
                            -Path $Config.FilterByDates `
                            -Encoding UTF8

    $xl                 = $Workbook.Parent
    $prevScreenUpdating = $xl.ScreenUpdating
    $prevEnableEvents   = $xl.EnableEvents
    $prevCalculation    = $xl.Calculation
    $prevAlerts         = $xl.DisplayAlerts
    $xl.ScreenUpdating  = $false
    $xl.EnableEvents    = $false
    $xl.Calculation     = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationManual

    try {
        $dates          = $Workbook.Sheets.Add()
        $dates.Name     = "Dates"

        # Pass the worksheet object and the array to the wrapper
        [ExcelHelper]::FastPaste($dates, [string[]]$DateRecords, 1, 1)
    } catch {
        & $Log "Error: $($_.Exception.Message)" "DarkRed" -Bold:$true
    } finally {
        $xl.Calculation     = $prevCalculation
        $xl.EnableEvents    = $prevEnableEvents
        $xl.ScreenUpdating  = $prevScreenUpdating
        $xl.DisplayAlerts   = $prevAlerts
    }
}

function Add-DateFilterFormula {
    param (
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$Workbook,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        $Config,
        [scriptblock]$Log
    )

    $fml    = ([regex]::Match((Get-Content `
                -Path $Config.DateFilterFormula `
                -Raw
            ), "(?s)``excel\s*(.*?)\s*``").Groups[1].Value)
    
    try {
        $res        = $Workbook.Sheets.Add()
        $res.Name   = "Upload"
        $res.Range('A1').Formula2 = ($fml -replace "`r?", "")

        # Get the Excel Application object from the workbook
        $xl = $Workbook.Application

        # Ensure a window Exists (normally does if Visible = true)
        $win = $xl.ActiveWindow

        # Zoom 80%
        $win.Zoom = 80

        # Scroll so A1 is visible
        # Most direct: set the top-left cell of the window using ScrollRow/ScrollColumn
        $win.ScrollRow      = 1
        $win.ScrollColumn   = 1

        # Alternatively (sometimes more robust): go to A1 (doesn't always guarantee top-left)
        # $xl.Goto($res.Rnge("A1"), $true)

        # Select A2 while keeping A1 visible
        $null = $res.Range("A2").Select()

        # Freeze panes (freezes rows above and columns left of active cell)
        # To freeze top row only, active cell must be A2 (row 2, col 1)
        if ($win.FreezePanes) { $win.FreezePanes = $false }
        $win.FreezePanes = $true

        # Autofit used range (rows + columns) ---
        $used = $res.UsedRange

        # Columns autofit
        $null = $used.EntireColumn.AutoFit()

        # Rows autofit
        $null = $used.EntireRow.AutoFit()

        # Optional niceties
        # Keep selection on A2 and ensure A1 still in view
        $win.ScrollRow      = 1
        $win.ScrollColumn   = 1
        $null = $res.range('A2').Select()
    } catch {
        & $Log "Error: $($_.Exception.Message)" "DarkRed" -Bold:$true
    }
}

Export-ModuleMember -Function Add-FilterByDates, Add-DateFilterFormula