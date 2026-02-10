@{
    RootModule          = 'IPI Load and Preserve CSV in Excel.psm1'
    ModuleVersion       = '1.0.0'
    Description         = 'Load CSV files into Excel and Preserve Long Digit/Leading Zero Numbers'
    PowerShellVersion   = '5.1'
    FunctionsToExport   = @('Set-FilePath', 'Get-Duration', 'Get-Headers', 'Get-SpecialColumns', 'Update-FormatAndData', 'Update-LogSpecialHeaders')
}