function Get-LiteralPath {
    param ([string]$path)

    (Get-Item -LiteralPath "$PSScriptRoot\$path")
}
function Get-ScriptPaths {
    [PSCustomObject]@{
        IconHelper      = (Get-LiteralPath "..\cs\IPI Windows Icon Helper.cs")
        FilterFormula   = (Get-LiteralPath "..\md\IPI Excel Filtering Formula.md")
        Engine          = (Get-LiteralPath "..\psm\IPI Load and Preserve CSV in Excel.psm1")
        IconPicker      = (Get-LiteralPath "..\psm\IPI Windows Icon Picker - Icons.psm1")
        MainWindow      = (Get-LiteralPath "..\xaml\IPI Load and Preserve CSV GUI.xaml")
        ZipWindow       = (Get-LiteralPath "..\xaml\IPI ZIP CSV Picker GUI.xaml")
    }
}
