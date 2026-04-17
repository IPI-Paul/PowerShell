function Get-LiteralPath {
    param ([string]$path)

    (Get-Item -LiteralPath "$PSScriptRoot\$path")
}
function Get-ScriptPaths {
    [PSCustomObject]@{
        IconHelper      = (Get-LiteralPath "..\cs\IPI Windows Icon Helper.cs")
        WatermarkSvc    = (Get-LiteralPath "..\cs\IPI WPF Place Holder.cs")
        FilterFormula   = (Get-LiteralPath "..\md\IPI Excel Filtering Formula.md")
        FolderSelect    = (Get-LiteralPath "..\psm\IPI Windows Folder Select.psm1")
        Engine          = (Get-LiteralPath "..\psm\IPI Load and Preserve CSV in Excel.psm1")
        Logger          = (Get-LiteralPath "..\psm\IPI Rich Text Logger.psm1")
        IconPicker      = (Get-LiteralPath "..\psm\IPI Windows Icon Picker - Icons.psm1")
        WebScraper      = (Get-LiteralPath "..\psm\IPI Web Scraper.psm1")
        PlaceHolder     = (Get-LiteralPath "..\psm\IPI WPF Place Holder.psm1")
        MainWindow      = (Get-LiteralPath "..\xaml\IPI Load and Preserve CSV GUI.xaml")
        ZipWindow       = (Get-LiteralPath "..\xaml\IPI ZIP CSV Picker GUI.xaml")
        Downloads       = (New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path
        Actions         = (Get-LiteralPath "..\actions")
    }
}

Export-ModuleMember -Function Get-ScriptPaths
