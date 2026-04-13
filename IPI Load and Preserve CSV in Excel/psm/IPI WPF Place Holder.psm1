function Enable-WpfWatermark {
    # Prevent multiple Add-Type calls
    if (-not ('WatermarkService' -as [type])) {
        Add-Type -TypeDefinition (Get-Content $Config.WatermarkSvc -Raw) -ReferencedAssemblies PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    }
}

function Set-WpfWatermark {
    param (
        [Parameter(Mandatory)]
        [string]$Text,
        [securestring]$Password,
        [System.Windows.Controls.TextBox]$TextBox,
        [System.Windows.Controls.PasswordBox]$PasswordBox
    )

    Enable-WpfWatermark
    if($TextBox) {
        [WatermarkService]::SetWatermark($TextBox, $Text)
    } elseif ($PasswordBox) {
        [WatermarkService]::SetWatermark($PasswordBox, $Password)
    }
}

Export-ModuleMember -Function Enable-WpfWatermark, Set-WpfWatermark