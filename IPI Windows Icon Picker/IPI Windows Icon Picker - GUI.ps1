# Ensure relative paths work
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Import-Module "$ScriptRoot\IPI Windows Icon Picker - Icons.psm1" -Force

# Icon Browser UI
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Shell32 Icon Browser"
        Height="600" Width="800"
        WindowStartupLocation="CenterScreen">
    <DockPanel>
        <TextBlock DockPanel.Dock="Bottom"
            Name="StatusText"
            Padding="6"
            FontWeight="Bold"
            Background="#EEE" />
        
        <ScrollViewer VerticalScrollBarVisibility="Auto">
            <WrapPanel Name="IconPanel" />
        </ScrollViewer>
    </DockPanel>
</Window>
"@

# Load window and populate icons
$reader = New-Object System.Xml.XmlNodeReader $XAML
$window = [Windows.Markup.XamlReader]::Load($reader)

$iconPanel = $window.FindName("IconPanel")
$statusText = $window.FindName("StatusText")

$statusText.Text = "Click and icon to see its index."

0..330 | ForEach-Object {
    $icon = Get-Shell32Icon -Index $_ -Size 32
    if ($null -ne $icon) {
        $img = New-Object System.Windows.Controls.image
        $img.Source = $icon
        $img.Width = 32
        $img.Height = 32
        $img.Margin = 6
        $img.ToolTip = "Index $_"
        $img.Tag = $_
        $img.Cursor = "Hand"

        $img.Add_MouseLeftButtonUp({
            param ($s, $e)
            $statusText.Text = "Selected icon index: $($s.Tag)"
        })

        $iconPanel.Children.Add($img) | Out-Null
    }
}

# Set Window icon
$window.Icon = Get-Shell32Icon 35

# Show the browser
$window.ShowDialog()