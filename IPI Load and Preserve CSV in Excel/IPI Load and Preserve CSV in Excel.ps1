param ($CSVFile)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# Ensure relative paths work
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Get Script Paths
Import-Module "$ScriptRoot\config\Config.psm1" -Force
$Config = Get-ScriptPaths

Import-Module $Config.IconPicker -Force

$Functions = @{
    ""                                          = 0
    "Load and Preserve CSV contents in Excel"   = 1
    "Author Repository"                         = 2
    "Clear Log"                                 = 3
}

# XAML UI
[xml]$Xaml = (Get-Content $Config.MainWindow -Raw)

# Load XAML
$Reader = New-Object System.Xml.XmlNodeReader $Xaml
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# Get controls
$BrowseButton       = $Window.FindName("BrowseButton")
$RunFunctions       = $Window.FindName("RunFunctions")
$FilePathBox        = $Window.FindName("FilePathBox")
$LogBox             = $Window.FindName("LogBox")
$StatusLabel        = $Window.FindName("StatusLabel")
$FilePathBox.Text   = $CSVFile
$LogBox.Document.Blocks.Clear()

$Functions.GetEnumerator() | Sort-Object -Property value | Select-Object -Property key | ForEach-Object {
    $RunFunctions.Items.Add(($_.Key)) | Out-Null
}

# Logging helper
function Write-Log {
    param (
        [string]$Text,
        [string]$Color = "Black",
        [switch]$Bold,
        [switch]$Italic
    )

    $Window.Dispatcher.Invoke([action]{
        $Paragraph = New-Object System.Windows.Documents.Paragraph
        $Paragraph.Margin = "0"     # Removes extra spacing
        $Paragraph.LineHeight = 12  # Adjust as needed (12-15 works well)

        if ($Text -match "https://") {
            # Create Hyperlink
            $hyperlink = New-Object System.Windows.Documents.Hyperlink
            $hyperlink.NavigateUri = [Uri]$Text
            $hyperlink.Inlines.Add($Text)
            $hyperlink.Cursor = [System.Windows.Input.Cursors]::Hand
            $hyperlink.Foreground = [System.Windows.Media.Brushes]::Blue
            $hyperlink.TextDecorations = [System.Windows.TextDecorations]::Underline

            $Paragraph.Inlines.Add($hyperlink)
        } else {
            $Run = New-Object System.Windows.Documents.Run($Text)
            $Run.Foreground = $Color
            if ($Bold) { $Run.FontWeight = "Bold" }
            if ($Italic) { $Run.FontStyle = "Italic" }

            $Paragraph.Inlines.Add($Run)
        }

        $LogBox.Document.Blocks.Add($Paragraph)
        $LogBox.ScrollToEnd()
    })
}

$LogBox.Add_PreviewMouseLeftButtonDown({
    param ($s, $e)

    $pointer = $LogBox.GetPositionFromPoint($e.GetPosition($LogBox), $true)
    if (-not $pointer) { return }

    # Get the inline element at the click
    $inline = $pointer.Parent
    while ($inline -and -not ($inline -is [System.Windows.Documents.Hyperlink])) {
        $inline = $inline.Parent
    }

    if ($inline -and $inline -is [System.Windows.Documents.Hyperlink]) {
        # Open the URI
        Start-Process $inline.NavigateUri.AbsoluteUri
        $e.Handled = $true
    }
})

# Browse for file
$BrowseButton.Add_Click({
    $Dialog             = New-Object Microsoft.Win32.OpenFileDialog
    $Dialog.Multiselect = $false

    if ($Dialog.ShowDialog()) {
        $FilePathBox.Text = $Dialog.FileName
    }
})

function Reset-Status {
    $StatusLabel.Content = "Status: Idle"
}

# Background Runspace
$RunFunctions.Add_SelectionChanged({
    if ($this.SelectedIndex -ne 0) {
        if ($this.SelectedIndex -eq 2) {
            Write-Log -Text "Author Repository:" -Color "DarkBlue" -Bold:$true
            Write-Log -Text "https://github.com/IPI-Paul/PowerShell/tree/main/IPI%20Load%20and%20Preserve%20CSV%20in%20Excel" `
                -Color "DarkBlue"
            $this.SelectedIndex = 0
            return
        } elseif ($this.SelectedIndex -eq 3) {
            $LogBox.Document.Blocks.Clear()
            Reset-Status
            $this.SelectedIndex = 0
            return
        }

        $FilePath = $FilePathBox.Text

        if ([string]::IsNullOrWhiteSpace($FilePath)) {
            Write-Log -Text "Error: No file selected." -Color "DarkRed" -Bold:$true
            Reset-Status
            $this.SelectedIndex = 0
            return
        }

        if (-not (Test-Path $FilePath)) {
            Write-Log -Text "Error: Please select a valid file." -Color "DarkRed" -Bold:$true
            Reset-Status
            $this.SelectedIndex = 0
            return
        }

        $StatusLabel.Content = "Status: Running..."

        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

        # Explicitly add the variable so that the module can see it upon import
        $iss.Variables.Add((New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "Config", $Config, "Global variable for module"))

        # Load the engine
        $iss.ImportPSModule($Config.Engine)

        $Runspace = [runspacefactory]::CreateRunspace($iss)
        $Runspace.ApartmentState = "STA"
        $Runspace.ThreadOptions = "ReuseThread"
        $Runspace.Open()

        $PowerShell = [powershell]::Create()
        $PowerShell.Runspace = $Runspace

        $PowerShell.AddScript({
            param ($Path, $LogCallback, $StatusCallback, $Icon)

            try {
                Set-FilePath $Path -Log $LogCallback -Icon $Icon
            } catch {
                & $LogCallback "Zip file extraction cancelled. `nError: $($_.Exception.Message)" "DarkRed" -Bold:$true
                & $StatusCallback "Status: Completed with errors."
                return
            }

            & $LogCallback "Started processing file:" "DarkBlue" -Bold:$true
            & $LogCallback $Path "Purple"

            $startTime = Get-Date

            & $LogCallback "Retrieving Headers" "DarkBlue" -Bold:$true
            $Result = Get-Headers $LogCallback
            & $LogCallback ("Column Count: $($Result.colCount)") "Black"

            & $LogCallback "Retrieving columns with Long Digit/Leading Zero numbers" "DarkBlue" -Bold:$true
            $Result = (Get-SpecialColumns -Headers $Result.Headers -colCount $Result.colCount -Log $LogCallback)
            
            & $LogCallback ("Special column indexes (0-based): $($Result.SpecialIndexes -join ', ')") "Black"
            Update-LogSpecialHeaders $Result.SpecialHeaders $LogCallback

            $durationStr = Get-Duration $startTime
            & $LogCallback "Duration to identify Long Digit/Leading Zero number columns: $durationStr" "DarkRed" -Bold:$true

            $startTime1 = Get-Date 
            & $LogCallback "Formatting columns with Long Digit/Leading Zero numbers and adding data!" "DarkBlue" -Bold:$true
            Update-FormatAndData -Excel $Result.Excel -SpecialIndexes $Result.SpecialIndexes -Headers $Result.Headers -ws1 $Result.ws1 `
                -schemaPath $Result.schemaPath -rowCount $Result.rowCount -colCount $Result.colCount -Log $LogCallback

            $durationStr = Get-Duration $startTime1
            & $LogCallback "Duration to format Long Digit/Leading Zero number columns: $durationStr" "DarkRed" -Bold:$true

            & $LogCallback "Processing completed succesfully." "DarkRed" -Bold:$true
            $durationStr = Get-Duration $startTime
            & $LogCallback "Total duration to identify, format and load data with Long Digit/Leading Zero number columns: $durationStr" "DarkRed" -Bold:$true

            & $StatusCallback "Status: Completed"
        }).AddArgument(
            $FilePath
        ).AddArgument({
            param ($msg, $color, $bold, $italic)
            Write-Log -Text $msg -Color $color -Bold:$bold -Italic:$italic
        }).AddArgument({
            param ($text)
            $Window.Dispatcher.Invoke([action]{
                $StatusLabel.Content = $text
            })
        }).AddArgument(
            (Get-Shell32Icon 146)
        )

        $PowerShell.BeginInvoke()
        $this.SelectedIndex = 0
    }
})

# Set Window icon
$Window.Icon = Get-Shell32Icon 146

# Set Window Top Most
$Window.Topmost = $true

# Show Window
$Window.ShowDialog() | Out-Null