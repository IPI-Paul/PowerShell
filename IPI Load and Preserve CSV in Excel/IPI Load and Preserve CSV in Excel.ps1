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
Import-Module $Config.PlaceHolder -Force
Import-Module $Config.Logger -Force
Import-Module $Config.WebScraper -Force
Import-Module $Config.FolderSelect -Force

$Functions = @{
    ""                                          = 0
    "Load and Preserve CSV contents in Excel"   = 1
    "Save all downloads"                        = 2
    "Save selected download"                    = 3
    "Author Repository"                         = 4
    "Clear Log"                                 = 5
}

# XAML UI
[xml]$Xaml = (Get-Content $Config.MainWindow -Raw)

# Load XAML
$Reader = New-Object System.Xml.XmlNodeReader $Xaml
$Window = [Windows.Markup.XamlReader]::Load($Reader)

# Get controls
$BrowseButton       = $Window.FindName("BrowseButton")
$SaveButton         = $Window.FindName("SaveButton")
$RunFunctions       = $Window.FindName("RunFunctions")
$FilePathBox        = $Window.FindName("FilePathBox")
$SavePathBox        = $Window.FindName("SavePathBox")
$LogBox             = $Window.FindName("LogBox")
$StatusLabel        = $Window.FindName("StatusLabel")
$FilterBox          = $Window.FindName("FilterBox")
$ResultsBox         = $Window.FindName("ResultsBox")
$DepthBox           = $Window.FindName("DepthBox")
$FilePathBox.Text   = $CSVFile
$Script:foundLinks  = @{}
$Script:Update      = $false
$LogBox.Document.Blocks.Clear()

Set-WpfWatermark -TextBox $FilePathBox -Text "Enter full path or use the Browse button to locate the file. Website address will return list of files in the dropdown list below."
Set-WpfWatermark -TextBox $FilterBox -Text "Enter strings to filter web scrape by (separated by +/plus signs) and how many levels deep in the drop down list."
Set-WpfWatermark -TextBox $SavePathBox -Text "Enter full path or use the Save Folder button to locate the folder to save the downloads to."

$Functions.GetEnumerator() | Sort-Object -Property value | Select-Object -Property key | ForEach-Object {
    $RunFunctions.Items.Add(($_.Key)) | Out-Null
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

$SaveButton.Add_Click({
    $Dialog                     = Get-SelectDialog
    $Dialog.InitialDirectory    = $Config.Downloads
    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $SavePathBox.Text = Split-Path $Dialog.FileName
    }
})

function Reset-Status {
    $StatusLabel.Content = "Status: Idle"
}

# Background Runspace
$RunFunctions.Add_SelectionChanged({
    if ($this.SelectedIndex -ne 0) {
        if ($this.SelectedIndex -match "2|3") {
            if ($Script:foundLinks.Count -eq 0) {
                Write-Log -Text "There were no found files" "DarkRed" -Bold:$true
                $this.SelectedIndex = 0
                return
            } 
            
            if ([string]::IsNullOrWhiteSpace($SavePathBox.Text)) {
                Write-Log "Please select a download folder to save to." "DarkRed" -Bold:$true
                $this.SelectedIndex = 0
                return
            } elseif ($this.SelectedIndex -eq 2) {
                $Script:foundLinks.GetEnumerator() | ForEach-Object {
                    $UrlPath = Join-Path $SavePathBox.Text ($_.Value.Split('/')[-1])

                    # Download the file
                    Invoke-WebRequest -Uri $_.Value -OutFile $UrlPath
                }
            } else {
                if ([string]::IsNullOrWhiteSpace($ResultsBox.Text)) {
                    Write-Log "Please select a file to download." "DarkRed" -Bold:$true
                    $this.SelectedIndex = 0
                    return
                } else {
                    $UrlPath = Join-Path $SavePathBox.Text $ResultsBox.Text

                    # Download the file
                    Invoke-WebRequest -Uri $Script:foundLinks[$ResultsBox.Text] -OutFile $UrlPath
                }
            }
            Write-Log "Download Complete." "DarkGreen" -Italic:$true
            $this.SelectedIndex = 0
            return
        }
        elseif ($this.SelectedIndex -eq 4) {
            Write-Log -Text "Author Repository:" -Color "DarkBlue" -Bold:$true
            Write-Log -Text "https://github.com/IPI-Paul/PowerShell/tree/main/IPI%20Load%20and%20Preserve%20CSV%20in%20Excel" `
                -Color "DarkBlue"
            $this.SelectedIndex = 0
            return
        } elseif ($this.SelectedIndex -eq 5) {
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

        if (-not (Test-Path $FilePath) -and -not ($FilePath -match "^http")) {
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
            param ($Path, $LogCallback, $StatusCallback, $Icon, $Scraper)
 
            if (-not ($Path -match "^http") -or -not ("$($Scraper.UrlPath.Trim())" -eq "")) {
                & $LogCallback "Started processing file:" "DarkBlue" -Bold:$true
                try {
                    if (($Path -match "^http") -and -not ($Scraper.UrlPath.Trim() -eq "")) {
                        # Extract the filename from the URL and combine it with the temp folder path
                        $UrlPath = Join-Path $env:TEMP ($Scraper.UrlPath.Split('/')[-1])

                        # Download the file
                        Invoke-WebRequest -Uri $Scraper.UrlPath -OutFile $UrlPath
                        Set-FilePath $UrlPath -Log $LogCallback -Icon $Icon 
                        & $LogCallback $UrlPath "Purple"
                    } else {
                        Set-FilePath $Path -Log $LogCallback -Icon $Icon
                        & $LogCallback $Path "Purple"
                    }
                } catch {
                    & $LogCallback "Zip file extraction cancelled. `nError: $($_.Exception.Message)" "DarkRed" -Bold:$true
                    & $StatusCallback "Status: Completed with errors."
                    return
                }


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

                if (($Path -match "^http") -and -not ($Scraper.UrlPath.Trim() -eq "")) {
                    & $LogCallback "Deleting temporary file: $UrlPath" "Brown" -Italic:$true
                    Remove-Item $UrlPath -Force
                }

                & $LogCallback "Processing completed succesfully." "DarkRed" -Bold:$true
                $durationStr = Get-Duration $startTime
                & $LogCallback "Total duration to identify, format and load data with Long Digit/Leading Zero number columns: $durationStr" "DarkRed" -Bold:$true
            } else {
                & $LogCallback "Started scraping web url:" "DarkBlue" -Bold:$true

                $startTime = Get-Date

                try {
                    if ($Scraper.Links) { $Scraper.Links.Clear() }

                    & $Scraper.Scrape `
                        -Results $Scraper.Results `
                        -url $Path `
                        -maxDepth $Scraper.Depth `
                        -filter $Scraper.Filter `
                        -Log $LogCallback

                    & $Scraper.Results -Update:$true
                } catch {
                    & $LogCallback "Error: $($_.Exception.Message) `nType: $($_.Exception.GetType().FullName) `nLine: $($_.InvocationInfo.ScriptLineNumber)" "DarkRed" -Bold:$true
                }

                & $LogCallback "Select a file from the dropdown list and run the function to load in Excel." "DarkGreen" -Italic:$true

                & $LogCallback "Processing completed succesfully." "DarkRed" -Bold:$true
                $durationStr = Get-Duration $startTime
                & $LogCallback "Total duration to retrieve files: $durationStr" "DarkRed" -Bold:$true
            }

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
        ).AddArgument(
            [PSCustomObject]@{
                Filter  = $FilterBox.Text
                Depth   = [int]$DepthBox.Text
                Results = ${function:Update-Results}
                Scrape  = ${function:Initialize-WebScrape}
                UrlPath = $Script:foundLinks[$ResultsBox.Text]
                Links   = $Script:foundLinks
            }
        )

        $PowerShell.BeginInvoke()
        $this.SelectedIndex = 0
    }
})

function Update-Results {
    param (
        [string]$Key,
        [string]$Value,
        [switch]$Update
    )
    
    if ($Update) {
        $ResultsBox.Dispatcher.Invoke([action]{
            $ResultsBox.Items.Clear()

            # Forces a visual reset
            $ResultsBox.SelectedIndex = -1
            
            # Forces WPF to re-read the collection
            $ResultsBox.Items.Refresh()
            $ResultsBox.Items.Add("")
            
            $Script:foundLinks.GetEnumerator() | Sort-Object Name | ForEach-Object{ 
                try {
                    [void]$ResultsBox.Items.Add($_.Name) 
                } catch {
                    Write-Log "Error: $($_.Exception.Message)" "Black"
                }
            }
        }) 
    } else {
        $Script:foundLinks[$Key] = $Value
    }
}

# Set Window icon
$Window.Icon = Get-Shell32Icon 146

# Set Window Top Most
$Window.Topmost = $true

# Show Window
$Window.ShowDialog() | Out-Null
