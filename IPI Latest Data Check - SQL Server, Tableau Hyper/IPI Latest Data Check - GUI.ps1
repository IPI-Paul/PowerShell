param($hyperFile)

function Get-ScriptPath {
    # Get the full path to the currently running script
    if ($PSCommandPath) {
        $ScriptPath = $PSCommandPath
    } elseif ($MyInvocation.MyCommand.Path) {
        $ScriptPath = $MyInvocation.MyCommand.Path
    } else {
        throw "Unable to determine script path (script may not be running interactively)."
    }

    # Derived values (optional but commonly useful)
    $ScriptDirectory    = Split-Path -Parent $ScriptPath
    $ScriptName         = Split-Path -Leaf $ScriptPath

    # Output results
    [PSCustomObject]@{
        ScriptPath      = $ScriptPath
        ScriptDirectory = $ScriptDirectory
        ScriptBaseName  = (Split-Path -Parent $ScriptDirectory)
        ScriptName      = $ScriptName
    }
}

$PathVars = (Get-ScriptPath)

# Ensure relative paths work
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Import-Module "$ScriptRoot\IPI Windows Icon Picker - Icons.psm1" -Force

# Parameters
$Servers = @{
    LocalDB='(LocalDB)\MSSQLLocalDB'
    Development='10.0.0.1'
    Production='10.0.0.2'
}
$Database = 'tempdb'

$SQLPath    = "$($PathVars.ScriptDirectory)\SQL Scripts\SQL\"
$HyperPath  = "$($PathVars.ScriptDirectory)\SQL Scripts\Hyper\"
$SQLQueries = @{}
$HyperQueries = @{}
$Functions = @{
    ""                  = 0
    "Get Hyper Schema"  = 1
    "Run SQL Queries"   = 2
    "Author Repository" = 3
    "Clear Log"         = 4
}

# Get all .sql files
(Get-ChildItem -Path $SQLPath -Filter *.sql) | ForEach-Object {
    $FullName = (Split-Path -Leaf $_.Name)
    $Name = ($FullName -replace ".sql", "")
    $SQLQueries.Add($Name, $FullName)
}

(Get-ChildItem -Path $HyperPath -Filter *.sql) | ForEach-Object {
    $FullName = (Split-Path -Leaf $_.Name)
    $Name = ($FullName -replace ".sql", "")
    $HyperQueries.Add($Name, $FullName)
}


Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore 
Add-Type -AssemblyName WindowsBase

# XAML UI
[xml]$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="IPI Latest Data Check - SQL Server/Tableau Hyper"
        Height="520"
        Width="1000"
        WindowStartupLocation="CenterScreen">

    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="190"/>
            <ColumnDefinition Width="360"/>
            <ColumnDefinition Width="151"/>
            <ColumnDefinition Width="34"/>
            <ColumnDefinition Width="80"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- SQL Section -->
        <ComboBox Name="cboSvr"
            Grid.Row="0"
            Grid.Column="0"
            VerticalAlignment="Center"
            Margin="0,2,0,2"/>
        
        <ComboBox Name="cboTable"
            Grid.Row="0"
            Grid.Column="1"
            VerticalAlignment="Center"
            Width="358"
            Margin="0,2,0,2"/>

        <TextBox Name="txtUser"
            Grid.Row="0"
            Grid.Column="2"
            Grid.ColumnSpan="3"
            Margin="0,2,0,2"/>

        <PasswordBox Name="txtPwd"
            Grid.Row="0"
            Grid.Column="5"
            Margin="1,2,0,2"/>

        <!-- Tableau Hyper Section -->
        <TextBox Name="FilePathBox"
            Grid.Row="1"
            Grid.Column="0"
            Grid.ColumnSpan="3"
            ToolTip="Choose a File Path to a Tableau Hyper file if you want to compare with the SQL ouptput!"
            Margin="0,2,0,2"/>

        <Button Name="BrowseButton"
            Grid.Row="1"
            Grid.Column="3"
            Content="..."
            Width="30"
            Margin="1,2,0,2"
            ToolTip="Browse your computer for a Tableau Hyper file!"
            HorizontalAlignment="Left"/>
        
        <ComboBox Name="cboHyper"
            Grid.Row="1"
            Grid.Column="4"
            Grid.ColumnSpan="2"
            VerticalAlignment="Center"
            Margin="0,2,0,2"/>

        <!-- Action ComboBox -->
        <ComboBox Name="RunFunctions"
            Grid.Row="2"
            Grid.Column="3"
            Grid.ColumnSpan="3"
            VerticalAlignment="Center"
            Margin="0,2,0,2"/>

        <!-- Status -->
        <Label Name="StatusLabel"
            Grid.Row="2"
            Grid.Column="0"
            Grid.ColumnSpan="2"
            Content="Status: Idle"
            FontWeight="Bold"
            Margin="2"/>

        <!-- Log Output -->
        <RichTextBox Name="LogBox"
            Grid.Row="3"
            Grid.Column="0"
            Grid.ColumnSpan="6"
            Margin="2"
            IsReadOnly="True"
            IsEnabled="True"
            VerticalScrollBarVisibility="Auto"/>
    </Grid>
</Window>    
"@

# Load XAML
$Reader = New-Object System.Xml.XmlNodeReader $Xaml
$Window = [Windows.Markup.XamlReader]::Load($Reader)

$cboSvr             = $Window.FindName("cboSvr")
$cboTable           = $Window.FindName("cboTable")
$txtUser            = $Window.FindName("txtUser")
$txtPwd             = $Window.FindName("txtPwd")
$FilePathBox        = $Window.FindName("FilePathBox")
$BrowseButton       = $Window.FindName("BrowseButton")
$cboHyper           = $Window.FindName("cboHyper")
$RunFunctions       = $Window.FindName("RunFunctions")
$LogBox             = $Window.FindName("LogBox")
$StatusLabel        = $Window.FindName("StatusLabel")

$Servers.GetEnumerator() | Sort-Object -Property key | Select-Object -Property key | ForEach-Object {
    $cboSvr.Items.Add(($_.Key)) | Out-Null
}

$SQLQueries.GetEnumerator() | Sort-Object -Property name | Select-Object -Property name | ForEach-Object {
    $cboTable.Items.Add(($_.Name)) | Out-Null
}

$HyperQueries.GetEnumerator() | Sort-Object -Property name | Select-Object -Property name | ForEach-Object {
    $cboHyper.Items.Add(($_.Name)) | Out-Null
}

$Functions.GetEnumerator() | Sort-Object -Property value | Select-Object -Property key | ForEach-Object {
    $RunFunctions.Items.Add(($_.Key)) | Out-Null
}

$txtUser.Text = ($Env:USERNAME).ToLower()
$txtPwd.PasswordChar = "*"

$cboSvr.SelectedIndex = 2
$cboTable.SelectedIndex = 0
$cboHyper.SelectedIndex = 0
$FilePathBox.Text = $hyperFile
$LogBox.Document.Blocks.Clear()
# $LogBox.Cursor = [System.Windows.Input.Cursors]::Arrow

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

# Convert String to date
function Convert-ToDate {
    param (
        [Parameter(Mandatory)]
        [string]$DateString
    )

    switch -regex ($DateString) {
        '^\d{4}-\d{2}-\d{2}$' {
            [datetime]::ParseExact($DateString, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture)
        }
        '^\d{2}/\d{2}/\d{4}$' {
            [datetime]::ParseExact($DateString, 'dd/MM/yyyy', [cultureinfo]::InvariantCulture)
        }
        '^\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}$' {
            [datetime]::ParseExact($DateString, 'yyyy-MM-dd HH:mm:ss', [cultureinfo]::InvariantCulture)
        }
        '^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}:\d{2}$' {
            [datetime]::ParseExact($DateString, 'dd/MM/yyyy HH:mm:ss', [cultureinfo]::InvariantCulture)
        }
        default {
            throw "Unsupported date format: $DateString"
        }
    }
}

# Logging Timer
function Get-Duration {
    param (
        $startTime
    )
    $endTime = Get-Date
    $duration = $endTime - $startTime
    return "{0:00}:{1:00}:{2:00}" -f $duration.Hours, $duration.Minutes, $duration.Seconds
}

function Get-MyRunspace {
    param (
        [int]$RunType
    )
    # Path tot the Tableau Hyper API .Net DLL
    $HyperDllPath = "$($PathVars.ScriptDirectory)\API\get_hyper_file_data.dll"
    [System.Environment]::CurrentdIRECTORY = (Split-Path $HyperDllPath)

    if (-not [string]::IsNullOrWhiteSpace($FilePathBox.Text)) {
        if (-not (Test-Path $FilePathBox.Text)) {
            Write-Log -Text "Error: Please select a valid file." -Color "DarkRed" -Bold
            Reset-Status
            return
        }
    }

    $StatusLabel.Content = "Status: Running..."

    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $pathEntry = New-Object System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList "PATH", "$(Split-Path $HyperDllPath);$($env:PATH)", "Updated PATH"
    $iss.EnvironmentVariables.Add($pathEntry)
    $iss.ImportPSModule(@(
        "$PSScriptRoot\IPI Latest Data Check - SQL Server.psm1"
        "$PSScriptRoot\IPI Latest Data Check - Tableau Hyper.psm1"
    ))

    $Runspace = [runspacefactory]::CreateRunspace($iss)
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()

    $PowerShell = [powershell]::Create()
    $PowerShell.Runspace = $Runspace

    $params = [PSCustomObject]@{
        Svr             = $Servers[$cboSvr.Text]
        Table           = $cboTable.Text
        Database        = $Database
        SQLPath         = $SQLPath
        SQLQueries      = $SQLQueries
        User            = $txtUser.Text
        Password        = $txtPwd.SecurePassword
        HyperFile       = $FilePathBox.Text
        Hyper           = $cboHyper.Text
        HyperDllPath    = $HyperDllPath
        HyperPath       = $HyperPath
        HyperQueries    = $HyperQueries
        LogCallback     = $null
        ToDate          = $null
        RunType         = $RunType
        PowerShell      = $PowerShell
    }
    
    return $params
}
function Reset-Status {
    $StatusLabel.Content = "Status: Idle"
}

# Logging Helper
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

# Browse for File
$BrowseButton.Add_Click({
    $Dialog = New-Object Microsoft.Win32.OpenFileDialog
    $Dialog.Multiselect = $false

    if ($Dialog.ShowDialog()) {
        $FilePathBox.Text = $Dialog.FileName
    }
})

# Background Runspace
$RunFunctions.Add_SelectionChanged({
    if ($this.SelectedIndex -ne 0) {
        if ($this.SelectedIndex -eq 4) {
            $LogBox.Document.Blocks.Clear()
            Reset-Status
            $this.SelectedIndex = 0
            return
        }
        $RunType = $this.SelectedIndex
        $params = (Get-MyRunspace -RunType $RunType)
        $FilePath = $params.HyperFile

        $PowerShell = $params.PowerShell

        $PowerShell.AddScript({
            param ($SelectedIndex, $Path, $Logcallback, $StatusCallback, $params, $GetDuration, $ConvertToDate)
            
            $handler = [System.Windows.Navigation.RequestNavigateEventHandler] {
                param ($s, $e)
                Write-Host "Navigated"
                Start-Process $e.Uri.AbsoluteUri
                $e.Handled = $true
            }

            $LogBox.AddHandler(
                [System.Windows.Documents.Hyperlink]::RequestNavigateEvent,
                $handler   
            )
            $params.LogCallback = $Logcallback
            $params.ToDate      = $ConvertToDate

            $startTime = Get-Date
            if ($SelectedIndex -eq 1) {
                & $Logcallback "Sarted processing Hyper file:" "DarkBlue" -Bold:$true
                & $Logcallback "$Path" "Blue" -Italic:$true
                
                (Start-Hyper $params) | Out-Null
                $durationStr = (& $GetDuration $startTime)
                & $Logcallback "Duration to retrieve data from Hyper file: $durationStr" "Purple" -Bold:$true
            } elseif ($SelectedIndex -eq 2){
                & $Logcallback "Sarted processing SQL file:" "DarkBlue" -Bold:$true
                & $Logcallback "$($params.SQLPath)$($params.SQLQueries[$params.Table])" "Blue" -Italic:$true

                $sqlResult = (Start-SQL $params)

                if ("$sqlResult" -ne "") {
                    $durationStr = (& $GetDuration $startTime)
                    & $Logcallback "Duration to retrieve data from SQL Server: $durationStr" "Purple" -Bold:$true

                    if ("$($params.HyperFile)" -ne "" -and "$($params.Hyper)" -ne "") {
                        $startTime1 = Get-Date
                        & $Logcallback "Started processing hyper file:" "DarkBlue" -Bold:$true
                        & $Logcallback "$($params.HyperFile)" "Blue" -Italic:$true

                        $hyperResult = (Start-Hyper $params)

                        if ("$hyperResult" -ne "") {
                            $durationStr = (& $GetDuration $startTime1)
                            & $Logcallback "Duration to retrieve data from Hyper file: $durationStr" "Purple" -Bold:$true

                            if ($sqlResult.Date -gt $hyperResult.Date) {
                                & $Logcallback "The Hyper file $($params.HyperFile) is out of date!" "Red" -Italic:$true
                            } else {
                                & $Logcallback "The Hyper file $($params.HyperFile) is up to date!" "Green" -Italic:$true
                            }

                            $durationStr = (& $GetDuration $startTime)
                            & $Logcallback "Total Duration to retrieve data from SQL Server and Hyper file: $durationStr" "DarkRed" -Bold:$true
                        }
                    }
                }
            } elseif ($SelectedIndex -eq 3) {
                & $Logcallback "Author Repoitory:" "DarkBlue" -Bold:$true

                (Start-Hyper $params) | Out-Null
                $durationStr = (& $GetDuration $startTime)
                & $Logcallback "Duration to retrieve data from Hyper file: $durationStr" "Purple" -Bold:$true
            }

            & $StatusCallback "Status: Completed"
        }).AddArgument($this.SelectedIndex).AddArgument($FilePath).AddArgument({
            param($msg, $color, $bold, $italic)
            Write-Log -Text $msg -Color $color -Bold:$bold -Italic:$italic
        }).AddArgument({
            param($text)
            $Window.Dispatcher.Invoke([action]{
                $StatusLabel.Content = $text
            })
        }).AddArgument($params).AddArgument({
            param ($startTime)
            Get-Duration -startTime $startTime
        }).AddArgument({
            param ($result)
            Convert-ToDate -DateString $result
        })

        $PowerShell.Begininvoke()
        $this.SelectedIndex = 0
    }
})

# Set Window icon
$Window.Icon = Get-Shell32Icon 213

# Set Window Top Most
$Window.Topmost = $true

# Show Window
$Window.ShowDialog() | Out-Null