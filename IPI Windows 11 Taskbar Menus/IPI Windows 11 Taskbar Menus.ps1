param (
    [string]$fPath, 
    [string]$Filter = 1,
    [int]$Icon,
    [int]$Close = 1
    )

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# Ensure relative paths work
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

Import-Module "$ScriptRoot\IPI Windows Icon Picker - Icons.psm1" -Force

function Get-CheckPath {
    param (
        [string]$fPath,
        $fName = $null
    )
    if ("$fPath" -eq "") {
        # Set folder path containing shortcuts
        $shortcutFolder = [System.Windows.Forms.FolderBrowserDialog]::new()
        $null = $shortcutFolder.ShowDialog()
        $folderPaths = @($shortcutFolder.SelectedPath)
        if (!$folderPaths) {
            [System.Windows.Forms.MessageBox]::Show("Selection Cancelled.", "Info", "OK", "Information")
            exit
        }
    } elseif ("$fPath" -eq "Desktop") {
        $folderPaths = @(
            [Environment]::GetFolderPath("Desktop"),
            [Environment]::GetFolderPath("CommonDesktopDirectory"),
            "C:\Users\$($env:USERNAME)\Desktop"
        ) | Select-Object -Unique
    } elseif ("$fPath" -eq "Applications") {
        # $appsFolder = $shell.NameSpace(42)
        # $appsFolder = $shell.NameSpace('shell:AppsFolder')
        # $appsFolder = $shell.NameSpace('shell:::{4234d49b-0245-4df3-b780-3893943456e1}')
        $folderPaths = @("shell:AppsFolder")
        $fName = "Applications"
    } elseif ("$fPath" -eq "Control Panel") {
        # The GUID below is the unique identifier for the Applications folder
        # $appsFolder = $shell.NameSpace('shell:::{26ee0668-a00a-44d7-9371-beb064c98683}') by Category only
        # $appsFolder = $shell.NameSpace('shell:ControlPanelFolder')
        # $appsFolder = $shell.NameSpace('shell:::{21ec2020-3aea-1069-a2dd-08002b30309d}')
        $folderPaths = @("shell:ControlPanelFolder")
        $fName = "Control Panel"
    } else {
        $folderPaths = @($fPath)
    }

    return [PSCustomObject]@{
        Paths   = $folderPaths
        Name    = $fName
    }
}

function Set-Listbox {
    param (
        $listbox
    )
    
    $count = 0
    $Script:width = 63
    $listbox.Items.Clear()
    # Add Shortcut names to the listbox
    foreach ($shortcut in ($Script:shortcuts | Sort-Object Name)) {
        if ($Script:Filter -eq 1) {
            $listbox.Items.Add(($shortcut.Name -replace ".lnk", "")) | Out-Null
        } else {
            $listbox.Items.Add($shortcut.Name) | Out-Null
        }
        if ($Script:width -lt ($shortcut.Name).Length){
            $Script:width = ($shortcut.Name).Length
        }
        $count += 1
    }
    $listbox.Items.Add("") | Out-Null
    $listbox.Refresh()
    $count += 1

    $Script:height = (340 / 22) * ($count, 50 | Measure-Object -Min).Minimum
    if ($Script:width -lt 63) {
        $Script:width = 360
    } else {
        $Script:width = (360 / 63) * $Script:width
    }

    $listbox.Size = New-Object System.Drawing.Size($Script:width, $Script:height)
    if ($form) { 
        $form.Size = New-Object System.Drawing.Size(($Script:width+10), ($Script:height+32)) 
        $screen = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
        $form.Top = $screen.Top + (($screen.Height - $form.Height) / 2)
        $form.Left = $screen.Left + (($screen.Width - $form.Width) / 2)
        Update-HoverItem $listbox
    }
}

function Set_Shortcuts {
    param (
        $folderPaths,
        $shortcuts = $null,
        $form = $null,
        $Filt = $null,
        $Icn = $null
    )

    $check          = Get-CheckPath -fPath $folderPaths
    $fName          = $check.Name
    $folderPaths    = $check.Paths

    if ($stack.Count -gt 0) {
        $shortcuts = @([PSCustomObject]@{
            Name = "..."
            FullName = " ..."
        })
    }
    
    $temp = Split-Path -Leaf $(if($folderPaths.Count -gt 1) { $folderPaths[0] } else { $folderPaths })
    if(!$fName -or ($fName -ne $temp -and -not ($temp -match "shell:"))) { $fName = $temp }
    
    if ($form) { $form.Text = "Taskbar for - $fName" }
    if ($Icn)  { $form.Icon = $Icn }
    if ($Filt) { $Script:Filter = $Filt }

    foreach ($folderPath in $folderPaths) {
        if (-not ($fName -eq "Applications" -or $fName -eq "Control Panel")) {
            if (-not (Test-Path $folderPath)) {
                [System.Windows.Forms.MessageBox]::Show("Folder not found!", "Error", "OK", "Error")
                exit
            }

            if ($Script:Filter -eq 1) {
                # Get all .lnk shortcut files
                $shortcuts += Get-ChildItem -Path $folderPath -Filter *.lnk
            } else {
                $shortcuts += Get-ChildItem -Path $folderPath
            }
        } else {
            $shell = New-Object -ComObject Shell.Application
            $appsFolder = $shell.NameSpace($folderPath)
            if ($fName -eq "Applications") { 
                $shortcuts += $appsFolder.Items() | Select-Object Name, Path | ForEach-Object {
                    if($_.Path -match ":") {
                        $path = $_.Path
                    } else {
                        $path = "$folderPath\$($_.Path)"
                    }
                    
                    [PSCustomObject]@{
                        FullName = $path
                        Name = $_.Name
                    }
                }
            } else {            
                $ns = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace"

                $shortcuts += Get-ChildItem $ns | ForEach-Object {
                    $guid   = $_.PSChildName
                    $key    = "HKLM:\SOFTWARE\Classes\CLSID\$guid"

                    if (Test-Path $key) {
                        $name = (Get-ItemProperty $key).'(default)'
                        # The entry below makes this very slow and is not needed
                        # $canonical = (Get-ItemProperty "$key\Shell\Open\Command" -ErrorAction SilentlyContinue).'(default)'

                        [PSCustomObject]@{
                            Name        = $name
                            Fullname    = "shell:::$guid"
                        }
                    }
                } | Sort-Object Name
            }
        }
    }
    $shortcuts
}

function Update-HoverItem {
    param (
        $listbox
    )

    # Re-evaluate hover item
    $pos = $listbox.PointToClient([System.Windows.Forms.Control]::MousePosition)
    $index = $listbox.IndexFromPoint($pos)
    if ($index -ge 0 -and $listbox.SelectedIndex -ne $index) {
        $listbox.SelectedIndex = $index
    }
}

# Stacks to keep track of folder navigation, icon and title
$stack = New-Object System.Collections.Stack
$icons = New-Object System.Collections.Stack
$titles = New-Object System.Collections.Stack

# Globals
$Script:shortcuts = $null
$Script:height = 0
$Script:width = 0
$Script:Filter = $Filter

# Get folder paths and title
$check          = Get-CheckPath -fPath $fPath
$fName          = $check.Name
$folderPaths    = $check.Paths

# Set title
if (!$fName) { $fName =  Split-Path -Leaf $(if($folderPaths.Count -gt 1) { $folderPaths[0] } else { $folderPaths }) }

# Get shortcuts
if ($shortcuts.Count -eq 0) {
    $Script:shortcuts = (Set_Shortcuts $folderPaths)
}

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Taskbar for - $fName"
$form.StartPosition = "CenterScreen"

# Create a listbox
$listbox = New-Object System.Windows.Forms.ListBox
$listbox.Location = New-Object System.Drawing.Point(0,0)
$listbox.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$listbox.SelectionMode = "One"

Set-Listbox $listbox

$toolTipOn = $true

$listbox.Add_Click({
    if ($Script:Filter -eq 1 -and -not ($fName -eq "Applications" -or $fName -eq "Control Panel")) {
        $selected = "$($listbox.SelectedItem).lnk"
    } else {
        $selected = $listbox.SelectedItem
    }
    if ($selected) {
        $shortcutPath = ($Script:shortcuts | Where-Object { $_.Name -eq $selected -or $_.FullName -eq $selected } | Select-Object FullName).FullName
        Write-Host $selected, $shortcutPath
        
        if ($shortcutPath -and -not ($shortcutPath -match "shell:")) { $item = Get-Item -LiteralPath $shortcutPath }
        if ((Test-Path $selected -PathType Container) -and -not ($selected -match "\.\.\.")) {
            $titles.Push($fName)
            $icons.Push($form.Icon)
            $stack.Push($Script:shortcuts)
            $Script:shortcuts = (Set_Shortcuts $selected -form $form)
            Set-Listbox $listbox
        } elseif ($item.PSIscontainer -and -not ($selected -match "\.\.\.")) {
            $titles.Push($fName)
            $icons.Push($form.Icon)
            $stack.Push($Script:shortcuts)
            $Script:shortcuts = (Set_Shortcuts $item -form $form)
            Set-Listbox $listbox
        } elseif ($selected -match "\.\.\.") {
            $Script:shortcuts = $stack.Pop()
            Set-Listbox $listbox
            if ($icons.Count -gt 0) { $form.Icon = $icons.Pop() }
            if ($titles.Count -gt 0) { $form.Text = "Taskbar for - $($titles.Pop())" }
        } else {        
            if ($shortcutPath -match ".lnk") {            
                $shortcutObject = $wshShell.CreateShortcut($shortcutPath)
                if ($shortcutObject.TargetPath -eq "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -and 
                    $shortcutObject.Arguments -match "Windows 11 Taskbar Menus.ps1"
                ){
                    $titles.Push(($form.Text -replace "Taskbar for - ", ""))
                    $icons.Push($form.Icon)
                    $src    = ($shortcutObject.Arguments -split "ps1`"")[-1]
                    $digits = ($src -split " ")[-2..-1]
                    $filt   = $digits[0]
                    $icn    = (Convert-WpfBitmapToIcon (Get-Shell32Icon $digits[1]))
                    $stack.Push($Script:shortcuts)
                    $Script:shortcuts = (Set_Shortcuts (($src -replace " $digits").Trim().Trim('"') -replace "%appdata%", $env:APPDATA) -form $form -Filt $filt -Icn $icn)
                    Set-Listbox $listbox
                } else {
                    $toolTipOn = $false 
                    try {
                        Start-Process -FilePath "$shortcutPath"
                    } catch {
                        try {
                            Start-Process -FilePath "$($shortcutObject.TargetPath)"
                        } catch {
                            Start-Process -FilePath "$($shortcutObject.TargetPath -replace ' \(x86\)', '')"
                        }
                    }
                    if ($Close -eq 1) { $form.Close() }
                }
            } else {
                $toolTipOn = $false 
                Start-Process -FilePath "$shortcutPath"
                if ($Close -eq 1) { $form.Close() }
            }
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a shortcut first.", "Warning", "OK", "Warning")
    }
})

# Create ToolTip object
$toolTip = New-Object System.Windows.Forms.ToolTip
$toolTip.AutoPopDelay = 5000
$toolTip.InitialDelay = 500
$toolTip.ReshowDelay = 500
$toolTip.ShowAlways = $true

# Track the last item index shown to avoid flickering
$lastindex = -1

# Create WScript.Shell COM object to work with shortcuts
$wshShell = New-Object -ComObject WScript.Shell

# MouseMove event handler to show tooltips based on item under mouse
$listbox.Add_MouseMove({
    param($s, $e)
    # Re-evaluate hover item
    $pos = $s.PointToClient([System.Windows.Forms.Control]::MousePosition)
    $index = $s.IndexFromPoint($pos)
    if ($index -ge 0 -and $s.SelectedIndex -ne $index -and $index -ne $lastindex -and ($Script:shortcuts | Sort-Object Name)[$index].Extension -eq ".lnk" -and $toolTipOn) {
        $s.SelectedIndex = $index
        $lastindex = $index
        $itemText = ($Script:shortcuts | Sort-Object Name)[$index].FullName
        $shortcutObject = $wshShell.CreateShortcut($itemText)
        $toolTip.SetToolTip($listbox, $shortcutObject.TargetPath)
    } elseif ($index -lt 0) {
        $toolTip.SetToolTip($listbox, "") # Clear tooltip if not over item
        $lastindex = -1
    }
})

$form.Size = New-Object System.Drawing.Size(($Script:width+10), ($Script:height+32))
$form.Controls.Add($listbox)
$form.TopMost = $true
$form.Add_Shown({$form.Activate()})
if ($Icon) { 
    $form.Icon = (Convert-WpfBitmapToIcon (Get-Shell32Icon $Icon)) 
}
[void]$form.ShowDialog()
