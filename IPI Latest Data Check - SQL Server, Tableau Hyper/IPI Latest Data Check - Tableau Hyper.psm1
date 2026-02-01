function Show-Object {
    param (
        [object]$Object,
        [int]$Indent = 0,
        [scriptblock]$Logcallback
    )

    $pad = ' ' * $Indent

    if ($Object -is [System.Collections.IEnumerable] -and -not ($Object -is [string])) {
        foreach ($item in $Object) {
            Show-Object $item $Indent $Logcallback
        }
        return
    }

    if ($Object -is [psobject]) {
        foreach ($prop in $Object.PSObject.Properties) {
            if ($prop.Value -is [psobject] -or ($prop.Value -is [System.Collections.IEnumerable] -and -not ($prop.Value -is [string]))) {
                & $Logcallback "$pad$($prop.Name):" "Green" -Bold:$true
                Show-Object $prop.Value ($Indent + 2) $Logcallback
            } else {
                if ($prop.Name -match "Column") {
                    & $Logcallback "$pad$($prop.Name): $($prop.Value)" "DarkGreen" -Italic:$true
                } elseif ($prop.Name -match "Table") {
                    & $Logcallback "$pad$($prop.Name): $($prop.Value)" "Brown" -Italic:$true
                } else {
                    & $Logcallback "$pad$($prop.Name): $($prop.Value)" "DarkSeaGreen" -Italic:$true
                }
            }
        }
    }
}

function Start-Hyper($params) {
    $HyperDllPath       = $params.HyperDllPath
    $HyperFile          = $params.HyperFile
    $Hyper              = $params.Hyper
    $HyperPath          = $params.HyperPath
    $HyperQueries       = $params.HyperQueries
    $Logcallback        = $params.LogCallback
    $ToDate             = $params.ToDate
    $RunType            = $params.RunType
    
    $hyperDate  = $null
    $sql        = ""
    $Script     = ""

    if ($RunType -ne 3) {
        try {
            $Script = $HyperQueries[$Hyper]
        }
        catch {
            & $Logcallback "Hyper`t: $Hyper `nError`t: The Queries object does not have an item matching $Hyper. Please update the GUI script." "DarkRed" -Bold:$true
            return
        }

        # SQL to get the max date
        if ("$Script" -ne "") {
            try {
                $sql = Get-Content -Path ($HyperPath + $Script) -ErrorAction Stop
            }
            catch {
                & $Logcallback "Hyper`t: $Hyper `nError`t: $($_.Exception.Message)" "DarkRed" -Bold:$true
                return
            }
        }
    }

    try {
        # Load Hyper API
        $dllPath = ($HyperDllPath -replace "\\", "\\")

        # Prevent multiple Add-Type calls
        if (-not ('HyperAPIWrapper' -as [type])) {        
            # Define the function using Add-Type
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class HyperAPIWrapper {
    [DllImport("$dllPath", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
    public static extern IntPtr read_hyper_file(string path, string sql, int schema);
    
    [DllImport("$dllPath", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.Cdecl)]
    public static extern void FreeResult(IntPtr ptr);
}
"@ -PassThru | Out-Null
        }

        # Call the DLL function and get pointer 
        $ptr = [HyperAPIWrapper]::read_hyper_file($HyperFile, $sql, $RunType)

        # Get the result string
        $result = [Runtime.InteropServices.Marshal]::PtrToStringAnsi($ptr)

        # Free unmanaged memory
        [HyperAPIWrapper]::FreeResult($ptr)

        if ($RunType -eq 1) {
            $obj = $result | ConvertFrom-Json 
            Show-Object $obj -Logcallback $Logcallback
        } elseif ($RunType -eq 2) {
            # Convert result to date and if Month Start change to Month End
            $hyperDate = (& $ToDate $result)
            if ($hyperDate.Day -eq 1) {
                $hyperDate = $hyperDate.AddMonths(1).AddDays(-1)
            }

            # Output
            & $Logcallback "Hyper`t: $Hyper `nResult`t: $($hyperDate.ToString('dd/MM/yyyy'))" "Green" -Italic:$true
            return $hyperDate
        } else {
            # Output
            & $Logcallback $result "Green" -Italic:$true
            return $result
        }
    }
    catch {
        & $Logcallback "Hyper`t: $Hyper `nError`t: $($_.Exception.Message)" "DarkRed" -Bold:$true
    }
}