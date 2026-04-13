function Initialize-WebScrape { 
    param (
        [Parameter(Mandatory)]
        $Results,
        [Parameter(Mandatory)]
        [string]$url,
        [Parameter(Mandatory)]
        [int]$maxDepth,
        [string]$filter,
        $Log
    )

    $url = $(if ($url -match "/$") { $url } else {"$url/"}) -replace "\\", "/"
    
    $Script:foundLinks  = @{}
    $Script:visited     = New-Object 'System.Collections.Generic.HashSet[string]'
    $Script:root        = $url

    try {
        Step-Recursive `
            -Results $Results `
            -url $Script:root `
            -depth 0 `
            -maxDepth $maxDepth `
            -filter $filter `
            -Log $Log
    } catch { 
        if ($Log)  {
            & $Log "Initialise Error: $($_.Exception.Message)" "DarkRed" -Bold:$true
        }
    }
}

function Step-Recursive {
    param (
        [Parameter(Mandatory)]
        $Results,
        [Parameter(Mandatory)]
        [string]$url,
        [Parameter(Mandatory)]
        [int]$depth,
        [Parameter(Mandatory)]
        [int]$maxDepth,
        [string]$filter,
        $Log
    )

    if ($depth -gt $maxDepth -or $Script:visited.Contains($url)) { return }
    $Script:visited.Add($url) | Out-Null

    try {
        if ($Log)  {
            & $Log "$url" "Black"
        }

        $page = Invoke-WebRequest -Uri $url -TimeoutSec 5 -UseBasicParsing -ErrorAction Inquire
        [int]$found = 0
        
        foreach($link in $page.Links.href) {
            $absUri = if ($link -match "^http") { $link } else { (New-Object System.Uri((New-Object System.Uri($url)), $link)).AbsoluteUri }

            if ($absUri -match "(\.csv|\.zip)") {
                $name = [System.IO.Path]::GetFileName($absUri).Split("?")[0]
                $filters = ($filter -split "\+" | ForEach-Object { $_.Trim() }) -join "|"
                
                if ($name -match "($filters)" -and -not $Script:foundLinks.ContainsKey($name)) {
                    $Script:foundLinks[$name] = $absUri
                    & $Results -Key $name -Value $absUri
                }

                if ($name -match "($filters)") { $found++ }
            } elseif ($absUri -match "^$([regex]::Escape($Script:root))") {
                Step-Recursive -Results $Results -url $absUri -depth ($depth + 1) -maxDepth $maxDepth -filter $filter -Log $Log
            } 
        }

        & $Log "Found: $found files." "DarkGreen" -Italic:$true
    } catch { 
        if ($Log)  {
            & $Log "Recursive Scan Error: $($_.Exception.Message)" "DarkRed" -Bold:$true
        }
    }
}

Export-ModuleMember -Function Initialize-WebScrape