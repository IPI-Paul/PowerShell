
function Convert-CentimetersToPoints 
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [double]
        $Centimeter
    )
    $point = $centimeter * 28.3464567;
    return $point;
}

function Convert-HTMLtoDataTable
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $Html
    )
    Begin
    {
        $doc = New-Object -ComObject 'HTMLFile'
        $tbl = New-Object System.Data.DataTable
    }
    Process
    {
        $doc.IHTMLDocument2_write(($Html -join "`r`n" -replace "<table>`r`n</table>`r`n", ''))
        $doc.all.tags('table') | % {
            for ($i = 0; $i -lt $_.Rows.length; $i++) {
                if (!$($_.rows[$i].style.display) -eq 'none') {
                    if ($i -ne 0) {
                        $row = $tbl.NewRow()
                    } 
                    for ($c = 0; $c -lt $_.rows[$i].Cells.length; $c++) {
                        if ($i -eq 0) {
                            $tbl.Columns.Add((New-Object System.Data.DataColumn $_.rows[$i].Cells[$c].innerText.trim()))
                        }
                        else {
                
                            $row[$_.rows[0].Cells[$c].innerText.trim()] = $_.rows[$i].Cells[$c].innerHTML.trim()
                        }
                    }
                    if ($i -ne 0) {
                        $tbl.Rows.Add($row)
                    }        
                }
            }
        }
    }
    End
    {
        return , $tbl 
    }
}

function Format-Html
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $Object
    )
    Begin
    {
    }
    Process
    {
        $body = ((((
            $Object | ConvertTo-Html
            ) -join '' -split '<body>')[1] -split '</body>')[0] -replace '<th>', '<th>¬' -replace '<td>', '<td>¬' -split '>' -join ">`r" -split '<' -join "`r<" -replace "`r`r", "`r"
        )
        $arr = @(
            @('<table>',"`t<table>"), @('</table>',"`t</table>"), @('<colgroup>',"`t`t<colgroup>"), @('</colgroup>', "`t`t</colgroup>"), @('<col>', "`t`t`t<col>"), 
            @('<col/>', "`t`t`t<col/>"), @('<tr>',"`t`t<tr>"), @('</tr>', "`t`t</tr>"), @('<th>', "`t`t`t<th>"), @('</th>', "`t`t`t</th>"), @('<td>', "`t`t`t<td>"), 
            @('</td>', "`t`t`t</td>"), @('¬', "`t`t`t`t")
        )
        $arr += @()
        foreach ($itm in $arr)
        {
            $body = ($body -replace $itm[0], $itm[1])
        }
    }
    End
    {
        return $body
    }
}

function Get-ColourFromRGB 
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $Red,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $Green,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        $Blue
    )
    Begin
    {
        '$Red' , '$Green', '$Blue' | % {
            $v = Invoke-Expression "$($_)"
            if ((!$v -or $v -gt 255 -or ($v -lt 0 -and -$v -gt 255)) -and $v -ne 0) 
            {
                Invoke-Expression "$_ = $([Long] 255)"
            }
            else 
            {
                Invoke-Expression "$_ = $([Long] [Math]::Abs($v))"
            }
        }
    }
    Process
    {
        $col = (($Red) + (($Green) * 256) + (($Blue) * 65536))
    }
    End
    {
        return $col
    }
}

function Get-ColoursFromCSS
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $Html
    )
    Begin
    {
        $doc = New-Object -ComObject 'HTMLFile'
    }
    Process
    {
        $doc.IHTMLDocument2_write("$Html")
        $css = Invoke-WebRequest -Uri ("file://$(Resolve-Path $doc.styleSheets[0].href)")
        $doc.styleSheets[0].cssText = $css.ToString().substring(3)
        $doc.styleSheets[0] | % {
            for ($i = 0; $i -lt $_.cssRules.length; $i++) 
            {
                $_.cssRules[$i]
            }} | Where-Object {
                $_.selectorText -eq 'th' -and $_.cssText -like '*color*'
                } | select -ExpandProperty cssText | % {
                    invoke-expression "`$cols=@{'$(($_ -split "{ " -replace ': ', "'='" -replace '; ', "'; '" -replace ' }', "'}")[1])"
                    } 
        $colours = @{
            'background-color' = $(
                if ($cols.'BACKGROUND-COLOR' -like 'RGB*') {
                    $($i=0;$cols.'BACKGROUND-COLOR' -replace 'rgb[(]', '' -replace '[)]', '' -split ',' | % {@{('R', 'G', 'B')[$i] = $_}; $i++})
                } elseif ($cols.'BACKGROUND-COLOR' -like '#*') {
                    ([System.Windows.Media.ColorConverter]::ConvertFromString($cols.'BACKGROUND-COLOR') | select R, G, B)
                } else {
                    ([System.Windows.Media.Colors]::($cols['background-color']) | select R, G, B)
                }
            );
            'color' = $(
                if ($cols.COLOR -like 'RGB*') { 
                    $($i=0;$cols.COLOR -replace 'rgb[(]', '' -replace '[)]', '' -split ',' | % {@{('R', 'G', 'B')[$i] = $_}; $i++})
                } elseif ($cols.COLOR -like '#*') {
                    ([System.Windows.Media.ColorConverter]::ConvertFromString($cols.COLOR) | select R, G, B)
                } else {
                    ([System.Windows.Media.Colors]::($cols['color']) | select R, G, B)
                }
            )
        }
    }
    End
    {
        return $colours
    }
}

function Get-FilterSQL
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [psobject]
        $Filter
    )
    Begin
    {
    }
    Process
    {
        $GetSql = {$(foreach ($itm in $_.Keys) {if ($_[$itm].length -gt 0) {"[$itm]" + " in ('" + $($_[$itm] -replace "'", "''" -join "', '") + "')"}}) -join ' and '}
    }
    End
    {
        return ($Filter | Foreach $GetSql | % {if ($_ -gt '') {' where ' + $_}})
    }
}

function Get-HyperlinkProperties
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $Html
    )
    Begin
    {
        $doc = New-Object -ComObject 'HTMLFile'
        $tbl = New-Object System.Data.DataTable
    }
    Process
    {
        $doc.IHTMLDocument2_write($Html)
        $tbl.Columns.Add((New-Object System.Data.DataColumn 'Anchor'))        
        $tbl.Columns.Add((New-Object System.Data.DataColumn 'Text'))
        $tbl.Columns.Add((New-Object System.Data.DataColumn 'URL'))
        $doc.all.tags('a') | % {
            for ($i = 0; $i -lt $_.length; $i++) {
                $row = $tbl.NewRow()
                $row['Anchor'] = $_.outerHTML.trim()
                $row['Text'] = $_.innerText.trim()
                $row['URL'] = $_.href
                $tbl.Rows.Add($row)
            }
        }
    }
    End
    {
        return , $tbl 
    }
}

function Get-Parameters
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CmdletName
    )
    Begin
    {
        <#
            Too Slow
        if ((Get-Module -ListAvailable | where {foreach ($itm in $_.ExportedCommands.Values) {if ($itm.Name -like 'set-contact') {$itm.Name}}}).Count -eq 0)
        {
            Write-Error 'Not Found' -ErrorAction stop
        }
        #>
    }
    Process
    {
        $GetParams = {
            foreach ($fld in ((("$(get-help $_)" -split "$_ ")[1] -split "Common")[0] -split ' ' -replace '[][]', '')) 
            {
                if($fld.IndexOf('-') -eq 0) 
                {
                    $fld.Substring(1)
                }
            }
        }
    }
    End
    {
        return ($CmdletName | Foreach $GetParams)
    }
}
