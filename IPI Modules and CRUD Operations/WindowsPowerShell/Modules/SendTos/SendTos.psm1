
function Format-PowerPointTable 
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $PptPresentation, 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        $SlideNumber,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [psobject]
        $GridResult,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [psobject]
        $Colours
    )
    Begin
    {
        if ($Colours)
        {
            $cboRed = @{'Text' = $Colours.'background-color'.R}
            $cboGreen = @{'Text' = $Colours.'background-color'.G}
            $cboBlue = @{'Text' = $Colours.'background-color'.B}
            $fgRed = $Colours.color.R
            $fgGreen = $Colours.color.G
            $fgBlue = $Colours.color.B
        }
        else
        {
            $fgRed = 255
            $fgGreen = 255
            $fgBlue = 255
        }
    }
    Process
    {
        $i = 1
        if ($Colours) {
            $rows = $GridResult.Rows.Count + 1
            $rw = 0
        } else {
            $rows = $GridResult.Rows.Count
            $rw = 1
        }
        foreach ($row in $GridResult.Rows) {
            if ((13 % $i) -eq 0) {
                $SlideNumber++
                $sld = $PptPresentation.Slides.Add($SlideNumber, 12)
                if ($rows -gt 12) {
                    $mrows = 12
                    $rows = $rows - 11
                } else {
                    $mrows = $rows
                }
                $tbl = $sld.Shapes.AddTable($mrows, $GridResult.Columns.Count)
                $tbl | ForEach-Object {
                    $_.Top = 20
                    $_.Left = 20
                    $_.Width = 920
                    $_.Height = (500 / 12) * $mrows
                }
                $i = 1
                $j = 1
                foreach ($col in $GridResult.Columns) {
                    $tbl.Table.Cell($i, $j).Shape.TextFrame.TextRange | ForEach-Object { 
                        if ($Colours) {
                            $_.Text = "$($col.ColumnName -replace '_', ' ')"
                        } else {
                            $_.Text = "$($col.Name)"
                        }
                        $_.Font.Size = 10
                    }
                    $tbl.Table.Cell($i, $j) | ForEach-Object {
                        $_.Borders(1).ForeColor.RGB = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                        $_.Borders(2).ForeColor.RGB = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                        $_.Borders(3).ForeColor.RGB = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                        $_.Borders(4).ForeColor.RGB = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                    }
                    $tbl.Table.Cell($i, $j).Shape.Fill.ForeColor.RGB = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                    $tbl.Table.Cell($i, $j).Shape.TextFrame.TextRange.Font.Color.RGB = [Long] (Get-ColourFromRGB -red $fgRed -green $fgGreen -blue $fgBlue)
                    $j++
                }
                $i++
            }
            $j = 1
            if ($rw -lt $GridResult.Rows.Count) {
                if ($Colours) {
                    $cells = $row.ItemArray | select @{L='Value'; E=({$_})}
                } else {
                    $cells = $row.cells
                }
                foreach ($cell in $cells) {
                    $val = $cell.Value
                    $tbl.Table.Cell($i, $j).Shape.TextFrame.TextRange | ForEach-Object {
                        $_.Text = "$val"
                        $_.Text = $_.Text.Replace("<BR>", "`n")
                        $_.Font.Size = 10
                        if("$val" -like '*</A>*') {
                            $aTbl = Get-HyperlinkProperties -Html $val
                            foreach ($r in $aTbl.Rows) {
                                $tbl.Table.Cell($i, $j).Shape.TextFrame.TextRange.Find($r['Anchor']) | ForEach-Object {
                                    $_.ActionSettings(1).Hyperlink.Address = $r['URL']
                                    $_.Replace("$(($r['Anchor'] -split '>">')[0])>`">", '')
                                    $_.Replace("$(($r['Anchor'] -split '>')[0])>", '')
                                    $_.Replace("$(($_.Text -split '<')[1])", '')
                                    $_.Replace('<', '')
                                }
                            }
                        }
                    }
                    $tbl.Table.Cell($i, $j) | ForEach-Object {
                            $_.Borders(1).ForeColor.RGB = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                            $_.Borders(2).ForeColor.RGB = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                            $_.Borders(3).ForeColor.RGB = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                            $_.Borders(4).ForeColor.RGB = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                    }
                    if ($i -eq 1) {
                        $tbl.Table.Cell($i, $j).Shape.TextFrame.TextRange.Font.Color.RGB = [Long] (Get-ColourFromRGB -red $fgRed -green $fgGreen -blue $fgBlue)
                    } else {
                        $tbl.Table.Cell($i, $j).Shape.Fill.ForeColor.RGB = [Long] (Get-ColourFromRGB -red 255 -green 255 -blue 255)
                    }
                    $j++
                }
                $i++
                $rw++
            }
            $tbl.Table.Rows | ForEach-Object {
                $_.Height = '0.5'
            }
        }
    }
    End
    {
        return $false
    }
}

function Format-WordTable 
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $WordDoc,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [psobject]
        $GridResult,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [psobject]
        $Colours
    )
    Begin
    {
        if ($Colours)
        {
            $cboRed = @{'Text' = $Colours.'background-color'.R}
            $cboGreen = @{'Text' = $Colours.'background-color'.G}
            $cboBlue = @{'Text' = $Colours.'background-color'.B}
            $fgRed = $Colours.color.R
            $fgGreen = $Colours.color.G
            $fgBlue = $Colours.color.B
        }
        else
        {
            $fgRed = 255
            $fgGreen = 255
            $fgBlue = 255
        }
    }
    Process
    {
        if ($Colours) {
            $rows = $GridResult.Rows.Count + 1
        } else {
            $rows = $GridResult.Rows.Count
        }
        $tbl = $WordDoc.Application.Selection.Tables.Add($WordDoc.Application.Selection.Range(), $rows, $GridResult.Columns.Count)
        $tbl | ForEach-Object { 
            $_.borders.InsideLineStyle = 1
            $_.borders.OutsideLineStyle = 1
            $_.borders.InsideColor = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
            $_.borders.OutsideColor = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
        }
        $pg = $tbl.Rows[1].Range.Information(3)
        $i = 1
        foreach ($row in $GridResult.Rows) { 
            $tbl.Rows[$i].AllowBreakAcrossPages = $false
            if ($i -le $rows) {
                if ($pg -lt $tbl.Rows[$i].Range.Information(3)) {
                    $pg = $tbl.Rows[$i].Range.Information(3)
                    $tbl.Rows.Add($tbl.Rows[$i])
                    $j = 1
                    foreach ($col in $GridResult.Columns) {
                        $tbl.Rows[$i].Cells[$j].range | ForEach-Object { 
                            if ($Colours) {
                                $_.text = "$($col.ColumnName -replace '_', ' ')"
                            } else {
                                $_.text = "$($col.Name)"
                            }
                            $_.Shading.BackgroundPatternColor = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                            $_.font.bold = $true
                            $_.font.color = [Long] (Get-ColourFromRGB -red $fgRed -green $fgGreen -blue $fgBlue)
                        }
                        $j++
                    }
                    $i++
                    $rows++
                }
                if ($i -eq 1) {
                    $j = 1
                    foreach ($col in $GridResult.Columns) {
                        $tbl.Rows[$i].Cells[$j].range | ForEach-Object { 
                            if ($Colours) {
                                $_.text = "$($col.ColumnName -replace '_', ' ')"
                            } else {
                                $_.text = "$($col.Name)"
                            }
                            $_.Shading.BackgroundPatternColor = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                            $_.font.bold = $true
                            $_.font.color = [Long] (Get-ColourFromRGB -red $fgRed -green $fgGreen -blue $fgBlue)
                        }
                        $j++
                    }
                    $i++
                }
                $j = 1
                if ($Colours) {
                    $cells = $row.ItemArray | select @{L='Value'; E=({$_})}
                } else {
                    $cells = $row.Cells
                }
                foreach ($cell in $cells) {
                    $val = $cell.Value
                    $tbl.Rows[$i].Cells[$j].Range | ForEach-Object { 
                        $_.Text = "$($val -replace '<BR>', "`n")"
                    }
                    if("$val" -like '*</A>*') {
                        $aTbl = Get-HyperlinkProperties -Html $val
                        foreach ($r in $aTbl.Rows) {
                            $tbl.Rows[$i].Cells[$j].Range | ForEach-Object {
                                $_.Find | ForEach-Object {
                                    $_.Forward = $true
                                    $_.Wrap = 0
                                    $_.Execute($r['Anchor'])
                                }
                                $_.Select
                                $_.Hyperlinks.Add($_, $r['URL'], '', '', $r['Text']) 
                            }
                        }
                    }
                    $j++
                }
                $tbl.Rows[$i].AllowBreakAcrossPages = $false
                if ($pg -lt $tbl.Rows[$i].Range.Information(3) -and $i -gt 2) {
                    $tbl.Rows.Add($tbl.Rows[$i])
                    if  ($pg -eq $tbl.Rows[$i - 1].Range.Information(3)) {
                        $tbl.Rows[$i - 1].SetHeight(22.4, 2)
                    }
                    while ($pg -eq $tbl.Rows[$i].Range.Information(3)) {
                        $tbl.Rows[$i - 1].SetHeight($tbl.Rows[$i - 1].Height + 5, 2)
                    }
                    $pg = $tbl.Rows[$i].Range.Information(3)
                    $j = 1
                    foreach ($col in $GridResult.Columns) {
                        $tbl.Rows[$i].Cells[$j].range | ForEach-Object { 
                            if ($Colours) {
                                $_.text = "$($col.ColumnName -replace '_', ' ')"
                            } else {
                                $_.text = "$($col.Name)"
                            }
                            $_.Shading.BackgroundPatternColor = [Long] (Get-ColourFromRGB -red $cboRed.Text -green $cboGreen.Text -blue $cboBlue.Text)
                            $_.font.bold = $true
                            $_.font.color = [Long] (Get-ColourFromRGB -red $fgRed -green $fgGreen -blue $fgBlue)
                        }
                        $j++
                    }
                    $i++
                    $rows++
                } elseif ($pg -lt $tbl.Rows[$i].Range.Information(3)) {
                        $pg = $tbl.Rows[$i].Range.Information(3)
                }
                $i++
            }
        }
    }
    End
    {
        return $false
    }
}

function Send-ToEmailNew
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Html
    )
    Begin
    {
    }
    Process
    {
        $ol = New-Object -ComObject Outlook.Application
        $oMail = $ol.CreateItem(0)
        $oMail.Display()
        $oMail.HTMLBody = "`r`n$Html"
    }
    End
    {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol)
    }
}

function Send-ToEmailOpen
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Html
    )
    Begin
    {
    }
    Process
    {
        $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        $ol.ActiveInspector().WordEditor.Application.Selection = "placeHere"
        $ol.ActiveInspector().CurrentItem.HTMLBody = $ol.ActiveInspector().CurrentItem.htmlBody.replace("placeHere", "$Html")
    }
    End
    {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol)
    }
}

function Send-ToPowerPointNew
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $GridResult,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [psobject]
        $Colours
    )
    Begin
    {
        $ppt = New-Object -ComObject Powerpoint.Application
        #$ppt.visible = $true
    }
    Process
    {
        $prs = $ppt.Presentations.Add()
        Format-PowerPointTable -PptPresentation $prs -SlideNumber 0 -GridResult $GridResult -Colours $Colours
    }
    End
    {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
    }
}

function Send-ToPowerPointOpen
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $GridResult,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [psobject]
        $Colours
    )
    Begin
    {
        $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Powerpoint.Application")
    }
    Process
    {
        $prs = $ppt.ActivePresentation
        Format-PowerPointTable -PptPresentation $prs -SlideNumber $ppt.ActiveWindow.Selection.SlideRange.SlideNumber -GridResult $GridResult -Colours $Colours
    }
    End
    {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
    }
}

function Send-ToWordNew
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $GridResult,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [psobject]
        $Colours
    )
    Begin
    {
        $wrd = New-Object -ComObject Word.Application
        $wrd.visible = $true
    }
    Process
    {
        $doc = $wrd.Documents.Add()
        $doc.PageSetup.Orientation = 1
        $doc.PageSetup | ForEach-Object {
            $_.TopMargin = Convert-CentimetersToPoints -centimeter 1.75
            $_.LeftMargin = Convert-CentimetersToPoints -centimeter 1.75
            $_.BottomMargin = Convert-CentimetersToPoints -centimeter 1.75
            $_.RightMargin = Convert-CentimetersToPoints -centimeter 1.75
        }
        Format-WordTable -WordDoc $doc -GridResult $GridResult -Colours $Colours
    }
    End
    {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wrd)
    }
}

function Send-ToWordOpen
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $GridResult,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [psobject]
        $Colours
    )
    Begin
    {
        $wrd = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    }
    Process
    {
        $doc = $wrd.ActiveDocument
        Format-WordTable -WordDoc $doc -GridResult $GridResult -Colours $Colours
    }
    End
    {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wrd)
    }
}