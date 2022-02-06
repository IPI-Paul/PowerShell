#[System.Windows.MessageBox]::Show($html);
function buildFilter() {
        $flt = $null;
        foreach ($row in $gridFilter.Rows) {
            if ($row.Cells[1].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + (convertType -row $row, -col ("[" + $row.Cells[0].Value + "]")) + " = ";
                $flt = $flt + (checkType -row $row) + $row.Cells[1].Value + (checkType -row $row);
            }
            if ($row.Cells[2].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + "[" + $row.Cells[0].Value + "] like '%" + $row.Cells[2].Value + "%'";
            }
            if ($row.Cells[3].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + "not [" + $row.Cells[0].Value + "] like '%" + $row.Cells[3].Value + "%'";
            }
            if ($row.Cells[4].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + (convertType -row $row, -col ("[" + $row.Cells[0].Value + "]")) + " in (";
                $flt = $flt + (checkType -row $row) + ($row.Cells[4].Value.split(",") -join (checkType -row $row) + "," + (checkType -row $row)) + (checkType -row $row) + ")";
            }
            if ($row.Cells[5].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + "not " + (convertType -row $row, -col ("[" + $row.Cells[0].Value + "]")) + " in (";
                $flt = $flt + (checkType -row $row) + ($row.Cells[5].Value.split(",") -join (checkType -row $row) + "," + (checkType -row $row)) + (checkType -row $row) + ")";
            }
            if ($row.Cells[6].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + (convertType -row $row, -col ("[" + $row.Cells[0].Value + "]")) + " >= ";
                $flt = $flt + (checkType -row $row) + $row.Cells[6].Value + (checkType -row $row);
            }
            if ($row.Cells[7].Value -ne $null) {
                $flt = $flt + [Environment]::NewLine + "and " + [Environment]::NewLine + (convertType -row $row, -col ("[" + $row.Cells[0].Value + "]")) + " <= ";
                $flt = $flt + (checkType -row $row) + $row.Cells[7].Value + (checkType -row $row);
            }
        }
        return $flt;
}
function buildQuery($idx, $rIdx) {
    if ($idx -ne 0) {
        $sql = $null;
        foreach ($row in $gridFilter.Rows) {
            if ($row.Cells[14].Value -eq $true) {
                if ($sql -gt $null) {
                    $sql = $sql + [Environment]::NewLine + ",";
                    $whr = $whr + [Environment]::NewLine;
                } else {
                    $sql = $sql + [Environment]::NewLine;
                    $whr = "where " + [Environment]::NewLine + "not " + [Environment]::NewLine;
                }
                $sql = $sql + (formatColumns -row $row -col (grouping -row $row -col ((convertType -row $row -col ("[" + $row.Cells[0].Value + "]")))));
                if ($row.Cells[14].Value -eq $true -and ($row.Cells[10].Value -in ("cdbl", "cdate") -or $row.Cells[11].Value -gt "" -or $row.Cells[12].Value -in ("avg", "first", "last", "min", "max", "sum"))) {
                    $sql = $sql + " as [" + $row.Cells[0].Value + "] ";
                }
                $whr = $whr + "[" + $row.Cells[0].Value + "] & ";
            }
        }
        $whr = $whr.Substring(0, ($whr.Length -2)) + " = """"";
        $sql = "Select $sql " + [Environment]::NewLine + [Environment]::NewLine + "from " + [Environment]::NewLine + "[" + $cboNames.Text + "]";
        $sql = $sql + [Environment]::NewLine + [Environment]::NewLine + $whr + (buildFilter) + (groupBy) + (having);
        if ($rIdx -eq $oIdx[2]) {
            $sql = $sql + (orderBY);
        } 
        $txtSQL.Text = $sql
    }
}
function CentimetersToPoints($centimeter) {
    $point = $centimeter * 28.3464567;
    return $point;
}
function checkType($row) {
    $flt = $null;
    if ($row.Cells[9].Value -notin ("Double", "DateTime", "Decimal", "Int16", "Int32", "Int64", "Single", "TimeSpan", "UInt16", "UInt32", "UInt64") -and 
            $row.Cells[10].Value -notin ("cdbl", "cdate")) {
        $flt = "'";
    }
    if ($row.Cells[9].Value -in ("DateTime", "TimeSpan") -or $row.Cells[10].Value -eq "cdate") {
        $flt = "#";
    }
    return $flt;
}
function cleanUpExcel($filepath) {
    try {
        $xls = Get-Process | where {$_.ProcessName -like "Excel"};
        foreach ($xl in $xls) {
            $me = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application");
            $me.Workbooks($filepath.Split("\")[$filepath.Split("\").length-1]).Close($false);
            if ($xl.MainWindowTitle -eq "") {
                spps -Id $xl.Id;
            } 
        }
    } catch {}
}
function clearNames() {
    $cboNames.Items.Clear();
    $cboNames.Items.Add("");
    $cboNames.SelectedIndex = 0;
}
function convertType($row, $col) {
    if ($row.Cells[10].Value -in ("cdbl", "cdate")) {
        return ($row.Cells[10].Value + "(iif($col = """", 0, iif(isnull($col) = true, 0, $col)))"); 
    } else {
        return $col; 
    }
}
function formatColumns($row, $col) {
    if ($row.Cells[11].Value -gt "") {
        return ("format($col,""" + $row.Cells[11].Value + """)");
    } else  {
        return $col;
    }
}
function Get-FileName($initialDirectory, $flt = $null) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")| Out-Null;

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog;
    $OpenFileDialog.InitialDirectory= $initialDirectory;
    $OpenFileDialog.Filter = "$flt" + "All Files (*.*)| *.*";
	$OpenFileDialog.ShowDialog()| Out-Null;
    $OpenFileDialog.FileName;
}
function getNames() {
        $filepath = $txtFilePath.Text;
        cleanUpExcel -filepath $filepath;
        $app = New-Object -ComObject Excel.Application;
        $wb = $app.Workbooks.Open($filepath);
        
        clearNames;
        foreach ($nm in $wb.Names) {
            if ($nm.Visible -eq $true -and $nm.Name -notcontains '!') {
                $cboNames.Items.Add($nm.Name);
            }
        }
        foreach ($sh in $wb.Sheets) {
            if ([System.IO.Path]::GetExtension($filepath) -eq ".csv") {
                $cboNames.Items.Add($sh.Name + ".csv");
            } else {
                $cboNames.Items.Add($sh.Name + "$");
            }
        }
                    
        $wb.Close($false);
        $app.Quit();
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app);
        cleanUpExcel -filepath $filepath;    
}
function groupBy() {
    $grp = $null;
    $grpd = $null;
    foreach ($row in $gridFilter.Rows) {
        if ($row.Cells[12].Value -gt "") {
            $grpd = $true;
            break;
        }
    }
    foreach ($row in $gridFilter.Rows) {
        if ($grpd -eq $true) {
            if ($row.Cells[12].Value -eq "group" -or ($row.Cells[12].Value -eq $null -and $row.Cells[14].Value -eq $true)) {
                if ($grp -gt $null) {
                    $grp = $grp + [Environment]::NewLine + ",";
                } else {
                    $grp = $grp + [Environment]::NewLine;
                }
                $grp = $grp + "[" + $row.Cells[0].Value + "]";
            } 
        }
    }
    if ($grpd -eq $true) {
        $grp = [Environment]::NewLine + [Environment]::NewLine + "Group By " + $grp;
    }
    return $grp;
}
function grouping($row, $col) {
    if ($row.Cells[14].Value -eq $true -and $row.Cells[12].Value -in ("avg", "first", "last", "min", "max", "sum")) {
        return ($row.Cells[12].Value + "($col)"); 
    } else {
        return $col;
    }
}
function having() {
    $hav = $null;
    foreach ($row in $gridFilter.Rows) {
        if ($row.Cells[8].Value -gt "") {
            if ($grp -gt $null) {
                $hav = $hav + [Environment]::NewLine + ",";
            } else {
                $hav = $hav + [Environment]::NewLine;
            }
            if ($row.Cells[8].Value -split "" -notcontains "=" -and $row.Cells[8].Value -split "" -notcontains "<" -and $row.Cells[8].Value -split "" -notcontains ">") {
                $hav = $hav + (grouping -row $row -col ((convertType -row $row -col ("[" + $row.Cells[0].Value + "]")))) + " = " + $row.Cells[8].Value;
            } else {
                $hav = $hav + (grouping -row $row -col ((convertType -row $row -col ("[" + $row.Cells[0].Value + "]")))) + $row.Cells[8].Value;
            }
        } 
    }
    if ($hav -gt $null) {
        $hav = [Environment]::NewLine + [Environment]::NewLine + "Having " + $hav;
    }
    return $hav;
}
function idxColour($col) {
    for ($i = 0; $i -lt 256; $i++) {
        $col.Items.Add($i);
    }
}
function loadFilter() {
    Set-Variable -Name "oIdx" -Value (0, 0, 0) -Scope Global;    
    $gridFilter.Rows.Clear();
    clearNames;
    $filePath = Get-FileName  -initialDirectory "$home\Documents" -flt "Grid TSV (*.grid)|*.grid|TSV (*.tsv)|*.tsv|";
    $grid = [IO.File]::ReadLines($filePath);
    $txtFilePath.Text = ($grid -split [Environment]::NewLine)[0];
    getNames;
    $i = 0;
    foreach($nm in $cboNames.Items) {
        if ($nm -eq ($grid -split [Environment]::NewLine)[1]) {
            $cboNames.SelectedIndex = $i;
            break;
        }
        $i++;
    }
    for ($i = 2; $i -lt ($grid -split [Environment]::NewLine).Length; $i++) {
        $cells = ($grid -split [Environment]::NewLine)[$i];
        foreach($row in $gridFilter.Rows) {
            if ($row.Cells[0].Value -eq $cells.split("`t")[0]) {
                for ($j = 1; $j -lt $row.Cells.Count; $j++) {
                    if ($cells.split("`t")[$j] -gt "" -and $j -ne 9) {
                        $row.Cells[$j].Value = $cells.split("`t")[$j];
                    }
                }
                break;
            }
        }
    }
}
function loadQuery() {
    $gridFilter.Rows.Clear();
    clearNames;
    $filePath = Get-FileName  -initialDirectory "$home\Documents" -flt "SQL (*.sql)|*.sql|";
    $SQL = [IO.File]::ReadLines($filePath);
    $txtSQL.Text = "";
    if (($SQL -split [Environment]::NewLine)[0] -like "# File Path =*") {
        $txtFilePath.Text = (($SQL -split [Environment]::NewLine)[0] -split "= ")[1];
        getNames;
        $j = 1;
    } else {
        $j = 0;
    }
    for ($i = $j; $i -lt ($SQL -split [Environment]::NewLine).Length; $i++) {
        $txtSQL.AppendText(($SQL -split [Environment]::NewLine)[$i] + [Environment]::NewLine);
    }
}
function orderBy() {
    if ($oIdx[0] -gt $null -and $oIdx[0] -gt "" -and $oIdx -lt $gridFilter.Rows[$oIdx[1]].Cells[13].Value) {
        foreach ($row in $gridFilter.Rows) {
            if ($row.Index -ne $oIdx[1] -and $row.Cells[13].Value -ne "" -and $row.Cells[13].Value -ne $null -and $row.Cells[13].Value -gt $oIdx[0]) {
                $oIdx[2] = $row.Index;
                $nVal = ($row.Cells[13].Value -as "Double") - 1;
                $row.Cells[13].Value = "$nVal";
            }
        }
    }
    foreach ($row in $gridFilter.Rows) {
        if ($row.Index -ne $oIdx[1] -and $row.Cells[13].Value -ne "" -and $row.Cells[13].Value -ne $null -and $row.Cells[13].Value -eq $gridFilter.Rows[$oIdx[1]].Cells[13].Value) {
            foreach ($row in $gridFilter.Rows) {
                if ($row.Index -ne $oIdx[1] -and $row.Cells[13].Value -ne "" -and $row.Cells[13].Value -ne $null -and $row.Cells[13].Value -ge $gridFilter.Rows[$oIdx[1]].Cells[13].Value) {
                    $oIdx[2] = $row.Index;
                    $nVal = ($row.Cells[13].Value -as "Double") + 1;
                    $row.Cells[13].Value = "$nVal";
                }
            }
        }
    }
    $ord = @{};
    foreach ($row in $gridFilter.Rows) {
        if ($row.Cells[13].Value -ne "" -and $row.Cells[13].Value -ne $null) {
            $ord[(($row.Cells[13].Value -as "int") - 1)] = $row.Cells[0].Value;
        }
    }
    if ($ord[0] -gt "") {
        $ordBy = $null;
        for ($i = 0; $i -lt $ord.Count; $i++) {
            if ($ordBy -gt $null) {
                $ordBy = $ordBy + [Environment]::NewLine + ",";
            } else {
                $ordBy = $ordBy + [Environment]::NewLine;
            }
            $ordBy = $ordBy + "[" + $ord[$i] + "]";
        }
        $ordBy = [Environment]::NewLine + [Environment]::NewLine + "Order By "+ $ordBy;
        return $ordBy;
    } else {
        return $null;
    }
}
function resizeItems() {
    $gridResult.Size = New-Object System.Drawing.Size(($objForm.Width - 20), ($objForm.Height - 70));
    if ($maxed -eq $null) { 
        $txtSQL.Size = New-Object System.Drawing.Size((($objForm.Width - 20) / 4), ($objForm.Height - 70));
        $gridFilter.Location = New-Object System.Drawing.Size(((($objForm.Width - 20) / 4) + 2),25);
        $gridFilter.Size = New-Object System.Drawing.Size(((($objForm.Width - 28) / 4) * 3), ($objForm.Height - 70));
    } elseif ($maxed -eq "txt") { 
        $txtSQL.Size = New-Object System.Drawing.Size(($objForm.Width - 20), ($objForm.Height - 70));
    } else { 
        $gridFilter.Location = New-Object System.Drawing.Size(2, 25);
        $gridFilter.Size = New-Object System.Drawing.Size(($objForm.Width - 20), ($objForm.Height - 70));
    }
}
function RGB($red, $green, $blue) {
    $col = (([Long] $red) + (([Long] $green) * 256) + (([Long] $blue) * 65536));
    return $col;
}
function runAction() {
    if ($cboAction.SelectedIndex -gt 0) {
            $cboNames.Visible = $true;
            $cboRed.Visible = $false;
            $cboGreen.Visible = $false;
            $cboBlue.Visible = $false;
        if ($cboAction.Text -eq "Edit Query") {
            $gridFilter.Visible = $true;
            $txtSQL.Visible = $true;
            $gridResult.Visible = $false;
        } elseif ($cboAction.Text -eq "Maximize Query Textbox" -or $cboAction.Text -eq "Load Query File") {
            $gridFilter.Visible = $false;
            $txtSQL.Visible = $true;
            $gridResult.Visible = $false;
            Set-Variable -Name "maxed" -Value ("txt") -Scope Global;
            resizeItems;
            if ($cboAction.Text -eq "Load Query File") {
                loadQuery;
            }
        } elseif ($cboAction.Text -eq "Maximize Filter Grid") {
            $gridFilter.Visible = $true;
            $txtSQL.Visible = $false;
            $gridResult.Visible = $false;
            Set-Variable -Name "maxed" -Value ("grid") -Scope Global;
            resizeItems;
        } elseif ($cboAction.Text -eq "Restore Editors" -or $cboAction.Text -eq "Load Filter File") {
            $gridFilter.Visible = $true;
            $txtSQL.Visible = $true;
            $gridResult.Visible = $false;
            Set-Variable -Name "maxed" -Value ($null) -Scope Global;
            resizeItems;
            if ($cboAction.Text -eq "Load Filter File") {
                loadFilter;
            }
        } elseif ($cboAction.Text -eq "Run Query") {
            $gridFilter.Visible = $false;
            $txtSQL.Visible = $false;
            $gridResult.Visible = $true;
            runUpdSQL;
        } elseif ($cboAction.Text -eq "Save Filter Grid to File") {
            saveGrid;
        } elseif ($cboAction.Text -eq "Save Query to File") {
            saveQuery;
        } elseif ($cboAction.Text -eq "Send To New Email") {
            sendToNewEmail;
        } elseif ($cboAction.Text -eq "Send To Open Email") {
            sendToOpenEmail;
        } elseif ($cboAction.Text -eq "Send To New Word Document") {
            sendToNewWord;
        } elseif ($cboAction.Text -eq "Send To Open Word Document") {
            sendToOpenWord;
        } elseif ($cboAction.Text -eq "Send To New PowerPoint Document") {
            sendToNewPPt;
        } elseif ($cboAction.Text -eq "Send To Open PowerPoint Document") {
            sendToOpenPPt;
        } elseif ($cboAction.Text -eq "Set Default Header RGB Colours") {
            $cboNames.Visible = $false;
            $cboRed.Visible = $true;
            $cboGreen.Visible = $true;
            $cboBlue.Visible = $true;
            $cboRed.Focus();
        } elseif ($cboAction.Text -eq "View Results Grid") {
            $gridFilter.Visible = $false;
            $txtSQL.Visible = $false;
            $gridResult.Visible = $true;
        }

        $cboAction.SelectedIndex = 0;
    }
}
function runUpdSQL() {
        $filepath = $txtFilePath.Text;
        if ($cboDriver.Text -eq "JET") {
            if ([System.IO.Path]::GetExtension($filepath) -eq ".csv") {
                $connString = "Provider = Microsoft.JET.OLEDB.4.0; Data Source=" + (Split-Path -Path $filepath) + ';Mode=Share Deny None; Extended Properties = "text;HDR=Yes;FMT=Delimited"';
            } else {
                cleanUpExcel -filepath $filepath;
                $app = New-Object -ComObject Excel.Application;
                $wb = $app.Workbooks.Open($filepath);
                $connString = "Provider = Microsoft.JET.OLEDB.4.0; Data Source=" + $filepath + ";Extended Properties=""Excel 8.0;IMEX=1;HDR=YES"";Persist Security Info=False;";
            }
        } else {
            if ([System.IO.Path]::GetExtension($filepath) -eq ".csv") {
                $connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + (Split-Path -Path $filepath) + ';Mode=Share Deny None; Extended Properties = "text;HDR=Yes;FMT=Delimited"';
            } else {
                $connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + $filepath + ";Extended Properties=""Excel 12.0;IMEX=1;HDR=YES"";Persist Security Info=False;";
            }
        }
        $conn = New-Object "System.Data.OleDb.OleDbConnection" $connString;
        $comm = New-Object "System.Data.OleDb.OleDbCommand";
        $commType = [System.Data.CommandType]"Text";
        $comm.CommandText = ($txtSQL.Text);
        $comm.Connection = $conn;

        $conn.Open();
        $adapter = New-Object "System.Data.OleDb.OleDbDataAdapter" $comm;
        $dt = New-Object System.Data.DataSet;
        $adapter.Fill($dt);
        $html = ($dt.Tables[0] | select * -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray | ConvertTo-Html -As Table -Fragment -Property *);
        Set-Variable -Name "html" -Value ($html) -Scope Global;

        if ($dt -ne $null){
            $gridResult.DataSource = $dt.Tables[0];
            $gridResult.Update();
        }

        $comm.Dispose();
        $conn.Close();
        $conn.Dispose();            
        if ($cboDriver.Text -eq "JET" -and [System.IO.Path]::GetExtension($filepath) -ne ".csv") {
            $wb.Close($false);
            $app.Quit();
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app);
            cleanUpExcel -filepath $filepath;
        }
}
function saveGrid() {
    $filePath = savePath("Grid TSV (*.grid)|*.grid|TSV (*.tsv)|*.tsv");
    if ($filePath) {
        $grid = $txtFilePath.Text + [Environment]::NewLine + $cboNames.Text + [Environment]::NewLine;
        foreach ($row in $gridFilter.Rows) {
            foreach ($cell in $row.Cells) {
                $grid = $grid + $cell.Value + "`t";
            }
            $grid = $grid.Substring(0, $grid.Length-1) + [Environment]::NewLine;
        }
        $grid > $filePath[1];
    }
}
function savePath($flt) {
    $dialogSave = New-Object System.Windows.Forms.SaveFileDialog;
    $dialogSave.InitialDirectory = "$home\Documents";
    $dialogSave.Filter = "$flt|All files (*.*)|*.*";
    $dialogSave.ShowDialog();
    $dialogSave.FileName;
}
function saveQuery() {
    $filePath = savePath("SQL (*.sql)|*.sql");
    if ($filePath) {
        "# File Path = " + $txtFilePath.Text + [Environment]::NewLine + $txtSQL.Text > $filePath[1];
    }
}
function sendToNewEmail() {
    $ol = New-Object -ComObject Outlook.Application;
    $oMail = $ol.CreateItem(0);
    $oMail.Display();
    $html = styleHTML;
    $oMail.HTMLBody = "`n`n$html";
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol);
}
function sendToOpenEmail() {
    $ol = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application");
    $ol.ActiveInspector().WordEditor.Application.Selection = "placeHere";
    $html = styleHTML;
    $ol.ActiveInspector().CurrentItem.HTMLBody = $ol.ActiveInspector().CurrentItem.htmlBody.replace("placeHere", "$html");
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ol);
}
function sendToNewPPt() {
    $ppt = New-Object -ComObject Powerpoint.Application;
    $ppt.visible = $true;
    $prs = $ppt.Presentations.Add();
    tblPPt -prs $prs -pos 0;
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt);
}
function sendToOpenPPt() {
    $ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Powerpoint.Application");
    $prs = $ppt.ActivePresentation;
    tblPPt -prs $prs -pos $ppt.ActiveWindow.Selection.SlideRange.SlideNumber;
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt);
}
function sendToNewWord() {
    $wrd = New-Object -ComObject Word.Application;
    $wrd.visible = $true;
    $doc = $wrd.Documents.Add();
    $doc.PageSetup.Orientation = 1;
    $doc.PageSetup | ForEach-Object {
        $_.TopMargin = CentimetersToPoints -centimeter 1.75;
        $_.LeftMargin = CentimetersToPoints -centimeter 1.75;
        $_.BottomMargin = CentimetersToPoints -centimeter 1.75;
        $_.RightMargin = CentimetersToPoints -centimeter 1.75;
    }
    tblWord -doc $doc;
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wrd);
}
function sendToOpenWord() {
    $wrd = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application");
    $doc = $wrd.ActiveDocument;
    tblWord -doc $doc;
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wrd);
}
function styleHTML() {
    $stl = "<style>`n";
    $stl = $stl + "table {`n";
    $stl = $stl + "`tborder-collapse: collapse;`n";
	$stl = $stl + "}`n";
    $stl = $stl + "table, th, td {`n";
    $stl = $stl + "`tpadding: 0px 5px 0px 5px;`n";
    $stl = $stl + "`tborder: 1 solid black;`n";
    $stl = $stl + "`tfont-size: 1em;`n";
	$stl = $stl + "}`n";
    $stl = $stl + "td {`n";
	$stl = $stl + "`tcolor: black;`n";
	$stl = $stl + "}`n";
    $stl = $stl + "th {`n";
	$stl = $stl + "`tcolor: white;`n";
	$stl = $stl + "`tbackground-color: rgb(" + $cboRed.Text + ", " + $cboGreen.Text + ", " + $cboBlue.Text + ");`n";
	$stl = $stl + "}`n";
    $stl = $stl + "</style>`n";
    $html = "$stl$html";
    return $html;
}
function tblPPt($prs, $pos) {
    $i = 1;
    $rows = $gridResult.Rows.Count;
    $rw = 1;
    foreach ($row in $gridResult.Rows) {
        if ((13 % $i) -eq 0) {
            $pos++;
            $sld = $prs.Slides.Add($pos, 12);
            if ($rows -gt 12) {
                $mrows = 12;
                $rows = $rows - 11;
            } else {
                $mrows = $rows;
            }
            $tbl = $sld.Shapes.AddTable($mrows, $gridResult.Columns.Count);
            $tbl | ForEach-Object {
                $_.Top = 20;
                $_.Left = 20;
                $_.Width = 920;
                $_.Height = (500 / 12) * $mrows;
            }
            $i = 1;
            $j = 1;
            foreach ($col in $gridResult.Columns) {
                $tbl.Table.Cell($i, $j).Shape.TextFrame.TextRange | ForEach-Object { 
                    $_.Text = $col.Name;
                    $_.Font.Size = 10;
                }
                $tbl.Table.Cell($i, $j) | ForEach-Object {
                    $_.Borders(1).ForeColor.RGB = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                    $_.Borders(2).ForeColor.RGB = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                    $_.Borders(3).ForeColor.RGB = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                    $_.Borders(4).ForeColor.RGB = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                }
                $tbl.Table.Cell($i, $j).Shape.Fill.ForeColor.RGB = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                $j++
            }
            $i++;
        }
        $j = 1;
        if ($rw -lt $gridResult.Rows.Count) {
            foreach ($cell in $row.cells) {
                $val = $cell.Value;
                $tbl.Table.Cell($i, $j).Shape.TextFrame.TextRange | ForEach-Object {
                    $_.Text = "$val";
                    $_.Font.Size = 10;
                }
                $tbl.Table.Cell($i, $j) | ForEach-Object {
                        $_.Borders(1).ForeColor.RGB = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                        $_.Borders(2).ForeColor.RGB = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                        $_.Borders(3).ForeColor.RGB = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                        $_.Borders(4).ForeColor.RGB = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                }
                if ($i -eq 1) {
                    $tbl.Table.Cell($i, $j).Shape.Fill.ForeColor.RGB = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                } else {
                    $tbl.Table.Cell($i, $j).Shape.Fill.ForeColor.RGB = [Long] (RGB -red 255 -green 255 -blue 255);
                }
                $j++;
            }
            $i++;
            $rw++;
        }
    }
}
function tblWord($doc) {
    $tbl = $doc.Application.Selection.Tables.Add($doc.Application.Selection.Range(), $gridResult.Rows.Count, $gridResult.Columns.Count);
    $tbl | ForEach-Object { 
        $_.borders.InsideLineStyle = 1;
        $_.borders.OutsideLineStyle = 1;
        $_.borders.InsideColor = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
        $_.borders.OutsideColor = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
    }
    $pg = $tbl.Rows[1].Range.Information(3);
    $i = 1;
    foreach ($row in $gridResult.Rows) { 
        if ($i -le $gridResult.Rows.Count) {
            if ($pg -lt $tbl.Rows[$i].Range.Information(3)) {
                $pg = $tbl.Rows[$i].Range.Information(3);
                $tbl.Rows.Add($tbl.Rows[$i]);
                $j = 1;
                foreach ($col in $gridResult.Columns) {
                    $tbl.Rows[$i].Cells[$j].range | ForEach-Object { 
                        $_.text = $col.Name;
                        $_.Shading.BackgroundPatternColor = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                        $_.font.bold = $true;
                    }
                    $j++;
                }
                $i++;
            }
            if ($i -eq 1) {
                $j = 1;
                foreach ($col in $gridResult.Columns) {
                    $tbl.Rows[$i].Cells[$j].range | ForEach-Object { 
                        $_.text = $col.Name;
                        $_.Shading.BackgroundPatternColor = [Long] (RGB -red $cboRed.Text-green $cboGreen.Text-blue $cboBlue.Text);
                        $_.font.bold = $true;
                    }
                    $j++;
                }
                $i++;
            }
            $j = 1;
            foreach ($cell in $row.Cells) {
                $val = $cell.Value;
                $tbl.Rows[$i].Cells[$j].range.text = "$val";
                $j++;
            }
            $i++;
        }
    }
}
function updateFilter() {
    if ($cboNames.SelectedIndex -eq 0) {
        $gridFilter.Visible = $true;
        $gridFilter.Rows.Clear();
        $txtSQL.Visible = $true;
        $gridResult.Visible = $false;
    } else {
        $gridFilter.Visible = $true;
        $txtSQL.Visible = $true;
        $gridResult.Visible = $false;
        Set-Variable -Name "maxed" -Value ($null) -Scope Global;
        resizeItems;
        $filepath = $txtFilePath.Text;
        if ($cboDriver.Text -eq "JET") {
            if ([System.IO.Path]::GetExtension($filepath) -eq ".csv") {
                $connString = "Provider = Microsoft.JET.OLEDB.4.0; Data Source=" + (Split-Path -Path $filepath) + ';Mode=Share Deny None; Extended Properties = "text;HDR=Yes;FMT=Delimited"';
            } else {
                cleanUpExcel -filepath $filepath;
                $app = New-Object -ComObject Excel.Application;
                $wb = $app.Workbooks.Open($filepath);
                $connString = "Provider = Microsoft.JET.OLEDB.4.0; Data Source=" + $filepath + ";Extended Properties=""Excel 8.0;IMEX=1;HDR=YES"";Persist Security Info=False;";
            }
        } else {
            if ([System.IO.Path]::GetExtension($filepath) -eq ".csv") {
                $connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + (Split-Path -Path $filepath) + ';Mode=Share Deny None; Extended Properties = "text;HDR=Yes;FMT=Delimited"';
            } else {
                $connString = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + $filepath + ";Extended Properties=""Excel 12.0;IMEX=1;HDR=YES"";Persist Security Info=False;";
            }
        }
        $conn = New-Object "System.Data.OleDb.OleDbConnection" $connString;
        $comm = New-Object "System.Data.OleDb.OleDbCommand";
        $commType = [System.Data.CommandType]"Text";
        $comm.CommandText = "select top 1 * from [" + $cboNames.Text + "]";
        $comm.Connection = $conn;

        $conn.Open();
        $adapter = New-Object "System.Data.OleDb.OleDbDataAdapter" $comm;
        $dt = New-Object System.Data.DataSet;
        $adapter.Fill($dt);

        $gridFilter.Rows.Clear();

        if ($dt -ne $null){
            $i = 1
            foreach ($col in $dt.Tables[0].Columns) {
                $cboIndex.Items.Add("$i");
                $i++;
            }
            foreach ($col in $dt.Tables[0].Columns) {
                $gridFilter.Rows.Add($col.ColumnName, $null, $null, $null, $null, $null, $null, $null, $null, $col.DataType.Name.ToString());
            }
        }

        $comm.Dispose();
        $conn.Close();
        $conn.Dispose();            
        if ($cboDriver.Text -eq "JET" -and [System.IO.Path]::GetExtension($filepath) -ne ".csv") {
            $wb.Close($false);
            $app.Quit();
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($app);
            cleanUpExcel -filepath $filepath;
        }
    }
}
function viewForm() {
    $objForm = New-Object System.Windows.Forms.Form;
    $objForm.text = "IPI International Excel Connections";
    $objForm.Size = New-Object System.Drawing.Size(1280,720);
    $objForm.StartPosition = "CenterScreen";

    $cboDriver = New-Object System.Windows.Forms.ComboBox;
    $cboDriver.Location = New-Object System.Drawing.Size(2, 2);
    $cboDriver.Size = New-Object System.Drawing.Size(49,20);
    $cboDriver.Items.Add('ACE');
    $cboDriver.Items.Add('JET');
    $cboDriver.SelectedIndex = 0;
    $objForm.Controls.Add($cboDriver);

    $txtFilePath = New-Object System.Windows.Forms.TextBox;
    $txtFilePath.Location = New-Object System.Drawing.Size(52,2);
    $txtFilePath.Size = New-Object System.Drawing.Size(660,25);
    $objForm.Controls.Add($txtFilePath);

    $btnUpdate = New-Object System.Windows.Forms.Button;
    $btnUpdate.Location = New-Object System.Drawing.Size(712,2);
    $btnUpdate.Size = New-Object System.Drawing.Size(30,21);
    $btnUpdate.Text = "...";
    $btnUpdate.Add_Click({$txtFilePath.Text = Get-FileName  -initialDirectory "$home\Documents"; getNames;});
    $objForm.Controls.Add($btnUpdate);

    $cboNames = New-Object System.Windows.Forms.ComboBox;
    $cboNames.Location = New-Object System.Drawing.Size(744,2);
    $cboNames.Size = New-Object System.Drawing.Size(300,21);
    $cboNames.Items.Add('');
    $cboNames.Add_SelectedIndexChanged({updateFilter;});
    $objForm.Controls.Add($cboNames);

    $cboRed = New-Object System.Windows.Forms.ComboBox;
    $cboRed.Location = New-Object System.Drawing.Size(744, 2);
    $cboRed.Size = New-Object System.Drawing.Size(98,20);
    $cboRed.Visible = $false;
    idxColour -col $cboRed;
    $cboRed.SelectedIndex = 0;
    $objForm.Controls.Add($cboRed);

    $cboGreen = New-Object System.Windows.Forms.ComboBox;
    $cboGreen.Location = New-Object System.Drawing.Size(842, 2);
    $cboGreen.Size = New-Object System.Drawing.Size(98,20);
    $cboGreen.Visible = $false;
    idxColour -col $cboGreen;
    $cboGreen.SelectedIndex = 0;
    $objForm.Controls.Add($cboGreen);

    $cboBlue = New-Object System.Windows.Forms.ComboBox;
    $cboBlue.Location = New-Object System.Drawing.Size(940, 2);
    $cboBlue.Size = New-Object System.Drawing.Size(98,20);
    $cboBlue.Visible = $false;
    idxColour -col $cboBlue;
    $cboBlue.SelectedIndex = 108;
    $objForm.Controls.Add($cboBlue);

    $cboAction = New-Object System.Windows.Forms.ComboBox;
    $cboAction.Location = New-Object System.Drawing.Size(1047,2);
    $cboAction.Size = New-Object System.Drawing.Size(210,21);
    $cboAction.Items.Add('');
    $cboAction.Items.Add('Edit Query');
    $cboAction.Items.Add('Run Query');
    $cboAction.Items.Add('View Results Grid');
    $cboAction.Items.Add('Maximize Query Textbox');
    $cboAction.Items.Add('Maximize Filter Grid');
    $cboAction.Items.Add('Restore Editors');
    $cboAction.Items.Add('Load Filter File');
    $cboAction.Items.Add('Save Filter Grid to File');
    $cboAction.Items.Add('Load Query File');
    $cboAction.Items.Add('Save Query to File');
    $cboAction.Items.Add('Send To New Email');
    $cboAction.Items.Add('Send To Open Email');
    $cboAction.Items.Add('Send To New Word Document');
    $cboAction.Items.Add('Send To Open Word Document');
    $cboAction.Items.Add('Send To New PowerPoint Document');
    $cboAction.Items.Add('Send To Open PowerPoint Document');
    $cboAction.Items.Add('Set Default Header RGB Colours');
    $cboAction.DropDownHeight = $cboAction.Items.Count * 24;
    $cboAction.Add_SelectedIndexChanged({runAction;});
    $objForm.Controls.Add($cboAction);

    $txtSQL = New-Object System.Windows.Forms.TextBox;
    $txtSQL.Multiline = $true;
    $txtSQL.Location = New-Object System.Drawing.Size(2,25);
    $txtSQL.Size = New-Object System.Drawing.Size((($objForm.Width - 20) / 4), ($objForm.Height - 70));
    $objForm.Controls.Add($txtSQL);

    $gridFilter = New-Object System.Windows.Forms.DataGridView;
    $gridFilter.Location = New-Object System.Drawing.Size(((($objForm.Width - 20) / 4) + 2),25);
    $gridFilter.Size = New-Object System.Drawing.Size(((($objForm.Width - 28) / 4) * 3), ($objForm.Height - 70));
    $gridFilter.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True;
    $gridFilter.AutoSize = $false;
    $gridFilter.AutoSizeRowsMode = "AllCells";
    $gridFilter.AutoSizeColumnsMode = "AllCells";
    $gridFilter.ColumnCount = 10;
    $gridFilter.Columns[0].Name = "Column Name";
    $gridFilter.Columns[1].Name = "Equals";
    $gridFilter.Columns[2].Name = "Like";
    $gridFilter.Columns[3].Name = "Not Like";
    $gridFilter.Columns[4].Name = "Is In";
    $gridFilter.Columns[5].Name = "Is Not In";
    $gridFilter.Columns[6].Name = "From";
    $gridFilter.Columns[7].Name = "To";
    $gridFilter.Columns[8].Name = "Having";
    $gridFilter.Columns[9].Name = "Type";
    $gridFilter.Add_Click({Set-Variable -Name "oIdx" -Value ($gridFilter.CurrentRow.Cells[13].Value, $gridFilter.CurrentRow.Index, $gridFilter.CurrentRow.Index) -Scope Global;});
    $gridFilter.add_CellValueChanged({buildQuery -idx $_.ColumnIndex -rIdx $gridFilter.CurrentRow.Index;});
    $cboConvert = New-Object System.Windows.Forms.DataGridViewComboBoxColumn;
    $cboConvert.Name = "Convert To";
    $cboConvert.Width = 50;
    $cboConvert.Items.Add("");
    $cboConvert.Items.Add("cdbl");
    $cboConvert.Items.Add("cdate");
    $gridFilter.Columns.Add($cboConvert);
    $cboFormat = New-Object System.Windows.Forms.DataGridViewComboBoxColumn;
    $cboFormat.Name = "Format";
    $cboFormat.Width = 50;
    $cboFormat.Items.Add("");
    $cboFormat.Items.Add("#,###,###,##0.00");
    $cboFormat.Items.Add("#,###,###,##0");
    $cboFormat.Items.Add("#,###,###,##0%");
    $cboFormat.Items.Add("#,###,###,##0.00%");
    $cboFormat.Items.Add("ddd dd mmm yyyy hh:mm");
    $cboFormat.Items.Add("ddd dd mmm yyyy");
    $cboFormat.Items.Add("dd/mm/yyyy hh:mm");
    $cboFormat.Items.Add("dd/mm/yyyy");
    $cboFormat.Items.Add("mmm-yy");
    $cboFormat.Items.Add("hh:mm");
    $cboFormat.DropDownHeight = $cboFormat.Items.Count * 24;
    $gridFilter.Columns.Add($cboFormat);
    $cboGroup = New-Object System.Windows.Forms.DataGridViewComboBoxColumn;
    $cboGroup.Name = "Group By";
    $cboGroup.Width = 50;
    $cboGroup.Items.Add("");
    $cboGroup.Items.Add("group");
    $cboGroup.Items.Add("first");
    $cboGroup.Items.Add("last");
    $cboGroup.Items.Add("min");
    $cboGroup.Items.Add("max");
    $cboGroup.Items.Add("avg");
    $cboGroup.Items.Add("sum");
    $gridFilter.Columns.Add($cboGroup);
    $cboIndex = New-Object System.Windows.Forms.DataGridViewComboBoxColumn;
    $cboIndex.Name = "Index";
    $cboIndex.Width = 50;
    $cboIndex.Items.Add("");
    $gridFilter.Columns.Add($cboIndex);
    $chkShow = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn;
    $chkShow.Name = "Show";
    $chkShow.Width = 30;
    $gridFilter.Columns.Add($chkShow);
    $objForm.Controls.Add($gridFilter);

    $gridResult = New-Object System.Windows.Forms.DataGridView;
    $gridResult.Location = New-Object System.Drawing.Size(2,25);
    $gridResult.Size = New-Object System.Drawing.Size(($objForm.Width - 20), ($objForm.Height - 70));
    $gridResult.DefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::True;
    $gridResult.AutoSize = $false;
    $gridResult.Visible = $false;
    $gridResult.AutoSizeRowsMode = "AllCells";
    $gridResult.AutoSizeColumnsMode = "AllCells";
    $objForm.Controls.Add($gridResult);

    $objForm.TopMost = $False;
    $objForm.Add_Shown({$objForm.Activate()});
    $objForm.Add_Resize({resizeItems;});
    [void]$objForm.ShowDialog();        
}
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing");
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms");

$html = $null;
$oIdx = $null;
$maxed = $null;
viewForm;
