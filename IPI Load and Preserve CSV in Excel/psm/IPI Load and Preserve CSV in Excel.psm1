# Add the COM object type
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

function Expand-ZipToCSV {
    param (
        [System.IO.Compression.ZipArchiveEntry]$CSV,
        [string]$TempPath,
        $Log
    )

    $csvPath = Join-Path $TempPath $CSV.Name
    
    [System.IO.Compression.ZipFileExtensions]::ExtractToFile(
        $CSV,
        $csvPath,
        $true
    )

    return $csvPath
}

function Get-CharacterSet {
    param ($path)

    # Check for UTF8 BOM
    $isUtf8 = (Get-Content $path -Encoding Byte -TotalCount 3) -join ',' -eq '239,187,191'

    # return Character Set
    if ($isUtf8) { "65001" } else { "ANSI" }
}

function Get-CSVColumnSamples {
    param (
        [string]$CsvPath,
        [int]$MaxValues = 50,
        [string]$Delimiter = ",",
        $Log
    )

    # Open StreamReader with UTF-8 encoding
    $sr = [System.IO.StreamReader]::new($CsvPath, [System.Text.Encoding]::UTF8)

    # Read header line
    $headerLine = $sr.ReadLine()
    $headers    = $headerLine -split $Delimiter
    $colCount   = $headers.Count

    # Initialise result hashtable
    $columnPipeStrings = @{}

    # Process each column individually
    for ($col = 0; $col -lt $colCount; $col++) {
        $values = @()

        # Rewind StreamReader for each column
        [void]($sr.BaseStream.Seek(0, [System.IO.SeekOrigin]::Begin))
        $sr.DiscardBufferedData()
        [void]($sr.ReadLine() | Out-Null) # skip header again

        while (-not $sr.EndOfStream -and $values.Count -lt $MaxValues) {
            $line   = $sr.ReadLine()
            $fields = $line -split $Delimiter

            if ($fields.Length -gt $col) {
                $val = $fields[$col].Trim()
                if ($val) { $values += $val }
            }
        }

        # Join collected values with pipe
        $columnPipeStrings[$headers[$col]] = ($values -join "|")
    }

    $sr.Close()
    return $columnPipeStrings
}

function Get-Duration {
    param (
        $startTime
    )
    $endTime    = Get-Date
    $duration   = $endTime - $startTime
    return "{0:00}:{1:00}:{2:00}" -f $duration.Hours, $duration.Minutes, $duration.Seconds
}

function Get-ExcelFormula {
    param (
        [Parameter(Mandatory)]
        [string]$path,
        [string]$colLetter
    )

    ([regex]::Match((Get-Content $path -Raw), "(?s)``excel\s*(.*?)\s*``").Groups[1].Value -replace "\([$]colLetter\)|[$]colLetter", $colLetter)
}

function Get-FilePath {
    param ([string]$Path, $Log)
    
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $csvPath    = $null
    $tempRoot   = $env:TEMP
    $zip        = [System.IO.Compression.ZipFile]::OpenRead($Path)
    $count      = 0
    $csvs       = @{}
    
    try {
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -match "\.csv$") {
                $count++
                $csvs.Add($count, $entry.FullName)
            }
        }
        if ($count -eq 1) {
            $csvPath = (Expand-ZipToCSV -CSV $zip.Entries[0] -TempPath $tempRoot -Log $Log)
        }elseif ($count -gt 1) {
            $selection = Show-CsvPicker -CSVs $csvs -Log $Log

            if ($selection) {
                foreach ($entry in $zip.Entries) {
                    if ($entry.FullName -eq $selection) {
                        $csvPath = Expand-ZipToCSV -CSV $entry -TempPath $tempRoot -Log $Log
                    }
                }
            }
        } 
    } finally {
        $zip.Dispose()
    }

    if ($count -ge 1) {
        return $csvPath
    } else {
        throw "No CSV found in ZIP!"
    }
}

function Get-Headers {
    param ($Log)

    $dest   = [System.IO.Path]::GetTempFileName()
    $fsIn   = [System.IO.File]::OpenRead($Script:FilePath)

    $Headers = (Invoke-CheckHeaders -Path $Script:FilePath)
    $Global:Headers = $Headers
    $colCount = $Headers.Count

    $hasSpace = $false
    foreach ($h in $Headers) {
        if ($h.IndexOf(' ') -ge 0) {
            $hasSpace = $true
            break
        }
    }

    try {
        $sr = New-Object System.IO.StreamReader($fsIn)

        # Read header
        $header = $sr.ReadLine()
        if ($hasSpace) {
            # Normalise headers for schema.ini
            $header = ($header -replace '\s+', '_')
        }
            
        $sr.Dispose()
        $fsIn.Dispose()

        # Write header 
        $header | Set-Content $dest

        # Write remaining data unchanged
        Get-Content $Script:FilePath | Select-Object -Skip 1 | Add-Content $dest
    } catch {
        & $Log $_.Exception.Message "DarkRed" -Bold:$true
    }
    
    $Headers = (Invoke-CheckHeaders -Path $dest)
    Set-FilePath -Path $dest -Temp $true

    $path = $Script:FilePath
    $Script:schemaPath = Join-Path (Split-Path $path) 'schema.ini'

    $schema = @()
    $schema += "[$([System.IO.Path]::GetFileName($path))]"
    $schema += "Format=CSVDelimited"
    $schema += "ColNameHeader=True"
    $schema += "MaxScanRows=0"
    $schema += "CharacterSet=$(Get-CharacterSet $path)"

    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $schema += "Col$($i + 1)=$($Headers[$i]) Text Width 255"
    }

    $schema | Set-Content -Encoding Ascii $Script:schemaPath

    return [PSCustomObject]@{
        Headers     = $Headers
        colCount    = $colCount
    }
}

function Open-Connection {
    # Create ADODB connection to read source file data
    $conn       = New-Object -ComObject ADODB.Connection
    $rs         = New-Object -ComObject ADODB.Recordset
    $connStr    = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=$Global:Folder;Extended Properties='Text;HDR=Yes;FMT=Delimited'"
    $conn.Open($connStr)

    return [PSCustomObject]@{
        conn    = $conn
        rs      = $rs
    }
}

function Get-SpecialColumns {
    param ($Headers, $colCount, $Log)
    
    # Open Excel
    $Excel          = New-Object -ComObject Excel.Application
    $Excel.Visible  = $true

    # Get Excel calculation type
    $xlCalcManual       = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationManual

    $wb         = $Excel.Workbooks.Add()
    $ws1        = $wb.Sheets.Item(1)
    $ws1.Name   = "Sheet1"

    # Add second sheet for filtered data
    $ws2            = $wb.Sheets.Add()
    $ws2.Name       = "Sheet2"

    # Write headers to Sheet1
    for ($i = 0; $i -lt $colCount; $i++) {
        $ws1.Cells.Item(1, $i + 1).Value2 = $Headers[$i]
    }

    # SQL query to get data
    $sql = "SELECT * FROM [$Global:FileName]"

    # Open recordset
    $conn = Open-Connection
    $conn.rs.Open($sql, $conn.conn, 3, 1)

    # Copy all rows into Sheet1 starting at A2
    $ws1.Range("A2").CopyFromRecordset($conn.rs)
    $rowCount = $ws1.UsedRange.Rows.Count() + 1

    # Close recordset and connection
    $conn.rs.Close()
    $conn.conn.Close()

    # Turn off screen updating, calculations and alerts
    $Excel.ScreenUpdating   = $false
    $Excel.Calculation      = $xlCalcManual
    $Excel.DisplayAlerts    = $false

    # Detect special columns using TrimRange/Filters in Sheet2
    for ($i = 1; $i -le $colCount; $i++) {
        # Use A.:.A notation for full column in Office 365 and Offsets in other versions
        $colLetter  = (($ws1.Cells.Item(1, $i).EntireColumn.Address($false, $false) -replace '\d', '') -split ":")[0]
        $formula    = (Get-ExcelFormula -path $Config.FilterFormula -colLetter $colLetter)
        $ws2.Cells.Item(1, $i).Formula2 = ($formula -replace "`r?`n", "")
    }

    # Force Excel to calculate formulae
    $Excel.Calculate()

    # Remove original schema
    Remove-Item $Script:schemaPath -ErrorAction SilentlyContinue

    # Read back results and detect special columns
    $SpecialIndexes = New-Object System.Collections.Generic.List[int]
    $ColumnTypes    = @{} # Header => ACE type (Text, Integer, Double, DateTine)
    $SpecialHeaders = @{}

    # Date regex patterns (very forgiving)
    $datePatterns = @(
        '^\d{4}[-/\.]\d{1,2}[-/\.]\d{1,2}$', # 2024-12-31, 2024/12/31, 2024.12.31
        '^\d{1,2}[-/\.]\d{1,2}[-/\.]\d{4}$', # 12/31/2024, 31.12.2024, etc.
        '^\d{1,2}[-/\.]\d{1,2}[-/\.]\d{2}$'  # 12/31/24
    )

    # Decimal number patters
    $decimalPatterns = @(
        '^\d+\.\d+$',   # 123.456
        '^0\.\d+$',     # 0.456
        '^\.\d+$',      # .25
        '^\d+\.$'       # 25.
    )

    for ($i = 1; $i -le $colCount; $i++) {
        $colHeader = $Headers[$i - 1]

        # Default fallback type:
        $ColumnTypes[$colHeader] = "Text"

        $res = $ws2.Cells.Item(1, $i).Value2
        if (-not $res) { continue }

        $vals = $res -split '\|'

        foreach ($v in $vals) {
            IF ($v.Trim() -match "[\+a-zA-Z]") { break }
            # --- Pattern Tests ---
            $isLeadingZero  = $v.Trim() -match '^0\d+$'
            $isLongNumber   = $v -match '^\d{11,}$'
            $isShortNumber  = $v -match '^\d{1,10}$' -and -not $isLeadingZero

            $isDate = $false
            foreach ($p in $datePatterns) {
                if ($v -match $p) { $isDate = $true; break }
            }

            # --- Decimal tests ---
            $isDecimal = $false
            foreach ($p in $decimalPatterns) {
                if ($v -match $p) { $isDecimal = $true; break }
            }

            # --- Determine ACE type ---
            if ($isLeadingZero) {
                $ColumnTypes[$colHeader] = "Text"
                $SpecialIndexes.Add($i - 1) # 0-based index
                $SpecialHeaders.Add($colHeader, "Leading Zeros")
                break
            } elseif ($isLongNumber) {
                $ColumnTypes[$colHeader] = "Text"   # must be text to preserve long numbers
                $SpecialIndexes.Add($i - 1)         # 0-based index
                $SpecialHeaders.Add($colHeader, "Long Digit Numbers")
                break
            } elseif ($isShortNumber -or $isDecimal) {
                $ColumnTypes[$colHeader] = "Double"
                break
            } elseif ($isDate) {
                $ColumnTypes[$colHeader] = "DateTime"
                break
            }
        }
    }

    # Write schema.ini
    $schemaPath = Join-Path $Global:Folder "schema.ini"

    $lines = @()
    $lines += "[$Global:FileName]"
    $lines += "Format=CSVDelimited"
    $lines += "ColNameHeader=True"
    $lines += "MaxScanRows=0"
    $lines += "CharacterSet=$(Get-CharacterSet $Script:FilePath)"

    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $colName = $Headers[$i]
        $colType = $ColumnTypes[$colName]

        # ACE format: Col1=Name Type
        $lines += "Col$($i + 1)= $colName $colType"
    }

    # Write file
    Set-Content -Path $schemaPath -Value $lines -Encoding ASCII

    # Delete Sheet12
    $ws2.Delete()

    # Clear formats only from Sheet1 used cells
    $ws1.UsedRange.ClearFormats()

    # Clear only data, leaving headers
    $ws1.UsedRange.Offset(1).Clear()

    return [PSCustomObject]@{
        Excel           = $Excel
        SpecialIndexes  = $SpecialIndexes
        SpecialHeaders  = $SpecialHeaders
        ws1             = $ws1
        Headers         = $Headers
        schemaPath      = $schemaPath
        rowCount        = $rowCount
        colCount        = $colCount
    }
}

function Invoke-CheckHeaders {
    param ($Path)

    # Read headers only
    $reader     = [System.IO.StreamReader]::new($Path, $true)
    $headerLine = $reader.ReadLine()
    $reader.Close()

    # Split headers using VB TextFieldParser (handles quoted fields)
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new([System.IO.StringReader]::new($headerLine))
    $parser.SetDelimiters(",")
    $parser.HasFieldsEnclosedInQuotes = $true
    $Headers = $parser.ReadFields()
    return $Headers
}

function Set-FilePath {
    param (
        $Path,
        $Log,
        $Icon,
        [switch]$Temp
    )
    
    if ("$([System.IO.Path]::GetExtension($Path))" -eq ".zip") {
        $Zip = (Get-FilePath -Path $Path -Log $Log)
        $Path = $Zip
        $Global:Zip = $Zip
    }

    $Script:FilePath    = $Path
    $Global:Folder      = Split-Path $Script:FilePath
    $Global:FileName    = Split-Path $Script:FilePath -Leaf
    $Global:Temp        = $Temp
    $Global:Icon        = $Icon
}

function Show-CsvPicker {
    param ([hashtable]$CSVs, $Log)

    Add-Type -AssemblyName PresentationFramework

    # Build display items
    $items = foreach ($key in ($CSVs.Keys | Sort-Object)) {
        [PSCustomObject]@{
            Number  = $key
            Path    = ($CSVs[$key])
        }
    }

    [xml]$xaml = (Get-Content $Config.ZipWindow -Raw)
    try {
        $reader = New-Object System.Xml.XmlNodeReader([xml]$xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)

        $listView       = $window.FindName("CsvList")
        $okButton       = $window.FindName("OkButton")
        $cancelButton   = $window.FindName("CancelButton")

        $listView.ItemsSource = $items
        $script:selectedValue = $null
        
        $okButton.Add_Click({
            if ($listView.SelectedItem) {
                $script:selectedValue = $listView.SelectedItem.Path
                $window.Close()
            } else {
                [System.Windows.MessageBox]::Show("Please select a CSV.", "Selection Required")
            }
        })

        $cancelButton.Add_Click({
            $window.Close()
        })

        $window.Icon    = $Global:Icon
        $window.Topmost = $true

        $window.ShowDialog() | Out-Null
        return $script:selectedValue
    } catch {
        & $Log "Error: $($_.Exception.Message)" "DarkRed" -Bold:$true
        throw "Error: $($_.Exception.Message)"
    }
}

function Update-FormatAndData {
    param (
        $Excel,
        $SpecialIndexes,
        $Headers,
        $ws1,
        $schemaPath,
        $rowCount,
        $colCount,
        $Log
    )
    
    # Get Excel calculation type
    $xlCalcAutomatic    = [Microsoft.Office.Interop.Excel.XlCalculation]::xlCalculationAutomatic

    # SQL query to get data
    $sql = "SELECT * FROM [$Global:FileName]"

    # Open recordset
    $conn = Open-Connection
    $conn.rs.Open($sql, $conn.conn, 3, 1)

    # Optional: format columns as Text for special indexes
    # Assume $SpecialIndexes contains 0-based column numbers
    foreach ($idx in $SpecialIndexes) {
        $col = $idx + 1  # Excel columns are 1-based
        $ws1.Range($ws1.Cells.Item(1, $col), $ws1.Cells.Item($rowCount, $col)).NumberFormat = "@"
    }

    # # Put back original headers
    for ($i = 0; $i -lt $colCount; $i++) {
        $ws1.Cells.Item(1, $i + 1).Value2 = $Global:Headers[$i]
    }

    # Copy all rows into Sheet1 starting at A2 and autofit
    $ws1.Range("A2").CopyFromRecordset($conn.rs)
    $ws1.UsedRange.EntireColumn.AutoFit()

    # Close recordset and connection
    $conn.rs.Close()
    $conn.conn.Close()

    # Remove schema file
    Remove-Item "$schemaPath" -ErrorAction SilentlyContinue

    # Restore Excel settings
    $Excel.Calculation      = $xlCalcAutomatic
    $Excel.ScreenUpdating   = $true

    # Release COM objects
    if ($null -ne $conn.rs) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($conn.rs) | Out-Null }
    if ($null -ne $conn.conn) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($conn.conn) | Out-Null }

    # Remove the PowerShell variables
    Remove-Variable rs, conn

    # For Excel, do not vall Quit(), just release the reference
    if ($null -ne $Excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null }
    Remove-Variable Excel

    # Remove temp file if it exists
    if ($Global:Temp -and (Test-Path $Script:FilePath)) {
        & $Log "Deleting temporary file: $($Script:FilePath)" "Brown" -Italic:$true
        Remove-Item $Script:FilePath -Force
    }

    # Remove unzipped file if it exists
    if (-not [string]::IsNullOrWhiteSpace($Global:Zip)) {
        if ((Test-Path $Global:Zip)) {
            & $Log "Deleting unzipped file: $($Global:Zip)" "Brown" -Italic:$true
            Remove-Item $Global:Zip -Force
        }
    }

    # Optional: force garbage collection
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
}

function Update-LogSpecialHeaders {
    param (
        $specialHeaders,
        $LogCallback
    )
    
    foreach ($key in $specialHeaders.Keys) {
        $value = $specialHeaders[$key]
        & $LogCallback "$($key): $value" "Black"
    }
}