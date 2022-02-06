
function CRUD-OLEDb
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Action,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CsvPath,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $SQL, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [string]
        $ExtendedProperties
    )
    Begin
    {
        if (!(Test-Path $CsvPath))
        {
            Write-Error -Message "The destination file '$CsvPath' could not be found!" -ErrorAction Stop
        }
        else
        {
            $connString = "Provider = Microsoft.ACE.OLEDB.15.0; Data Source=$CsvPath; Persist Security Info=False; $ExtendedProperties"
            $conn = New-Object System.Data.OleDb.OleDbConnection $connString
            $conn.Open();
        }
    }
    Process
    {
        $cmd = $conn.CreateCommand();
        $cmd.CommandText = $SQL
        if ($Action -iin @('Insert', 'Update', 'Delete'))
        {
            $_ = $cmd.ExecuteNonQuery()
        }
        else 
        {
            $adapter = New-Object System.Data.OleDb.OleDbDataAdapter $cmd;
            $dt = New-Object System.Data.DataSet;
            $adapter.Fill($dt);
            $obj = $dt.Tables[0]
        }
    }
    End
    {
        $conn.Close()
        $conn.Dispose()
        if (!($Action -iin @('Insert', 'Update', 'Delete')))
        {
            return $obj
        }
    }
}

function Get-OLEDbCSV
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CsvPath,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $SQL
    )
    CRUD-OLEDb -Action 'Select' -CsvPath "$(split-path (Resolve-Path $CsvPath))" -SQL $SQL -ExtendedProperties 'Extended Properties="Text;HDR=Yes;IMEX=1;FMT=Delimited(,)"'
}

function Insert-OLEDbCSV
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CsvPath,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $SQL
    )
    CRUD-OLEDb -Action 'Insert' -CsvPath "$(split-path (Resolve-Path $CsvPath))" -SQL $SQL -ExtendedProperties 'Extended Properties="Text;HDR=Yes;IMEX=1;FMT=Delimited(,)"'
}

function Insert-OLEDbXlsx
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $XlsxPath,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $SQL
    )
    CRUD-OLEDb -Action 'Insert' -CsvPath (resolve-path $XlsxPath) -SQL $SQL -ExtendedProperties 'Extended Properties="Excel 8.0;IMEX=1;HDR=YES;"'
}

