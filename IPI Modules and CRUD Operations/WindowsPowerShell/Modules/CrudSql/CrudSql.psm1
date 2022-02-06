
function CRUD-SQL
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
        $Server,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Database,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Table,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=4)]
        [ValidateNotNullOrEmpty()]
        [string]
        $SQL
    )
    Begin
    {
        $connString = "Server=$Server;Database=$Database;Trusted_Connection=yes;Integrated Security=True;";
        $conn = New-Object System.Data.SqlClient.SqlConnection $connString;
        $conn.Open();
        $cmd = $conn.CreateCommand();
        $cmd.CommandText = $SQL
        $dt = New-Object System.Data.DataTable;
    }
    Process
    {
        if ($Action -iin @('Delete', 'Insert', 'Update'))
        {
            $_ = $cmd.ExecuteNonQuery()
        }
        else 
        {
            $rdr = $cmd.ExecuteReader();
            $dt.Load($rdr);
            $obj = $dt
        }
    }
    End
    {
        $conn.Close()
        $conn.Dispose()
        if (!$skip)
        {
            if (!($Action -iin @('Delete', 'Insert', 'Update')))
            {
                return $obj
            }
        }
    }
}