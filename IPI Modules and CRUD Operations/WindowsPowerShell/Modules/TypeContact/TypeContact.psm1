[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

function Check-ContactSQL
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Database,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Table
    )
    Begin
    {
        $connString = "Server=$Server;Database=master;Trusted_Connection=yes;Integrated Security=True;";
        $conn = New-Object System.Data.SqlClient.SqlConnection $connString;
        $conn.Open();
        $cmd = $conn.CreateCommand();
        $chk = ([IO.File]::ReadLines("$(Split-Path((Get-Module TypeContact).Path))\CheckExists.sql"))
        $cmd.CommandText = $chk.replace('PowerShellModulesDb', $Database).replace('psContacts', $Table)
        $dt = New-Object System.Data.DataTable;
        $rdr = $cmd.ExecuteReader();
        $dt.Load($rdr);
        $obj = ($dt | select DatabaseExists, ContactsExist)
        $skip = $false
        if (!($obj.DatabaseExists) -or !($obj.ContactsExist))
        {
            $db = "The database [$Database] $(if(!($obj.DatabaseExists)) {'does not exist'} else {'exists'})"
            $tbl = "The table [$Table] $(if(!($obj.ContactsExist)) {'does not exist'} else {'exists'})"
            $msg = "$db`r$tbl"
            $wrn = "Add $(if(!($obj.DatabaseExists)) {`"database [$Database] and `"})$(if(!($obj.ContactsExist)) {`"table [$Table]`"})"
            if ([System.Windows.Forms.MessageBox]::Show("$msg`r`rDo you want to continue?",'Database and/or Table Do Not Exist', 4, 64) -eq "Yes") 
            {
                Write-host "$wrn Approved!" -ForegroundColor Cyan
            }
            else
            {
                Write-Warning -Message $msg -WarningAction Continue
                if("Stop option." -iin ($err[0].Message).Split(' ',10)) 
                {
                    Write-host "$($wrn): $($err[0].Message)" -ForegroundColor Cyan
                }
                $skip = $true
            }
        }
    }
    Process
    {
        if (!$skip) 
        {
            if (!($obj.DatabaseExists)) 
            {
                $NewDb = "CREATE DATABASE [$Database]"
                $cmd.CommandText = $NewDb
                $_ = $cmd.ExecuteNonQuery()
                $conn.Close()

                $connString = "Server=$Server;Database=$Database;Trusted_Connection=yes;Integrated Security=True;";
                $conn = New-Object System.Data.SqlClient.SqlConnection $connString;
                $conn.Open();
                $cmd = $conn.CreateCommand();
                $NewSp = ([IO.File]::ReadLines("$(Split-Path((Get-Module TypeContact).Path))\spEntityCache_Fix.sql"))
                $NewSp = $NewSp.replace('PowerShellModulesDb', $Database)
                $cmd.CommandText = $NewSp
                $_ = $cmd.ExecuteNonQuery()
                $ExecSp = "exec [$Database].dbo.[spEntityCache_Fix]"
                $cmd.CommandText = $ExecSp
                $_ = $cmd.ExecuteNonQuery()
            }
            if (!($obj.ContactsExist)) 
            {
                $NewTbl = ([IO.File]::ReadLines("$(Split-Path((Get-Module TypeContact).Path))\CreateTable.sql"))
                $NewTbl = $NewTbl.replace('PowerShellModulesDb', $Database).replace('psContacts', $Table)
                $cmd.CommandText = $NewTbl
                $_ = $cmd.ExecuteNonQuery()
            }
        }
    }
    End
    {
        if (!$skip)
        {
            return $true
        } 
        else
        {
            return $false
        }
    }
}

function Convert-ContactCSVtoHTML
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
        $HtmlPath, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [string]
        $Head, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [string]
        $Title
    )
    Begin
    {
        if (!(Test-Path $CsvPath) -and !(Test-Path $HtmlPath)) 
        {
            Write-Error -Message "The source file '$CsvPath' and destination file '$HtmlPath' could not be found!" -ErrorAction Stop
        } 
        elseif (!(Test-Path $CsvPath))
        {
            Write-Error -Message "The source file '$CsvPath' could not be found!" -ErrorAction Stop
        } 
        elseif (!(Test-Path $HtmlPath))
        {
            Write-Error -Message "The destination file '$HtmlPath' could not be found!" -ErrorAction Stop
        }
    }
    Process
    {
        $html = (Import-Csv $CsvPath | ConvertTo-Html -Head "$Head" -Title $Title)
        $html | Out-File $HtmlPath -Encoding utf8
    }
    End
    {
        return $html
    }
}

function Convert-ContactSQLtoHTML
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Database,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Table,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [ValidateNotNullOrEmpty()]
        [string]
        $DestPath, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=4)]
        [string]
        $Head, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=5)]
        [string]
        $Title
    )
    Begin
    { 
        if (!(Test-Path $DestPath))
        {
            Write-Error -Message "The destination file '$DestPath' could not be found!" -ErrorAction Stop
        }
        $obj = Get-ContactSQL -Server $Server -Database $Database -Table $Table
    }
    Process
    {
        if (!$skip)
        {
            $html = ($obj | ConvertTo-Html -Head "$Head" -Title $Title)
            $html | Out-File $DestPath -Encoding utf8
        }
    }
    End
    {
        if (!$skip)
        {
            return $html
        }
    }
}

function Delete-ContactCSV
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Id, 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CsvPath
    )
    
    Begin
    {
        if (Test-Path $CsvPath)
        {
            $Dest = ([IO.File]::ReadLines("$(resolve-path $CsvPath)") | ConvertFrom-Csv)
            if (($Id | % {($_ -split '\| ')} | where {$_ -notin $Dest.Id})) 
            {
                $msg = "Contacts with Id: '$(($Id | % {($_ -split '\| ')} | where {$_ -notin $Dest.Id}) -join ', ')' do not exist!"
                $skip = $false
                if ([System.Windows.Forms.MessageBox]::Show("$msg`r`rDo you want to continue?",'Contact Does Not Exist', 4, 64) -eq "Yes")                
                {
                    Write-host "Ignored Contacts with Id: '$(($Id | % {($_ -split '\| ')} | where {$_ -notin $Dest.Id}) -join ', ')'!" -ForegroundColor Cyan
                }
                else 
                {
                    Write-Warning -Message $msg -WarningAction Continue
                    if("Stop option." -iin ($err[0].Message).Split(' ',10)) 
                    {
                        Write-host "Delete Contacts: $($err[0].Message)" -ForegroundColor Cyan
                    }
                    $skip = $true
                }
            }
        } 
        elseif (!(Test-Path $CsvPath))
        {
            Write-Error -Message "The source file '$CsvPath' could not be found!" -ErrorAction Stop
        }
    }
    Process
    {
        $ids = (($Id | % {($_ -split '\| ')} | where {$_ -in $Dest.Id}) -join ', ')
        $msg = "Contacts with Id: '$ids' will be deleted!"
        if ([System.Windows.Forms.MessageBox]::Show("$msg`r`rDo you want to continue?",'Delete Contact Confirmation', 4, 64) -eq "Yes") 
        {
            Write-host "Delete Contacts with Id: '$ids' Approved!" -ForegroundColor Cyan
        }
        else
        {
            Write-Warning -Message $msg -WarningAction Continue
            if("Stop option." -iin ($err[0].Message).Split(' ',10)) 
            {
                Write-host "Delete Contacts: $($err[0].Message)" -ForegroundColor Cyan
            }
            $skip = $true
        }
    }
    End
    {
        if(!$skip) 
        {
            $Id | % {($_ -split '\| ')} | where {$_ -in $Dest.Id} | % {Update-ContactCSV -Id $_ -CsvPath $CsvPath -Object @() -Delete $true}
        }
    }
}

function Delete-ContactSQL
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Id, 
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
        $Table
    )
    
    Begin
    {
        $SQL = "SELECT * FROM [$Database].dbo.[$Table]"
        $dt = (CRUD-SQL -Action 'Select' -Server $Server -Database $Database -Table $Table -SQL $SQL)
        $Dest = ($dt | select Id, Contact_Title, First_Name, Last_Name, Phones, Emails -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray)
        if (($Id | % {($_ -split '\| ')} | where {$_ -notin $Dest.Id})) 
        {
            $msg = "Contacts with Id: '$(($Id | % {($_ -split '\| ')} | where {$_ -notin $Dest.Id}) -join ', ')' do not exist!"
            $skip = $false
            if ([System.Windows.Forms.MessageBox]::Show("$msg`r`rDo you want to continue?",'Contact Does Not Exist', 4, 64) -eq "Yes") 
            {
                Write-host "Ignored Contacts with Id: '$(($Id | % {($_ -split '\| ')} | where {$_ -notin $Dest.Id}) -join ', ')'!" -ForegroundColor Cyan
            }
            else 
            {
                Write-Warning -Message $msg -WarningAction Continue
                if("Stop option." -iin ($err[0].Message).Split(' ',10)) 
                {
                    Write-host "Delete Contacts: $($err[0].Message)" -ForegroundColor Cyan
                }
                $skip = $true
            }
        }
        if(!$skip)
        {
            $ids = (($Id | % {($_ -split '\| ')} | where {$_ -in $Dest.Id}) -join ', ')
            $msg = "Contacts with Id: '$ids' will be deleted!"
            if ([System.Windows.Forms.MessageBox]::Show("$msg`r`rDo you want to continue?",'Delete Contact Confirmation', 4, 64) -eq "Yes") 
            {
                Write-host "Delete Contacts with Id: '$ids' Approved!" -ForegroundColor Cyan
            }
            else
            {
                Write-Warning -Message $msg -WarningAction Continue
                if("Stop option." -iin ($err[0].Message).Split(' ',10)) 
                {
                    Write-host "Delete Contacts: $($err[0].Message)" -ForegroundColor Cyan
                }
                $skip = $true
            }
        }
    }
    Process
    {
        if(!$skip) 
        {
            $SQL = "DELETE FROM [$Database].dbo.[$Table] where [Id] in ($(($Id | % {($_ -split '\| ')} | where {$_ -in $Dest.Id}) -join ', '))"
        }
    }
    End
    {
        if(!$skip) 
        {
            CRUD-SQL -Action 'Delete' -Server $Server -Database $Database -Table $Table -SQL $SQL
        }
    }
}

function Get-ContactCSV
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $CsvPath
    )
    Begin
    { 
        if (!(Test-Path $CsvPath) -and !(Test-Path $HtmlPath)) 
        {
            Write-Error -Message "The source file '$CsvPath' and destination file '$HtmlPath' could not be found!" -ErrorAction Stop
        } 
        elseif (!(Test-Path $CsvPath))
        {
            Write-Error -Message "The source file '$CsvPath' could not be found!" -ErrorAction Stop
        } 
        elseif (!(Test-Path $HtmlPath))
        {
            Write-Error -Message "The destination file '$HtmlPath' could not be found!" -ErrorAction Stop
        }
    }
    Process
    {
        $obj = (Import-Csv $CsvPath | select Id, Contact_Title, First_Name, Last_Name, Phones, Emails -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray)
    }
    End
    {
        return $obj
    }
}

function Get-ContactOLECSV
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
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [psobject]
        $Filter
    )    
    Begin
    {        
        $skip = $false
        if (!(Test-Path $CsvPath))
        {
            try
            {
                Write-Error -Message "The source file '$CsvPath' could not be found!" -ErrorAction Stop
            }
            catch
            {
                $skip = $true
            }
        } 
    }
    Process
    {
        $cols = @(@("\[Id]", "[ï»¿]"), @('Contact_Title', 'F2'), @('First_Name', 'F3'), @('Last_Name',  'F4'), @('Phones', 'F5'), @('Emails', 'F6'))
        $where = Get-FilterSQL $Filter
        foreach ($itm in $cols) 
        {
            $where = $where -replace $itm[0], $itm[1]
        }
        $sql = "SELECT * FROM [$(Split-Path $CsvPath -Leaf)]$where"
        $dt = Get-OLEDbCSV -CsvPath (Resolve-Path $CsvPath) -SQL $sql
        $obj = ($dt | 
            select @{L='Id';E=({$_.'ï»¿'})}, @{L='Contact_Title';E=({$_.F2})}, @{L='First_Name';E=({$_.F3})}, @{L='Last_Name';E=({$_.F4})}, @{L='Phones';E=({$_.F5})}, 
                @{L='Emails';E=({$_.F6})} -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray | where {$_.Id -gt ''})
    }
    End
    {
        return $obj
    }
}

function Get-ContactSQL
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Database,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Table,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [psobject]
        $Filter
    )
    Begin
    {
        $skip = $false
        if (!(Check-ContactSQL -Server $Server -Database $Database -Table $Table ))
        {
            $skip = $true
        }
    }
    Process
    {
        if (!$skip)
        {
            $sql = "SELECT * FROM [$Database].dbo.[$Table]"
            $where = Get-FilterSQL $Filter
            $dt = (CRUD-SQL -Action 'Select' -Server $Server -Database $Database -Table $Table -SQL "$SQL$where")
            $obj = ($dt | select Id, Contact_Title, First_Name, Last_Name, Phones, Emails -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray) 
        }
    }
    End
    {
        if (!$skip)
        {
            return $obj
        }
    }
}

function Insert-ContactCSV
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
        [psobject]
        $Object
    )
    
    Begin
    {
        if (Test-Path $CsvPath)
        {
            $Dest = ([IO.File]::ReadLines("$(resolve-path $CsvPath)") | ConvertFrom-Csv)
            $Id = [int64](($Dest | Measure-Object -Property Id -Maximum).Maximum + 1)
            if ($Object.First_Name + ' ' + $Object.Last_Name -in ($Dest | % {$_.First_Name + ' ' + $_.Last_Name})) 
            {
                $msg = "A contact with First Name '$($Object.First_Name)' and Last Name '$($Object.Last_Name)' already exists!"
                $skip = $false
                if ([System.Windows.Forms.MessageBox]::Show("$msg`r`rDo you want to continue and add the record?",'Contact Already Exists', 4, 64) -eq "Yes") 
                {
                    Write-host "Add Contact Approved: For 'Id: $Id $($Object.First_Name) $($Object.Last_Name)'" -ForegroundColor Cyan
                }
                else 
                {
                    Write-Warning -Message $msg -WarningAction Continue
                    if("Stop option." -iin ($err[0].Message).Split(' ',10)) 
                    {
                        Write-host "Add Contact: $($err[0].Message)" -ForegroundColor Cyan
                    }
                    $skip = $true
                }
            }
        }
    }
    Process
    {
        if(!$skip) 
        {
            if (Test-Path $CsvPath)
            {
                $obj = ($Object | select @{L='Id';E=({$Id})}, Contact_Title, First_Name, Last_Name, @{L='Phones';E=({$_.Phones -join ', '})}, @{L='Emails';E=({$_.Emails -join ', '})} | ConvertTo-Csv)[2]
            } else
            {
                $Id = [int64]1
                $obj = ($Object | select @{L='Id';E=({$Id})}, Contact_Title, First_Name, Last_Name, @{L='Phones';E=({$_.Phones -join ', '})}, @{L='Emails';E=({$_.Emails -join ', '})} | ConvertTo-Csv -NoTypeInformation)
            }
        }
    }
    End
    {
        if(!$skip) 
        {
            $obj | Out-File $CsvPath -Encoding utf8 -Append
        }
    }
}

function Insert-ContactOLEDBtoCSV
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
        [psobject]
        $Object
    )    
    Begin
    {        
        if (Test-Path $CsvPath)
        {
            $Dest = Get-ContactOLECSV -CsvPath $CsvPath
            $Id = [int64](($Dest | Measure-Object -Property Id -Maximum).Maximum + 1)
            if ($Object.First_Name + ' ' + $Object.Last_Name -in ($Dest | % {$_.First_Name + ' ' + $_.Last_Name})) 
            {
                $msg = "A contact with First Name '$($Object.First_Name)' and Last Name '$($Object.Last_Name)' already exists!"
                $skip = $false
                if ([System.Windows.Forms.MessageBox]::Show("$msg`r`rDo you want to continue and add the record?",'Contact Already Exists', 4, 64) -eq "Yes") 
                {
                    Write-host "Add Contact Approved: For 'Id: $Id $($Object.First_Name) $($Object.Last_Name)'" -ForegroundColor Cyan
                }
                else 
                {
                    Write-Warning -Message $msg -WarningAction Continue
                    if("Stop option." -iin ($err[0].Message).Split(' ',10)) 
                    {
                        Write-host "Add Contact: $($err[0].Message)" -ForegroundColor Cyan
                    }
                    $skip = $true
                }
            }
        }
    }
    Process
    {
        if(!$skip) 
        {
            
            if (Test-Path $CsvPath)
            {
                $obj = ($Object | select @{L='Id';E=({$Id})}, Contact_Title, First_Name, Last_Name, @{L='Phones';E=({$_.Phones -join ', '})}, @{L='Emails';E=({$_.Emails -join ', '})})
            } else
            {
                $Id = [int64]1
                $obj = ($Object | select @{L='Id';E=({$Id})}, Contact_Title, First_Name, Last_Name, @{L='Phones';E=({$_.Phones -join ', '})}, @{L='Emails';E=({$_.Emails -join ', '})})
            }
        }
    }
    End
    {
        if(!$skip) 
        {
            $ins = @("INSERT INTO [$(Split-Path $CsvPath -Leaf)] ([ï»¿], [F2], [F3], [F4], [F5], [F6])
                VALUES ('$($obj.Id)', '$($obj.Contact_Title)', '$($obj.First_Name)', '$($obj.Last_Name)', '$($obj.Phones)', '$($obj.Emails)')
            ")
            Insert-OLEDbCSV -CsvPath $CsvPath -SQL "$ins"
        }
    }
}

function Insert-ContactSQL
{
    [CmdletBinding()]
    Param (  
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Database,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Table, 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $Object
    )    
    Begin
    { 
        $sql = (Send-ContactSQL -Server $Server -Database $Database -Table $Table -Object $Object)
        $skip = $false
        if (!$sql)
        {
            $msg = "Unable to update the server '$Server' database '$Database' table '$Table' with this contact!"
            try {
                Write-Error -Message $msg -ErrorAction Stop -ErrorVariable err
            }
            catch
            {
                Write-Warning -Message $msg -WarningAction Continue
                if("Stop option." -iin ($err[0].Message).Split(' ',10)) 
                {
                    Write-host "Add Contact: $($err[0].Message)" -ForegroundColor Cyan
                }
                $skip = $true
            }
        }
    }
    Process
    {
        if (!$skip)
        {
            CRUD-SQL -Action 'Insert' -Server $Server -Database $Database -Table $Table -SQL $sql
        }
    }
    End
    {
        return $Object
    }
}

function Send-ContactSQL
{
    [CmdletBinding()]
    Param (  
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Database,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Table, 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $Object
    )    
    Begin
    {   
        if (!(Check-ContactSQL -Server $Server -Database $Database -Table $Table ))
        {
            $skip = $true
        }
        else
        {
            $SQL = "SELECT * FROM [$Database].dbo.[$Table]"
            $dt = (CRUD-SQL -Action 'Select' -Server $Server -Database $Database -Table $Table -SQL $SQL)
            $Dest = ($dt | select Id, Contact_Title, First_Name, Last_Name, Phones, Emails -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray)
            $Id = [int64](($Dest | Measure-Object -Property Id -Maximum).Maximum + 1)
            if ($Object.First_Name + ' ' + $Object.Last_Name -in ($Dest | % {$_.First_Name + ' ' + $_.Last_Name})) 
            {
                $msg = "A contact with First Name '$($Object.First_Name)' and Last Name '$($Object.Last_Name)' already exists!"
                $skip = $false
                if ([System.Windows.Forms.MessageBox]::Show("$msg`r`rDo you want to continue and add the record?",'Duplicate Contact Exists', 4, 64) -eq "Yes")                
                {
                    Write-host "Add Contact Approved: For 'Id: $Id $($Object.First_Name) $($Object.Last_Name)'" -ForegroundColor Cyan
                }
                else 
                {
                    Write-Warning -Message $msg -WarningAction Continue
                    if("Stop option." -iin ($err[0].Message).Split(' ',10)) 
                    {
                        Write-host "Add Contact: $($err[0].Message)" -ForegroundColor Cyan
                    }
                    $skip = $true
                }
            }
        }
    }
    Process
    {
        if(!$skip) 
        {
            
            $obj = ($Object | select Contact_Title, First_Name, Last_Name, @{L='Phones';E=({$_.Phones -join ', '})}, @{L='Emails';E=({$_.Emails -join ', '})})
        }
    }
    End
    {
        if(!$skip) 
        {
            return @("INSERT INTO [$Database].dbo.[$Table] ([Contact_Title], [First_Name], [Last_Name], [Phones], [Emails])
                VALUES ( '$($obj.Contact_Title)', '$($obj.First_Name)', '$($obj.Last_Name)', '$($obj.Phones)', '$($obj.Emails)')
            ")
        }
    }
}

function Send-ContactsEmail
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Body
    )
    Begin
    { 
        $html = "$Body"
    }
    Process
    {
        Send-ToEmailOpen $html
    }
    End
    {
        return $false
    }
}

function Send-ContactsEmailNew
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Body,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [string]
        $Head
    )
    Begin
    { 
        $html = "$Head`r`n$Body"
    }
    Process
    {
        Send-ToEmailNew $html
    }
    End
    {
        return $false
    }
}

function Send-ContactsPowerPointNew
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Body,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Head
    )
    Begin
    { 
        $cols = Get-ColoursFromCSS $Head
        $tbl = Convert-HTMLtoDataTable $Body
    }
    Process
    {
        Send-ToPowerPointNew -GridResult $tbl -Colours $cols | Out-Null
    }
    End
    {
        return $false
    }
}

function Send-ContactsPowerPointOpen
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Body,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Head
    )
    Begin
    { 
        $cols = Get-ColoursFromCSS $Head
        $tbl = Convert-HTMLtoDataTable $Body
    }
    Process
    {
        Send-ToPowerPointOpen -GridResult $tbl -Colours $cols | Out-Null
    }
    End
    {
        return $false
    }
}

function Send-ContactsWordNew
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Body,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Head
    )
    Begin
    { 
        $cols = Get-ColoursFromCSS $Head
        $tbl = Convert-HTMLtoDataTable $Body
    }
    Process
    {
        Send-ToWordNew -GridResult $tbl -Colours $cols | Out-Null
    }
    End
    {
        return $false
    }
}

function Send-ContactsWordOpen
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Body,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Head
    )
    Begin
    { 
        $cols = Get-ColoursFromCSS $Head
        $tbl = Convert-HTMLtoDataTable $Body
    }
    Process
    {
        Send-ToWordOpen -GridResult $tbl -Colours $cols | Out-Null
    }
    End
    {
        return $false
    }
}

function Set-Contact 
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $First_Name,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)] 
        [ValidateNotNullOrEmpty()]
        [string]
        $Last_Name,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [string[]]
        $Phones, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [string[]]
        $Emails, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=4)]
        [string]
        $Contact_Title
        )
    Begin
    {
        class Record
        {
            $Contact_Title
            $First_Name
            $Last_Name
            $Phones
            $Emails
        }
    }
    Process
    {
        $Record = [Record]@{
            Contact_Title = $Contact_Title
            First_Name = $First_Name
            Last_Name = $Last_Name
            Phones = $Phones
            Emails = $Emails
        }
    }
    End
    {
        $obj = New-Object -TypeName psobject $Record
        Write-Output $obj
    }
}

function Show-ContactCSV
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
        $HtmlPath,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [string]
        $Head, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [string]
        $Title, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=4)]
        [bool]
        $HtmlOnly = $false,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=5)]
        [psobject]
        $Filter
    )
    Begin
    { 
        if (!$Filter)
        {
            $obj = Get-ContactCSV -CsvPath $CsvPath
        }
        else
        {
            $obj = Get-ContactOLECSV $CsvPath -Filter $Filter
        }
    }
    Process
    {
        $body = (Format-Html $obj)
        $arr = @('../css/', '../js/')
        foreach ($itm in $arr)
        {
            $Head = ($Head -replace $itm, (Resolve-Path $itm))
        }
        $rDefs = '
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        '
        $cDefs = '
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />    
            <ColumnDefinition Width="Auto" />    
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
        ';
        $fields = @(
            @('Id', 60),
            @('Contact_Title', 80),
            @('First_Name', 120),
            @('Last_Name', 200),
            @('Phones', 200),
            @('Emails', 300)
        )
        foreach ($lbl in $($fields|select @{L='Fields'; E=({$_; $fields.IndexOf($_); 0; $_ -replace '_', ' '})}))
        {
            $xBody += "
                <Label 
                    Name = `"lbl$($lbl.Fields[0])`" 
                    Width = `"$($lbl.Fields[1])`" 
                    Grid.Column = `"$($lbl.Fields[2])`" 
                    Grid.Row = `"$($lbl.Fields[3])`" 
                    Content = `"$($lbl.Fields[4])`"
                /> 
            "
        }
        foreach ($tbx in $($fields|select @{L='Fields'; E=({$_; $fields.IndexOf($_); 1; 1})}))
        {
            $xBody += "
                <TextBox 
                    Name = `"$($tbx.Fields[0])`" 
                    Width = `"$($tbx.Fields[1])`" 
                    Grid.Column = `"$($tbx.Fields[2])`" 
                    Grid.Row = `"$($tbx.Fields[3])`" 
                    Margin = `"$($tbx.Fields[4])`"
                /> 
            "
        }
        $xBody += '
            <ComboBox 
                Name = "cboUpdate" 
                Grid.Column = "6" 
                Grid.Row = "1" 
                Margin = "1"
            />
        '
        $bColSpan = 'Grid.ColumnSpan="7"';
        $bCol = '';
        $bRowSpan = '';
        $bRow = 'Grid.Row="1"';
        $fWidth = 280;
        $fHeight = 280;
        $fTopMost = $false;
        $bHeight = ($fHeight * 2.5);
        $bWidth = ($fWidth * 3.9);
        $objects = @(@('Empty', ''));
        $dSource = @{}
        'CsvPath', 'HtmlPath' | select @{E=({$dSource.Add(($_), (Invoke-Expression "`$$_"))})} > $null
        $jQuery = @('Clear Higlighting', 'Clear JQuery Filter', 'Filter Using JQuery')
        $sources = @(
            @('cboUpdate',
                @('Select a Function', 'Add Contact', 'Clear Fields', 'Clear Higlighting', 'Clear JQuery Filter', 'Delete Contact', 'Get All Contacts', 
                    'Filter Using JQuery', 'Filter Using OLEDB', 'Send to Email New', 'Send to Email Open', 'Send to PowerPoint New', 'Send to PowerPoint Open', 
                    'Send to Word New', 'Send to Word Open', 'Update Contact'), 
                @('Id', 'Contact_Title', 'First_Name', 'Last_Name', 'Phones', 'Emails'),
                @(
                    @(''), 
                    @('Set-Contact', 'Insert-ContactCSV', 'Show-ContactCSV'), 
                    @(''),
                    @('clearHighlight'),
                    @('clearFilter'),
                    @('Delete-ContactCSV', 'Show-ContactCSV'), 
                    @('Show-ContactCSV'),  
                    @('filterRows'),
                    @('Show-ContactCSV'),
                    @('Send-ContactsEmailNew'), 
                    @('Send-ContactsEmail'),  
                    @('Send-ContactsPowerPointNew'), 
                    @('Send-ContactsPowerPointOpen'),  
                    @('Send-ContactsWordNew'), 
                    @('Send-ContactsWordOpen'),
                    @('Set-Contact', 'Update-ContactCSV', 'Show-ContactCSV' )
                ) 
                @()
            ), 
            @()
        )
    }
    End
    {
        if (!$HtmlOnly) 
        {
            $prm = @{}
            'head', 'title', 'body', 'cDefs', 'xBody', 'bColSpan', 'objects', 'fWidth', 'fHeight', 'fTopMost', 'sources', 'dSource', 'jQuery' | select @{E=({$prm.Add(($_), (Invoke-Expression "`$$_"))})} > $null
            Show-WPFWebViewForm -object (Set-WebForm @prm)
        }
        else
        {
            return ((
                ConvertTo-Html -Body $body -Head $Head -Title $Title) -replace '<html xmlns="http://www.w3.org/1999/xhtml">', 
                    '<html xmlns="http://www.w3.org/1999/xhtml"><meta http-equiv="x-ua-compatible" content="IE=11">'
            )
        }
    }
}

function Show-ContactSQL
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Server,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Database,
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        [ValidateNotNullOrEmpty()]
        [string]
        $Table,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [string]
        $Head, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=4)]
        [string]
        $Title, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=5)]
        [bool]
        $HtmlOnly = $false,
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=6)]
        [psobject]
        $Filter
    )
    Begin
    { 
        $obj = Get-ContactSQL -Server $Server -Database $Database -Table $Table -Filter $Filter
    }
    Process
    {
        $body = (Format-Html $obj)
        $arr = @('../css/', '../js/')
        foreach ($itm in $arr)
        {
            $Head = ($Head -replace $itm, (Resolve-Path $itm))
        }
        $rDefs = '
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        '
        $cDefs = '
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />    
            <ColumnDefinition Width="Auto" />    
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />
        ';
        $fields = @(
            @('Id', 60),
            @('Contact_Title', 80),
            @('First_Name', 120),
            @('Last_Name', 200),
            @('Phones', 200),
            @('Emails', 300)
        )
        foreach ($lbl in $($fields|select @{L='Fields'; E=({$_; $fields.IndexOf($_); 0; $_ -replace '_', ' '})}))
        {
            $xBody += "
                <Label 
                    Name = `"lbl$($lbl.Fields[0])`" 
                    Width = `"$($lbl.Fields[1])`" 
                    Grid.Column = `"$($lbl.Fields[2])`" 
                    Grid.Row = `"$($lbl.Fields[3])`" 
                    Content = `"$($lbl.Fields[4])`"
                /> 
            "
        }
        foreach ($tbx in $($fields|select @{L='Fields'; E=({$_; $fields.IndexOf($_); 1; 1})}))
        {
            $xBody += "
                <TextBox 
                    Name = `"$($tbx.Fields[0])`" 
                    Width = `"$($tbx.Fields[1])`" 
                    Grid.Column = `"$($tbx.Fields[2])`" 
                    Grid.Row = `"$($tbx.Fields[3])`" 
                    Margin = `"$($tbx.Fields[4])`"
                /> 
            "
        }
        $xBody += '
            <ComboBox 
                Name = "cboUpdate" 
                Grid.Column = "6" 
                Grid.Row = "1" 
                Margin = "1"
            />
        '
        $bColSpan = 'Grid.ColumnSpan="7"';
        $bCol = '';
        $bRowSpan = '';
        $bRow = 'Grid.Row="1"';
        $fWidth = 280;
        $fHeight = 280;
        $fTopMost = $false;
        $bHeight = ($fHeight * 2.5);
        $bWidth = ($fWidth * 3.9);
        $objects = @(@('Empty', ''));
        $dSource = @{}
        'Server', 'Database', 'Table' | select @{E=({$dSource.Add(($_), (Invoke-Expression "`$$_"))})} > $null
        $jQuery = @('Clear Higlighting', 'Clear JQuery Filter', 'Filter Using JQuery')
        $sources = @(
            @('cboUpdate',
                @('Select a Function', 'Add Contact', 'Clear Fields', 'Clear Higlighting', 'Clear JQuery Filter', 'Delete Contact', 'Get All Contacts', 
                    'Filter Using JQuery', 'Filter Using SQL', 'Send to Email New', 'Send to Email Open', 'Send to PowerPoint New', 'Send to PowerPoint Open', 
                    'Send to Word New', 'Send to Word Open', 'Update Contact'), 
                @('Id', 'Contact_Title', 'First_Name', 'Last_Name', 'Phones', 'Emails'),
                @(
                    @(''), 
                    @('Set-Contact', 'Insert-ContactSQL', 'Show-ContactSQL'), 
                    @(''),
                    @('clearHighlight'),
                    @('clearFilter'),
                    @('Delete-ContactSQL', 'Show-ContactSQL'), 
                    @('Show-ContactSQL'), 
                    @('filterRows'),
                    @('Show-ContactSQL'),
                    @('Send-ContactsEmailNew'),
                    @('Send-ContactsEmail'),
                    @('Send-ContactsPowerPointNew'),
                    @('Send-ContactsPowerPointOpen'),
                    @('Send-ContactsWordNew'), 
                    @('Send-ContactsWordOpen'),
                    @('Set-Contact', 'Update-ContactSQL', 'Show-ContactSQL' )
                ) 
                @()
            ), 
            @()
        )
    }
    End
    {
        if (!$HtmlOnly) 
        {
            $prm = @{}
            'head', 'title', 'body', 'cDefs', 'xBody', 'bColSpan', 'objects', 'fWidth', 'fHeight', 'fTopMost', 'sources', 'dSource', 'jQuery' | select @{E=({$prm.Add(($_), (Invoke-Expression "`$$_"))})} > $null
            Show-WPFWebViewForm -object (Set-WebForm @prm)
        }
        else
        {
            return ((
                ConvertTo-Html -Body $body -Head $Head -Title $Title) -replace '<html xmlns="http://www.w3.org/1999/xhtml">', 
                    '<html xmlns="http://www.w3.org/1999/xhtml"><meta http-equiv="x-ua-compatible" content="IE=11">'
            )
        }
    }
}

function Update-ContactCSV
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $Id, 
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
        [psobject]
        $Object, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        [bool]
        $Delete
    )
    
    Begin
    {
        if (Test-Path $CsvPath)
        {
            $Dest = ([IO.File]::ReadLines("$(resolve-path $CsvPath)") | ConvertFrom-Csv)
            if (($id | where {$_ -notin $Dest.Id})) 
            {
                $msg = "Contacts with Id: '$($id | where {$_ -notin $Dest.Id})' do not exist!"
                $skip = $false
                try
                {
                    Write-Warning -Message $msg -WarningAction Stop -WarningVariable wrn
                }
                catch 
                {
                    $skip = $true
                }
            }
        } 
        elseif (!(Test-Path $CsvPath))
        {
            Write-Error -Message "The source file '$CsvPath' could not be found!" -ErrorAction Stop
        }
    }
    Process
    {
        if(!$skip) 
        {
            $obj = @()
            $obj += ($Dest | Where-Object {$_.Id -notin $Id})
            if (!$Delete) 
            {
                $obj += ($Object | select @{L='Id';E=({$Id})}, Contact_Title, First_Name, Last_Name, @{L='Phones';E=({$_.Phones -join ', '})}, @{L='Emails';E=({$_.Emails -join ', '})})
            }
            $obj = ($obj | Sort-Object -Property Id)
        }
    }
    End
    {
        if(!$skip) 
        {
            $obj | ConvertTo-Csv -NoTypeInformation | Out-File $CsvPath -Encoding utf8 -Force
        }
        return $true
    }
}

function Update-ContactSQL
{
    [CmdletBinding()]
    Param ( 
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [int64]
        $Id, 
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
        [psobject]
        $Object
    )
    
    Begin
    {
        $SQL = "SELECT * FROM [$Database].dbo.[$Table]"
        $dt = (CRUD-SQL -Action 'Select' -Server $Server -Database $Database -Table $Table -SQL $SQL)
        $Dest = ($dt | select Id, Contact_Title, First_Name, Last_Name, Phones, Emails -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray)
        if (!($Id -in $Dest.Id)) 
        {
            $msg = "A contact with Id: '$Id' does not exist!"
            $skip = $false
            try
            {
                Write-Warning -Message $msg -WarningAction Stop -WarningVariable wrn
            }
            catch 
            {
                $skip = $true
            }
        }
    }
    Process
    {
        if(!$skip) 
        {
            $obj = ($Object | select @{L='Id';E=({$Id})}, Contact_Title, First_Name, Last_Name, @{L='Phones';E=({$_.Phones -join ', '})}, @{L='Emails';E=({$_.Emails -join ', '})})
            $SQL = "UPDATE [$Database].dbo.[$Table] 
                SET [Contact_Title] = '$($obj.Contact_Title)', 
                [First_Name] = '$($obj.First_Name)', 
                [Last_Name] = '$($obj.Last_Name)', 
                [Phones] = '$($obj.Phones)', 
                [Emails] = '$($obj.Emails)' 
                where [Id] = $Id"
        }
    }
    End
    {
        if(!$skip) 
        {
            CRUD-SQL -Action 'Update' -Server $Server -Database $Database -Table $Table -SQL $SQL
        }
    }
}