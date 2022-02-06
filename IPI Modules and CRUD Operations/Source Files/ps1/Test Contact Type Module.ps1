param ($First_Name, $Last_Name, $Phones, $Emails, $Contact_Title, $ShowSql, $ShowCsv)

$csv = '..\csv\psContacts.csv'
$html = '..\html\psContacts.htm'
$style = '<link media="screen" rel="stylesheet" href="../css/psContacts.css" /><link rel="stylesheet" href="../css/bootstrap.min.css" />'
$script = '<script src="../js/psContacts.js"></script>'
$jQuery = '<script src="../js/jquery-3.5.1.bom.js"></script>'
$title = 'Contacts'
$Server = '(LocalDB)\MSSQLLocalDB'
$Database = 'PowerShellModulesDb'
$Table = 'psContacts'

function Test-WebForm
{
    $object = Set-WebForm
    Show-WPFWebViewForm -object $object
}

function Test-GetContactsCSV
{
    $tbl = (Convert-ContactCSVtoHTML -CsvPath $csv -HtmlPath $html -Head "$style`r$jQuery`r$script" -Title $title)
    Invoke-item $html
}

function Test-InsertCSV
{
    $test = Set-Contact -First_Name $First_Name -Last_Name $Last_Name -Phones $Phones -Emails $Emails -Contact_Title $Contact_Title
    if ($test)
    {
        Insert-ContactCSV -CsvPath $csv -Object $test
        Test-GetContactsCSV
    }
}

function Test-InsertOLEDbCSV
{
    $test = Set-Contact -First_Name $First_Name -Last_Name $Last_Name -Phones $Phones -Emails $Emails -Contact_Title $Contact_Title
    if ($test)
    {
        Insert-ContactOLEDBtoCSV -CsvPath $csv -Object $test
        Test-GetContactsCSV
    }
}

function Test-UpdateCSV
{
    Param($Id = 1)
    $test = Set-Contact -First_Name $First_Name -Last_Name $Last_Name -Phones $Phones -Emails $Emails -Contact_Title $Contact_Title
    if ($test)
    {
        Update-ContactCSV -Id $Id -CsvPath $csv -Object $test
        Test-GetContactsCSV
    }
}

function Test-DeleteCSV
{
    Param($Id = 2)
    Delete-ContactCSV -Id $Id -CsvPath $csv
    Test-GetContactsCSV
}

function Test-ShowCSV
{
    Show-ContactCSV -CsvPath $csv -HtmlPath $html -Head "$style`r$jQuery`r$script" -Title $title
}

function Test-GetContactsSQL
{
    $tbl = (Convert-ContactSQLtoHTML -Server $Server -Database $Database -Table $Table -HtmlPath $html -Head "$style`r$jQuery`r$script" -Title $title)
    if ($tbl)
    {
        Invoke-item $html
    }
}

function Test-InsertSQL
{
    $test = Set-Contact -First_Name $First_Name -Last_Name $Last_Name -Phones $Phones -Emails $Emails -Contact_Title $Contact_Title
    if ($test)
    {
        Insert-ContactSQL -Server $Server -Database $Database -Table $Table -Object $test
        Test-ShowSQL
    }
}

function Test-UpdateSQL
{
    Param($Id = 1)
    $test = Set-Contact -First_Name $First_Name -Last_Name $Last_Name -Phones $Phones -Emails $Emails -Contact_Title $Contact_Title
    if ($test)
    {
        Update-ContactSQL -Id $Id -Server $Server -Database $Database -Table $Table -Object $test
        Test-ShowSQL
    }
}

function Test-DeleteSQL
{
    Param($Id = 2)
    Delete-ContactSQL -Id $Id -Server $Server -Database $Database -Table $Table
    Test-ShowSQL
}

function Test-ShowSQL
{
    Show-ContactSQL -Server $Server -Database $Database -Table $Table -Head "$style`r$jQuery`r$script" -Title $title
}

if($ShowSql) 
{
    Test-ShowSQL
}

if($ShowCsv) 
{
    Test-ShowCSV
}

cd "C:$env:HOMEPATH\Documents\Source Files\ps1"
if (!$First_Name) {
    . .\'Test Contact Type Module.ps1' Paul Ighofose 01932,777 paul@home, paul@work 'Mr.'
}