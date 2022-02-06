Add-Type -AssemblyName PresentationFramework

function Show-WPFWebViewForm {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        [ValidateNotNullOrEmpty()]
        [psobject]
        $object
    )
    $innerHTML = '
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
    <meta http-equiv="x-ua-compatible" content="IE=11">
    <head>
        <title>
            ' + $object.Title + '
        </title>
        <script>
            ' + $object.Script + '
        </script>
        <style>
            ' + $object.Style + '
        </style>
        ' + $object.Head + '
    </head>
    <body>
        ' + $object.Body + '
    </body>
</html>
'
    [xml]$XAML = '
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="' + $object.Title + '"
    WindowStartupLocation="' + $object.WindowLocation + '"
    SizeToContent="' + $object.WindowSize + '"
    Top="' + $object.WindowTop + '">
    <Grid>
        <Grid.RowDefinitions>
            ' + $object.RowDefinitions + '
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            ' + $object.ColumnDefinitions + '
        </Grid.ColumnDefinitions>
        ' + $object.XAMLBody + '
        <WebBrowser
            HorizontalAlignment="Left"
            Margin="10,10,10,10"
            VerticalAlignment="Top"
            Name="WebBrowser"
            ' + $object.BrowserColSpan + '
            ' + $object.BrowserColumn + '
            ' + $object.BrowserRowSpan + '
            ' + $object.BrowserRow + ' 
            />
    </Grid>
</Window>
    '; 
    $reader = New-Object System.Xml.XmlNodeReader($XAML)
    $_ = ConvertFrom-Json "{'LoadIt': 0}"
    $Form = [Windows.Markup.XamlReader]::Load($reader)
    $Form.Width = $object.FormWidth
    $Form.Topmost = $object.FormTopMost
    $WebBrowser = $Form.FindName('WebBrowser')
    $WebBrowser.Width = $object.BrowserWidth
    $WebBrowser.Height = $object.BrowserHeight
    $WebBrowser.NavigateToString($innerHTML)
    $WebBrowser.Add_Navigated(
        {
            if($loaded -gt 0) 
            {
                Set-Variable -Name "selected" -Value "$($WebBrowser.Document.Title)" -Scope Global
                $ht = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($selected, [System.Collections.Hashtable])
                foreach($k in $ht.Keys)
                {
                    $Form.FindName($k).Text = ($ht[$k] -join '| ')
                }
            } 
            else 
            {
                Set-Variable -Name "loaded" -Value 1 -Scope Global
            }
        }
    )

    if ($object.Sources[0] -ne 'Empty') {
        #$GetParams = {foreach ($fld in ((("$(get-help $_)" -split "$_ ")[1] -split "Common")[0] -split ' ' -replace '[][]', '')) {if($fld.IndexOf('-') -eq 0) {$fld.Substring(1)}}}
        $i = 0
        $fields = @()
        $names = @()
        $cbo = @()
        $cmds = @()

        foreach ($src in $object.Sources) 
        {
            if ($src.length -gt 0) 
            {
                $cbo += $Form.FindName($src[0])
                $cbo[$i].SelectedIndex = 0
                foreach ($itm in $src[1]) 
                {
                    $cbo[$i].Items.Add($itm)  | Out-Null
                }
                foreach ($fld in $src[2]) 
                {
                    $fields += $Form.FindName($fld)
                    $names += $Form.FindName($fld).Name
                }
                $cmds = $src[3]
                $cbo[$i].Add_SelectionChanged({
                    if ($this.SelectedIndex -ne 0) 
                    {
                        
                        $flts = @{}
                        for ($j = 0; $j -lt $fields.Length; $j++) 
                        {
                            $flts.Add($names[$j], ($fields[$j] | % {$_.Text -split '\| '}))
                        }
                        $cmdArr = $cmds[$this.SelectedIndex]
                        if ($this.SelectedItem -iin $object.JQuery)
                        {
                            $WebBrowser.InvokeScript($cmdArr)
                        }
                        elseif ($this.SelectedItem -eq 'Clear Fields')
                        {
                            $fields | % {$_.Text = ''}
                        }
                        else
                        {
                            $val = ($fields|select @{L='Text';E=({$_.Text -split ', '})})
                            $prms = @{}
                            $prms.Add('Title', $object.Title)
                            $prms.Add('Head', $object.Head)
                            $outer = $WebBrowser.InvokeScript('getHTML')
                            $prms.Add('Body', "$outer")
                            $prms.Add('HtmlOnly', $true)
                            $prms.Add('Object', $null)
                            if ($this.SelectedItem -like 'Filter*')
                            {
                                $prms.Add('Filter', $flts)
                            }
                            for ($j = 0; $j -lt $val.Length; $j++) 
                            {
                                $prms.Add($names[$j], ($val[$j]|select -ExpandProperty Text))
                            }
                            foreach ($itm in $object.DataSource) 
                            {
                                $prms.Add($itm.Key, $itm.Value)
                            }
                            foreach ($cmd in $cmdArr)
                            {
                                $prm = @{}
                                foreach ($itm in $prms.Keys) 
                                {
                                    foreach ($param in (Get-Parameters $cmd)) 
                                    {
                                        if ($itm -eq $param)
                                        {
                                            $prm.Add($itm, $prms[$itm])
                                        }
                                 
                                    }
                                }
                                $prms.Object = (Invoke-Expression "$cmd @prm")
                            }
                            Set-Variable -Name "loaded" -Value 0 -Scope Global
                            if ($prms.Object[0])
                            {
                                $WebBrowser.NavigateToString($prms.Object)
                            }
                        }
                        $this.SelectedIndex = 0
                    }
                })
                $i += 1
            }
        }
    }

    if ($object.Objects[0] -ne 'Empty') {
        $i = 0
        $field = @()
        $button = @()
        forEach ($action in $object.Objects) {
            if ($action.Length -gt 0) {
                $field += $Form.FindName($action[1])
                $button += $Form.FindName($action[0])
                $button[$i].Add_Click({
                    $WebBrowser.NavigateToString($field[((0..($object.Objects.Count-1)) | where {$object.Objects[$_] -eq $this.Name})].Text)
                })          
                $i += 1
            }
        }
    }
    $Form.ShowDialog()
}