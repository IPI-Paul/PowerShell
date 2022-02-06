function Set-WebForm 
{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=0)]
        $head, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=1)]
        $title , 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=2)]
        $script, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=3)]
        $style, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=4)]
        $body, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=5)]
        $wLocation, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=6)]
        $wSize, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=7)]
        $wTop, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=8)]
        $rDefs, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=9)]
        $cDefs, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=10)]
        $xBody, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=11)]
        $bColSpan, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=12)]
        $bCol, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=13)]
        $bRowSpan, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=14)]
        $bRow, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=15)]
        $xScript, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=16)]
        $fWidth, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=17)]
        $fHeight, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=18)]
        $fTopMost, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=19)]
        $bHeight, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=20)]
        $bWidth, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=21)]
        $objects, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=22)]
        $sources, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=23)]
        $dSource, 
        [Parameter(Mandatory=$false,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true,
                   Position=24)]
        $jQuery
    )
    class Structure
    {
        [string]$Head
        [string]$Title
        [string]$Script
        [string]$Style
        [string]$Body
        [string]$WindowLocation
        [string]$WindowSize
        [string]$WindowTop
        [string]$RowDefinitions
        [string]$ColumnDefinitions
        [string]$XAMLBody
        [string]$BrowserColSpan
        [string]$BrowserColumn
        [string]$BrowserRowSpan
        [string]$BrowserRow
        [string]$XAMLScript
        [int64]$FormWidth
        [int64]$FormHeight
        [boolean]$FormTopMost
        [int64]$BrowserHeight
        [int64]$BrowserWidth
        [System.Collections.ArrayList]$Objects
        [System.Collections.ArrayList]$Sources
        [System.Collections.ArrayList]$DataSource
        [System.Collections.ArrayList]$JQuery
    }
    
    if ("$head"-eq "") {
        $head = ''
    }
    if ("$title"-eq "") {
        $title = 'WPF Web Browser Test'
    } 
    if ("$body"-eq "") {
        $body = 'Hello World'
    } 
    if ("$wLocation"-eq "") {
        $wLocation = 'Manual'
    } 
    if ("$wSize"-eq "") {
        $wSize = 'WidthAndHeight'
    } 
    if ("$wTop"-eq "") {
        $wTop = '0'
    } 
    if ("$rDefs"-eq "") {
        $rDefs = '
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        '
    } 
    if ("$cDefs"-eq "") {
        $cDefs = '
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="Auto" />    
            <ColumnDefinition Width="Auto" />    
        ';
    } 
    if ("$xBody"-eq "") {
        $xBody = '
            <Label 
                Name = "lbl1" 
                Width = "300" 
                Grid.Column = "0" 
                Grid.Row = "0" 
                Content = "Put Text in here and Click Update to view in browser!"
            /> 
            <TextBox 
                Name = "txt1" 
                Width = "300" 
                Grid.Column = "0" 
                Grid.Row = "1" 
                Margin = "1"
            /> 
            <Button 
                Name = "btnUpdate" 
                Content = "Update" 
                Grid.Column = "1" 
                Grid.Row = "1" 
                Margin = "1"
            />
            <Button 
                Name = "btnUpdate1" 
                Content = "BUpdate" 
                Grid.Column = "2" 
                Grid.Row = "1" 
                Margin = "1"
            />
        ';
    } 
    if ("$bColSpan"-eq "") {
        $bColSpan = 'Grid.ColumnSpan="3"'
    } 
    if ("$bCol"-eq "") {
        $bCol = ''
    } 
    if ("$bRowSpan"-eq "") {
        $bRowSpan = ''
    } 
    if ("$bRow"-eq "") {
        $bRow = 'Grid.Row="2"'
    } 
    if ("$fWidth"-eq "") {
        $fWidth = 100
    } 
    if ("$fHeight"-eq "") {
        $fHeight = 50
    } 
    if ("$fTopMost"-eq "") {
        $fTopMost = $true;
    } 
    if ($bHeight -eq $null) {
        $bHeight = ($fHeight * 2.5)
    } 
    if ($bWidth -eq $null) {
        $bWidth = ($fWidth * 3.9)
    } 
    if ($objects.Count -eq 0) {
        $objects = @(@('btnUpdate', 'txt1'), @('btnUpdate1', 'txt1'), @())
    } 
    if ($sources.Count -eq 0) {
        $sources = @(@('Empty', ''))
    }  
    if ($dSource.Count -eq 0) {
        $dSource = @(@())
    } 
    return [Structure]@{
            Head = $head
            Title = $title
            Script = $script
            Style = $style
            Body = $body
            WindowLocation = $wLocation
            WindowSize = $wSize
            WindowTop = $wTop
            RowDefinitions = $rDefs
            ColumnDefinitions = $cDefs
            XAMLBody = $xBody
            BrowserColspan = $bColSpan
            BrowserColumn = $bCol
            BrowserRowspan = $bRowSpan
            BrowserRow = $bRow
            XAMLScript = $xScript
            FormWidth = $fWidth
            FormHeight = $fHeight
            FormTopmost = $fTopMost
            BrowserHeight = $bHeight
            BrowserWidth = $bWidth
            Objects = $objects
            Sources = $sources
            DataSource = $dSource
            JQuery = $jQuery
        }
}