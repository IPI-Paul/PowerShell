# Logging helper
function Write-Log {
    param (
        [string]$Text,
        [string]$Color = "Black",
        [switch]$Bold,
        [switch]$Italic
    )

    $Window.Dispatcher.Invoke([action]{
        $Paragraph = New-Object System.Windows.Documents.Paragraph
        $Paragraph.Margin = "0"     # Removes extra spacing
        $Paragraph.LineHeight = 12  # Adjust as needed (12-15 works well)

        if ($Text -match "^http") {
            # Create Hyperlink
            $hyperlink = New-Object System.Windows.Documents.Hyperlink
            $hyperlink.NavigateUri = [Uri]$Text
            $hyperlink.Inlines.Add($Text)
            $hyperlink.Cursor = [System.Windows.Input.Cursors]::Hand
            $hyperlink.Foreground = [System.Windows.Media.Brushes]::Blue
            $hyperlink.TextDecorations = [System.Windows.TextDecorations]::Underline

            $Paragraph.Inlines.Add($hyperlink)
        } else {
            $Run = New-Object System.Windows.Documents.Run($Text)
            $Run.Foreground = $Color
            if ($Bold) { $Run.FontWeight = "Bold" }
            if ($Italic) { $Run.FontStyle = "Italic" }

            $Paragraph.Inlines.Add($Run)
        }

        $LogBox.Document.Blocks.Add($Paragraph)
        $LogBox.ScrollToEnd()
    })
}

Export-ModuleMember -Function Write-Log