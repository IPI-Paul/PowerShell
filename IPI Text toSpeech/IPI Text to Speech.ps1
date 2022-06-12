[System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null


Function Get-FileName($initialDirectory)
{  
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog;
    $OpenFileDialog.initialDirectory = $initialDirectory;
    $OpenFileDialog.filter = “All files (*.*)| *.*”;
    $OpenFileDialog.ShowDialog() | Out-Null;
    $file = $OpenFileDialog.filename;
    $file;
} 

Function New-WPFMessageBox {

    [cmdletbinding(DefaultParameterSetName = "standard")]
    [alias("nmb")]
    [Outputtype([int], [boolean], [string])]

    Param(
        [string]$Message,
        [string]$Title = "Play Text to Speech"
    )

    if ($PSEdition -eq 'Core') {
        Write-Warning "Sorry. This command will not run on PowerShell Core."
        #bail out
        Return
    }

    # It may not be necessary to add these types but it doesn't hurt to include them
    # but if they can't be loaded then this function will never work anwyway
    Try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction stop
        Add-Type -assemblyName PresentationCore -ErrorAction stop
        Add-Type -AssemblyName WindowsBase -ErrorAction stop
        Add-Type -AssemblyName System.speech
        $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
        $speak.Rate = -1 # -10 is slowest, 10 is fastest
    }
    Catch {
        Throw $_
        #make sure we abort
        return
    }
    
    $form = New-Object System.Windows.Window
    #define what it looks like
    $form.Title = $Title
    $form.Height = 80
    $form.Width = 300
    $form.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen

    if ($Background) {
        Try {
            $form.Background = $Background
        }
        Catch {
            Write-Warning "See https://docs.microsoft.com/en-us/dotnet/api/system.windows.media.brushes?view=netframework-4.7.2 for help on selecting a proper color. You can enter a color by name or X11 color value."
            Throw $_
        }
    }

    $grid = New-Object System.Windows.Controls.Grid

    if ($prm -eq 'open') {
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = "_Clipboard"
        $btn.Width = 75
        $btn.Height = 25
        $btn.Margin = "-175,5,0,0"
        $btn.Add_click( {
                $script:r = "Clipboard";
                $form.Close();
            })
        $btn1 = New-Object System.Windows.Controls.Button
        $btn1.Content = "_File"
        $btn1.Width = 75
        $btn1.Height = 25

        $btn1.Margin = "0,5,0,0"
        $btn1.Add_click( {
                $script:r = "File";
                $form.Close();
            })
        $btn2 = New-Object System.Windows.Controls.Button
        $btn2.Content = "_Web Page"
        $btn2.Width = 75
        $btn2.Height = 25
        $btn2.Margin = "175,5,0,0"
        $btn2.Add_click( {
                $script:r = "Web Page";
                $form.Close();
            })
    } elseif ($prm -eq 'run') {
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = "_Start"
        $btn.Width = 55
        $btn.Height = 25
        $btn.Margin = "-200,5,0,0"
        $btn.Add_click( {
                $script:r = 'Start';
                $speak.SpeakAsync($Content);
            })
        $btn1 = New-Object System.Windows.Controls.Button
        $btn1.Content = "_Pause"
        $btn1.Width = 55
        $btn1.Height = 25

        $btn1.Margin = "-70,5,0,0"
        $btn1.Add_click( {
                $script:r = 'Pause'
                $speak.Pause()
            })
        $btn2 = New-Object System.Windows.Controls.Button
        $btn2.Content = "_Resume"
        $btn2.Width = 55
        $btn2.Height = 25
        $btn2.Margin = "65,5,0,0"
        $btn2.Add_click( {
                $script:r = 'Resume'
                $speak.Resume()
            })
        $btn3 = New-Object System.Windows.Controls.Button
        $btn3.Content = "_Restart"
        $btn3.Width = 55
        $btn3.Height = 25
        $btn3.Margin = "200,5,0,0"
        $btn3.Add_click( {
                $script:r = 'Restart'
                $form.Close()
            })
    }
    $grid.AddChild($btn)
    $grid.AddChild($btn1)
    $grid.AddChild($btn2)
    if ($prm -eq 'run') {
        $grid.AddChild($btn3)
    }
    $form.add_Loaded( { $btn.Focus() })

    #display the form
    $form.AddChild($grid)
    #$form.Add_Load({ $form.Activate() })
    $form.ShowDialog() #| Out-Null

    if ("$script:r" -eq "Restart") {
        ReadOutLoud
    }

    #write the button result to the pipeline if not using -Quiet
    if (-Not $Quiet) {
        return $script:r;
    }
} #end function

function ReadOutLoud 
{
    $prm = 'open';
    $selection = (New-WPFMessageBox)[1];
    $prm = 'run';
    #$choices = @('Clipboard', 'File', 'Web Page');
    #$selection = $choices | Out-GridView -Title 'Select Type' -OutputMode Single;

    if ("$selection" -eq "File") {
        $file = Get-FileName -initialDirectory "H:”;
    } elseif ("$selection" -eq "Web Page") {
        [void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic');
        $title = 'Web Page Link';
        $msg   = 'Please paste the link in here:';
        $file = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title);
    } elseif ("$selection" -eq "Clipboard") {
        $file = "";
        $Content = (Get-Clipboard);
        New-WPFMessageBox;
    } else {
        $file = "";
        $title = 'Error!';
        $msg   = 'No Selection was made';
        [System.Windows.Forms.MessageBox]::Show($msg, $title);
    }
    if ("$file" -ne "") { 
        if ($file.Substring(($file.Length - 4),4) -in @(".log", ".txt")) {
            $content = Get-Content "$file";
        }
        if ("$file" -match ".docx" -or "$file" -match ".pdf" -or "$file" -match "http") {
                #Instance of word
            $Word=NEW-Object –comobject Word.Application
            $Word.visible = $false

            #open file and take content of word file
            $Document=$Word.documents.open($file, $false, $true)
            $content = $document.content.Text

            $word.ActiveDocument.Close($true)
            if ($word.Documents.Count -eq 0) {
                $word.Quit()
            }
        }
        $Content = $content;
        New-WPFMessageBox;
     }
}

ReadOutLoud