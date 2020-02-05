function Find-Folders {
    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
    $browse.RootFolder = [System.Environment+SpecialFolder]'MyComputer'
    $browse.ShowNewFolderButton = $false
    $browse.Description = "Select a directory"

    $loop = $true
    while($loop)
    {
        if ($browse.ShowDialog() -eq "OK")
        {
        $loop = $false
		$directory = $browse.SelectedPath
		#Insert your script here
                    Get-ChildItem  $directory |`
            ForEach-Object{
                $file = $_.FullName
                #Opens, reads, and searches .docx files
                if($file -match '.docx' -or $file -match '.doc'){
                    $word = New-Object -ComObject Word.application
                    $document = $word.Documents.Open($file, $false, $true)
                    if($document.content.Text -match 'test'){
                    $file
                    }
                    $document.close()
                    $word.Quit()
                }
                elseif($File -match '.xlsx'){
                    $Excel = New-Object -ComObject Excel.Application
                    $Workbook = $Excel.Workbooks.Open($File)
                    for($i = 1; $i -lt $($Workbook.Sheets.Count() + 1); $i++){
                        $Range = $Workbook.Sheets.Item($i).Range("A:Z")
                        $Target = $Range.Find($SearchString)
                        if($null -ne $Target){
                            $File
                        }
                    }
                    $Excel.Quit()
                   }
                #  elseif($file -match '.pdf'){
                #     $document = $word.Documents.Open("D:\Test\PDFTest.pdf", $false, $true, $false,"" ,"", $false, "", "", 15)  
                #  }
                $content = Get-Content $_.FullName
                if($content -match 'test'){
                $file
                }
            }
        } else
        {
            $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
            if($res -eq "Cancel")
            {
                #Ends script
                return
            }
        }
    }
    $browse.Dispose()
} Find-Folders