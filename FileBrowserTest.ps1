$wordToSearch = @("PASSPORT", "DEPENDENTS", "EFMP", "EXCEPTIONAL FAMILY MEMBER", "DEROS", "OUT OF CYCLE", "ATAAPS", "SOCIAL", "CARDS", "SPOUSE", "SIGNIFICANT OTHER", "DRIVERS LICENSE NUMBER", "OPR", "EPR", "SSN", "SSAN", "SOCIAL ROSTER", "RECALL ROSTER", "ALPHA ROSTER", "DOB", "DATE OF BIRTH", "BANK ROUTING NUMBER", "GAINS ROSTER", "LOSSES", "INSURANCE", "RATER", "RATEE", "UMPR", "REPORTS", "DD577", "AF910", "AF 910", "AF911", "AF 911", "AF912", "AF 912", "LEAVE", "AF707", "AF 707", "AF780", "AF 780", "ADDITIONAL DUTY", "TEST")
$out = @() 
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
                    Get-ChildItem  -Path $directory -Include "*.txt", "*.docx", "*.doc", "*.xlsx" -Recurse -ErrorAction SilentlyContinue -Force |`
            ForEach-Object{
                $file = $_.FullName
                #Opens, reads, and searches .docx files
                if($file -match '.docx' -or $file -match '.doc'){
                    $word = New-Object -ComObject Word.application
                    $document = $word.Documents.Open($file, $false, $true)
                    $content = $document.content.Text
                    foreach ($elem in $wordToSearch) 
                        { 
                            if ($null -ne $content -and $content.ToUpper().Contains($elem))
                            #if ($content.contains($elem)) 
                            { 
                                $result = New-Object psobject -Property @{
                                    Location = $file
                                    Type = $elem 
                                } 
                                $out += $result   
                                break                            
                            } 
                        }
                    # if($document.content.Text.ToUpper.contains()){
                    # $file
                    # }
                    $document.close()
                    $word.Quit()
                }
                elseif($File -match '.xlsx'){
                    $Excel = New-Object -ComObject Excel.Application
                    $Workbook = $Excel.Workbooks.Open($File, $false, $true)
                    for($i = 1; $i -lt $($Workbook.Sheets.Count() + 1); $i++){
                        $Range = $Workbook.Sheets.Item($i).Range("A:Z")
                        foreach ($elem in $wordToSearch) 
                        { 
                            if ($Range.Find($elem)) 
                            { 
                                $result = New-Object psobject -Property @{
                                    Location = $file
                                    Type = $elem 
                                } 
                                $out += $result                               
                            } 
                        }
                        #$Target = $Range.Find("TeSt")
                        # if($null -ne $Target){
                        #     $File
                        # }
                    }
                    $Workbook.close()
                    $Excel.Quit()
                   }
                #  elseif($file -match '.pdf'){
                #     $document = $word.Documents.Open("D:\Test\PDFTest.pdf", $false, $true, $false,"" ,"", $false, "", "", 15)  
                #  }
                elseif($File -match '.txt'){
                    $content = Get-Content $_.FullName
                    foreach ($elem in $wordToSearch) 
                        { 
                            if ($null -ne $content -and $content.ToUpper().Contains($elem))
                            #if ($content.contains($elem)) 
                            { 
                                $result = New-Object psobject -Property @{
                                    Location = $file
                                    Type = $elem 
                                } 
                                # $result = New-Object System.Object
                                # $result = Add-Member NoteProperty -Name "Type" -Value ($elem)
                                # $result = Add-Member NoteProperty -Name "Location" -Value ($file)
                                $out += $result 
                                break                              
                            } 
                        }
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
        $out | Format-Table
    }
    $browse.Dispose()
} Find-Folders
