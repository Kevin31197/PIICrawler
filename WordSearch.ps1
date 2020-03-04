#This script crawls down a given directory and searches all compatible file types for
#possible PII from the given wordbank.
#Author: Amn Arriaga, Kevin
#Last Revision: 3/4/2020
$wordToSearch = @("PASSPORT", "DEPENDENTS", "EFMP", "EXCEPTIONAL FAMILY MEMBER", "DEROS", "OUT OF CYCLE", "ATAAPS", "SOCIAL", "CARDS", "SPOUSE", "SIGNIFICANT OTHER", "DRIVERS LICENSE NUMBER", "OPR", "EPR", "SSN", "SSAN", "SOCIAL ROSTER", "RECALL ROSTER", "ALPHA ROSTER", "DOB", "DATE OF BIRTH", "BANK ROUTING NUMBER", "GAINS ROSTER", "LOSSES", "INSURANCE", "RATER", "RATEE", "UMPR", "REPORTS", "DD577", "AF910", "AF 910", "AF911", "AF 911", "AF912", "AF 912", "LEAVE", "AF707", "AF 707", "AF780", "AF 780", "ADDITIONAL DUTY")
$wordDocs = @()
$excelDocs = @()
$pdfDocs = @()
$txtDocs = @()
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
		#Script
          Get-ChildItem  -Path $directory -Include "*.txt", "*.doc", "*.docx", "*.xlsx", "*.pdf" -Recurse -ErrorAction SilentlyContinue |`
          ForEach-Object{
            $file = $_.FullName
            #Seperates file locations into their respective array
            if($file -match "(\.docx)$" -or $file -match "(\.doc)$"){
              $wordDocs += $file
            }
            elseif($File -match "(\.xlsx)$"){
              $excelDocs += $file
            }
            elseif($file -match "(\.pdf)$"){
              $pdfDocs += $file
            }
            elseif($File -match "(\.txt)$"){
              $txtDocs += $file
            }
            Write-Progress -Activity "Searching for files..." -Status "$($_.FullName)"
          }
          #Searches files for a match from the word bank
          if($wordDocs){
            Write-Output "`n`n`nTotal Word Documents Found: $($wordDocs.count)"
            $t=0
            $word = New-Object -ComObject Word.application
            $word.WordBasic.DisableAutoMacros()
            $word.Application.AutomationSecurity = 2
            foreach ($target in $wordDocs){
              try{

              $document = $word.Documents.Open($target, $false, $true, $false, "pass")
                    $content = $document.content.Text
                    foreach ($elem in $wordToSearch) 
                        { 
                            if ($null -ne $content -and $content.ToUpper().Contains($elem)) 
                            { 
                                $result = New-Object psobject -Property @{
                                    Location = $target
                                    Type = $elem
                                } 
                                $out = $out + $result   
                                break                            
                            } 
                        }
                    $document.close($false)
              }
              catch [System.Runtime.InteropServices.COMException]{
                $result = New-Object psobject -Property @{
                  Location = $target
                  Type = $_.Exception.Message
              }
              $out = $out + $result          
              }
              #Creates progress bar
              $t = $t + 1
              Write-Progress -Activity "Processing Word Documents" -Status "File $t of $($wordDocs.count)" -PercentComplete ($t/$wordDocs.count*100)
            }
            $word.Quit()
          }
          if($excelDocs){
            Write-Output "Total Excel Documents Found: $($excelDocs.count)"
            $t=0
            $Excel = New-Object -ComObject Excel.Application
            foreach ($target in $excelDocs){
              try{
              $Workbook = $Excel.Workbooks.Open($target, $false, $true, 5, "pass")
                    for($i = 1; $i -lt $($Workbook.Sheets.Count() + 1); $i++){
                        $Range = $Workbook.Sheets.Item($i).Range("A:Z")
                        foreach ($elem in $wordToSearch) 
                        { 
                            if ($Range.Find($elem)) 
                            { 
                                $result = New-Object psobject -Property @{
                                    Location = $target
                                    Type = $elem 
                                }     
                                $out = $out + $result                            
                            } 
                        }
                    }
                    $Workbook.close($false)              
              }
              catch [System.Runtime.InteropServices.COMException]{
                $result = New-Object psobject -Property @{
                  Location = $targetgw
                  Type = $_.Exception.Message
              }
              $out = $out + $result
              } 
              #Creates progress bar  
              $t = $t + 1
              Write-Progress -Activity "Processing Excel Documents" -Status "File $t of $($excelDocs.count)" -PercentComplete ($t/$excelDocs.count*100)  
            }
            $Excel.Quit()
          }
          if($pdfDocs){
            Write-Output "Total PDF Documents Found: $($pdfDocs.count)"
            $t=0
            $adobe = New-Object -ComObject AcroExch.App
            foreach ($target in $pdfDocs){
              try {
                $PDdoc = New-Object -ComObject AcroExch.PDDoc
                $PDdoc.Open($target) 
                $AVdoc = $PDdoc.OpenAVDoc("") 
                foreach ($elem in $wordToSearch){
                  if ($AVDoc.FindText($elem, 0, 0, 1)){
                      $result = New-Object psobject -Property @{
                          Location = $target
                          Type = $elem
                      } 
                      $out = $out + $result 
                      break
                  }                  
                }
                if(!$PDdoc.close()){
                  $adobe.CloseAllDocs()
                  $adobe.exit()
                  $adobe = New-Object -ComObject AcroExch.App
                }
                $PDdoc.close()
              }
              catch [System.Runtime.InteropServices.COMException]{
                $result = New-Object psobject -Property @{
                  Location = $target
                  Type = $_.Exception.Message
              }
              $out = $out + $result
              $PDdoc.close()
              $adobe.CloseAllDocs()
              }
              #Creates progress bar
              $t = $t + 1
              Write-Progress -Activity "Processing PDF Documents" -Status "File $t of $($pdfDocs.count)" -PercentComplete ($t/$pdfDocs.count*100)
            }
            $adobe.CloseAllDocs()
            $adobe.exit()
          }
          if($txtDocs){
            Write-Output "Total Text Documents Found: $($txtDocs.count)"
            $t=0
            foreach ($target in $txtDocs){
              try{
                if(test-path -path $target){
                  $content = Get-Content $target
                }
                #$content = Get-Content $target
                foreach ($elem in $wordToSearch){ 
                  if ($null -ne $content -and $content.ToUpper().Contains($elem))
                  { 
                      $result = New-Object psobject -Property @{
                          Location = $target
                          Type = $elem 
                      } 
                      $out = $out + $result 
                      break                              
                  } 
                }
              }
              catch [Microsoft.PowerShell.Commands.GetContentCommand]{
                $result = New-Object psobject -Property @{
                  Location = $target
                  Type = "ERROR" 
                }  
                $out = $out + $result
              }
              #Creates progress bar
              $t = $t + 1
              Write-Progress -Activity "Processing Text Documents" -Status "File $t of $($txtDocs.count)" -PercentComplete ($t/$txtDocs.count*100)
            }
          }
        } else #Cancels script 
        {
            $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
            if($res -eq "Cancel")
            {
                #Ends script
                return
            }
        }
        #$out | Sort-Object -Property directory  | Format-Table -AutoSize -Property Type, Location | Out-File "$($env:USERPROFILE)\Documents\Output.txt" -Width 300
        $out |Export-Csv "$($env:USERPROFILE)\Documents\Output $(Get-Date -Format MM-dd_hhmm).csv"  -NoTypeInformation
    }
    $browse.Dispose()
} Find-Folders