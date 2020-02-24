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
          Get-ChildItem  -Path $directory -Include "*.txt", "*.docx", "*.doc", "*.xlsx", "*.pdf" -Recurse -ErrorAction SilentlyContinue |`
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
          }
          #Searches files for a match from the word bank
          if($wordDocs){
            $word = New-Object -ComObject Word.application
            foreach ($target in $wordDocs){
              try{

              $document = $word.Documents.Open($target, $false, $true, $false, "ttt")
                    $content = $document.content.Text
                    foreach ($elem in $wordToSearch) 
                        { 
                            if ($null -ne $content -and $content.ToUpper().Contains($elem)) 
                            { 
                                $result = New-Object psobject -Property @{
                                    Location = $target
                                    Type = $elem
                                    Format = "Word" 
                                } 
                                #$out += $result
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
                  Format = "Word"
              }
              $out = $out + $result          
              }
            }
            $word.Quit()
          }
          if($excelDocs){
            $Excel = New-Object -ComObject Excel.Application
            foreach ($target in $excelDocs){
              try{

              $Workbook = $Excel.Workbooks.Open($target, $false, $true)
                    for($i = 1; $i -lt $($Workbook.Sheets.Count() + 1); $i++){
                        $Range = $Workbook.Sheets.Item($i).Range("A:Z")
                        foreach ($elem in $wordToSearch) 
                        { 
                            if ($Range.Find($elem)) 
                            { 
                                $result = New-Object psobject -Property @{
                                    Location = $target
                                    Type = $elem 
                                    Format = "Excel"
                                } 
                                #$out += $result    
                                $out = $out + $result                            
                            } 
                        }
                    }
                    $Workbook.close($false)              
              }
              catch [System.Runtime.InteropServices.COMException]{
                $result = New-Object psobject -Property @{
                  Location = $target
                  Type = $_.Exception.Message
                  Format = "Excel"
              }
              $out = $out + $result
              }     
            }
            $Excel.Quit()
          }
          if($pdfDocs){
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
                          Format = "PDF"
                      } 
                      #$out += $result
                      $out = $out + $result 
                      break
                  }                  
                }
                $PDdoc.close()
              }
              catch [System.Runtime.InteropServices.COMException]{
                $result = New-Object psobject -Property @{
                  Location = $target
                  Type = $_.Exception.Message
                  Format = "PDF"
              }
              $out = $out + $result
              $PDdoc.close()
              $adobe.CloseAllDocs()
              }
            }
            $adobe.CloseAllDocs()
            $adobe.exit()
          }
          if($txtDocs){
            foreach ($target in $txtDocs){
              try{
              $content = Get-Content $target
              foreach ($elem in $wordToSearch){ 
                if ($null -ne $content -and $content.ToUpper().Contains($elem))
                { 
                    $result = New-Object psobject -Property @{
                        Location = $target
                        Type = $elem 
                        Format = "TXT"
                    } 
                    #$out += $result 
                    $out = $out + $result 
                    break                              
                } 
              }
              }
              catch [System.Management.Automation.ItemNotFoundException]{
                $result = New-Object psobject -Property @{
                  Location = $target
                  Type = "ERROR" 
                  Format = "TXT"
                }  
                $out = $out + $result
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
        $out | Sort-Object -Property Location  | Format-Table -AutoSize -Property Type, Location
    }
    $browse.Dispose()
} Find-Folders
