#K.V.Iankovich 2024
#А script for generating documents based on text file raw data

function Write-Log {
    param (
    [string]$logString,
    $runPath
    )
    [string]$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    [string]$logFile = $runPath + "\" + $stamp.Remove(13) + ".log"
    [string]$logMessage = "$stamp $logString"
    Add-content $logFile -value $logMessage
}

Function Select-File ([string]$title, [string]$filter)
{
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = $title
    $OpenFileDialog.InitialDirectory = ".\"
    $OpenFileDialog.filter = $filter
    $openFileDialog.ShowHelp = $true
    If ($OpenFileDialog.ShowDialog() -eq "Cancel")
    {
        [System.Windows.Forms.MessageBox]::Show("No File Selected. Please select a file !", "Error", 0, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
        Return null
    }
        $Global:SelectedFile = $OpenFileDialog.FileName
        Return $SelectedFile #add this return
}

$todocConf = Select-File -title "Выберите файл конфигурации" -filter "Configuration file (*.conf) | *.conf"
#Adding variables from the configuration file
foreach ($i in $(Get-Content -Path $todocConf)){
    Set-Variable -Name $i.split("=")[0] -Value $i.split("=",2)[1]
}

$dataFile = Select-File -title "Выберите файл с исходными данными" -filter "Txt file (*.txt) | *.txt" 
$templateFile = Select-File -title "Выберите файл с шаблоном" -filter "Docx file (*.docx) | *.docx" 
[string[]]$arrString = Get-Content $dataFile
$runPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

$hashsetEmployees = @{}
if($arrString.Count -ge 0) {     
    $rawFields = $pattern_search | Select-String "(?<=<)\w+(?<!>)" -AllMatches
    [string[]]$fields = $rawFields.Matches.Value
    $arrString | Select-String -Pattern $pattern_search | ForEach-Object {
        $hashSetFields = @{}
        for ($i = 0; $i -le $fields.Length; $i++) {
            $n = $i + 1
            [string]$stringValue = $($_.matches.groups[$n])
            $stringValue = $stringValue.TrimStart("")
            $stringValue = $stringValue.TrimEnd("")
            $stringValue = $stringValue -replace '\s{2,}', ' '
            if ($stringValue) {
                $hashSetFields.add($fields[$i], $stringValue)
            }
        }
        $hashsetEmployees.add([guid]::NewGuid().ToString(), $hashSetFields)
    }

    Write-Host "Выбраных записей -" $hashSetEmployees.Count

    #filter by object fields
    if($pattern_filter_enable -eq "true") { 
        $removeList = [System.Collections.Generic.List[string]]::new()
        if ($pattern_filter_invers -eq "true") {
            foreach ($itemKey in $hashsetEmployees.Keys) {
                [string]$hashsetEmployees.$itemKey.Values | Where-Object { $_ -cmatch $pattern_filter } | ForEach-Object { $removeList.Add($itemKey) }
            }
        } elseif ($pattern_filter_invers -eq "false") {
            foreach ($itemKey in $hashsetEmployees.Keys) {
                [string]$hashsetEmployees.$itemKey.Values | Where-Object { $_ -cnotmatch $pattern_filter } | ForEach-Object { $removeList.Add($itemKey) }
            }
        }       
        if ($removeList.Count -gt 0) {
            $removeList | ForEach-Object { $hashsetEmployees.Remove($_) }
        }
    }
    
    #sorting by a specific field
    $sortedSetEmployees = [System.Collections.SortedList]::new()
    foreach ($key in $hashSetEmployees.Keys) {
        try {
            $sorting_filed
            $sortedSetEmployees.add($hashSetEmployees.$key.$sorting_field, $key)
        }
        catch {
            [string]$str = $hashSetEmployees.$key.$sorting_filed
            Write-Warning  "$str - дублирующая запись"
            # $PSItem.Exception
        }
    }
    Write-Host "Отфильтрованных, отсортированных записей -" $sortedSetEmployees.Count

    #Paramets for Find.Execute (Word)
    $MatchCase = $false
    $MatchWholeWorld = $true
    $MatchWildcards = $false
    $MatchSoundsLike = $false
    $MatchAllWordForms = $false
    $Forward = $false
    $Wrap = 1
    $Format = $false
    $Replace = 2

    $word = NEW-Object –comobject Word.Application
    $word.visible = $False

    foreach ($employeeKey in $sortedSetEmployees.Keys) {
        $hashKeyFio = $sortedSetEmployees.$employeeKey
        $hashtableRecord = $hashsetEmployees.$hashKeyFio
              
        $document = $word.documents.open($templateFile, $false, $true)
        $employeeKey
        $missingFields = (Compare-Object -ReferenceObject ($fields) -DifferenceObject ([System.Collections.ArrayList]$hashtableRecord.Keys) | Where-Object{$_.sideIndicator -eq "<="}).InputObject
        
        if ($missingFields.Count -gt 1) {
            [string]$missingString = $missingFields -join ', '
            Write-Log "$employeeKey - отсутствует: $missingString" -runPath $runPath
            Write-Host -Message "$employeeKey - отсутствует: $missingString" -ForegroundColor Red
        }
       
        foreach ($recordKey in $hashtableRecord.Keys) {
            [string]$replaceText = $hashtableRecord.$recordKey
            [string]$findPatern = "<" + $recordKey.ToUpper() + ">"  
            $document.Content.Find.Execute($findPatern, $MatchCase, $MatchWholeWorld, $MatchWildcards, $MatchSoundsLike, $MatchAllWordForms, $Forward, $Wrap, $Format, $replaceText, $Replace) | Out-Null
        }
        
        # Save as and close the document (directory of the PowerShell script)
        $day = Get-Date -Format "yyyyMMdd"

        New-Item -Path $runPath -Name $day -ItemType Directory -ErrorAction SilentlyContinue -InformationAction SilentlyContinue | Out-Null

        $saltNameFile = Split-Path $templateFile -leaf 
        [string]$fileName = $employeeKey + "-" + $saltNameFile
        $document.SaveAs([ref]"$runPath\$day\$fileName")
        $document.Close(-1)
    }
    $word.Quit()
    # freeing up resources
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}





    
