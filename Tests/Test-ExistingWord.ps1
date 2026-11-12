$ErrorActionPreference='Stop'
if(@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count){throw 'Close actual Word sessions before running this generated-session test.'}
$names=@('PrintHiddenText','UpdateFieldsAtPrint','UpdateLinksAtPrint','UpdateLinksAtOpen','SaveNormalPrompt','ShowMarkupOpenSave')
$appRoot=Split-Path -Parent $PSScriptRoot
$python=Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$word=$null;$doc=$null;$baseline=@{}
try {
 $word=New-Object -ComObject Word.Application;$word.Visible=$false;$word.DisplayAlerts=0
 $userProcess=@(Get-Process WINWORD)[0]
 foreach($name in $names){$baseline[$name]=[bool]$word.Options.$name}
 $doc=$word.Documents.Add();$doc.Content.Text='Unsaved document in the existing user session.'
 $text=$doc.Content.Text;$saved=$doc.Saved;$count=$word.Documents.Count
 $output=& $python -X utf8 (Join-Path $PSScriptRoot 'existing_word_child.py') $userProcess.Id 2>&1
 if($LASTEXITCODE -ne 0){throw ($output|Out-String)}
 if(!(Get-Process -Id $userProcess.Id -ErrorAction SilentlyContinue)){throw 'Existing Word process was closed.'}
 if($doc.Content.Text -ne $text -or $doc.Saved -ne $saved -or $word.Documents.Count -ne $count){throw 'Existing unsaved document was changed.'}
 foreach($name in $names){if([bool]$word.Options.$name -ne $baseline[$name]){throw ('Existing session preference changed: '+$name)}}
 $other=@(Get-Process WINWORD|Where-Object {$_.Id -ne $userProcess.Id})
 if($other.Count){throw 'Conversion or recovery Word process was left running.'}
 $output|Write-Output
 Write-Output 'PASS: existing unsaved document, process, document count and all six options preserved after normal conversion and forced worker recovery.'
} finally {
 if($doc){$doc.Close(0);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc)}
 if($word){foreach($name in $baseline.Keys){$word.Options.$name=$baseline[$name]};$save=0;$word.Quit([ref]$save);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)}
}
