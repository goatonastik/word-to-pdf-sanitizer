# Settings are read and changed only on the proven conversion application.
function Add-WordGuardEvent($audit,$stage,$name,$result) {
 if(!$audit.ContainsKey('wordGuardEvents')){$audit.wordGuardEvents=@()}
 $audit.wordGuardEvents+=,@{stage=$stage;setting=$name;result=$result}
 if(Test-Path Function:\Trace-Word){Trace-Word 'Word export guard event' @{checkpoint=$stage;setting=$name;result=$result}}
}
function Assert-WordGuardSettings($application,$audit,$stage,[bool]$repair=$true) {
 $required=@{PrintHiddenText=$false;UpdateFieldsAtPrint=$false;UpdateLinksAtPrint=$false;UpdateLinksAtOpen=$false;SaveNormalPrompt=$false;ShowMarkupOpenSave=$false;AutomationSecurity=3}
 foreach($name in $required.Keys){
  $holder=if($name -eq 'AutomationSecurity'){$application}else{$application.Options}
  try {$actual=$holder.$name} catch {Add-WordGuardEvent $audit $stage $name 'Read failed';throw "Cannot verify Word setting $name at $stage. PDF withheld."}
  if($null -eq $actual){Add-WordGuardEvent $audit $stage $name 'Read returned no value';throw "Cannot verify Word setting $name at $stage. PDF withheld."}
  $valid=if($name -eq 'AutomationSecurity'){[int]$actual -eq 3}else{[bool]$actual -eq [bool]$required[$name]}
  if(!$valid){
   Add-WordGuardEvent $audit $stage $name 'Unexpected value detected'
   if(!$repair){throw "Word setting $name changed while opening the document. Its effect cannot be undone safely. PDF withheld."}
   try {$holder.$name=$required[$name];$actual=$holder.$name} catch {Add-WordGuardEvent $audit $stage $name 'Correction failed';throw "Cannot restore Word setting $name at $stage. PDF withheld."}
   $valid=$null -ne $actual -and $(if($name -eq 'AutomationSecurity'){[int]$actual -eq 3}else{[bool]$actual -eq [bool]$required[$name]})
   if(!$valid){Add-WordGuardEvent $audit $stage $name 'Correction readback failed';throw "Word setting $name did not retain its required value at $stage. PDF withheld."}
   Add-WordGuardEvent $audit $stage $name 'Corrected and verified'
  }
 }
}
function Assert-WordExportGuard($application,$ownerPid,$ownerStarted,$excludedPids,$audit,$stage,[bool]$repair=$true) {
 $clock=[Diagnostics.Stopwatch]::StartNew()
 try {
  $current=Get-Process -Id $ownerPid -ErrorAction Stop
  if($excludedPids -contains $ownerPid -or $current.ProcessName -ne 'WINWORD' -or $current.StartTime.ToUniversalTime().ToString('o') -ne $ownerStarted){throw 'Word process ownership no longer matches.'}
  $boundPid=[WordPdfProcessIdentity]::ForWindow([int]$application.Hwnd)
  if($boundPid -gt 0 -and $boundPid -ne $ownerPid){throw 'Word automation is bound to a different process.'}
 } catch {
  Add-WordGuardEvent $audit $stage 'Process ownership' 'Verification failed'
  throw 'Cannot verify ownership of the conversion Word instance. PDF withheld; no settings changed.'
 }
 try {
  Assert-WordGuardSettings $application $audit $stage $repair
  $audit.wordGuardChecks=1+[int]$audit.wordGuardChecks
 } finally {$audit.wordGuardSeconds=[double]$audit.wordGuardSeconds+$clock.Elapsed.TotalSeconds}
}
