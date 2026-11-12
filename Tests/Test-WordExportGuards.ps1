$ErrorActionPreference='Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'WordExportGuards.ps1')
Add-Type -TypeDefinition @'
using System;
public class GuardProbeOptions {
 public bool PrintHiddenText{get;set;}
 public bool UpdateFieldsAtPrint{get;set;}
 public bool UpdateLinksAtPrint{get;set;}
 public bool UpdateLinksAtOpen{get;set;}
 public bool SaveNormalPrompt{get;set;}
 public bool ShowMarkupOpenSave{get;set;}
}
public class GuardProbeApp {
 public GuardProbeOptions Options{get;set;}
 public GuardProbeApp(){Options=new GuardProbeOptions();}
 public bool RefuseMacroCorrection{get;set;}
 private int security=3;
 public int AutomationSecurity {get{return RefuseMacroCorrection?1:security;}set{security=value;}}
}
'@
foreach($name in 'PrintHiddenText','UpdateFieldsAtPrint','UpdateLinksAtPrint','UpdateLinksAtOpen','SaveNormalPrompt','ShowMarkupOpenSave'){
 $app=New-Object GuardProbeApp;$app.Options.$name=$true;$audit=@{}
 Assert-WordGuardSettings $app $audit 'Before open'
 if($app.Options.$name -or @($audit.wordGuardEvents|Where-Object {$_.result -eq 'Corrected and verified'}).Count -ne 1){throw ('Correction failed for '+$name)}
}
$app=New-Object GuardProbeApp;$app.AutomationSecurity=1;$audit=@{}
Assert-WordGuardSettings $app $audit 'Before open'
if($app.AutomationSecurity -ne 3){throw 'Macro protection was not corrected.'}
$app=New-Object GuardProbeApp;$app.RefuseMacroCorrection=$true;$audit=@{};$failed=$false
try {Assert-WordGuardSettings $app $audit 'Before export'} catch {$failed=$_.Exception.Message -like '*did not retain*'}
if(!$failed){throw 'Uncorrectable macro protection did not block export.'}
$app=New-Object GuardProbeApp;$app.Options.UpdateLinksAtOpen=$true;$audit=@{};$failed=$false
try {Assert-WordGuardSettings $app $audit 'After open' $false} catch {$failed=$_.Exception.Message -like '*changed while opening*'}
if(!$failed -or !$app.Options.UpdateLinksAtOpen){throw 'After-open drift was silently repaired rather than withheld.'}
$app=New-Object GuardProbeApp;$app.AutomationSecurity=1;$audit=@{};$failed=$false
try {Assert-WordExportGuard $app $PID 'invalid start time' @() $audit 'Before open'} catch {$failed=$_.Exception.Message -like '*ownership*'}
if(!$failed -or $app.AutomationSecurity -ne 1){throw 'Settings were accessed or changed before ownership was established.'}
Write-Output 'PASS: six option corrections, macro protection, failed correction refusal, after-open drift refusal, and ownership-before-settings checks.'
