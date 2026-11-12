param([string]$Workspace, [string]$SettingsFile, [int]$ParentPid=0, [switch]$RestoreOnly)
$ErrorActionPreference = 'Stop'
$word=$null; $owned=$false; $saved=@{}; $process=$null;$debugWriter=$null
. (Join-Path $PSScriptRoot 'PropertyInventory.ps1')
. (Join-Path $PSScriptRoot 'WordSessionSupport.ps1')
. (Join-Path $PSScriptRoot 'WordExportGuards.ps1')
$optionsPath=Join-Path $Workspace 'options.json'
$parent=if($ParentPid -gt 0) { Get-Process -Id $ParentPid -ErrorAction SilentlyContinue } else { $null }
$debugPath=$null
function Trace-Word($stage,$details=@{}) {
 if(!$debugPath) { return }
 try {
  $record=@{timestamp_utc=[DateTime]::UtcNow.ToString('o');pid=$PID;stage=$stage;details=$details}
  if($job) { $record.working_path=$job.source; $record.path_characters=([string]$job.source).Length; $record.document=$job.originalName }
  $line=($record | ConvertTo-Json -Depth 10 -Compress)
  if($debugWriter) { $debugWriter.WriteLine($line);$debugWriter.Flush() }
  else { [IO.File]::AppendAllText($debugPath,($line+[Environment]::NewLine),(New-Object Text.UTF8Encoding($false))) }
 } catch {} # Debug output must not change conversion behavior.
}
function Trace-Layout($document,$stage) {
 if(!$debugPath) { return }
 $metrics=@{}
 try { $metrics=@{compatibility_mode=$document.CompatibilityMode;pages=$document.ComputeStatistics(2);page_width=$document.PageSetup.PageWidth;page_height=$document.PageSetup.PageHeight;shapes=$document.Shapes.Count;fields=$document.Fields.Count} }
 catch { $metrics.statistics_unavailable=$_.Exception.GetType().Name }
 Trace-Word $stage $metrics
}
function Reply($path,$value) {
 $temp=$path+'.tmp'
 $value | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $temp -Encoding UTF8
 Move-Item -LiteralPath $temp -Destination $path -Force
}
try {
 $config=Get-Content -LiteralPath $SettingsFile -Raw | ConvertFrom-Json
 $debugPath=$config._wordDebug
 if($debugPath) { $debugWriter=New-Object IO.StreamWriter($debugPath,$true,(New-Object Text.UTF8Encoding($false)));$debugWriter.AutoFlush=$true }
 if($ParentPid -gt 0 -and !$parent) { throw 'Conversion parent process is no longer running.' }
 Trace-Word 'starting Word session'
 $before=@(Get-Process WINWORD -ErrorAction SilentlyContinue)
 if ($before.Count) { Trace-Word 'existing Word processes detected; creating isolated conversion instance' @{existing_pids=@($before.Id)} }
 $word=New-Object -ComObject Word.Application
 $ownedPid=[WordPdfProcessIdentity]::ForWindow([int]$word.Hwnd)
 $newProcesses=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before.Id -notcontains $_.Id })
 if($ownedPid -le 0 -and $newProcesses.Count -eq 1) { $ownedPid=$newProcesses[0].Id }
 if($ownedPid -le 0 -or $newProcesses.Id -notcontains $ownedPid) { throw 'Could not unambiguously identify the Word instance created for conversion.' }
 $owned=$true
 $process=Get-Process -Id $ownedPid
 $ownerStarted=$process.StartTime.ToUniversalTime().ToString('o')
 Reply (Join-Path $Workspace 'owner.json') @{pid=$process.Id; started=$process.StartTime.ToUniversalTime().ToString('o')}
 $word.Visible=$false; $word.DisplayAlerts=0; $word.AutomationSecurity=3
 Trace-Word 'configuring Word session' @{word_version=$word.Version;word_pid=$process.Id}
 if(Test-Path -LiteralPath $optionsPath) {
  $previous=Get-Content -LiteralPath $optionsPath -Raw | ConvertFrom-Json
  foreach($property in $previous.PSObject.Properties) { $word.Options.($property.Name)=$property.Value }
 }
 if($RestoreOnly) { return }
 foreach($name in 'PrintHiddenText','UpdateFieldsAtPrint','UpdateLinksAtPrint','UpdateLinksAtOpen','SaveNormalPrompt','ShowMarkupOpenSave') { $saved[$name]=$word.Options.$name }
 Reply $optionsPath $saved
 foreach($name in $saved.Keys) { $word.Options.$name=$false }
 Reply (Join-Path $Workspace 'ready.json') @{ok=$true; version=$word.Version; pid=$process.Id;existingWordCount=$before.Count}
 while (!(Test-Path -LiteralPath (Join-Path $Workspace 'stop'))) {
  if($parent -and $parent.HasExited) { Trace-Word 'parent process exited'; break }
  $jobFile=Join-Path $Workspace 'job.json'
  if (!(Test-Path -LiteralPath $jobFile)) { Start-Sleep -Milliseconds 10; continue }
    $job=Get-Content -LiteralPath $jobFile -Raw | ConvertFrom-Json
  Remove-Item -LiteralPath $jobFile
  $doc=$null; $audit=@{wordGuardChecks=0;wordGuardSeconds=0.0;wordGuardEvents=@()}; $timing=@{}; $clock=[Diagnostics.Stopwatch]::StartNew()
  try {
   $m=[Type]::Missing; $password=[Guid]::NewGuid().ToString()
   Assert-WordExportGuard $word $ownedPid $ownerStarted @($before.Id) $audit 'Before document open'
   Trace-Word 'opening working document'
   $doc=$word.Documents.Open($job.source,$false,$true,$false,$password,$password,$false,$password,$password,$m,$m,$true,$false)
   Assert-WordExportGuard $word $ownedPid $ownerStarted @($before.Id) $audit 'After document open' $false
   $doc.Activate()
   Trace-Word 'inspecting filename fields and saved display'
   Assert-WorkingFields $doc ([IO.Path]::GetFileName($job.source)) $job.originalName
   $timing.open=$clock.Elapsed.TotalSeconds; $clock.Restart()
   $initialAudit=@{commentsFound=$doc.Comments.Count; revisionsFound=(Get-StoryRevisionCount $doc); customPropertiesFound=$doc.CustomDocumentProperties.Count; commentsRemoved=0; revisionsAccepted=0; hiddenRangesRemoved=0; propertiesExcluded=[bool]$config.excludeProperties; tagsExcluded=[bool]$config.excludeTags}
   foreach($key in $initialAudit.Keys){$audit[$key]=$initialAudit[$key]}
   $inventory=[WordPdfPropertyInventory]::Read($doc)
   foreach($key in $inventory.Keys) { $audit[$key]=$inventory[$key] }
   $view=$doc.ActiveWindow.View
   $audit.savedMarkup=[int]$view.RevisionsFilter.Markup
   $audit.savedRevisionView=[int]$view.RevisionsFilter.View
   $audit.savedShowRevisions=[bool]$view.ShowRevisionsAndComments
   $audit.savedInsertionsDeletions=[bool]$view.ShowInsertionsAndDeletions
   $audit.savedFormatting=[bool]$view.ShowFormatChanges
   $audit.savedShowComments=[bool]$view.ShowComments
   $audit.reviewReasons=@()
   $audit.visibleRevisionsPreserved=$false; $audit.visibleCommentsPreserved=$false
   $hiddenRevisions=(!$audit.savedShowRevisions -or $audit.savedMarkup -eq 0 -or (!$audit.savedInsertionsDeletions -and !$audit.savedFormatting))
   $discardAll=!$config.preserveVisibleRevisions
   $acceptHidden=$audit.revisionsFound -gt 0 -and ($discardAll -or ($hiddenRevisions -and $audit.savedRevisionView -eq 0))
   $audit.revisionPolicy='No tracked changes'
   if($audit.revisionsFound -gt 0) {
    if($audit.savedRevisionView -ne 0 -and !$discardAll) { throw 'Document is saved in Original revision view. PDF withheld for manual review rather than accepting edits or changing that view.' }
    if($acceptHidden) {
     $audit.revisionPolicy='Hidden revisions accepted in export copy'
     if($discardAll) { $audit.revisionPolicy='All revisions accepted; history discarded by selected policy' }
    } else {
     $audit.revisionPolicy='Saved revision display preserved'
     $audit.visibleRevisionsPreserved=$true
     $audit.reviewReasons+='Tracked changes preserved using saved display settings; review visible redlines and reviewer labels.'
    }
   }
   $baselineItem=0
   if($audit.savedShowRevisions -and $audit.savedMarkup -ne 0) { $baselineItem=7 }
   $savedVisibleComments=$audit.savedShowRevisions -and $audit.savedShowComments -and $audit.savedMarkup -ne 0
   if($discardAll) {
    $view.RevisionsFilter.View=0
    $view.ShowInsertionsAndDeletions=$false;$view.ShowFormatChanges=$false
    if(!$config.removeAllComments -and $savedVisibleComments) {
     $view.ShowRevisionsAndComments=$true;$view.ShowComments=$true;$view.RevisionsFilter.Markup=$audit.savedMarkup;$baselineItem=7
    } else { $view.RevisionsFilter.Markup=0;$baselineItem=0 }
   }
   Trace-Word 'preparing comment-free reference' @{comments_found=$audit.commentsFound;revisions_found=$audit.revisionsFound;saved_markup=$audit.savedMarkup;saved_revision_view=$audit.savedRevisionView;revision_policy=$audit.revisionPolicy}
   # Removing visible comments is an explicit policy choice. Reference the
   # comment-free view while retaining the selected tracked-change display.
   if($doc.TrackRevisions -and $doc.ProtectionType -eq -1) { $doc.TrackRevisions=$false }
   if($config.removeAllComments -and $audit.commentsFound -gt 0) {
    $doc.DeleteAllComments()
    if($doc.Comments.Count -ne 0) { throw 'Could not remove all comments; PDF withheld.' }
    $audit.commentsRemoved=$audit.commentsFound
   }
   Trace-Word 'exporting appearance reference'
   if($doc.ProtectionType -eq -1) { $doc.Repaginate() }
   Trace-Layout $doc 'reference layout ready'
   Assert-WordExportGuard $word $ownedPid $ownerStarted @($before.Id) $audit 'Before reference export'
   $doc.ExportAsFixedFormat(($job.output+'.baseline.pdf'),17,$false,0,0,1,1,$baselineItem,$false,$true,0,$false,$true,$false)
   $timing.reference_export=$clock.Elapsed.TotalSeconds; $clock.Restart()
    $referenceMarkedClean=$false
    try { $doc.Saved=$true;$referenceMarkedClean=[bool]$doc.Saved } catch { $referenceMarkedClean=$false }
   if($doc.TrackRevisions -and $doc.ProtectionType -eq -1) { $doc.TrackRevisions=$false }
   Trace-Word 'applying selected revision policy'
   if($acceptHidden) { Accept-StoryRevisions $doc }
   $visibleComments=$savedVisibleComments
   if(!$config.removeAllComments -and $visibleComments -and $doc.Comments.Count -lt $audit.commentsFound) { throw 'Accepting tracked edits would remove visible comments while comment removal is disabled. PDF withheld; original retained.' }
   if(!$config.removeAllComments -and $visibleComments -and $audit.commentsFound -gt 0) {
    $audit.visibleCommentsPreserved=$true
    $audit.reviewReasons+='Visible comments preserved to maintain saved appearance; review reviewer identities.'
   }
     $audit.commentsRemoved=$audit.commentsFound-$doc.Comments.Count
     if($acceptHidden) { $audit.revisionsAccepted=$audit.revisionsFound-(Get-StoryRevisionCount $doc) }
     else { $audit.revisionsAccepted=0 }
   # Explicitly delete or reveal hidden-font text in this unsaved document only.
   # This avoids depending on Word's global hidden-text export preference.
   Trace-Word 'preparing hidden-font text'
    $hasHiddenText=([bool]$config.hideText) -and (Test-StoryHiddenText $doc)
    $audit.hiddenSearchSkipped=([bool]$config.hideText) -and !$hasHiddenText
    if($hasHiddenText) { Prepare-HiddenText $doc $true $audit }
   $timing.prepare=$clock.Elapsed.TotalSeconds; $clock.Restart()
   $item=$baselineItem
   # PDF, print quality, all pages, no source properties by default. Preserve IRM.
    Assert-WordExportGuard $word $ownedPid $ownerStarted @($before.Id) $audit 'Before prepared export'
    $sameExportOptions=[bool]$config.excludeProperties -and [bool]$config.excludeTags -and $item -eq $baselineItem
    $noPostReferenceChanges=($audit.revisionsAccepted -eq 0 -and $audit.hiddenRangesRemoved -eq 0)
    $reuseReference=$referenceMarkedClean -and [bool]$doc.Saved -and $sameExportOptions -and $noPostReferenceChanges
    $audit.secondExportSkipped=$reuseReference
    if($reuseReference) {
     Trace-Word 'reusing unchanged appearance reference as prepared PDF'
     [IO.File]::Copy(($job.output+'.baseline.pdf'),$job.output,$false)
    } else {
     Trace-Word 'exporting prepared PDF'
     if($doc.ProtectionType -eq -1) { $doc.Repaginate() }
     Trace-Layout $doc 'prepared layout ready'
     Assert-WordExportGuard $word $ownedPid $ownerStarted @($before.Id) $audit 'Before prepared PDF publication'
     $doc.ExportAsFixedFormat($job.output,17,$false,0,0,1,1,$item,(!$config.excludeProperties),$true,0,(!$config.excludeTags),$true,$false)
    }
   $timing.export=$clock.Elapsed.TotalSeconds
   $reply=@{ok=$true; timings=$timing; audit=$audit}
  } catch {
   Trace-Word 'document error' @{exception=$_.Exception.ToString();position=$_.InvocationInfo.PositionMessage}
   $reply=@{ok=$false; error=($_.Exception.ToString() + ' at ' + $_.InvocationInfo.PositionMessage); timings=$timing; audit=$audit}
  } finally {
   $timing.word_guard_checks=[double]$audit.wordGuardSeconds
   Trace-Word 'Word export guard summary' @{checks=$audit.wordGuardChecks;seconds=$audit.wordGuardSeconds;events=$audit.wordGuardEvents}
   Trace-Word 'closing document without saving'
   if($doc) { $doc.Close(0); [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) }
  }
  Reply (Join-Path $Workspace 'result.json') $reply
 }
} catch { Trace-Word 'session error' @{exception=$_.Exception.ToString();position=$_.InvocationInfo.PositionMessage}; Reply (Join-Path $Workspace 'ready.json') @{ok=$false; error=$_.Exception.ToString()} }
finally {
 Trace-Word 'shutting down Word session'
 if($word -and $owned) {
  $restored=$true
  foreach($name in $saved.Keys) { try { $word.Options.$name=$saved[$name] } catch { $restored=$false;Trace-Word 'Word option restore failed' @{option=$name;error=$_.Exception.GetType().Name} } }
  if($restored -and (Test-Path -LiteralPath $optionsPath)) { Remove-Item -LiteralPath $optionsPath -ErrorAction SilentlyContinue }
  try {
   if($word.Documents.Count -gt 0) {
    $word.Visible=$true;$owned=$false
    if($process) { Reply (Join-Path $Workspace 'owner.json') @{pid=$process.Id;started=$process.StartTime.ToUniversalTime().ToString('o');released=$true} }
    Trace-Word 'Word left visible because documents remain open'
   } else { $word.Quit(0) }
  } catch { Trace-Word 'Word shutdown error' @{error=$_.Exception.GetType().Name} }
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
 } elseif($word) {
  # Never Quit or change settings on an instance whose ownership was not proven.
  [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
 }
 [GC]::Collect(); [GC]::WaitForPendingFinalizers()
 if($owned -and $process -and !$process.WaitForExit(2000)) { Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue;Wait-Process -Id $process.Id -Timeout 15 -ErrorAction SilentlyContinue }
 if($debugWriter) { $debugWriter.Dispose();$debugWriter=$null }
}
