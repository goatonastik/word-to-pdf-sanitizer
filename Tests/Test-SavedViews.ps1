param([string]$Folder)
$ErrorActionPreference='Stop'
if(@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count) { throw 'Close Word before testing.' }
$word=New-Object -ComObject Word.Application
$ownedProcess=@(Get-Process WINWORD)[0]
$word.Visible=$false; $word.DisplayAlerts=0
$old=$word.Options.ShowMarkupOpenSave
$word.Options.ShowMarkupOpenSave=$false
try {
 foreach($entry in @(@('Hidden',0,0),@('Visible',2,0),@('Simple',1,0),@('Original',0,1),@('CommentsVisible',2,0))) {
  $path=Join-Path $Folder ($entry[0]+'.docx')
  $m=[Type]::Missing; $password=[Guid]::NewGuid().ToString()
  $doc=$word.Documents.Open($path,$false,$false,$false,$password,$password,$false,$password,$password,$m,$m,$true,$false)
  $view=$doc.ActiveWindow.View
  $view.ShowRevisionsAndComments=$true
  $view.ShowInsertionsAndDeletions=$true
  $view.ShowFormatChanges=$true
  $view.ShowComments=($entry[0] -in @('Visible','CommentsVisible'))
  $view.RevisionsFilter.Markup=[int]$entry[1]
  $view.RevisionsFilter.View=[int]$entry[2]
  $doc.Saved=$false
  $doc.Save(); $noSave=0; $doc.Close([ref]$noSave)
 }
} finally {
 $word.Options.ShowMarkupOpenSave=$old
 try { $noSave=0; $word.Quit([ref]$noSave) } catch {}
 [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
 [GC]::Collect(); [GC]::WaitForPendingFinalizers()
 if(!$ownedProcess.WaitForExit(500)) { Stop-Process -Id $ownedProcess.Id -Force }
}
