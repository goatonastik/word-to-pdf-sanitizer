Add-Type -AssemblyName PresentationFramework
$ErrorActionPreference='Stop'
$appRoot=Split-Path -Parent $PSScriptRoot
$text=[IO.File]::ReadAllText((Join-Path $appRoot 'Interface.ps1'))
$start=$text.IndexOf('<Window xmlns=');$end=$text.IndexOf("`n'@",$start)
[xml]$xml=$text.Substring($start,$end-$start)
$w=[Windows.Markup.XamlReader]::Load((New-Object Xml.XmlNodeReader $xml))
$w.FindName('Root').Text='C:\Export\Sample Client'
$w.FindName('Archive').Text='C:\Export\#ORIGINALS'
$w.FindName('Destination').Text="Archive: C:\Export\#ORIGINALS\Sample Client\[same subfolders]`nReview: C:\Export\#REVIEW\Sample Client\[same subfolders]"
$root=$w.Content; $root.Background=$w.Background
foreach($size in @(@(1072,813),@(932,673))) {
 $root.Measure((New-Object Windows.Size($size[0],$size[1])))
 $root.Arrange((New-Object Windows.Rect(0,0,$size[0],$size[1])))
 $root.UpdateLayout()
 $scroll=$w.FindName('OptionsScroll')
 foreach($name in 'Progress','Status','Counts','Reports','OpenReport','OpenDetailedReport','Log','Start') {
  $control=$w.FindName($name)
  if($scroll.IsAncestorOf($control)) { throw "$name is scrollable" }
  $point=$control.TransformToAncestor($root).Transform((New-Object Windows.Point(0,0)))
  if($point.Y -lt 0 -or ($point.Y+$control.ActualHeight) -gt ($size[1]+1)) { throw "$name is clipped at $size" }
 }
 $progress=$w.FindName('Progress');$before=$progress.TransformToAncestor($root).Transform((New-Object Windows.Point(0,0))).Y
 $scroll.ScrollToBottom();$root.UpdateLayout()
 $after=$progress.TransformToAncestor($root).Transform((New-Object Windows.Point(0,0))).Y
 if($before -ne $after) { throw 'Progress moved with options' }
 # Changing filenames/counts must not resize the status rows or shift progress.
 $status=$w.FindName('Status'); $counts=$w.FindName('Counts')
 $status.Text='Ready'; $counts.Text='1 of 1'
 $root.UpdateLayout()
 $initialY=$progress.TransformToAncestor($root).Transform((New-Object Windows.Point(0,0))).Y
 $initialHeight=$status.ActualHeight; $countHeight=$counts.ActualHeight
 $status.Text='Converting C:\Export\'+('Long client folder and document name '*30)+'.docx'
 $counts.Text=('Converted: 1000 | Skipped: 1000 | Failed: 1000 | Review: 1000 '*10)
 $root.UpdateLayout()
 $longY=$progress.TransformToAncestor($root).Transform((New-Object Windows.Point(0,0))).Y
 if($initialY -ne $longY -or $status.ActualHeight -ne $initialHeight -or $counts.ActualHeight -ne $countHeight) { throw 'Long text shifted progress or enlarged status rows' }
 foreach($control in @($status,$counts)) {
  if($control.TextWrapping -ne 'NoWrap' -or !$control.ClipToBounds) { throw 'Status text must clip without wrapping' }
  if($control.ActualWidth -gt $root.ActualWidth) { throw 'Long text widened the layout' }
 }
 Write-Output 'PASS: long filenames/counts clip without shifting progress or adding scrollbars.'
 Write-Output ('PASS: fixed progress/status/log at client size '+$size+'; option scroll height '+$scroll.ActualHeight)
}

