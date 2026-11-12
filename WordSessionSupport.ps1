# Shared Word safety checks; never save the opened document.
if (-not ('WordPdfProcessIdentity' -as [type])) {
 Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class WordPdfProcessIdentity {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
 public static int ForWindow(int window) { uint process; GetWindowThreadProcessId(new IntPtr(window),out process); return (int)process; }
}
'@
}
function Get-StoryRevisionCount($document) {
 $count=0
 foreach($first in $document.StoryRanges) { $story=$first; while($null -ne $story) { $count+=$story.Revisions.Count; $story=$story.NextStoryRange } }
 return $count
}
function Accept-StoryRevisions($document) {
 foreach($first in $document.StoryRanges) { $story=$first; while($null -ne $story) { if($story.Revisions.Count -gt 0) { $story.Revisions.AcceptAll() }; $story=$story.NextStoryRange } }
}
$script:DynamicFieldTypes=@{21='CREATEDATE';22='SAVEDATE';30='TEMPLATE';69='FILESIZE';23='PRINTDATE';31='DATE';32='TIME';60='USERNAME';61='USERINITIALS';62='USERADDRESS'}
function Assert-LinkedShape($shape,$directory) {
 $children=$null
 if($shape.Type -eq 6) { try { $children=$shape.GroupItems } catch {} }
 if($children) { foreach($child in $children) { Assert-LinkedShape $child $directory } }
 $link=$null;$source=$null;$cached=$false
 try { $link=$shape.LinkFormat;$source=[string]$link.SourceFullName } catch { return }
 if([string]::IsNullOrWhiteSpace($source)) { return }
 try { $cached=[bool]$link.SavePictureWithDocument } catch {}
 if($cached) { return }
 try {
  $resolved=if([IO.Path]::IsPathRooted($source)) { $source } else { Join-Path $directory $source }
  $available=[IO.File]::Exists($resolved)
 } catch { $available=$false }
 if(!$available) { throw 'Uncached linked content is unavailable from the TEMP working copy. PDF withheld to prevent changed or missing visible content.' }
}
function Assert-WorkingFields($document,$workingName,$originalName) {
 foreach($first in $document.StoryRanges) {
  $story=$first
  while($null -ne $story) {
   foreach($field in $story.Fields) {
    if($script:DynamicFieldTypes.ContainsKey([int]$field.Type) -and !$field.Locked) { throw ('Unfrozen '+$script:DynamicFieldTypes[[int]$field.Type]+' field cannot be guaranteed to retain its saved display. PDF withheld; original retained.') }
    if([int]$field.Type -eq 29 -and !$field.Locked) {
     if($field.Code.Text -match '(?i)(^|\s)\\p(\s|$)') { throw 'Full-path FILENAME field would change in the TEMP working copy. PDF withheld to preserve original visible text.' }
     if($workingName -ne $originalName) { throw 'FILENAME field requires the original filename; shortened TEMP names cannot preserve its display. PDF withheld.' }
    }

   }
   foreach($shape in $story.InlineShapes) { Assert-LinkedShape $shape $document.Path }
   $floating=@();try { $floating=@($story.ShapeRange) } catch {}
   foreach($shape in $floating) { Assert-LinkedShape $shape $document.Path }
   $story=$story.NextStoryRange
  }
 }
}
function Prepare-HiddenText($document,$remove,$audit) {
 if(!$remove) { return } # Preserve original hidden formatting; never reveal hidden text.
 foreach($first in $document.StoryRanges) {
  $story=$first
  while($null -ne $story) {
   $iterations=0;$limit=[Math]::Max(1,$story.Characters.Count+1)
   $search=$story.Duplicate;$find=$search.Find;$find.ClearFormatting();$find.Text='';$find.Font.Hidden=-1;$find.Format=$true;$find.Wrap=0
   while($find.Execute()) {
    $iterations++
    if($iterations -gt $limit) { throw 'Hidden-text preparation exceeded its progress limit. PDF withheld; original retained.' }
    if($remove) {
     $length=$story.End-$story.Start
     [void]$search.Delete()
     if(($story.End-$story.Start) -ge $length) { throw 'Word cannot remove a hidden structural character without changing the document. PDF withheld; original retained.' }
     $audit.hiddenRangesRemoved++
    } else { $search.Font.Hidden=0 }
    $search=$story.Duplicate;$find=$search.Find;$find.ClearFormatting();$find.Text='';$find.Font.Hidden=-1;$find.Format=$true;$find.Wrap=0
   }
   $story=$story.NextStoryRange
  }
 }
}
function Test-StoryHiddenText($document) {
 foreach($first in $document.StoryRanges) {
  $story=$first
  while($null -ne $story) {
   # 0 means every character in this story is visibly formatted. -1 means
   # hidden and Word's mixed-value sentinel means at least one hidden range.
   if([int]$story.Font.Hidden -ne 0) { return $true }
   $story=$story.NextStoryRange
  }
 }
 return $false
}
