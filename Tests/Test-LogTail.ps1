$ErrorActionPreference='Stop'
$appRoot=Split-Path -Parent $PSScriptRoot
. (Join-Path $appRoot 'LogTail.ps1')
$folder=Join-Path ([IO.Path]::GetTempPath()) ('word-pdf-log-test-'+[Guid]::NewGuid())
[void][IO.Directory]::CreateDirectory($folder)
function Assert($condition,$message) { if (!$condition) { throw $message } }
function Append-Bytes($path,[byte[]]$bytes) {
 $stream=[IO.File]::Open($path,[IO.FileMode]::Append,[IO.FileAccess]::Write,[IO.FileShare]::ReadWrite)
 try { $stream.Write($bytes,0,$bytes.Length) } finally { $stream.Dispose() }
}
try {
 $path=Join-Path $folder 'utf8.log'
 $tail=New-Object WordPdfLogTail($path)
 Assert ($tail.ReadNew($false) -eq '') 'Missing file should be harmless'
 $bytes=[Text.Encoding]::UTF8.GetBytes("First line`n")
 Append-Bytes $path $bytes
 Assert ($tail.ReadNew($false) -eq "First line`n") 'Initial text differs'
 for($i=0;$i -lt 20;$i++) { Assert ($tail.ReadNew($false) -eq '') 'Old text was read again' }
 Assert ($tail.BytesRead -eq $bytes.Length) 'Polling reread existing bytes'
 Append-Bytes $path ([byte[]]@(0xe2,0x82))
 Assert ($tail.ReadNew($false) -eq '') 'Incomplete Unicode should wait'
 Append-Bytes $path ([byte[]]@(0xac,10))
 Assert ($tail.ReadNew($true) -eq ([char]0x20ac + "`n")) 'Split Unicode was corrupted'
 [IO.File]::WriteAllText($path,'new',[Text.Encoding]::UTF8)
 Assert ($tail.ReadNew($true) -eq 'new') 'Truncated file should restart'
 $path2=Join-Path $folder 'utf16.log'
 $tail2=New-Object WordPdfLogTail($path2)
 Append-Bytes $path2 ([byte[]]@(0xff))
 Assert ($tail2.ReadNew($false) -eq '') 'Partial BOM should wait'
 Append-Bytes $path2 ([byte[]]@(0xfe,65,0))
 Assert ($tail2.ReadNew($false) -eq 'A') 'UTF16 BOM detection failed'
 Append-Bytes $path2 ([byte[]]@(66))
 Assert ($tail2.ReadNew($false) -eq '') 'Partial UTF16 character should wait'
 Append-Bytes $path2 ([byte[]]@(0))
 Assert ($tail2.ReadNew($true) -eq 'B') 'Split UTF16 character failed'
 $large=Join-Path $folder 'large.log'
 [IO.File]::WriteAllText($large,('x'*150000),[Text.Encoding]::ASCII)
 $reader=New-Object WordPdfLogTail($large)
 $first=$reader.ReadNew($true)
 Assert ($first.Length -eq 65536 -and $reader.HasMore) 'Per-refresh limit failed'
 $total=$first.Length
 while($reader.HasMore) { $total += $reader.ReadNew($true).Length }
 Assert ($total -eq 150000) 'Final drain lost content'
 Assert ($reader.BytesRead -eq 150000) 'Log bytes were reread'
 'PASS: appended-only reads, shared files, Unicode splits, truncation, bounded refreshes and final drain.'
} finally {
 # Delete only the individual files created by this test.
 foreach($name in 'utf8.log','utf16.log','large.log') {
  $file=Join-Path $folder $name
  if([IO.File]::Exists($file)) { [IO.File]::Delete($file) }
 }
 [IO.Directory]::Delete($folder)
}
