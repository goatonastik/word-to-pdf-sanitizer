$ErrorActionPreference='Stop'
$appRoot=Split-Path -Parent $PSScriptRoot
. (Join-Path $appRoot 'ProcessLauncher.ps1')
$folder=Join-Path ([IO.Path]::GetTempPath()) ('Word Launcher [test] '+[Guid]::NewGuid())
[void][IO.Directory]::CreateDirectory($folder)
$python=Join-Path $env:USERPROFILE '.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$scriptPath=Join-Path $folder 'sample [worker].py'
$output=Join-Path $folder 'Root [2026-11-09 14-00-00.000] - output FAILED.log'
$errors=Join-Path $folder 'Root [2026-11-09 14-00-00.000] - errors FAILED.log'
try {
 @'
import sys
print("worker started: \u20ac")
print("test diagnostic", file=sys.stderr)
print('QUEUE_RESULT: {"queueCompleted": true, "queueErrors": false}')
assert sys.argv[1] == "folder [with brackets] and spaces"
'@ | Set-Content -LiteralPath $scriptPath -Encoding UTF8
 $arguments='-X utf8 -u "'+$scriptPath+'" "folder [with brackets] and spaces"'
 $child=New-Object WordPdfChild($python,$arguments,$folder,$output,$errors)
 $child.WaitForExit()
 if($child.ExitCode -ne 0) { throw 'Worker failed' }
 if($child.LoggingError) { throw $child.LoggingError }
 if(!($child.QueueResultJson | ConvertFrom-Json).queueCompleted) { throw 'Canonical completion result was not captured.' }
 $text=[IO.File]::ReadAllText($output)
 if($text.Contains('QUEUE_RESULT:')) { throw 'Internal completion data leaked into the displayed log.' }
 if(!$text.Contains('worker started: '+[char]0x20ac)) { throw 'Unicode output was lost' }
 if(![IO.File]::ReadAllText($errors).Contains('test diagnostic')) { throw 'stderr was lost' }
 # Make sure completion closes the streams so final status renaming works.
 [IO.File]::Move($output,$output+'.done')
 [IO.File]::Move($errors,$errors+'.done')
 'PASS: bracketed paths, spaces, UTF-8 output, stderr and completed-log renaming.'
} finally {
 foreach($path in @($scriptPath,$output,$errors,($output+'.done'),($errors+'.done'))) {
  if([IO.File]::Exists($path)) { [IO.File]::Delete($path) }
 }
 [IO.Directory]::Delete($folder)
}
