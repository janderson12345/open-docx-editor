# One full comparison with Word: build the demo page, print every document through it, extract the lines of both
# sets of PDFs, and write the report. Word's PDFs are made once (word-print.ps1) and kept; ours go to a new folder per
# run, so two runs can be compared (diff.py).
# usage: pwsh tools/word-compare/run.ps1 -Docs <folder of .docx> -Word <folder of Word's PDFs> -Out <work folder> -Run <name>
param(
  [Parameter(Mandatory)][string]$Docs,
  [Parameter(Mandatory)][string]$Word,
  [Parameter(Mandatory)][string]$Out,
  [Parameter(Mandatory)][string]$Run
)
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path "$PSScriptRoot\..\..").Path
New-Item -ItemType Directory -Force $Out | Out-Null
$Out = (Resolve-Path $Out).Path
# Each step has a time limit: a step that hangs stops the run instead of waiting forever.
function Step($exe, $argList, $minutes, $log) {
  $p = Start-Process -FilePath $exe -ArgumentList $argList -NoNewWindow -PassThru -RedirectStandardOutput "$log.out" -RedirectStandardError "$log.err" -WorkingDirectory $repo
  if (-not $p.WaitForExit($minutes * 60000)) { $p.Kill($true); throw "timeout: $exe $argList" }
  if ($p.ExitCode -ne 0) { Get-Content "$log.err" -Tail 20; throw "failed ($($p.ExitCode)): $exe $argList" }
}
Step 'node' @('build.mjs') 3 "$Out\build"
if (-not (Test-Path "$Out\lines-word")) { Step 'node' @('tools\word-compare\lines.cjs', $Word, "$Out\lines-word") 20 "$Out\lines-word" }
Step 'node' @('tools\word-compare\print-ours.mjs', $Docs, "$Out\pdf-$Run") 60 "$Out\print-$Run"
Step 'node' @('tools\word-compare\lines.cjs', "$Out\pdf-$Run", "$Out\lines-$Run") 20 "$Out\lines-$Run"
Step 'python' @('tools\word-compare\compare.py', "$Out\lines-word", "$Out\lines-$Run") 5 "$Out\compare-$Run"
Copy-Item "$Out\compare-$Run.out" "$Out\report-$Run.txt"
Get-Content "$Out\report-$Run.txt" -Tail 1
