# Has one Word, started for this run and used only by it, print a list of .docx files to PDF: each opened read-only
# and invisibly, exported, closed without saving. No Word setting is changed. Progress goes to a status file, line by
# line (PID <word's process id>, START <name>, DONE <name>, FAIL <name> <why>, END), which run.mjs watches: a document
# that takes too long is abandoned by ending this script and the Word it started, never any other Word.
# usage: pwsh tools/doc-generator/word-pdf.ps1 -List <file of .docx paths> -Out <pdf folder> -Status <status file>
param([Parameter(Mandatory)][string]$List, [Parameter(Mandatory)][string]$Out, [Parameter(Mandatory)][string]$Status)
$ErrorActionPreference = 'Stop'
Add-Type -Namespace OdxGen -Name Win -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);'
function Log([string]$s) { [IO.File]::AppendAllText($Status, $s + [Environment]::NewLine) }
New-Item -ItemType Directory -Force $Out | Out-Null

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$w.DisplayAlerts = 0
$wordPid = 0
try {
  foreach ($doc in [IO.File]::ReadAllLines($List)) {
    if (-not $doc) { continue }
    $name = [IO.Path]::GetFileNameWithoutExtension($doc)
    $pdf = Join-Path $Out "$name.pdf"
    if (Test-Path -LiteralPath $pdf) { continue }
    Log "START $name"
    $d = $null
    try {
      # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles
      $d = $w.Documents.Open($doc, $false, $true, $false)
      if (-not $wordPid) {
        # The process of this Word, from its window: what the watcher may end if a document hangs it.
        try { $p = [uint32]0; [void][OdxGen.Win]::GetWindowThreadProcessId([IntPtr][long]$d.ActiveWindow.Hwnd, [ref]$p); $wordPid = $p; Log "PID $wordPid" } catch { Log "PID 0" ; $wordPid = -1 }
      }
      $d.ExportAsFixedFormat($pdf, 17)
      $d.Close(0)
      Log "DONE $name"
    } catch {
      Log ("FAIL $name " + ($_.Exception.Message -replace '\s+', ' '))
      try { if ($d) { $d.Close(0) } } catch {}
      # A Word that has stopped answering can't print the rest: end, and the watcher starts another.
      try { [void]$w.Version } catch { Log 'DEAD'; exit 3 }
    }
  }
} finally {
  try { $w.Quit(0) } catch {}
  try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch {}
}
Log 'END'
