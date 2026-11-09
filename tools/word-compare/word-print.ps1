# Has Word print each .docx in a folder to PDF: opened read-only, closed without saving. Changes no settings.
# Documents already printed are skipped, so a rerun picks up where it stopped. Each document gets a time limit.
# usage: pwsh wordpdf.ps1 <docxDir> <pdfDir> [secondsPerDoc]
param([Parameter(Mandatory)][string]$In, [Parameter(Mandatory)][string]$Out, [int]$Seconds = 120)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $Out | Out-Null
$In = (Resolve-Path $In).Path; $Out = (Resolve-Path $Out).Path
$files = Get-ChildItem $In -Filter *.docx | Where-Object { -not (Test-Path (Join-Path $Out ($_.BaseName + '.pdf'))) }
foreach ($f in $files) {
  $pdf = Join-Path $Out ($f.BaseName + '.pdf')
  # One Word per document, in its own job, so a document that hangs Word can be abandoned after the time limit.
  $job = Start-Job -ArgumentList $f.FullName, $pdf -ScriptBlock {
    param($doc, $pdf)
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false; $w.DisplayAlerts = 0
    try {
      $d = $w.Documents.Open($doc, $false, $true, $false)
      $d.ExportAsFixedFormat($pdf, 17)
      $d.Close(0)
    } finally { $w.Quit(0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w) }
  }
  if (Wait-Job $job -Timeout $Seconds) { Receive-Job $job -ErrorAction SilentlyContinue | Out-Null; if ($job.State -eq 'Failed') { "FAILED $($f.Name): $($job.ChildJobs[0].JobStateInfo.Reason.Message)" } }
  else { "TIMEOUT $($f.Name)"; Stop-Job $job }
  Remove-Job $job -Force
}
"printed $((Get-ChildItem $Out -Filter *.pdf).Count)"
