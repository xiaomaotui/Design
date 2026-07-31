param(
  [Parameter(Mandatory=$true)][string]$InputDoc,
  [Parameter(Mandatory=$true)][string]$OutputPdf
)
$ErrorActionPreference = "Stop"
$outDir = Split-Path -Parent $OutputPdf
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.AutomationSecurity = 3
  $word.Options.UpdateLinksAtOpen = $false
  $doc = $word.Documents.Open($InputDoc, $false, $true)
  $doc.ExportAsFixedFormat($OutputPdf, 17, $false, 0, 0, 1, 999, 0, $true, $true, 0, $true, $true, $false)
  Write-Output $OutputPdf
}
finally {
  if ($doc -ne $null) { $doc.Close($false) }
  if ($word -ne $null) { $word.Quit() }
}
