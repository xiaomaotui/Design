param(
  [Parameter(Mandatory=$true)][string]$InputDoc
)
$ErrorActionPreference = "Stop"
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $word.AutomationSecurity = 3
  $word.Options.UpdateLinksAtOpen = $false
  $doc = $word.Documents.Open($InputDoc, $false, $false)
  foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
  foreach ($tof in $doc.TablesOfFigures) { $tof.Update() }
  $doc.Fields.Update() | Out-Null
  $doc.Save()
  Write-Output $InputDoc
}
finally {
  if ($doc -ne $null) { $doc.Close($false) }
  if ($word -ne $null) { $word.Quit() }
}
