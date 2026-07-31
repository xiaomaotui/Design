$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$path = (Get-ChildItem -LiteralPath $root -Recurse -File -Filter '*2026-08-01.docx' |
    Where-Object { $_.Length -gt 250000 } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1).FullName
if (-not $path) { throw 'Target DOCX not found.' }
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $document = $word.Documents.Open($path, $false, $false)
    foreach ($toc in $document.TablesOfContents) {
        $toc.Update()
    }
    foreach ($story in $document.StoryRanges) {
        $range = $story
        while ($null -ne $range) {
            if ($range.Fields.Count -gt 0) {
                $range.Fields.Update() | Out-Null
            }
            $range = $range.NextStoryRange
        }
    }
    $document.Save()
    $document.Close()
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
