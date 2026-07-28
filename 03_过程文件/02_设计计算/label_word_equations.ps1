param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath
)

$ErrorActionPreference = "Stop"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($DocumentPath)
    $count = 0
    for ($p = 1; $p -le $doc.Paragraphs.Count; $p++) {
        $paragraph = $doc.Paragraphs.Item($p)
        if ($paragraph.Range.OMaths.Count -eq 0) {
            continue
        }
        $count++
        $section = $paragraph.Range.Sections.Item(1)
        $usableWidth = $section.PageSetup.PageWidth - $section.PageSetup.LeftMargin - $section.PageSetup.RightMargin
        $paragraph.Format.TabStops.ClearAll()
        [void]$paragraph.Format.TabStops.Add($usableWidth / 2, 1, 0)
        [void]$paragraph.Format.TabStops.Add($usableWidth, 2, 0)
        $paragraph.Format.FirstLineIndent = 0
        $paragraph.Format.LeftIndent = 0
        $paragraph.Format.RightIndent = 0

        $insertAt = $paragraph.Range.End - 1
        $labelRange = $doc.Range($insertAt, $insertAt)
        $label = "`t" + [char]0xFF08 + [char]0x5F0F + $count + [char]0xFF09
        $labelRange.InsertAfter($label)
    }

    if ($count -ne 41) {
        throw "Expected 41 equations, found $count."
    }
    $doc.Save()
    Write-Output "labeled=$count"
}
finally {
    if ($doc) {
        $doc.Close()
    }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
