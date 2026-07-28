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
    $converted = 0

    for ($i = $doc.Paragraphs.Count; $i -ge 1; $i--) {
        $paragraph = $doc.Paragraphs.Item($i)
        $raw = $paragraph.Range.Text.TrimEnd("`r", [char]7)
        if ($raw -match '^\[\[EQ\|(\d+)\|(.+)\]\]$') {
            $number = [int]$Matches[1]
            $linear = $Matches[2]

            $contentRange = $paragraph.Range.Duplicate
            $contentRange.End = $contentRange.End - 1
            $contentRange.Text = $linear
            $paragraph.Format.FirstLineIndent = 0
            $paragraph.Format.LeftIndent = 0
            $paragraph.Format.RightIndent = 0
            $paragraph.Format.TabStops.ClearAll()

            $section = $paragraph.Range.Sections.Item(1)
            $usableWidth = $section.PageSetup.PageWidth - $section.PageSetup.LeftMargin - $section.PageSetup.RightMargin
            [void]$paragraph.Format.TabStops.Add($usableWidth / 2, 1, 0)
            [void]$paragraph.Format.TabStops.Add($usableWidth, 2, 0)

            $formulaRange = $paragraph.Range.Duplicate
            $formulaRange.End = $formulaRange.End - 1
            [void]$doc.OMaths.Add($formulaRange)
            $formulaRange.OMaths.Item(1).BuildUp()
            $converted++
        }
    }

    $doc.Save()
    Write-Output "converted=$converted"
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
