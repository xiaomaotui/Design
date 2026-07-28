$ErrorActionPreference = "Stop"

$matches = @(Get-ChildItem -LiteralPath (Get-Location).Path -Recurse -File -Filter "*2026-07-28.docx")
if ($matches.Count -ne 1) {
    throw "Expected one dated draft, found $($matches.Count)."
}
$docx = $matches[0].FullName
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($docx)

    $replacements = @(
        @("39 573.19", "39 572.62"),
        @("24 732.99", "24 732.88"),
        @("39,573.19", "39,572.62"),
        @("24,732.99", "24,732.88"),
        @("123.93", "124.10"),
        @("6 984", "6 985")
    )

    foreach ($pair in $replacements) {
        $range = $doc.Content.Duplicate
        $find = $range.Find
        $find.ClearFormatting()
        $find.Replacement.ClearFormatting()
        $find.Text = $pair[0]
        $find.Replacement.Text = $pair[1]
        [void]$find.Execute(
            $pair[0], $false, $false, $false, $false, $false,
            $true, 1, $false, $pair[1], 2
        )
    }

    # Remove the second sentence from the initial-rain paragraph to avoid a one-line orphan page.
    $rainRange = $doc.Content.Duplicate
    $rainFind = $rainRange.Find
    $rainFind.Text = "202.5"
    if ($rainFind.Execute()) {
        $paragraphRange = $rainRange.Paragraphs.Item(1).Range
        $paragraphText = $paragraphRange.Text
        $firstStop = $paragraphText.IndexOf([char]0x3002)
        $valueStart = $paragraphText.IndexOf("202.5")
        if (($firstStop -ge 0) -and ($valueStart -ge 0)) {
            $prefix = $paragraphText.Substring(0, 4) + [char]0x91CF + [char]0x4E3A
            $valueAndResult = $paragraphText.Substring($valueStart, $firstStop - $valueStart + 1)
            $keptText = $prefix + $valueAndResult
            $paragraphRange.End = $paragraphRange.End - 1
            $paragraphRange.Text = $keptText
        }
    }

    # The 2026 document is used only as the typography/layout template.
    # All visible headers must identify the student's actual 2027 cohort.
    foreach ($section in $doc.Sections) {
        foreach ($header in $section.Headers) {
            if ($header.Exists) {
                $headerRange = $header.Range
                if ($headerRange.Text -like "*2026*") {
                    $headerRange.Text = $headerRange.Text.Replace("2026", "2027")
                }
            }
        }
    }
    foreach ($story in $doc.StoryRanges) {
        $storyRange = $story
        while ($null -ne $storyRange) {
            if (($storyRange.StoryType -ge 6) -and ($storyRange.StoryType -le 11) -and ($storyRange.Text -like "*2026*")) {
                $storyRange.Text = $storyRange.Text.Replace("2026", "2027")
            }
            $storyRange = $storyRange.NextStoryRange
        }
    }

    foreach ($toc in $doc.TablesOfContents) {
        $toc.Update()
    }
    $doc.Fields.Update() | Out-Null
    $doc.Save()
    $doc.Close()
}
finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

Write-Output "final consistency fixes applied"
