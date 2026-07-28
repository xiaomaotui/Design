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
    $sourceDate = $doc.Paragraphs.Item(18).Range.Text.TrimEnd("`r", [char]7)
    if ([string]::IsNullOrWhiteSpace($sourceDate)) {
        $sourceDate = $doc.Paragraphs.Item(16).Range.Text.TrimEnd("`r", [char]7)
    }
    if ([string]::IsNullOrWhiteSpace($sourceDate)) {
        for ($i = 1; $i -le $doc.Shapes.Count; $i++) {
            $candidate = $doc.Shapes.Item($i)
            if ($candidate.Name -eq "SubmissionDateTextBox") {
                $sourceDate = $candidate.TextFrame.TextRange.Text.TrimEnd("`r", [char]7)
                break
            }
        }
    }
    $targetRange = $doc.Paragraphs.Item(16).Range.Duplicate
    $targetRange.End = $targetRange.End - 1
    $targetRange.Text = ""
    $oldDateRange = $doc.Paragraphs.Item(18).Range.Duplicate
    $oldDateRange.End = $oldDateRange.End - 1
    $oldDateRange.Text = ""
    foreach ($index in @(17, 18)) {
        $paragraph = $doc.Paragraphs.Item($index)
        $paragraph.Format.SpaceBefore = 0
        $paragraph.Format.SpaceAfter = 0
        $paragraph.Format.LineSpacingRule = 0
        $paragraph.Range.Font.Size = 1
    }
    $dateParagraph = $doc.Paragraphs.Item(16)
    $dateParagraph.Format.SpaceBefore = 0
    $dateParagraph.Format.SpaceAfter = 0
    $dateParagraph.Format.LineSpacingRule = 0
    $dateParagraph.Range.Font.Size = 1
    foreach ($index in 9..15) {
        $doc.Paragraphs.Item($index).Format.SpaceBefore = 0
    }
    for ($i = $doc.Shapes.Count; $i -ge 1; $i--) {
        $shape = $doc.Shapes.Item($i)
        if ($shape.Name -eq "SubmissionDateTextBox") {
            $shape.Delete()
        }
    }
    $anchor = $doc.Paragraphs.Item(1).Range
    $box = $doc.Shapes.AddTextbox(1, 198, 795, 200, 16, $anchor)
    $box.Name = "SubmissionDateTextBox"
    $box.RelativeHorizontalPosition = 1
    $box.RelativeVerticalPosition = 1
    $box.Left = 198
    $box.Top = 795
    $box.WrapFormat.Type = 3
    $box.Line.Visible = 0
    $box.Fill.Visible = 0
    $box.TextFrame.TextRange.Text = $sourceDate
    $box.TextFrame.TextRange.ParagraphFormat.Alignment = 1
    $box.TextFrame.TextRange.Font.Name = "Times New Roman"
    $box.TextFrame.TextRange.Font.Size = 10.5
    $box.TextFrame.MarginTop = 0
    $box.TextFrame.MarginBottom = 0
    $doc.Save()
    Write-Output "front-layout-fixed"
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
