
param (
    [string]$DocPath
)

$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open($DocPath)

foreach ($paragraph in $doc.Paragraphs) {

    $spacing = $paragraph.Format.SpaceAfter
    $content = $paragraph.Range.Text

    # Check if the paragraph's "after line spacing" is equal to the desired value
    if ($spacing -ne 0 -and $content -notlike "//*" -and $content -notlike "{*") {
        $newText = $content + "//`r`n"
        $paragraph.Range.Text = $newText
    }
}

$doc.Save()
$doc.Close()
$word.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null