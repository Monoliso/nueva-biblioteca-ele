
param (
    [string]$docPath
)

if (-not (Test-Path -PathType Leaf -Filter "*.rtf" $docPath)) {
    Write-Error "El archivo no existe o no se trata de un .rtf."
    return $false
}

$doc_abs_path = Resolve-Path $docPath # Path absoluto

$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open($doc_abs_path)


# Iterate through the paragraphs
for ($i = 0; $i -le $doc.Paragraphs.Count; $i++) {
    # Access the text content of the paragraph
    $paragraph = $doc.Paragraphs[$i]
    $paragraphText = $paragraph.Range.Text
    
    # Check if the paragraph starts with a number (indicating a new verse)
    if ($paragraphText -match "^\d+") {
        # Replace the new paragraph separator with a line break
        echo $paragraphText
        echo "------"
        $paragraph.Range.Text = "`v" + $paragraphText
    } else {
        echo "x"
        $paragraph.Range.Text = $paragraphText + "`v"
    }
}


# Save the modified document to a new file
# $doc.SaveAs([System.Object]$newFilePath, [ref]3)  # 3 represents RTF format

$doc.Save
$doc.Close()
$word.Quit()

# $doc_name = Split-Path $doc_abs_path -Leaf
# $doc_dir = Split-Path -Parent $doc_ab_path

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
