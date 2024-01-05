
param (
    [string]$docPath
)

if (-not (Test-Path -PathType Leaf -Filter "*.rtf" $docPath)) {
    Write-Error "El archivo no existe o no se trata de un .rtf."
    return $false
}

$doc_abs_path = Resolve-Path $docPath # Tipo: PathInfo

$wordApp = New-Object -ComObject Word.Application
$document = $wordApp.Documents.Open($doc_abs_path.Path)

# Perform the find and replace
$findText = "^p" # FUNCA
# $findPattern = "^(?![0-9])"
$replaceText = "^l"
$replaceAll = 2  # Constant for replace all

$wordApp.Selection.Find.Execute($findText, $false, $false, $false, $false, $false, $true, $false, $false, $replaceText, $replaceAll)



$document.Save
$document.Close()
$wordApp.Quit()

# $doc_name = Split-Path $doc_abs_path -Leaf
# $doc_dir = Split-Path -Parent $doc_ab_path

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
