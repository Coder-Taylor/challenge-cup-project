$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath '医影知微_科技说明_初版.docx').Path
$outputFolder = Join-Path (Get-Location).Path '.tmp_pdf\docx_render'
New-Item -ItemType Directory -Path $outputFolder -Force | Out-Null
$outputPath = Join-Path $outputFolder '医影知微_科技说明_初版.pdf'
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($inputPath, $false, $true)
    $doc.Repaginate()
    Write-Output ('word_count=' + $doc.ComputeStatistics(0))
    Write-Output ('far_east_characters=' + $doc.ComputeStatistics(6))
    Write-Output ('pages=' + $doc.ComputeStatistics(2))
    $doc.ExportAsFixedFormat($outputPath, 17)
    Write-Output ('pdf=' + $outputPath)
} finally {
    if ($null -ne $doc) { $doc.Close(0) }
    if ($null -ne $word) { $word.Quit() }
}
