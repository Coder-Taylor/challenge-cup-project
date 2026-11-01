<#
.SYNOPSIS
    文件用途：用本机 Microsoft Word 复核 DOCX 的字数和页数，并导出用于检查分页的 PDF。
.DESCRIPTION
    输入：项目根目录的“医影知微_科技说明_初版.docx”。
    输出：控制台打印 Word 字数、东亚字符数和页数；PDF 写入 .tmp_pdf/docx_render/。
    依赖：Windows、本机 Microsoft Word 及可用的 Word COM 会话。
    使用位置：从项目根目录调用 tools/word_qa.ps1，路径按当前工作目录解析。
    执行影响：只读打开 DOCX，关闭时不保存；同名检查 PDF 会重新导出。
    当前状态：只检查已有 Word 文件；不会将最新 Markdown 转为 Word。
#>

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
