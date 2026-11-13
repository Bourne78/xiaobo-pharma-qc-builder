# export_pdf.ps1
# ---------------------------------------------------------------
# ★ 按环境分支（本机实测）：Word COM 自动化被沙箱拦截，WINWORD.EXE 无法启动、
# phantom v12 注册返回 E_FAIL。本脚本当前不可执行；DOCX→PDF 须由用户在 Word
# 手动「另存为 / 导出为 PDF」，可选 WPS 或 soffice --convert-to pdf 兜底。
# 脚本保留为「环境修复后可用」的实现，勿当作当前交付步骤。
# 详见 references/python_docx_pitfalls.md §10 / §12。
# ---------------------------------------------------------------
# 用 Word COM 把 .docx 导出为 .pdf。
# 关键：导出前必须 $doc.Fields.Update()，否则页码域（PAGE / NUMPAGES）
# 会显示为 0 或 1。Word COM 的 Quit() 偶发 RPC 异常无害，已 try/finally 兜底。
#
# 用法（PowerShell）：
#   .\export_pdf.ps1 -DocxPath "手册.docx"
#   .\export_pdf.ps1 -DocxPath "手册.docx" -PdfPath "手册.pdf"
# ---------------------------------------------------------------
param(
    [Parameter(Mandatory=$true)]
    [string]$DocxPath,
    [string]$PdfPath = ""
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $DocxPath)) {
    Write-Error "找不到文件: $DocxPath"
    exit 1
}

if ($PdfPath -eq "") {
    $PdfPath = [System.IO.Path]::ChangeExtension($DocxPath, "pdf")
}

# 绝对路径，避免 Word 相对路径解析到错误目录
$DocxPath = Resolve-Path $DocxPath
$PdfPath  = [System.IO.Path]::GetFullPath($PdfPath)

$word = $null
$doc  = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    # Open(FileName, ConfirmConversions, ReadOnly)
    $doc = $word.Documents.Open($DocxPath, $false, $true)
    # 必须先更新域，页码/PAGE/NUMPAGES 才正确
    $doc.Fields.Update() | Out-Null
    # ExportAsFixedFormat(FileName, ExportFormat=17=wdExportFormatPDF)
    $doc.ExportAsFixedFormat($PdfPath, 17)
    Write-Output "PDF_EXPORTED: $PdfPath"
}
catch {
    Write-Error "导出失败: $_"
    exit 2
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close($false) } catch { }
    }
    if ($word -ne $null) {
        try { $word.Quit() } catch { }   # Quit 偶发 RPC 异常，忽略
    }
}
