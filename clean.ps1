param()

$targets = @(
    (Join-Path $PSScriptRoot "publish"),
    (Join-Path $PSScriptRoot "installer-output"),
    (Join-Path $PSScriptRoot "bin_verify_inspector"),
    (Join-Path $PSScriptRoot "obj_cli_runner"),
    (Join-Path $PSScriptRoot "obj_verify_inspector"),
    (Join-Path $PSScriptRoot "tmp")
)

Get-ChildItem -Path (Join-Path $PSScriptRoot "TranslatorApp") -Directory -Recurse -Force |
    Where-Object { $_.Name -in @("bin", "obj") } |
    ForEach-Object {
        Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }

foreach ($target in $targets) {
    if (Test-Path $target) {
        Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$toolCache = Join-Path $PSScriptRoot "tools\PdfBilingualInspector\tools"
if (Test-Path $toolCache) {
    Remove-Item -LiteralPath $toolCache -Recurse -Force -ErrorAction SilentlyContinue
}

$tmpDoc = Join-Path $PSScriptRoot "tmp-bilingual-check.docx"
if (Test-Path $tmpDoc) {
    Remove-Item -LiteralPath $tmpDoc -Force -ErrorAction SilentlyContinue
}

Write-Host "Clean completed."
