param(
    [string]$Runtime = "win-x64",
    [string]$Output = "publish-singlefile"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$project = Join-Path $PSScriptRoot "TranslatorApp\TranslatorApp.csproj"
$projectDirectory = Join-Path $PSScriptRoot "TranslatorApp"
$outputPath = Join-Path $PSScriptRoot $Output
$projectObjPath = Join-Path $projectDirectory "obj"
$projectBinPath = Join-Path $projectDirectory "bin"

if (Test-Path $outputPath) {
    Remove-Item -LiteralPath $outputPath -Recurse -Force
}

if (Test-Path $projectObjPath) {
    Remove-Item -LiteralPath $projectObjPath -Recurse -Force
}

if (Test-Path $projectBinPath) {
    Remove-Item -LiteralPath $projectBinPath -Recurse -Force
}

dotnet publish $project `
  -c Release `
  -r $Runtime `
  --self-contained true `
  /p:PublishSingleFile=true `
  /p:IncludeNativeLibrariesForSelfExtract=true `
  /p:DebugSymbols=false `
  /p:DebugType=None `
  -o $outputPath

if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed with exit code $LASTEXITCODE."
}

Write-Host ""
Write-Host "单文件发布完成：" -ForegroundColor Green
Write-Host "  $outputPath\TranslatorApp.exe"
Write-Host ""
Write-Host "说明：" -ForegroundColor Yellow
Write-Host "  1. 此输出不再额外携带 appsettings.json。"
Write-Host "  2. 此输出默认不打包 tessdata；程序可独立启动。"
Write-Host "  3. 如果要对扫描版 PDF 使用 OCR，请在界面里指定外部 tessdata 目录。"
