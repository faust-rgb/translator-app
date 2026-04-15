$iscc = Get-Command ISCC -ErrorAction SilentlyContinue
if (-not $iscc) {
    Write-Host "未找到 Inno Setup 编译器 ISCC。请先安装 Inno Setup 后重试。"
    exit 1
}

powershell -ExecutionPolicy Bypass -File "$PSScriptRoot\\publish.ps1"
& $iscc.Source "$PSScriptRoot\\installer.iss"
