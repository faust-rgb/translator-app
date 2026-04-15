param()

$targets = @(
    (Join-Path $PSScriptRoot "publish"),
    (Join-Path $PSScriptRoot "installer-output")
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

Write-Host "Clean completed."
