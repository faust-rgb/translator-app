param(
    [string]$Runtime = "win-x64"
)

$project = Join-Path $PSScriptRoot "TranslatorApp\\TranslatorApp.csproj"
$output = Join-Path $PSScriptRoot "publish"
$tessdata = Join-Path $PSScriptRoot "tessdata"

dotnet publish $project `
  -c Release `
  -r $Runtime `
  --self-contained true `
  /p:PublishSingleFile=true `
  /p:IncludeNativeLibrariesForSelfExtract=true `
  -o $output

if (Test-Path $tessdata) {
    $target = Join-Path $output "tessdata"
    New-Item -ItemType Directory -Force -Path $target | Out-Null
    Copy-Item -Path (Join-Path $tessdata "*") -Destination $target -Recurse -Force
}
