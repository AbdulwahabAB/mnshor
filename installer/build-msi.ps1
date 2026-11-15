param(
    [string]$Version = "0.1.0",
    [ValidateSet("alpha", "preview", "beta", "rc", "stable")]
    [string]$Stage = "alpha",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"

$RootDir = Resolve-Path (Join-Path $PSScriptRoot "..")
$InstallerDir = Join-Path $PSScriptRoot "Mnshor.Installer"
$PublishDir = Join-Path $InstallerDir "obj\publish\app"
$LauncherDir = Join-Path $InstallerDir "obj\publish\launcher"
$DistDir = Join-Path $RootDir "dist"

$StageLabel = switch ($Stage) {
    "alpha"   { "Alpha" }
    "preview" { "Preview" }
    "beta"    { "Beta" }
    "rc"      { "RC" }
    "stable"  { "" }
}

$ProductName = if ($Stage -eq "stable") { "Mnshor" } else { "Mnshor ($StageLabel)" }
$MsiFileName = if ($Stage -eq "stable") {
    "Mnshor-$Version-win-x64.msi"
} else {
    "Mnshor-$Version-$Stage-win-x64.msi"
}
$MsiOutputPath = Join-Path $DistDir $MsiFileName

Write-Host "=== Building $ProductName ($Version) ===" -ForegroundColor Cyan

if (-not $SkipTests) {
    Write-Host "[1/4] Running unit tests..." -ForegroundColor Yellow
    dotnet test (Join-Path $RootDir "Mnshor.sln") -c Release
    if ($LASTEXITCODE -ne 0) { throw "Tests failed." }
}

Write-Host "[2/4] Publishing Mnshor (win-x64 self-contained)..." -ForegroundColor Yellow
if (Test-Path $PublishDir) { Remove-Item $PublishDir -Recurse -Force }
dotnet publish (Join-Path $RootDir "src\Mnshor\Mnshor.csproj") `
    -c Release -r win-x64 --self-contained true `
    -p:Version=$Version `
    -o $PublishDir
if ($LASTEXITCODE -ne 0) { throw "Mnshor publish failed." }

Write-Host "[3/4] Publishing Mnshor.Launcher (win-x64 single-file)..." -ForegroundColor Yellow
if (Test-Path $LauncherDir) { Remove-Item $LauncherDir -Recurse -Force }
dotnet publish (Join-Path $RootDir "src\Mnshor.Launcher\Mnshor.Launcher.csproj") `
    -c Release -r win-x64 --self-contained true `
    -p:PublishSingleFile=true `
    -p:Version=$Version `
    -o $LauncherDir
if ($LASTEXITCODE -ne 0) { throw "Mnshor.Launcher publish failed." }

Write-Host "[4/4] Building MSI with WiX..." -ForegroundColor Yellow
New-Item -ItemType Directory -Force -Path $DistDir | Out-Null

Push-Location $InstallerDir
try {
    wix build Package.wxs `
        -arch x64 `
        -ext WixToolset.UI.wixext `
        -d ProductName="$ProductName" `
        -d ProductVersion="$Version" `
        -d PublishDir="$PublishDir" `
        -d LauncherDir="$LauncherDir" `
        -o $MsiOutputPath
    if ($LASTEXITCODE -ne 0) { throw "WiX build failed." }
}
finally {
    Pop-Location
}

Write-Host "=== MSI Ready: $MsiOutputPath ===" -ForegroundColor Green
