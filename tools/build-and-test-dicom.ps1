[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",
    [ValidateSet("x64", "AnyCPU")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"
$dicomRepositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).ProviderPath
$dicomMsbuild = & (Join-Path $PSScriptRoot "resolve-msbuild.ps1")
if ([string]::IsNullOrWhiteSpace($dicomMsbuild) -or -not (Test-Path -LiteralPath $dicomMsbuild -PathType Leaf)) {
    throw "The MSBuild resolver did not return an existing executable."
}

# Lock files pin the exact direct/transitive versions and package content hashes.
# This step must run before the legacy solution build on a fresh checkout.
& $dicomMsbuild (Join-Path $dicomRepositoryRoot "ClearPlan.Dicom.Tests\ClearPlan.Dicom.Tests.csproj") `
    /restore /t:Build /p:RestoreLockedMode=true `
    "/p:Configuration=$Configuration" "/p:Platform=$Platform" /v:minimal /nologo
if ($LASTEXITCODE -ne 0) { throw "Locked RTPLAN dependency restore/build failed with exit code $LASTEXITCODE." }

$dicomTests = Join-Path $dicomRepositoryRoot "artifacts\dicom-tests\$Configuration\ClearPlan.Dicom.Tests.exe"
if (-not (Test-Path -LiteralPath $dicomTests -PathType Leaf)) { throw "RTPLAN test executable is missing after the build." }
& $dicomTests
if ($LASTEXITCODE -ne 0) { throw "RTPLAN synthetic DICOM tests failed with exit code $LASTEXITCODE." }
Write-Host "Locked RTPLAN build and synthetic-file tests passed. No clinical inputs were used."
