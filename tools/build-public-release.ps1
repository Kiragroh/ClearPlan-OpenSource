[CmdletBinding()]
param(
    [string]$Root = "",
    [string]$Version = "3.1.0",
    [switch]$SkipSimulatorSmoke
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}
$resolvedRoot = [IO.Path]::GetFullPath($Root)
$msbuild = & (Join-Path $PSScriptRoot "resolve-msbuild.ps1")

foreach ($target in @(
    "ClearPlan_Core_Tests",
    "ClearPlan_Simulator"
)) {
    & $msbuild `
        (Join-Path $resolvedRoot "ClearPlan.sln") `
        "/t:$target" `
        "/p:Configuration=Release" `
        "/p:Platform=x64" `
        /m
    if ($LASTEXITCODE -ne 0) {
        throw "Public Release build target '$target' failed."
    }
}

$testExecutable = Join-Path `
    $resolvedRoot "artifacts\bin\Release\ClearPlan.Core.Tests.exe"
& $testExecutable
if ($LASTEXITCODE -ne 0) {
    throw "ClearPlan.Core.Tests failed."
}

& python (Join-Path $PSScriptRoot "validate-constraint-workbook.py")
if ($LASTEXITCODE -ne 0) {
    throw "Public workbook validation failed."
}

& python -m unittest discover `
    -s (Join-Path $resolvedRoot "examples\raystation\tests") `
    -v
if ($LASTEXITCODE -ne 0) {
    throw "Illustrative RayStation adapter tests failed."
}

& (Join-Path $PSScriptRoot "test-raystation-contract.ps1") `
    -Root $resolvedRoot `
    -TestExecutable $testExecutable
if ($LASTEXITCODE -ne 0) {
    throw "RayStation-to-C# contract test failed."
}

& (Join-Path $PSScriptRoot "validate-public-release.ps1") `
    -Root $resolvedRoot
if ($LASTEXITCODE -ne 0) {
    throw "Public-tree privacy validation failed."
}

& (Join-Path $PSScriptRoot "validate-vendor-free-assemblies.ps1") `
    -Root $resolvedRoot
if ($LASTEXITCODE -ne 0) {
    throw "Vendor-free assembly validation failed."
}

& (Join-Path $PSScriptRoot "test-public-simulator-package-input.ps1") `
    -Root $resolvedRoot
if ($LASTEXITCODE -ne 0) {
    throw "Simulator package-input tests failed."
}

& (Join-Path $PSScriptRoot "validate-release-version.ps1") `
    -Root $resolvedRoot `
    -ExpectedVersion $Version `
    -PublicBinariesOnly
if ($LASTEXITCODE -ne 0) {
    throw "Public release version validation failed."
}

if (-not $SkipSimulatorSmoke) {
    & (Join-Path $PSScriptRoot "test-simulator.ps1") `
        -SimulatorDirectory (
            Join-Path $resolvedRoot "artifacts\simulator\Release")
    if ($LASTEXITCODE -ne 0) {
        throw "Public simulator smoke test failed."
    }
}

$determinismRoot = Join-Path `
    $resolvedRoot "artifacts\package-determinism"
$packageOutputs = @(
    (Join-Path $determinismRoot "first"),
    (Join-Path $determinismRoot "second")
)
$packageHashes = @()
foreach ($output in $packageOutputs) {
    & (Join-Path $PSScriptRoot "package-public-simulator.ps1") `
        -Root $resolvedRoot `
        -Version $Version `
        -OutputDirectory $output | Out-Host
    $archivePath = Join-Path `
        $output "ClearPlan-Simulator-v$Version-win-x64.zip"
    $packageHashes += (
        Get-FileHash `
            -LiteralPath $archivePath `
            -Algorithm SHA256).Hash
}
if (-not [string]::Equals(
    $packageHashes[0],
    $packageHashes[1],
    [StringComparison]::OrdinalIgnoreCase)) {
    throw "Repeated packaging did not produce a byte-identical archive."
}

& (Join-Path $PSScriptRoot "package-public-simulator.ps1") `
    -Root $resolvedRoot `
    -Version $Version | Out-Host

Write-Host "PASS public Release build and deterministic package."
Write-Host "Package SHA256: $($packageHashes[0].ToLowerInvariant())"
