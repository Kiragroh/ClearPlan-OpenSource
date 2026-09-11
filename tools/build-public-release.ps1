[CmdletBinding()]
param(
    [string]$Root = "",
    [string]$Version = "3.2.0",
    [switch]$SkipSimulatorSmoke,
    [string]$PythonExecutable = "python",
    [string]$EvidenceOutputPath = "",
    [switch]$DevelopmentPreview
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}
$resolvedRoot = [IO.Path]::GetFullPath($Root)
$msbuild = & (Join-Path $PSScriptRoot "resolve-msbuild.ps1")
if ([string]::IsNullOrWhiteSpace($EvidenceOutputPath)) {
    $EvidenceOutputPath = Join-Path $resolvedRoot 'artifacts\public-release-test-evidence.json'
}
elseif (-not [IO.Path]::IsPathRooted($EvidenceOutputPath)) {
    $EvidenceOutputPath = Join-Path $resolvedRoot $EvidenceOutputPath
}
Push-Location -LiteralPath $resolvedRoot
try {
& $PythonExecutable -c 'import openpyxl'
if ($LASTEXITCODE -ne 0) { throw 'Python with openpyxl is required; pass -PythonExecutable explicitly if needed.' }
& $PythonExecutable (Join-Path $PSScriptRoot 'collect-paper-evidence.py') --self-test
if ($LASTEXITCODE -ne 0) { throw 'Execution-evidence parser tests failed.' }

# Optional RTPLAN enrichment has its own locked vendor-free dependency closure.
# Build/test it separately; it is deliberately not added to the simulator ZIP.
& (Join-Path $resolvedRoot 'tools\build-and-test-dicom.ps1') -Configuration Release -Platform x64
if ($LASTEXITCODE -ne 0) { throw 'Locked synthetic DICOM build/tests failed.' }

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
# This executes (not merely counts) Core, synthetic DICOM, RayStation and stock
# suites, checks the default workbook and inventories both distributed XLSX files.
# Missing private stock intermediates produce explicit skipped-group evidence.
& $PythonExecutable (Join-Path $resolvedRoot 'tools\collect-paper-evidence.py') `
    --configuration Release --release-version $Version --include-dicom --output $EvidenceOutputPath
if ($LASTEXITCODE -ne 0) {
    throw "Current software-execution evidence could not be collected."
}

& (Join-Path $PSScriptRoot "test-raystation-contract.ps1") `
    -Root $resolvedRoot `
    -TestExecutable $testExecutable `
    -PythonExecutable $PythonExecutable
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
    -PublicBinariesOnly `
    -IncludeDicomBinaries `
    -ExecutionEvidencePath $EvidenceOutputPath `
    -DevelopmentPreview:$DevelopmentPreview
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
else {
    Write-Warning 'SKIPPED simulator UI/report smoke; this invocation is not a complete release verification.'
}

$determinismRoot = Join-Path `
    $resolvedRoot "artifacts\package-determinism"
$packageOutputs = @(
    (Join-Path $determinismRoot "first"),
    (Join-Path $determinismRoot "second")
)
$packageHashes = @()
$packageArchiveName = "ClearPlan-Simulator-v$Version-win-x64.zip"
if ($DevelopmentPreview) {
    $packageArchiveName = "ClearPlan-Simulator-v$Version-development-preview-win-x64.zip"
}
foreach ($output in $packageOutputs) {
    & (Join-Path $PSScriptRoot "package-public-simulator.ps1") `
        -Root $resolvedRoot `
        -Version $Version `
        -PythonExecutable $PythonExecutable `
        -Preview:$DevelopmentPreview `
        -ExecutionEvidencePath $EvidenceOutputPath `
        -OutputDirectory $output | Out-Host
    $archivePath = Join-Path `
        $output $packageArchiveName
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
    -Version $Version `
    -PythonExecutable $PythonExecutable `
    -Preview:$DevelopmentPreview `
    -ExecutionEvidencePath $EvidenceOutputPath | Out-Host

Write-Host "PASS vendor-free build, recorded software suites and deterministic packaging of this build."
Write-Host "Executed counts and explicit skips: $EvidenceOutputPath"
Write-Host 'Repeated packaging is not independent clean-build reproducibility; final publication/artifact review remains separate.'
if ($DevelopmentPreview) { Write-Warning 'DevelopmentPreview: manuscript/release evidence gates were not completed.' }
Write-Host "Package SHA256: $($packageHashes[0].ToLowerInvariant())"
}
finally {
    Pop-Location
}
