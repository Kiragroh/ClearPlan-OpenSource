[CmdletBinding()]
param(
    [string]$Root = "",
    [string]$Version = "3.2.0",
    [string]$OutputDirectory,
    [switch]$Preview,
    [string]$PythonExecutable = "python",
    [string]$ExecutionEvidencePath = ""
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}
$resolvedRoot = [IO.Path]::GetFullPath($Root)
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $resolvedRoot "artifacts\release"
}
$resolvedOutput = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $resolvedOutput -Force | Out-Null

& (Join-Path $PSScriptRoot "validate-release-version.ps1") `
    -Root $resolvedRoot `
    -ExpectedVersion $Version `
    -PublicBinariesOnly `
    -ExecutionEvidencePath $ExecutionEvidencePath `
    -DevelopmentPreview:$Preview

& (Join-Path $PSScriptRoot "validate-vendor-free-assemblies.ps1") `
    -Root $resolvedRoot
if ($LASTEXITCODE -ne 0) {
    throw "Vendor-free assembly validation failed with exit code $LASTEXITCODE."
}

$simulatorDirectory = Join-Path `
    $resolvedRoot "artifacts\simulator\Release"
$requiredRuntimeFiles = @(
    "ClearPlan.Simulator.exe",
    "ClearPlan.Core.dll",
    "ClearPlan.Rendering.dll",
    "ClearPlan.Presentation.dll",
    "ClearPlan.Reporting.dll",
    "ClearPlan.Reporting.MigraDoc.dll",
    "MigraDoc.DocumentObjectModel.dll",
    "MigraDoc.Rendering.dll",
    "Newtonsoft.Json.dll",
    "OxyPlot.dll",
    "OxyPlot.Wpf.dll",
    "PdfSharp.Charting.dll",
    "PdfSharp.dll"
)
foreach ($fileName in $requiredRuntimeFiles) {
    $path = Join-Path $simulatorDirectory $fileName
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Simulator package input is missing: $path"
    }
}

$forbiddenFiles = Get-ChildItem -LiteralPath $simulatorDirectory -Recurse -File |
    Where-Object {
        $_.Extension -eq ".pdb" -or
        $_.Name -match "^(VMS\.|EsapiEssentials)" -or
        $_.FullName -match "[\\/](VMS\.|EsapiEssentials)"
    }
if (@($forbiddenFiles).Count -gt 0) {
    throw (
        "Forbidden proprietary or symbol file found in simulator output: " +
        (($forbiddenFiles | Select-Object -ExpandProperty FullName) -join ", ")
    )
}

$scenarioDirectory = Join-Path $simulatorDirectory "Scenarios"
$scenarioFiles = @(
    Get-ChildItem -LiteralPath $scenarioDirectory -File -Filter "*.json"
)
& (Join-Path $PSScriptRoot "validate-simulator-package-input.ps1") `
    -Root $resolvedRoot `
    -ScenarioDirectory $scenarioDirectory

$stageName = ".stage-" + [Guid]::NewGuid().ToString("N")
$stageRoot = [IO.Path]::GetFullPath((Join-Path $resolvedOutput $stageName))
if (-not $stageRoot.StartsWith(
    $resolvedOutput.TrimEnd("\") + "\",
    [StringComparison]::OrdinalIgnoreCase)) {
    throw "Refusing to create a staging directory outside the release output."
}

$archiveName = "ClearPlan-Simulator-v$Version-win-x64.zip"
if ($Preview) { $archiveName = "ClearPlan-Simulator-v$Version-development-preview-win-x64.zip" }
$archivePath = Join-Path $resolvedOutput $archiveName
$hashPath = $archivePath + ".sha256"

try {
    New-Item -ItemType Directory -Path $stageRoot | Out-Null
    foreach ($fileName in $requiredRuntimeFiles) {
        Copy-Item `
            -LiteralPath (Join-Path $simulatorDirectory $fileName) `
            -Destination (Join-Path $stageRoot $fileName)
    }

    $stageScenarios = Join-Path $stageRoot "Scenarios"
    New-Item -ItemType Directory -Path $stageScenarios | Out-Null
    foreach ($scenarioFile in $scenarioFiles) {
        Copy-Item `
            -LiteralPath $scenarioFile.FullName `
            -Destination (Join-Path $stageScenarios $scenarioFile.Name)
    }
    Copy-Item `
        -LiteralPath (Join-Path `
            $resolvedRoot "ClearPlan.Simulator\Scenarios\README.md") `
        -Destination (Join-Path $stageScenarios "README.md")
    & (Join-Path $PSScriptRoot "validate-simulator-package-input.ps1") `
        -Root $resolvedRoot `
        -ScenarioDirectory $stageScenarios

    foreach ($relativePath in @(
        "README.md",
        "LICENSE",
        "CITATION.cff",
        "THIRD-PARTY-NOTICES.md"
    )) {
        Copy-Item `
            -LiteralPath (Join-Path $resolvedRoot $relativePath) `
            -Destination (Join-Path $stageRoot $relativePath)
    }

    if ($Preview) {
        Copy-Item -LiteralPath (Join-Path $resolvedRoot "docs\releases\v$Version.md") `
            -Destination (Join-Path $stageRoot 'START-HERE.md')
    }

    & $PythonExecutable `
        (Join-Path $PSScriptRoot "create-deterministic-zip.py") `
        --source $stageRoot `
        --output $archivePath
    if ($LASTEXITCODE -ne 0) {
        throw "Deterministic ZIP creation failed with exit code $LASTEXITCODE."
    }

    $stream = [IO.File]::Open(
        $archivePath,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        [IO.FileShare]::Read)
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try {
        $hashBytes = $algorithm.ComputeHash($stream)
        $hash = ([BitConverter]::ToString($hashBytes)).Replace("-", "").ToLowerInvariant()
    }
    finally {
        $algorithm.Dispose()
        $stream.Dispose()
    }
    Set-Content `
        -LiteralPath $hashPath `
        -Value ("{0}  {1}" -f $hash, $archiveName) `
        -Encoding ASCII
}
finally {
    if (Test-Path -LiteralPath $stageRoot -PathType Container) {
        $resolvedStage = [IO.Path]::GetFullPath($stageRoot)
        if ($resolvedStage.StartsWith(
            $resolvedOutput.TrimEnd("\") + "\.stage-",
            [StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $resolvedStage -Recurse -Force
        }
        else {
            throw "Refusing to remove an unexpected staging directory."
        }
    }
}

$archive = Get-Item -LiteralPath $archivePath
Write-Host "PASS public simulator package"
Write-Host "Archive: $($archive.FullName)"
Write-Host "Bytes: $($archive.Length)"
Write-Host "SHA256: $hash"
Write-Output $archivePath
Write-Output $hashPath
