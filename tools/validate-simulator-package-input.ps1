[CmdletBinding()]
param(
    [string]$Root = "",
    [Parameter(Mandatory = $true)]
    [string]$ScenarioDirectory
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}
$resolvedRoot = [IO.Path]::GetFullPath($Root)
$resolvedScenarioDirectory =
    [IO.Path]::GetFullPath($ScenarioDirectory)
$sourceDirectory = Join-Path `
    $resolvedRoot "ClearPlan.Simulator\Scenarios"

if (-not (Test-Path `
    -LiteralPath $resolvedScenarioDirectory `
    -PathType Container)) {
    throw "Scenario package input does not exist: $resolvedScenarioDirectory"
}
if (-not (Test-Path `
    -LiteralPath $sourceDirectory `
    -PathType Container)) {
    throw "Checked-in scenario source does not exist: $sourceDirectory"
}

$sourceFiles = @(
    Get-ChildItem `
        -LiteralPath $sourceDirectory `
        -File `
        -Filter "*.json" |
        Sort-Object Name
)
$candidateFiles = @(
    Get-ChildItem `
        -LiteralPath $resolvedScenarioDirectory `
        -File `
        -Filter "*.json" |
        Sort-Object Name
)

if ($sourceFiles.Count -ne 7 -or
    $candidateFiles.Count -ne 7) {
    throw "The public simulator package must contain exactly seven scenarios."
}

$nameDifferences = @(
    Compare-Object `
        -ReferenceObject $sourceFiles.Name `
        -DifferenceObject $candidateFiles.Name `
        -CaseSensitive
)
if ($nameDifferences.Count -gt 0) {
    throw (
        "Scenario names differ from the checked-in release set: " +
        (($nameDifferences |
          ForEach-Object {
              $_.InputObject + " " + $_.SideIndicator
          }) -join ", "))
}

$forbiddenPatterns = @(
    @{ Label = "UNC path"; Pattern = '\\\\[^\\\r\n"]+\\' },
    @{ Label = "drive path"; Pattern = '(?i)\b[A-Z]:\\' },
    @{ Label = "DICOM UID"; Pattern = '\b(?:\d+\.){4,}\d+\b' },
    @{
        Label = "email address"
        Pattern = '(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b'
    }
)

foreach ($candidateFile in $candidateFiles) {
    $sourcePath = Join-Path $sourceDirectory $candidateFile.Name
    $sourceHash = (
        Get-FileHash `
            -LiteralPath $sourcePath `
            -Algorithm SHA256).Hash
    $candidateHash = (
        Get-FileHash `
            -LiteralPath $candidateFile.FullName `
            -Algorithm SHA256).Hash
    if (-not [string]::Equals(
        $sourceHash,
        $candidateHash,
        [StringComparison]::OrdinalIgnoreCase)) {
        throw (
            "Packaged scenario differs from checked-in source: " +
            $candidateFile.Name)
    }

    $raw = Get-Content -LiteralPath $candidateFile.FullName -Raw
    try {
        $scenario = $raw | ConvertFrom-Json
    }
    catch {
        throw (
            "Scenario is not valid JSON: " +
            $candidateFile.Name +
            ". " +
            $_.Exception.Message)
    }

    if ($scenario.synthetic -ne $true) {
        throw (
            "Scenario is not explicitly synthetic: " +
            $candidateFile.Name)
    }
    if (-not [string]::Equals(
        [IO.Path]::GetFileNameWithoutExtension(
            $candidateFile.Name),
        [string]$scenario.scenarioId,
        [StringComparison]::Ordinal)) {
        throw (
            "Scenario ID does not match its file name: " +
            $candidateFile.Name)
    }
    if (-not [string]::Equals(
        [string]$scenario.patientDisplayLabel,
        "Synthetic demonstration",
        [StringComparison]::Ordinal) -or
        [string]::IsNullOrWhiteSpace(
            [string]$scenario.provenanceText) -or
        ([string]$scenario.provenanceText).IndexOf(
            "no clinical source data",
            [StringComparison]::OrdinalIgnoreCase) -lt 0) {
        throw (
            "Scenario synthetic provenance is incomplete: " +
            $candidateFile.Name)
    }

    foreach ($entry in $forbiddenPatterns) {
        if ($raw -match $entry.Pattern) {
            throw (
                "Scenario contains a forbidden " +
                $entry.Label +
                ": " +
                $candidateFile.Name)
        }
    }
}

Write-Host (
    "Simulator package input validation passed: " +
    $candidateFiles.Count +
    " byte-identical synthetic scenarios.")
