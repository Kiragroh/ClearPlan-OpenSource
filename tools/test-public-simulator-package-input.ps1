[CmdletBinding()]
param(
    [string]$Root = ""
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}
$resolvedRoot = [IO.Path]::GetFullPath($Root)
$validator = Join-Path $PSScriptRoot "validate-simulator-package-input.ps1"
$sourceDirectory = Join-Path $resolvedRoot "ClearPlan.Simulator\Scenarios"
$temporaryRoot = Join-Path (
    [IO.Path]::GetTempPath()) (
    "clearplan-package-input-" + [Guid]::NewGuid().ToString("N"))

try {
    New-Item -ItemType Directory -Path $temporaryRoot | Out-Null
    Get-ChildItem `
        -LiteralPath $sourceDirectory `
        -File `
        -Filter "*.json" |
        Copy-Item -Destination $temporaryRoot

    & $validator `
        -Root $resolvedRoot `
        -ScenarioDirectory $temporaryRoot

    $tamperedPath = Join-Path $temporaryRoot "mixed-review.json"
    $tampered = (
        Get-Content -LiteralPath $tamperedPath -Raw).Replace(
            '"synthetic": true',
            '"synthetic": false')
    Set-Content `
        -LiteralPath $tamperedPath `
        -Value $tampered `
        -Encoding UTF8

    $rejected = $false
    try {
        & $validator `
            -Root $resolvedRoot `
            -ScenarioDirectory $temporaryRoot
    }
    catch {
        $rejected = $true
    }
    if (-not $rejected) {
        throw "A tampered non-synthetic scenario was accepted."
    }
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot -PathType Container) {
        $resolvedTemporary = [IO.Path]::GetFullPath($temporaryRoot)
        $expectedPrefix = [IO.Path]::GetFullPath(
            [IO.Path]::GetTempPath()).TrimEnd("\") +
            "\clearplan-package-input-"
        if (-not $resolvedTemporary.StartsWith(
            $expectedPrefix,
            [StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to remove an unexpected package-test directory."
        }
        Remove-Item -LiteralPath $resolvedTemporary -Recurse -Force
    }
}

Write-Host "Public simulator package-input tests passed."
