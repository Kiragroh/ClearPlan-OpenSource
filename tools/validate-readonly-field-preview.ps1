[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"
$files = @(
    "ClearPlan.Core/Fields/FieldNameSuggester.cs",
    "ClearPlan.Script/Calculators/FieldNamingPreviewCalculator.cs"
)
$forbiddenPatterns = @(
    "BeginModifications\s*\(",
    "ApplyParameters\s*\(",
    "SaveModifications\s*\(",
    "AddSetupBeam\s*\(",
    "\bbeam\.Id\s*=",
    "\bbeam\.Name\s*="
)
$failed = $false

foreach ($relativePath in $files) {
    $path = Join-Path $RepositoryRoot $relativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        Write-Error "Missing read-only field-preview source: $relativePath"
        $failed = $true
        continue
    }

    $content = Get-Content -LiteralPath $path -Raw
    foreach ($pattern in $forbiddenPatterns) {
        if ($content -match $pattern) {
            Write-Error "$relativePath contains forbidden mutation pattern: $pattern"
            $failed = $true
        }
    }
}

if ($failed) {
    exit 1
}

Write-Host "Read-only field preview validation passed."
