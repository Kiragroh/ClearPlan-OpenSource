[CmdletBinding()]
param(
    [string]$Root = "",
    [string]$TestExecutable,
    [string]$PythonExecutable = "python"
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}
$resolvedRoot = [IO.Path]::GetFullPath($Root)
if ([string]::IsNullOrWhiteSpace($TestExecutable)) {
    $TestExecutable = Join-Path `
        $resolvedRoot "artifacts\bin\Release\ClearPlan.Core.Tests.exe"
}
$resolvedTestExecutable = [IO.Path]::GetFullPath($TestExecutable)
if (-not (Test-Path -LiteralPath $resolvedTestExecutable -PathType Leaf)) {
    throw "Core contract validator executable is missing: $resolvedTestExecutable"
}

$outputDirectory = Join-Path $resolvedRoot "artifacts\raystation-contract"
$outputPath = Join-Path $outputDirectory "fake-context-review.json"
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null

& $PythonExecutable `
    (Join-Path `
        $resolvedRoot `
        "examples\raystation\tests\export_contract_fixture.py") `
    $outputPath
if ($LASTEXITCODE -ne 0) {
    throw "RayStation fake-context export failed."
}

& $resolvedTestExecutable ValidateSnapshot $outputPath
if ($LASTEXITCODE -ne 0) {
    throw "The RayStation export did not satisfy the C# ReviewSnapshot contract."
}

Write-Host "PASS RayStation Python export -> C# snapshot contract validation."
