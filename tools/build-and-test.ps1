[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [ValidateSet("x64", "AnyCPU")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..")).ProviderPath
$msbuild = & (Join-Path $PSScriptRoot "resolve-msbuild.ps1")

if ([string]::IsNullOrWhiteSpace($msbuild) -or
    -not (Test-Path -LiteralPath $msbuild -PathType Leaf)) {
    throw "The MSBuild resolver did not return an existing executable."
}

& $msbuild `
    (Join-Path $root "ClearPlan.sln") `
    /t:Build `
    "/p:Configuration=$Configuration" `
    "/p:Platform=$Platform" `
    /m
if ($LASTEXITCODE -ne 0) {
    throw "ClearPlan solution build failed with exit code $LASTEXITCODE."
}

$testExecutable = Join-Path $root "artifacts\bin\$Configuration\ClearPlan.Core.Tests.exe"
if (-not (Test-Path -LiteralPath $testExecutable -PathType Leaf)) {
    throw "Core test executable is missing after the build: $testExecutable"
}

& $testExecutable
if ($LASTEXITCODE -ne 0) {
    throw "ClearPlan.Core.Tests failed with exit code $LASTEXITCODE."
}

$workbookValidator = Join-Path $PSScriptRoot "validate-constraint-workbook.py"
if (Test-Path -LiteralPath $workbookValidator -PathType Leaf) {
    & python $workbookValidator
    if ($LASTEXITCODE -ne 0) {
        throw "Constraint workbook validation failed with exit code $LASTEXITCODE."
    }
}
else {
    Write-Host "Constraint workbook validation skipped until the validator is introduced."
}

$conversionTests = Join-Path $PSScriptRoot "test-constraint-csv-conversion.py"
& python $conversionTests
if ($LASTEXITCODE -ne 0) {
    throw "Constraint CSV conversion tests failed with exit code $LASTEXITCODE."
}

$rayStationTests = Join-Path $root "examples\raystation\tests"
& python -m unittest discover -s $rayStationTests -v
if ($LASTEXITCODE -ne 0) {
    throw "Illustrative RayStation adapter tests failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "test-raystation-contract.ps1") `
    -Root $root `
    -TestExecutable $testExecutable
if ($LASTEXITCODE -ne 0) {
    throw "RayStation-to-C# contract test failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-public-release.ps1") -Root $root
if ($LASTEXITCODE -ne 0) {
    throw "Public release validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-vendor-free-assemblies.ps1") -Root $root
if ($LASTEXITCODE -ne 0) {
    throw "Vendor-free assembly validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-ui-distinctiveness.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "UI distinctiveness validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-readonly-field-preview.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "Read-only field preview validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-runner-skin.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "Essentials runner skin validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-runner-resilience.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "Runner path resilience validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-pqm-mousemove-safety.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "PQM mouse initialization safety validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-simulator-actions.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "Simulator action validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-responsive-review.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "Responsive review validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-clinical-plancheck-build.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "Clinical PlanCheck build provenance validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-clinical-review-snapshot-builder.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "Clinical review snapshot builder validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-clinical-shared-workspace.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "Clinical shared workspace validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-network-source-bounds.ps1")
if ($LASTEXITCODE -ne 0) {
    throw "Network source bounds validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "validate-pqm-transaction-safety.ps1")
if ($LASTEXITCODE -ne 0) {
    throw "PQM transaction safety validation failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "test-public-simulator-package-input.ps1") `
    -Root $root
if ($LASTEXITCODE -ne 0) {
    throw "Simulator package-input tests failed with exit code $LASTEXITCODE."
}

& (Join-Path $PSScriptRoot "test-settings-ini-integration.ps1") -RepositoryRoot $root
if ($LASTEXITCODE -ne 0) {
    throw "Settings INI integration failed with exit code $LASTEXITCODE."
}

if ($Configuration -eq "Release") {
    $simulatorDirectory = Join-Path $root "artifacts\simulator\Release"
    & (Join-Path $PSScriptRoot "test-simulator.ps1") `
        -SimulatorDirectory $simulatorDirectory
    if ($LASTEXITCODE -ne 0) {
        throw "Simulator smoke test failed with exit code $LASTEXITCODE."
    }

    & python (Join-Path $PSScriptRoot "collect-paper-evidence.py")
    if ($LASTEXITCODE -ne 0) {
        throw "Paper evidence collection failed with exit code $LASTEXITCODE."
    }

    & python (Join-Path $PSScriptRoot "build-paper-figures.py")
    if ($LASTEXITCODE -ne 0) {
        throw "Paper figure build failed with exit code $LASTEXITCODE."
    }

    & python (Join-Path $PSScriptRoot "build-technical-note-docx.py")
    if ($LASTEXITCODE -ne 0) {
        throw "Technical-note DOCX build failed with exit code $LASTEXITCODE."
    }

    & (Join-Path $PSScriptRoot "validate-release-version.ps1") -Root $root
    if ($LASTEXITCODE -ne 0) {
        throw "Release version validation failed with exit code $LASTEXITCODE."
    }
}
else {
    & (Join-Path $PSScriptRoot "validate-release-version.ps1") `
        -Root $root `
        -SkipBinaries
    if ($LASTEXITCODE -ne 0) {
        throw "Source version validation failed with exit code $LASTEXITCODE."
    }
}

Write-Host "ClearPlan build and test pipeline passed."
