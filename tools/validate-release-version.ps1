[CmdletBinding()]
param(
    [string]$Root = "",
    [string]$ExpectedVersion = "3.2.0",
    [switch]$SkipBinaries,
    [switch]$PublicBinariesOnly,
    [switch]$DevelopmentPreview,
    [string]$ExecutionEvidencePath = "",
    [switch]$IncludeDicomBinaries
)

$ErrorActionPreference = "Stop"
if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = Split-Path -Parent $PSScriptRoot
}
$Root = [IO.Path]::GetFullPath($Root)
$expectedFileVersion = $ExpectedVersion + ".0"
$escapedVersion = [regex]::Escape($ExpectedVersion)
$escapedFileVersion = [regex]::Escape($expectedFileVersion)

function Assert-FileContains {
    param(
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [Parameter(Mandatory = $true)][string]$Pattern,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $path = Join-Path $Root $RelativePath
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Release version validation is missing $RelativePath."
    }
    $text = Get-Content -LiteralPath $path -Raw
    if ($text -notmatch $Pattern) {
        throw "$RelativePath does not declare $Description."
    }
}

Assert-FileContains `
    -RelativePath "CITATION.cff" `
    -Pattern "(?m)^version:\s*`"$escapedVersion`"\s*$" `
    -Description "CFF version $ExpectedVersion"
Assert-FileContains `
    -RelativePath "CITATION.cff" `
    -Pattern ([regex]::Escape(
        "https://github.com/Kiragroh/ClearPlan-OpenSource/releases/tag/v$ExpectedVersion")) `
    -Description "the v$ExpectedVersion release URL"
Assert-FileContains `
    -RelativePath "ClearPlan.Script\Distribution\CHANGELOG.md" `
    -Pattern "(?m)^##\s+\[$escapedFileVersion\]" `
    -Description "changelog version $expectedFileVersion"
Assert-FileContains `
    -RelativePath "README.md" `
    -Pattern ([regex]::Escape("releases/tag/v$ExpectedVersion")) `
    -Description "the v$ExpectedVersion release URL"
# A registration count is an inventory, never proof that tests ran. README has
# no hardcoded pass-count requirement; final counts come from fresh run evidence.
if (-not $DevelopmentPreview) {
    Assert-FileContains `
        -RelativePath "paper\ClearPlan_ZMP_short_communication.md" `
        -Pattern ([regex]::Escape("releases/tag/v$ExpectedVersion")) `
        -Description "the manuscript release URL"
}
if (-not $DevelopmentPreview -or -not [string]::IsNullOrWhiteSpace($ExecutionEvidencePath)) {
    if ([string]::IsNullOrWhiteSpace($ExecutionEvidencePath)) {
        $ExecutionEvidencePath = Join-Path $Root "paper\evidence\technical_note_evidence.json"
    }
    elseif (-not [IO.Path]::IsPathRooted($ExecutionEvidencePath)) {
        $ExecutionEvidencePath = Join-Path $Root $ExecutionEvidencePath
    }
    if (-not (Test-Path -LiteralPath $ExecutionEvidencePath -PathType Leaf)) {
        throw "Executed technical-note evidence is missing; collect it from the current build first."
    }
    $evidence = Get-Content -LiteralPath $ExecutionEvidencePath -Raw | ConvertFrom-Json
    if ($evidence.schemaVersion -lt 2 -or [string]$evidence.repositoryRelease -ne "v$ExpectedVersion") {
        throw "Evidence must use the executed-run schema and target v$ExpectedVersion."
    }
    $suiteNames = @('csharpTests', 'illustrativeRayStationAdapterTests', 'stockCatalogTests')
    if (Test-Path -LiteralPath (Join-Path $Root 'ClearPlan.Dicom\ClearPlan.Dicom.csproj')) {
        $suiteNames += 'syntheticDicomTests'
    }
    foreach ($suiteName in $suiteNames) {
        $suite = $evidence.softwareVerification.$suiteName
        if ($null -eq $suite -or $null -eq $suite.exitCode -or $suite.exitCode -ne 0 -or
            $null -eq $suite.failures -or $suite.failures -ne 0 -or $suite.total -le 0 -or
            $suite.passed -le 0 -or $suite.total -ne ($suite.passed + $suite.skippedTests) -or
            [string]$suite.status -notin @('passed', 'passed-with-skips') -or
            [string]$suite.outputSha256 -notmatch '^[a-f0-9]{64}$' -or
            [string]::IsNullOrWhiteSpace([string]$suite.startedUtc) -or
            [string]::IsNullOrWhiteSpace([string]$suite.completedUtc)) {
            throw "Missing or unreconciled execution evidence for $suiteName."
        }
        if ([DateTimeOffset]::Parse($suite.completedUtc) -lt [DateTimeOffset]::Parse($suite.startedUtc) -or
            [DateTimeOffset]::Parse($evidence.generatedUtc) -lt [DateTimeOffset]::Parse($suite.completedUtc)) {
            throw "Invalid execution chronology for $suiteName."
        }
        foreach ($property in $suite.inputSha256.PSObject.Properties) {
            $inputPath = [IO.Path]::GetFullPath((Join-Path $Root $property.Name))
            if (-not $inputPath.StartsWith($Root.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase) -or
                -not (Test-Path -LiteralPath $inputPath -PathType Leaf) -or
                (Get-FileHash -LiteralPath $inputPath -Algorithm SHA256).Hash -ne $property.Value) {
                throw "Executed input no longer matches the release tree ($suiteName)."
            }
        }
        Write-Host "Execution evidence: $suiteName — $($suite.passed) passed, $($suite.skippedTests) skipped tests, $($suite.skippedGroups) skipped groups."
    }
    $coreRun = $evidence.softwareVerification.csharpTests
    $registeredTestCount = [regex]::Matches(
        (Get-Content -LiteralPath (Join-Path $Root 'ClearPlan.Core.Tests\Program.cs') -Raw), 'new\s+TestCase\s*\(').Count
    if ($coreRun.total -ne $registeredTestCount -or @($coreRun.passedTestNames | Select-Object -Unique).Count -ne $coreRun.total -or
        $coreRun.skippedTests -ne 0 -or $coreRun.skippedGroups -ne 0 -or
        @($coreRun.inputSha256.PSObject.Properties).Count -eq 0) {
        throw "The actual C# execution does not cover the current registered suite and hashed inputs."
    }
}

$primaryAssemblyInfoFiles = @(
    "ClearPlan.Script\Properties\AssemblyInfo.cs",
    "ClearPlan.Runner\Properties\AssemblyInfo.cs",
    "ClearPlan.Simulator\Properties\AssemblyInfo.cs"
)
foreach ($relativePath in $primaryAssemblyInfoFiles) {
    Assert-FileContains `
        -RelativePath $relativePath `
        -Pattern "AssemblyVersion\(`"$escapedFileVersion`"\)" `
        -Description "assembly version $expectedFileVersion"
}

$releaseAssemblyInfoFiles = @(
    "ClearPlan.Script\Properties\AssemblyInfo.cs",
    "ClearPlan.Runner\Properties\AssemblyInfo.cs",
    "ClearPlan.Simulator\Properties\AssemblyInfo.cs",
    "ClearPlan.Core\Properties\AssemblyInfo.cs",
    "ClearPlan.Rendering\Properties\AssemblyInfo.cs",
    "ClearPlan.Presentation\Properties\AssemblyInfo.cs",
    "ClearPlan.Reporting\Properties\AssemblyInfo.cs",
    "ClearPlan.Reporting.MigraDoc\Properties\AssemblyInfo.cs"
)
foreach ($relativePath in $releaseAssemblyInfoFiles) {
    Assert-FileContains `
        -RelativePath $relativePath `
        -Pattern "AssemblyFileVersion\(`"$escapedFileVersion`"\)" `
        -Description "file version $expectedFileVersion"
    Assert-FileContains `
        -RelativePath $relativePath `
        -Pattern "AssemblyInformationalVersion\(`"$escapedVersion`"\)" `
        -Description "informational version $ExpectedVersion"
}

$dicomProjectPath = Join-Path $Root 'ClearPlan.Dicom\ClearPlan.Dicom.csproj'
if (Test-Path -LiteralPath $dicomProjectPath -PathType Leaf) {
    [xml]$dicomProject = Get-Content -LiteralPath $dicomProjectPath -Raw
    $dicomVersion = @($dicomProject.Project.PropertyGroup.Version | Where-Object { $_ })
    $dicomFileVersion = @($dicomProject.Project.PropertyGroup.FileVersion | Where-Object { $_ })
    if ($dicomVersion.Count -ne 1 -or $dicomVersion[0] -ne $ExpectedVersion -or
        $dicomFileVersion.Count -ne 1 -or $dicomFileVersion[0] -ne $expectedFileVersion) {
        throw "Optional DICOM source must declare Version $ExpectedVersion and FileVersion $expectedFileVersion."
    }
}

if (-not $SkipBinaries) {
    # Validate the outputs produced by the Release|x64 solution build. The
    # built\ directory belongs to the separate clinical deployment workflow
    # and may intentionally contain the last approved workstation build.
    $publicBinaryPaths = @(
        "artifacts\bin\Release\ClearPlan.Core.dll",
        "artifacts\bin\Release\ClearPlan.Rendering.dll",
        "artifacts\bin\Release\ClearPlan.Presentation.dll",
        "debug\ClearPlan.Reporting.dll",
        "debug\ClearPlan.Reporting.MigraDoc.dll",
        "artifacts\simulator\Release\ClearPlan.Simulator.exe",
        "artifacts\simulator\Release\ClearPlan.Core.dll",
        "artifacts\simulator\Release\ClearPlan.Rendering.dll",
        "artifacts\simulator\Release\ClearPlan.Presentation.dll",
        "artifacts\simulator\Release\ClearPlan.Reporting.dll",
        "artifacts\simulator\Release\ClearPlan.Reporting.MigraDoc.dll"
    )
    $binaryPaths = if ($PublicBinariesOnly) {
        $publicBinaryPaths
    }
    else {
        @(
            "debug\ClearPlan.esapi.dll",
            "ClearPlan.Runner\bin\x64\Release\ClearPlan.Runner.exe"
        ) + $publicBinaryPaths
    }
    if ($IncludeDicomBinaries) {
        $binaryPaths += "artifacts\dicom\Release\ClearPlan.Dicom.dll"
    }
    foreach ($relativePath in $binaryPaths) {
        $path = Join-Path $Root $relativePath
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Release binary is missing: $relativePath"
        }
        $versionInfo = [Diagnostics.FileVersionInfo]::GetVersionInfo($path)
        if ([string]$versionInfo.FileVersion -ne $expectedFileVersion) {
            throw (
                "{0} has file version '{1}', expected '{2}'." -f
                $relativePath,
                $versionInfo.FileVersion,
                $expectedFileVersion
            )
        }
    }
}

$versionScope = if ($SkipBinaries) { 'source only; binary checks skipped' } else { 'source and built binaries' }
Write-Host (
    "PASS version consistency ({2}): v{0} / file version {1}; publication and clinical acceptance are separate gates." -f
    $ExpectedVersion,
    $expectedFileVersion,
    $versionScope
)
