[CmdletBinding()]
param(
    [string]$Root = "",
    [string]$ExpectedVersion = "3.1.0",
    [switch]$SkipBinaries,
    [switch]$PublicBinariesOnly
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
Assert-FileContains `
    -RelativePath "README.md" `
    -Pattern "125 C# executable tests" `
    -Description "the verified C# test count"
Assert-FileContains `
    -RelativePath "paper\ClearPlan_ZMP_short_communication.md" `
    -Pattern ([regex]::Escape("releases/tag/v$ExpectedVersion")) `
    -Description "the manuscript release URL"

$evidencePath = Join-Path `
    $Root "paper\evidence\technical_note_evidence.json"
if (-not (Test-Path -LiteralPath $evidencePath -PathType Leaf)) {
    throw "The technical-note evidence file is missing."
}
$evidence = Get-Content -LiteralPath $evidencePath -Raw | ConvertFrom-Json
if ([string]$evidence.repositoryRelease -ne "v$ExpectedVersion") {
    throw "Evidence release '$($evidence.repositoryRelease)' is not v$ExpectedVersion."
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

if (-not $SkipBinaries) {
    # Validate the outputs produced by the Release|x64 solution build. The
    # built\ directory belongs to the separate clinical deployment workflow
    # and may intentionally contain the last approved workstation build.
    $publicBinaryPaths = @(
        "artifacts\bin\Release\ClearPlan.Core.dll",
        "artifacts\bin\Release\ClearPlan.Presentation.dll",
        "debug\ClearPlan.Reporting.dll",
        "debug\ClearPlan.Reporting.MigraDoc.dll",
        "artifacts\simulator\Release\ClearPlan.Simulator.exe",
        "artifacts\simulator\Release\ClearPlan.Core.dll",
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

Write-Host (
    "PASS release version consistency: v{0} / file version {1}" -f
    $ExpectedVersion,
    $expectedFileVersion
)
