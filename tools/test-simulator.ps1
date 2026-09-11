[CmdletBinding()]
param(
    [string]$SimulatorDirectory,
    [string]$ArtifactsDirectory
)

$ErrorActionPreference = "Stop"

$repoRoot = [System.IO.Path]::GetFullPath(
    (Join-Path $PSScriptRoot ".."))
if ([string]::IsNullOrWhiteSpace($SimulatorDirectory)) {
    $SimulatorDirectory = Join-Path `
        $repoRoot "artifacts\simulator\Release"
}
if ([string]::IsNullOrWhiteSpace($ArtifactsDirectory)) {
    $runId = [Guid]::NewGuid().ToString("N")
    $ArtifactsDirectory = Join-Path `
        $repoRoot ("artifacts\simulator-smoke\" + $runId)
}

$SimulatorDirectory = [System.IO.Path]::GetFullPath($SimulatorDirectory)
$ArtifactsDirectory = [System.IO.Path]::GetFullPath($ArtifactsDirectory)
$exePath = Join-Path $SimulatorDirectory "ClearPlan.Simulator.exe"

$smokeSource = Get-Content -LiteralPath $PSCommandPath -Raw
$moduleHashCommand = "Get-" + "FileHash"
if ($smokeSource.IndexOf(
    $moduleHashCommand,
    [StringComparison]::OrdinalIgnoreCase) -ge 0) {
    throw "Simulator smoke hashing must use System.Security.Cryptography directly."
}

$mainWindowSourcePath = Join-Path `
    $repoRoot "ClearPlan.Simulator\MainWindow.xaml.cs"
$mainWindowSource = Get-Content -LiteralPath $mainWindowSourcePath -Raw
if ($mainWindowSource -notmatch
    'SimulatorOutputPaths\.GetDefaultReportDirectory' -or
    $mainWindowSource -notmatch
    'InitialDirectory\s*=\s*defaultReportDirectory') {
    throw "Interactive reports must default to artifacts/simulator-reports."
}

$releaseProjectPaths = @(
    "ClearPlan.Core\ClearPlan.Core.csproj",
    "ClearPlan.Presentation\ClearPlan.Presentation.csproj",
    "ClearPlan.Reporting\ClearPlan.Reporting.csproj",
    "ClearPlan.Reporting.MigraDoc\ClearPlan.Reporting.MigraDoc.csproj",
    "ClearPlan.Simulator\ClearPlan.Simulator.csproj"
)
foreach ($relativeProjectPath in $releaseProjectPaths) {
    $releaseProjectPath = Join-Path $repoRoot $relativeProjectPath
    [xml]$releaseProject = Get-Content `
        -LiteralPath $releaseProjectPath `
        -Raw
    $releaseGroups = @(
        $releaseProject.Project.PropertyGroup |
        Where-Object {
            ([string]$_.Condition).IndexOf(
                "Release",
                [StringComparison]::OrdinalIgnoreCase) -ge 0
        }
    )
    if ($releaseGroups.Count -eq 0) {
        throw "Release configuration is missing: $relativeProjectPath"
    }
    foreach ($releaseGroup in $releaseGroups) {
        if ([string]$releaseGroup.DebugType -ne "none" -or
            [string]$releaseGroup.DebugSymbols -ne "false") {
            throw "Release symbols expose build paths: $relativeProjectPath"
        }
    }
}

if (-not (Test-Path -LiteralPath $exePath -PathType Leaf)) {
    throw "Simulator executable is missing: $exePath"
}

$layoutProbeDirectory = Join-Path `
    $ArtifactsDirectory "layout-probe"
& (Join-Path $PSScriptRoot "test-responsive-layout.ps1") `
    -SimulatorDirectory $SimulatorDirectory `
    -ArtifactsDirectory $layoutProbeDirectory

$scenarioIds = @(
    "baseline-pass",
    "target-underdose",
    "oar-overdose",
    "metadata-plancheck",
    "field-and-mapping",
    "optional-path-fallback",
    "mixed-review"
)
$tabIds = @("overview", "pqm", "plancheck", "fields", "dvh", "parameters", "comparison", "bev")

$scenarioDirectory = Join-Path $SimulatorDirectory "Scenarios"
foreach ($scenarioId in $scenarioIds) {
    $scenarioPath = Join-Path $scenarioDirectory ($scenarioId + ".json")
    if (-not (Test-Path -LiteralPath $scenarioPath -PathType Leaf)) {
        throw "Checked-in simulator scenario is missing from output: $scenarioPath"
    }
}

$forbiddenFile = Get-ChildItem `
    -LiteralPath $SimulatorDirectory `
    -File `
    -Recurse |
    Where-Object {
        $_.Name -match "(?i)^(VMS\.|EsapiEssentials)" -or
        $_.FullName -match "(?i)\\(VMS\.|EsapiEssentials)"
    } |
    Select-Object -First 1
if ($null -ne $forbiddenFile) {
    throw "Proprietary dependency leaked into simulator output: $($forbiddenFile.FullName)"
}

$forbiddenReferenceTokens = @(
    "VMS.TPS.",
    "EsapiEssentials"
)
foreach ($assembly in Get-ChildItem `
    -LiteralPath $SimulatorDirectory `
    -File |
    Where-Object { $_.Extension -in @(".dll", ".exe") }) {
    $assemblyText = [Text.Encoding]::ASCII.GetString(
        [IO.File]::ReadAllBytes($assembly.FullName))
    foreach ($token in $forbiddenReferenceTokens) {
        if ($assemblyText.IndexOf(
            $token,
            [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            throw "Proprietary assembly reference '$token' found in $($assembly.Name)."
        }
    }
}

$privateBuildPathPattern = "[A-Za-z]:\\Users\\"
foreach ($binary in Get-ChildItem `
    -LiteralPath $SimulatorDirectory `
    -File |
    Where-Object { $_.Extension -in @(".dll", ".exe", ".pdb") }) {
    $binaryBytes = [IO.File]::ReadAllBytes($binary.FullName)
    $binaryTexts = @(
        [Text.Encoding]::ASCII.GetString($binaryBytes),
        [Text.Encoding]::Unicode.GetString($binaryBytes),
        [Text.Encoding]::BigEndianUnicode.GetString($binaryBytes)
    )
    foreach ($binaryText in $binaryTexts) {
        if ($binaryText -match $privateBuildPathPattern) {
            throw "Private build path found in public simulator output: $($binary.Name)"
        }
    }
}

$simulatorText = [Text.Encoding]::ASCII.GetString(
    [IO.File]::ReadAllBytes($exePath))
foreach ($marker in @(
    "SYNTHETIC DEMONSTRATION",
    "NOT FOR CLINICAL USE"
)) {
    if ($simulatorText.IndexOf(
        $marker,
        [StringComparison]::Ordinal) -lt 0) {
        throw "Required safety marker '$marker' is absent from the executable."
    }
}

$publishSafeFiles = Get-ChildItem `
    -LiteralPath $scenarioDirectory `
    -File `
    -Filter "*.json"
$privatePatterns = @(
    "[A-Za-z]:\\\\",
    "\\\\\\\\[^\\]+\\\\",
    "(?<![\d.])\d+(?:\.\d+){4,}(?![\d.])"
)
foreach ($file in $publishSafeFiles) {
    $content = Get-Content -LiteralPath $file.FullName -Raw
    foreach ($pattern in $privatePatterns) {
        if ($content -match $pattern) {
            throw "Private path or UID pattern found in $($file.Name)."
        }
    }
}

New-Item `
    -ItemType Directory `
    -Path $ArtifactsDirectory `
    -Force | Out-Null

function Invoke-Simulator {
    param([string[]]$Arguments)

    $quotedArguments = foreach ($argument in $Arguments) {
        if ($argument -match '\s|"') {
            '"' + $argument.Replace('"', '\"') + '"'
        }
        else {
            $argument
        }
    }
    $process = Start-Process `
        -FilePath $exePath `
        -ArgumentList $quotedArguments `
        -WindowStyle Hidden `
        -PassThru `
        -Wait
    if ($process.ExitCode -ne 0) {
        throw "Simulator failed with exit code $($process.ExitCode) for: $($Arguments -join ' ')"
    }
}

function Assert-Png {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Expected PNG was not created: $Path"
    }
    $file = Get-Item -LiteralPath $Path
    if ($file.Length -lt 50000) {
        throw "PNG is implausibly small ($($file.Length) bytes): $Path"
    }

    Add-Type -AssemblyName System.Drawing
    $image = [System.Drawing.Image]::FromFile($Path)
    try {
        if ($image.Width -ne 1600 -or $image.Height -ne 1000) {
            throw "PNG dimensions must be 1600x1000, found $($image.Width)x$($image.Height): $Path"
        }
    }
    finally {
        $image.Dispose()
    }
}

function Assert-DvhPng {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Expected DVH PNG was not created: $Path"
    }
    $file = Get-Item -LiteralPath $Path
    if ($file.Length -lt 20000) {
        throw "DVH PNG is implausibly small ($($file.Length) bytes): $Path"
    }

    Add-Type -AssemblyName System.Drawing
    $image = [System.Drawing.Image]::FromFile($Path)
    try {
        if ($image.Width -ne 1400 -or $image.Height -ne 800) {
            throw "DVH PNG dimensions must be 1400x800, found $($image.Width)x$($image.Height): $Path"
        }
    }
    finally {
        $image.Dispose()
    }
}

function Get-Sha256Hex {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $stream = [IO.File]::Open(
        $Path,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        [IO.FileShare]::Read)
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try {
        $hash = $algorithm.ComputeHash($stream)
        return ([BitConverter]::ToString($hash)).Replace("-", "")
    }
    finally {
        $algorithm.Dispose()
        $stream.Dispose()
    }
}

$validationDirectory = Join-Path $ArtifactsDirectory "scenario-validation"
New-Item `
    -ItemType Directory `
    -Path $validationDirectory `
    -Force | Out-Null
foreach ($scenarioId in $scenarioIds) {
    $validationPng = Join-Path `
        $validationDirectory ($scenarioId + "-overview.png")
    Invoke-Simulator @(
        "--scenario", $scenarioId,
        "--tab", "overview",
        "--capture", $validationPng
    )
    Assert-Png $validationPng
}

$firstDirectory = Join-Path $ArtifactsDirectory "capture-first"
$secondDirectory = Join-Path $ArtifactsDirectory "capture-second"
Invoke-Simulator @(
    "--scenario", "field-and-mapping",
    "--capture-all", $firstDirectory
)
Invoke-Simulator @(
    "--scenario", "field-and-mapping",
    "--capture-all", $secondDirectory
)

foreach ($tabId in $tabIds) {
    $firstPng = Join-Path $firstDirectory ($tabId + ".png")
    $secondPng = Join-Path $secondDirectory ($tabId + ".png")
    Assert-Png $firstPng
    Assert-Png $secondPng
    $firstHash = Get-Sha256Hex -Path $firstPng
    $secondHash = Get-Sha256Hex -Path $secondPng
    if ($firstHash -ne $secondHash) {
        throw "Capture '$tabId' is not byte-identical across two runs."
    }
}

$firstDvhExport = Join-Path `
    $firstDirectory "clearplan-synthetic-dvh.png"
$secondDvhExport = Join-Path `
    $secondDirectory "clearplan-synthetic-dvh.png"
Assert-DvhPng $firstDvhExport
Assert-DvhPng $secondDvhExport
if ((Get-Sha256Hex -Path $firstDvhExport) -ne
    (Get-Sha256Hex -Path $secondDvhExport)) {
    throw "DVH PNG export is not byte-identical across two runs."
}

$reportPath = Join-Path $firstDirectory "clearplan-synthetic-report.pdf"
if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) {
    throw "Synthetic PDF report was not generated: $reportPath"
}
$reportBytes = [IO.File]::ReadAllBytes($reportPath)
if ($reportBytes.Length -lt 10000 -or
    [Text.Encoding]::ASCII.GetString($reportBytes, 0, 5) -ne "%PDF-") {
    throw "Generated report is not a plausible PDF: $reportPath"
}

Write-Host "PASS ClearPlan simulator smoke test"
Write-Host "Simulator: $exePath"
Write-Host "Artifacts: $ArtifactsDirectory"
Write-Host "Scenarios: $($scenarioIds.Count)"
Write-Host "Deterministic captures: $($tabIds.Count)"
Write-Host "Deterministic DVH PNG: $firstDvhExport"
Write-Host "PDF: $reportPath"
