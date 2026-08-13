[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceRoot,

    [Parameter(Mandatory = $true)]
    [string]$TargetRoot,

    [Parameter(Mandatory = $true)]
    [string]$PrivateSettingsPath,

    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"
$isWhatIf = [bool]$WhatIfPreference
$resolvedSource = (Resolve-Path -LiteralPath $SourceRoot).ProviderPath.TrimEnd("\")
$resolvedTarget = (Resolve-Path -LiteralPath $TargetRoot).ProviderPath.TrimEnd("\")
$resolvedPrivateSettings =
    (Resolve-Path -LiteralPath $PrivateSettingsPath).ProviderPath
$privatePathSettingsCandidate =
    [System.IO.Path]::ChangeExtension($resolvedPrivateSettings, ".ini")
$resolvedPrivatePathSettings =
    (Resolve-Path -LiteralPath $privatePathSettingsCandidate).ProviderPath

if ([System.IO.Path]::GetFileName($resolvedTarget) -ne "ClearPlan_OpenSource") {
    throw "Refusing deployment: target must resolve to the exact ClearPlan_OpenSource directory."
}

$settingsObject = Get-Content -LiteralPath $resolvedPrivateSettings -Raw |
    ConvertFrom-Json
if ($null -eq $settingsObject.constraintSource) {
    throw "Private settings do not contain constraintSource."
}

function Get-IniPathValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Key
    )

    $matchingLine = Get-Content -LiteralPath $Path |
        Where-Object { $_ -match ("^\s*" + [regex]::Escape($Key) + "\s*=") } |
        Select-Object -First 1
    if ($null -eq $matchingLine) {
        throw "Private path settings do not contain '$Key': $Path"
    }

    return [string](($matchingLine -split "=", 2)[1]).Trim()
}

$privatePathSettingsText =
    Get-Content -LiteralPath $resolvedPrivatePathSettings -Raw
if ($privatePathSettingsText -match "(?i)variancom") {
    throw "Private path settings still contain the retired variancom host."
}

$configuredRefDb = Get-IniPathValue `
    -Path $resolvedPrivatePathSettings `
    -Key "RefDbJsonPath"
if (-not [string]::IsNullOrWhiteSpace($configuredRefDb) -and
    -not (Test-Path -LiteralPath $configuredRefDb -PathType Leaf)) {
    throw "Configured clinical RefDB is not readable: $configuredRefDb"
}

$privateWorkbook = Join-Path ([System.IO.Path]::GetDirectoryName($resolvedPrivateSettings)) "ClearPlan_ClinicalConstraints.xlsx"
if (-not (Test-Path -LiteralPath $privateWorkbook -PathType Leaf)) {
    throw "Private clinical workbook is missing: $privateWorkbook"
}

$clinicalErrorCalculator = Join-Path $resolvedTarget "ClearPlan.Script\Calculators\ErrorCalculator.cs"
$savedWhatIfPreference = $WhatIfPreference
$WhatIfPreference = $false
try {
    $planCheckBefore = & (Join-Path $resolvedSource "tools\compare-plancheck-surface.ps1") `
        -ErrorCalculatorPath $clinicalErrorCalculator
}
finally {
    $WhatIfPreference = $savedWhatIfPreference
}

$clinicalArtifactNames = @(
    "ClearPlan.Core.dll",
    "ClearPlan.Core.pdb",
    "ClearPlan.Presentation.dll",
    "ClearPlan.Presentation.pdb",
    "ClearPlan.esapi.dll",
    "ClearPlan.esapi.pdb",
    "ClearPlan.Reporting.dll",
    "ClearPlan.Reporting.pdb",
    "ClearPlan.Reporting.MigraDoc.dll",
    "ClearPlan.Reporting.MigraDoc.pdb",
    "ClearPlan.Runner.exe",
    "ClearPlan.Runner.exe.config",
    "ClearPlan.Runner.pdb"
)
$clinicalBuildManifestPath =
    Join-Path $resolvedSource "built\clinical-build-manifest.json"

function Get-Sha256Hex {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $stream = [System.IO.File]::Open(
        $Path,
        [System.IO.FileMode]::Open,
        [System.IO.FileAccess]::Read,
        [System.IO.FileShare]::Read)
    $algorithm = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString(
            $algorithm.ComputeHash($stream))).Replace("-", "")
    }
    finally {
        $algorithm.Dispose()
        $stream.Dispose()
    }
}

function Get-ClinicalArtifactRecords {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BuildRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$Names
    )

    foreach ($name in $Names) {
        $path = Join-Path $BuildRoot $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Clinical build artifact is missing: $path"
        }

        [ordered]@{
            name = $name
            sha256 = Get-Sha256Hex -Path $path
            bytes = (Get-Item -LiteralPath $path).Length
        }
    }
}

function Assert-ClinicalBuildManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ManifestPath,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedPlanCheckHash,

        [Parameter(Mandatory = $true)]
        [string]$BuildRoot,

        [Parameter(Mandatory = $true)]
        [string[]]$ExpectedArtifactNames
    )

    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
        throw "Clinical build manifest is missing. Run deployment without -SkipBuild."
    }

    $buildManifest = Get-Content -LiteralPath $ManifestPath -Raw |
        ConvertFrom-Json
    if ($buildManifest.planCheckSourceSha256 -ne $ExpectedPlanCheckHash) {
        throw "Clinical build manifest does not match the preserved PlanCheck source."
    }

    $manifestArtifactNames = @(
        $buildManifest.artifacts |
            ForEach-Object { [string]$_.name }
    )
    foreach ($expectedArtifactName in $ExpectedArtifactNames) {
        if ($manifestArtifactNames -notcontains $expectedArtifactName) {
            throw "Clinical build manifest omits required artifact: $expectedArtifactName"
        }
    }

    foreach ($artifact in $buildManifest.artifacts) {
        $artifactPath = Join-Path $BuildRoot $artifact.name
        if (-not (Test-Path -LiteralPath $artifactPath -PathType Leaf)) {
            throw "Manifest artifact is missing: $artifactPath"
        }

        $actualHash = Get-Sha256Hex -Path $artifactPath
        if ($actualHash -ne $artifact.sha256) {
            throw "Clinical build artifact hash mismatch: $($artifact.name)"
        }
    }

    return $buildManifest
}

function Assert-CompiledTargetsReplaceable {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.IEnumerable]$CopyItems
    )

    foreach ($item in $CopyItems | Where-Object {
        $_.role -eq "compiled" -and
        (Test-Path -LiteralPath $_.target -PathType Leaf)
    }) {
        $stream = $null
        try {
            $stream = [System.IO.File]::Open(
                $item.target,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::ReadWrite,
                [System.IO.FileShare]::None)
        }
        catch {
            throw "Deployment target is in use: $($item.target). Close all ClearPlan instances before deployment; no deployment copy was started."
        }
        finally {
            if ($null -ne $stream) {
                $stream.Dispose()
            }
        }
    }
}

if (-not $SkipBuild) {
    $buildWhatIfPreference = $WhatIfPreference
    $WhatIfPreference = $false
    try {
        & (Join-Path $resolvedSource "tools\build-and-test.ps1")
        if ($LASTEXITCODE -ne 0) {
            throw "Build and test pipeline failed before deployment."
        }

        $msbuild = & (Join-Path $resolvedSource "tools\resolve-msbuild.ps1")
        if ([string]::IsNullOrWhiteSpace($msbuild) -or
            -not (Test-Path -LiteralPath $msbuild -PathType Leaf)) {
            throw "The MSBuild resolver did not return an existing executable."
        }

        & $msbuild `
            (Join-Path $resolvedSource "ClearPlan.sln") `
            /t:Build `
            "/p:Configuration=Debug" `
            "/p:Platform=x64" `
            "/p:ErrorCalculatorSource=$clinicalErrorCalculator" `
            /m
        if ($LASTEXITCODE -ne 0) {
            throw "Clinical PlanCheck build failed with exit code $LASTEXITCODE."
        }

        $clinicalBuildManifest = [ordered]@{
            createdAt = (Get-Date).ToString("o")
            configuration = "Debug"
            platform = "x64"
            planCheckSource = $clinicalErrorCalculator
            planCheckSourceSha256 = $planCheckBefore.sha256
            checkCallCount = $planCheckBefore.checkCallCount
            artifacts = @(
                Get-ClinicalArtifactRecords `
                    -BuildRoot (Join-Path $resolvedSource "built") `
                    -Names $clinicalArtifactNames
            )
        }
        $clinicalBuildManifest | ConvertTo-Json -Depth 6 |
            Set-Content -LiteralPath $clinicalBuildManifestPath -Encoding UTF8
    }
    finally {
        $WhatIfPreference = $buildWhatIfPreference
    }
}

$verifiedClinicalBuild = Assert-ClinicalBuildManifest `
    -ManifestPath $clinicalBuildManifestPath `
    -ExpectedPlanCheckHash $planCheckBefore.sha256 `
    -BuildRoot (Join-Path $resolvedSource "built") `
    -ExpectedArtifactNames $clinicalArtifactNames

$scriptFiles = @(
    "ClearPlan.sln",
    ".gitignore",
    "README.md",
    "CONTRIBUTING.md",
    "CITATION.cff",
    "THIRD-PARTY-NOTICES.md",
    "docs\releases\v3.1.0.md",
    "ClearPlan.Script\ClearPlan.csproj",
    "ClearPlan.Script\Script.cs",
    "ClearPlan.Script\Properties\AssemblyInfo.cs",
    "ClearPlan.Script\MainView.xaml",
    "ClearPlan.Script\MainView.xaml.cs",
    "ClearPlan.Script\PlanSelectView.xaml",
    "ClearPlan.Script\PlanSelectView.xaml.cs",
    "ClearPlan.Script\Review\ClinicalReviewWorkspaceHost.cs",
    "ClearPlan.Script\Review\EsapiReviewSnapshotBuilder.cs",
    "ClearPlan.Script\Helpers\ClearPlanSettings.cs",
    "ClearPlan.Script\Calculators\PQMSummaryCalculator.cs",
    "ClearPlan.Script\Calculators\FieldNamingPreviewCalculator.cs",
    "ClearPlan.Script\Mappers\PqmObjectiveViewModelMapper.cs",
    "ClearPlan.Script\ViewModels\ConstraintListViewModel.cs",
    "ClearPlan.Script\ViewModels\ConstraintViewModel.cs",
    "ClearPlan.Script\ViewModels\MainViewModel.cs",
    "ClearPlan.Script\ViewModels\PQMSummaryViewModel.cs",
    "ClearPlan.Script\ViewModels\DvhStructureViewModel.cs",
    "ClearPlan.Script\ViewModels\PlanSelectViewModel.cs",
    "ClearPlan.Script\ViewModels\FieldNamePreviewViewModel.cs",
    "ClearPlan.Script\ViewModels\ConstraintSourceStatusViewModel.cs",
    "ClearPlan.Script\ViewModels\OverviewViewModel.cs",
    "ClearPlan.Script\ViewModels\SettingsViewModel.cs",
    "ClearPlan.Script\Styles\ClinicalBlueprint.xaml",
    "ClearPlan.Script\Views\SettingsView.xaml",
    "ClearPlan.Script\Views\SettingsView.xaml.cs",
    "ClearPlan.Script\Distribution\CHANGELOG.md",
    "ClearPlan.Script\Distribution\ConstraintTemplates\ClearPlan_DefaultConstraints.xlsx",
    "ClearPlan.Runner\ClearPlan.Runner.csproj",
    "ClearPlan.Runner\App.xaml",
    "ClearPlan.Runner\App.xaml.cs",
    "ClearPlan.Runner\RunnerFailureReporter.cs",
    "ClearPlan.Runner\RunnerWindowStyler.cs",
    "ClearPlan.Runner\Properties\AssemblyInfo.cs",
    "ClearPlan.Runner\Styles\ClinicalBlueprintRunner.xaml"
)

$requiredCoreSourceFiles = @(
    "ClearPlan.Core\Review\ClinicalReviewValueMapper.cs"
)
foreach ($relativePath in $requiredCoreSourceFiles) {
    $requiredPath = Join-Path $resolvedSource $relativePath
    if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
        throw "Required clinical review core source is missing: $relativePath"
    }
}

$copyItems = New-Object System.Collections.Generic.List[object]
foreach ($relativePath in $scriptFiles) {
    $copyItems.Add([pscustomobject]@{
        source = Join-Path $resolvedSource $relativePath
        target = Join-Path $resolvedTarget $relativePath
        role = "source"
    })
}

$coreRoot = Join-Path $resolvedSource "ClearPlan.Core"
$coreFiles = Get-ChildItem -LiteralPath $coreRoot -Recurse -File |
    Where-Object {
        $_.FullName -notmatch '[\\/](bin|obj)[\\/]'
    }
foreach ($file in $coreFiles) {
    $relative = $file.FullName.Substring($resolvedSource.Length).TrimStart("\")
    $copyItems.Add([pscustomobject]@{
        source = $file.FullName
        target = Join-Path $resolvedTarget $relative
        role = "core"
    })
}

$presentationRoot = Join-Path $resolvedSource "ClearPlan.Presentation"
$presentationFiles =
    Get-ChildItem -LiteralPath $presentationRoot -Recurse -File |
        Where-Object {
            $_.FullName -notmatch '[\\/](bin|obj)[\\/]'
        }
foreach ($file in $presentationFiles) {
    $relative =
        $file.FullName.Substring($resolvedSource.Length).TrimStart("\")
    $copyItems.Add([pscustomobject]@{
        source = $file.FullName
        target = Join-Path $resolvedTarget $relative
        role = "presentation"
    })
}

$additionalRecursiveRoots = @(
    "ClearPlan.Core.Tests",
    "ClearPlan.Reporting",
    "ClearPlan.Reporting.MigraDoc",
    "ClearPlan.Simulator",
    "examples\raystation",
    "paper",
    "tools"
)
foreach ($relativeRoot in $additionalRecursiveRoots) {
    $sourceRoot = Join-Path $resolvedSource $relativeRoot
    if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
        throw "Required deployment source directory is missing: $relativeRoot"
    }
    $sourceFiles = Get-ChildItem -LiteralPath $sourceRoot -Recurse -File |
        Where-Object {
            $_.FullName -notmatch '[\\/](bin|obj|__pycache__)[\\/]' -and
            $_.Extension -ne ".pyc"
        }
    foreach ($file in $sourceFiles) {
        $relative =
            $file.FullName.Substring($resolvedSource.Length).TrimStart("\")
        $copyItems.Add([pscustomobject]@{
            source = $file.FullName
            target = Join-Path $resolvedTarget $relative
            role = "support"
        })
    }
}

$builtFiles = @(
    "ClearPlan.Core.dll",
    "ClearPlan.Core.pdb",
    "ClearPlan.Presentation.dll",
    "ClearPlan.Presentation.pdb",
    "ClearPlan.esapi.dll",
    "ClearPlan.esapi.pdb",
    "ClearPlan.Reporting.dll",
    "ClearPlan.Reporting.pdb",
    "ClearPlan.Reporting.MigraDoc.dll",
    "ClearPlan.Reporting.MigraDoc.pdb",
    "ClearPlan.Runner.exe",
    "ClearPlan.Runner.exe.config",
    "ClearPlan.Runner.pdb",
    "clinical-build-manifest.json",
    "CHANGELOG.md"
)
foreach ($fileName in $builtFiles) {
    $copyItems.Add([pscustomobject]@{
        source = Join-Path $resolvedSource ("built\" + $fileName)
        target = Join-Path $resolvedTarget ("built\" + $fileName)
        role = "compiled"
    })
}

$copyItems.Add([pscustomobject]@{
    source = $resolvedPrivateSettings
    target = Join-Path $resolvedTarget "built\settings.json"
    role = "private-settings"
})
$copyItems.Add([pscustomobject]@{
    source = $resolvedPrivateSettings
    target = Join-Path $resolvedTarget "ClearPlan.Script\Distribution\settings.json"
    role = "private-settings"
})
$copyItems.Add([pscustomobject]@{
    source = $resolvedPrivatePathSettings
    target = Join-Path $resolvedTarget "built\settings.ini"
    role = "private-path-settings"
})
$copyItems.Add([pscustomobject]@{
    source = $resolvedPrivatePathSettings
    target = Join-Path $resolvedTarget "ClearPlan.Script\Distribution\settings.ini"
    role = "private-path-settings"
})
$copyItems.Add([pscustomobject]@{
    source = $privateWorkbook
    target = Join-Path $resolvedTarget "built\ConstraintTemplates\ClearPlan_ClinicalConstraints.xlsx"
    role = "clinical-workbook"
})
$copyItems.Add([pscustomobject]@{
    source = $privateWorkbook
    target = Join-Path $resolvedTarget "ClearPlan.Script\Distribution\ConstraintTemplates\ClearPlan_ClinicalConstraints.xlsx"
    role = "clinical-workbook"
})

$forbiddenTarget = [System.IO.Path]::GetFullPath($clinicalErrorCalculator)
$forbiddenCopies = @($copyItems | Where-Object {
    [System.IO.Path]::GetFullPath($_.target) -eq $forbiddenTarget
})
if ($forbiddenCopies.Count -gt 0) {
    throw "Deployment manifest must never copy clinical ErrorCalculator.cs."
}

$missingSources = @($copyItems | Where-Object {
    -not (Test-Path -LiteralPath $_.source -PathType Leaf)
})
if ($missingSources.Count -gt 0) {
    throw "Deployment source files are missing: $($missingSources.source -join ', ')"
}

$retireRelativeFiles = @(
    "built\ConstraintTemplates\T_1Fx.csv",
    "built\ConstraintTemplates\T_3Fx.csv",
    "built\ConstraintTemplates\T_5Fx.csv",
    "built\ConstraintTemplates\T_8Fx.csv",
    "built\ConstraintTemplates\T_10Fx.csv",
    "built\ConstraintTemplates\T_15Fx.csv",
    "built\ConstraintTemplates\T_20Fx.csv",
    "built\ConstraintTemplates\T_Conv.csv",
    "built\ConstraintTemplates\T_Conv_Sum.csv",
    "ClearPlan.Script\Distribution\ConstraintTemplates\T_1Fx.csv",
    "ClearPlan.Script\Distribution\ConstraintTemplates\T_3Fx.csv",
    "ClearPlan.Script\Distribution\ConstraintTemplates\T_5Fx.csv",
    "ClearPlan.Script\Distribution\ConstraintTemplates\T_8Fx.csv",
    "ClearPlan.Script\Distribution\ConstraintTemplates\T_10Fx.csv",
    "ClearPlan.Script\Distribution\ConstraintTemplates\T_15Fx.csv",
    "ClearPlan.Script\Distribution\ConstraintTemplates\T_20Fx.csv",
    "ClearPlan.Script\Distribution\ConstraintTemplates\T_Conv.csv",
    "ClearPlan.Script\Distribution\ConstraintTemplates\T_Conv_Sum.csv"
)
$retireFiles = @($retireRelativeFiles |
    ForEach-Object { Join-Path $resolvedTarget $_ } |
    Where-Object { Test-Path -LiteralPath $_ -PathType Leaf })

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$prospectiveBackup = Join-Path ([System.IO.Path]::GetDirectoryName($resolvedTarget)) `
    "ClearPlan_OpenSource_backup_$stamp"
$manifest = [ordered]@{
    createdAt = (Get-Date).ToString("o")
    mode = if ($isWhatIf) { "WhatIf" } else { "Deploy" }
    sourceRoot = $resolvedSource
    targetRoot = $resolvedTarget
    backup = $prospectiveBackup
    planCheck = [ordered]@{
        path = $clinicalErrorCalculator
        sha256 = $planCheckBefore.sha256
        methods = $planCheckBefore.methods
        calculateReferences = $planCheckBefore.calculateReferences
        calculate2References = $planCheckBefore.calculate2References
        checkCallCount = $planCheckBefore.checkCallCount
        checkDescriptionCount = $planCheckBefore.checkDescriptionCount
        clinicalBuildCreatedAt = $verifiedClinicalBuild.createdAt
        preservation = "required"
    }
    copy = @($copyItems | ForEach-Object {
        [ordered]@{
            role = $_.role
            source = $_.source
            target = $_.target
        }
    })
    retireAfterBackup = $retireFiles
}

if ($isWhatIf) {
    $manifest | ConvertTo-Json -Depth 8
    exit 0
}

Assert-CompiledTargetsReplaceable -CopyItems $copyItems

$backupPath = & (Join-Path $resolvedSource "tools\backup-internal-clearplan.ps1") `
    -TargetRoot $resolvedTarget
$manifest.backup = $backupPath

foreach ($item in $copyItems) {
    $targetParent = [System.IO.Path]::GetDirectoryName($item.target)
    if (-not (Test-Path -LiteralPath $targetParent -PathType Container)) {
        New-Item -ItemType Directory -Path $targetParent | Out-Null
    }
    Copy-Item -LiteralPath $item.source -Destination $item.target -Force
}

foreach ($retireFile in $retireFiles) {
    Remove-Item -LiteralPath $retireFile -Force
}

$planCheckAfter = & (Join-Path $resolvedSource "tools\compare-plancheck-surface.ps1") `
    -ErrorCalculatorPath $clinicalErrorCalculator `
    -ExpectedHash $planCheckBefore.sha256
$manifest.planCheck.afterSha256 = $planCheckAfter.sha256
$manifest.planCheck.preservation = "verified"
$manifestPath = Join-Path $resolvedTarget "deployment-manifest-$stamp.json"
$manifest | ConvertTo-Json -Depth 8 |
    Set-Content -LiteralPath $manifestPath -Encoding UTF8

Write-Output $manifestPath
