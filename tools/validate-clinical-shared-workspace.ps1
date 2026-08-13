[CmdletBinding()]
param(
    [string]$RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
)

$ErrorActionPreference = "Stop"

$paths = [ordered]@{
    mainXaml = Join-Path $RepositoryRoot "ClearPlan.Script\MainView.xaml"
    mainCode = Join-Path $RepositoryRoot "ClearPlan.Script\MainView.xaml.cs"
    host = Join-Path $RepositoryRoot "ClearPlan.Script\Review\ClinicalReviewWorkspaceHost.cs"
    controller = Join-Path $RepositoryRoot "ClearPlan.Presentation\ViewModels\ReviewWorkspaceHostController.cs"
    scriptProject = Join-Path $RepositoryRoot "ClearPlan.Script\ClearPlan.csproj"
    presentationProject = Join-Path $RepositoryRoot "ClearPlan.Presentation\ClearPlan.Presentation.csproj"
    deploy = Join-Path $RepositoryRoot "tools\deploy-internal-clearplan.ps1"
}
foreach ($path in $paths.Values) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Clinical shared workspace file is missing: $path"
    }
}

$mainXaml = Get-Content -LiteralPath $paths.mainXaml -Raw
$mainCode = Get-Content -LiteralPath $paths.mainCode -Raw
$hostSource = Get-Content -LiteralPath $paths.host -Raw
$controller = Get-Content -LiteralPath $paths.controller -Raw
$scriptProject = Get-Content -LiteralPath $paths.scriptProject -Raw
$presentationProject = Get-Content -LiteralPath $paths.presentationProject -Raw
$deploy = Get-Content -LiteralPath $paths.deploy -Raw

foreach ($requiredPattern in @(
    'x:Name="SharedReviewWorkspace"',
    'x:Name="LegacyCompatibilitySurface"',
    'review:ReviewWorkspaceView',
    'x:Name="SharedConstraintComboBox"',
    'Click="SharedSettings_Click"'
)) {
    if ($mainXaml -notmatch $requiredPattern) {
        throw "Clinical XAML does not expose the shared workspace contract: $requiredPattern"
    }
}

foreach ($requiredPattern in @(
    'SharedReviewWorkspace\.Visibility\s*=\s*Visibility\.Visible',
    'LegacyCompatibilitySurface\.Visibility\s*=\s*Visibility\.Collapsed',
    'SharedReviewWorkspace\.Visibility\s*=\s*Visibility\.Collapsed',
    'LegacyCompatibilitySurface\.Visibility\s*=\s*Visibility\.Visible',
    'HandleSharedReport\s*\(',
    'HandleSharedOpenPlan\s*\(',
    'HandleSharedDvhReset\s*\(',
    'HandleSharedDvhExport\s*\(',
    'HandleSharedStructureMapping\s*\(',
    'HandleSharedSettings\s*\(',
    'SyncSharedDvhSelections\s*\(',
    'OpenPlanningItem\s*\('
)) {
    if ($mainCode -notmatch $requiredPattern) {
        throw "Clinical shared action/fallback is missing: $requiredPattern"
    }
}

foreach ($requiredPattern in @(
    'ReviewWorkspaceHostController',
    'EsapiReviewSnapshotBuilder',
    'TryRefresh\s*\(',
    'Dispose\s*\(',
    'OnReportRequested',
    'OnOpenPlanRequested',
    'OnDvhResetRequested',
    'OnDvhExportRequested',
    'OnStructureMappingRequested',
    'OnSettingsRequested'
)) {
    if ($hostSource -notmatch $requiredPattern) {
        throw "Clinical workspace host contract is missing: $requiredPattern"
    }
}

foreach ($requiredPattern in @(
    'ReviewSnapshotValidator\.Validate\s*\(',
    'ReviewWorkspaceViewModel\s+previous\s*=\s*CurrentViewModel',
    'CurrentViewModel\s*=\s*next',
    'Unsubscribe\s*\(\s*previous\s*\)',
    'Unsubscribe\s*\(\s*CurrentViewModel\s*\)',
    'ReportRequested\s*\+=',
    'ReportRequested\s*-='
)) {
    if ($controller -notmatch $requiredPattern) {
        throw "Shared host lifecycle contract is missing: $requiredPattern"
    }
}

if ($scriptProject -notmatch
    '<Compile\s+Include="Review\\ClinicalReviewWorkspaceHost\.cs"') {
    throw "ClearPlan.Script does not compile ClinicalReviewWorkspaceHost.cs."
}
if ($presentationProject -notmatch
    '<Compile\s+Include="ViewModels\\ReviewWorkspaceHostController\.cs"') {
    throw "ClearPlan.Presentation does not compile ReviewWorkspaceHostController.cs."
}
foreach ($relativePath in @(
    "ClearPlan.Script\Review\ClinicalReviewWorkspaceHost.cs",
    "tools\validate-clinical-shared-workspace.ps1"
)) {
    $isDirectlyListed =
        $deploy -match [regex]::Escape('"' + $relativePath + '"')
    $isCoveredByRecursiveTools =
        $relativePath.StartsWith(
            "tools\",
            [StringComparison]::OrdinalIgnoreCase) -and
        $deploy -match [regex]::Escape('"tools"')
    if (-not $isDirectlyListed -and -not $isCoveredByRecursiveTools) {
        throw "Clinical deployment omits required source: $relativePath"
    }
}

$clinicalCode = $hostSource + [Environment]::NewLine + $controller
foreach ($forbidden in @(
    '\bBeginModifications\s*\(',
    '\bSaveModifications\s*\(',
    '\bPatient\s*\.\s*Save\s*\(',
    '\bTask\s*\.\s*Run\s*\(',
    '\bnew\s+Thread\s*\('
)) {
    if ($clinicalCode -match $forbidden) {
        throw "Clinical shared workspace contains a forbidden operation: $($Matches[0])"
    }
}

Write-Host "Clinical shared workspace validation passed."
