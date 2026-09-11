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
    'x:Name="SyntheticDemoToggle"',
    'x:Name="ClinicalExtrasMenuItem"',
    'Click="SyntheticDemoToggle_Click"',
    'SYNTHETIC DEMONSTRATION — NOT FOR CLINICAL USE',
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
    'HandleSharedSyntheticReport\s*\(',
    'SetSyntheticDemoPresentation\s*\(',
    'ClinicalExtrasMenuItem\.Visibility\s*=\s*enabled',
    'window\.Title\s*=\s*"ClearPlan .* Synthetische Demonstration"',
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
    'TrySetSyntheticDemo\s*\(',
    'SyntheticScenarioFactory\.Create\("mixed-review"\)',
    'controller\.TryShowSnapshot\s*\(',
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
    'TryShowSnapshot\s*\(',
    'CurrentSnapshot\s*=\s*snapshot',
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

# File parsing is the single reviewed worker exception in the host. Its lambda
# receives a detached path only, never PlanSetup/Beam/Structure. Preserve the
# blanket rejection for any other worker added to this clinical controller.
$safeReadPattern = 'Task\.Run\(\(\) => RtPlanReader\.Read\(selectedPath\), analysisCancellation\.Token\)'
if ([regex]::Matches($hostSource, $safeReadPattern).Count -ne 1) {
    throw "Clinical host must contain exactly one detached RTPLAN read worker."
}
foreach ($requiredPattern in @(
    'Task\.WhenAny\(readTask, Task\.Delay\(-1, analysisCancellation\.Token\)\)',
    'RtPlanReader\.MatchAndApply',
    'AnalysisContextIsCurrent\(original, expectedUid\)',
    'analysisCancellation\.Token\.ThrowIfCancellationRequested\(\)',
    'controller\.TryShowSnapshot\(updated, out failure\)',
    'ReferenceEquals\(controller\.CurrentSnapshot, original\)',
    'PlanningItemUID != expectedUid'
)) {
    if ($hostSource -notmatch $requiredPattern) {
        throw "RTPLAN worker lacks a reviewed timeout/staleness/transaction guard: $requiredPattern"
    }
}
if ($hostSource -match 'original\.PlanAnalysis\s*=') {
    throw "An optional analysis must not mutate the last valid snapshot before validation."
}
$clinicalCode = [regex]::Replace($hostSource, $safeReadPattern, 'ReviewedDetachedFileRead') + [Environment]::NewLine + $controller
if ($mainCode -match '_vm\.ActivePlanningItem\s*=\s*selectedPlanningItem') {
    throw "Plan replacement must not mutate the previous live context before construction succeeds."
}
foreach ($pattern in @(
    'new MainView\(mainViewModel, comparisonReference\)',
    'window\.Content\s*=\s*replacement',
    'CanExportCurrentSnapshot\(snapshot\)'
)) {
    if ($mainCode -notmatch $pattern) { throw "Clinical plan replacement/report guard missing: $pattern" }
}
foreach ($pattern in @('ReviewPlanContextGuard', 'planContext\.Matches', 'planContext\.Commit\(updated, expectedUid, source\.ActivePlanningItem\.PlanningItemObject\)')) {
    if ($hostSource -notmatch $pattern) { throw "Host-only committed plan identity guard missing: $pattern" }
}
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
