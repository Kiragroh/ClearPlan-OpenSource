$ErrorActionPreference = "Stop"

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).ProviderPath
$mainViewModelPath = Join-Path $repoRoot "ClearPlan.Script\ViewModels\MainViewModel.cs"
$pqmViewModelPath = Join-Path $repoRoot "ClearPlan.Script\ViewModels\PQMSummaryViewModel.cs"
$mainViewPath = Join-Path $repoRoot "ClearPlan.Script\MainView.xaml.cs"

$mainViewModel = Get-Content -LiteralPath $mainViewModelPath -Raw
$pqmViewModel = Get-Content -LiteralPath $pqmViewModelPath -Raw
$mainView = Get-Content -LiteralPath $mainViewPath -Raw

if ($mainViewModel -notmatch "nextPqmSummaries" -or
    $mainViewModel -notmatch "PqmSummaries\s*=\s*nextPqmSummaries") {
    throw "PQM table recalculation is not committed transactionally."
}

if ($mainViewModel -notmatch "CapturePqmReviewState" -or
    $mainViewModel -notmatch "RestorePqmReviewState") {
    throw "PQM review state cannot be restored after a derived-view failure."
}

if ($pqmViewModel -notmatch "EvaluateStructure" -or
    $pqmViewModel -notmatch "ApplyStructureEvaluation" -or
    $pqmViewModel -notmatch "CaptureStructureEvaluation") {
    throw "PQM structure mapping does not expose a prepare/commit boundary."
}

if ($mainView -notmatch "preparedMappings" -or
    $mainView -notmatch "Previous\s*=\s*pqm\.CaptureStructureEvaluation" -or
    $mainView -notmatch "ApplyStructureEvaluation") {
    throw "Shared structure mappings are not prepared fully before commit."
}

if ($mainView -notmatch
    "if\s*\(\s*!RefreshClinicalReviewWorkspace\(\)\s*\)" -or
    $mainView -notmatch
    "RestorePqmReviewState") {
    throw "Snapshot failures do not trigger a complete PQM rollback."
}

if ($mainView -match "workspace\.MarkStructureMappingApplied") {
    throw "Structure mapping is committed to the visible snapshot before refresh succeeds."
}

if ($mainView -notmatch "catch\s*\(Exception exception\)" -or
    $mainView -notmatch "ShowSharedWorkspaceFailure\(\s*exception,\s*true") {
    throw "Clinical review actions do not retain the last-good workspace on failure."
}

Write-Host "PQM transaction safety validation passed."
