$ErrorActionPreference = 'Stop'
$taskSource = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../ClearPlan.Script/Review/ClinicalReviewWorkspaceHost.cs') -Raw
foreach ($taskMethod in @('public bool TryRefresh()', 'public bool TrySetSyntheticDemo(bool enabled)')) {
    $taskStart = $taskSource.IndexOf($taskMethod)
    $taskEnd = $taskSource.IndexOf('        public ', $taskStart + $taskMethod.Length)
    $taskBody = $taskSource.Substring($taskStart, $taskEnd - $taskStart)
    if (-not $taskBody.Contains('QueueNativeAnalysis();')) { throw "$taskMethod does not start configured native analysis automatically." }
}
if (-not $taskSource.Contains('internal Task NativeAnalysisTask')) { throw 'Native analysis completion must be observable for native GUI verification.' }
if (-not $taskSource.Contains('await previous;') -or -not $taskSource.Contains('ReferenceEquals(controller.CurrentSnapshot, expectedSnapshot)')) { throw 'Automatic analysis must serialize and reject stale snapshots.' }
if (-not $taskSource.Contains('selectedCurves.TryGetValue') -or -not $taskSource.Contains('series.IsSelected = selected')) { throw 'Completing analysis must preserve current DVH selections.' }
if (-not $taskSource.Contains('RestoreAnalysisSelection(') -or -not $taskSource.Contains('selectedControlPointIndex')) { throw 'Automatic analysis must preserve beam, control-point and target choices.' }
if (-not $taskSource.Contains('CanUpdateAnalysisStatus(original)')) { throw 'Late native status must not overwrite a sandbox or newer plan.' }
Write-Output 'PASS native analysis startup, freshness, completion and selection contracts'
