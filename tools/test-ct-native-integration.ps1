$ErrorActionPreference = 'Stop'
$taskRepo = Split-Path $PSScriptRoot -Parent
$taskView = Get-Content -Raw -LiteralPath (Join-Path $taskRepo 'ClearPlan.Presentation/Views/ReviewWorkspaceView.xaml')
$taskHost = Get-Content -Raw -LiteralPath (Join-Path $taskRepo 'ClearPlan.Script/Review/ClinicalReviewWorkspaceHost.cs')
$taskMain = Get-Content -Raw -LiteralPath (Join-Path $taskRepo 'ClearPlan.Script/MainView.xaml.cs')
$taskBuilder = Get-Content -Raw -LiteralPath (Join-Path $taskRepo 'ClearPlan.Script/Review/EsapiPlanImageBuilder.cs')
if (-not $taskView.Contains('x:Name="PlanImagesTab"') -or -not $taskView.Contains('views:PlanImagesView')) { throw 'Three-plane CT view is not reachable from the workspace.' }
foreach ($taskContract in @('PreparePlanImagesAsync', 'CanExportCurrentSnapshot(snapshot)', 'EsapiPlanImageBuilder.BuildAsync', 'PlanImagesTask', 'CurrentViewModel.PlanImages', '"PlanImagesTab"')) {
    if (-not $taskHost.Contains($taskContract)) { throw "Missing native CT integration: $taskContract" }
}
if ($taskMain.Contains('snapshot.PlanImages.Clear();')) { throw 'Report export still discards the current GUI CT cache.' }
if (-not $taskMain.Contains('PreparePlanImagesAsync(snapshot)')) { throw 'PDF export must share the guarded CT capture/cache.' }
foreach ($taskContract in @('CancellationToken', 'Func<bool>', 'BuildAsync', 'OperationCanceledException')) {
    if (-not $taskBuilder.Contains($taskContract)) { throw "Missing bounded asynchronous capture: $taskContract" }
}
if ($taskBuilder.Contains('Task.Run(')) { throw 'Native CT capture must never access ESAPI on a worker thread.' }
$taskCaptureStart = $taskHost.IndexOf('private async Task<System.Collections.Generic.List<ClearPlan.Core.Review.ReviewPlanImage>> CapturePlanImagesAsync')
$taskNativeRead = $taskHost.IndexOf('var images = await EsapiPlanImageBuilder.BuildAsync', $taskCaptureStart)
$taskEntry = $taskHost.Substring($taskCaptureStart, $taskNativeRead - $taskCaptureStart)
if (-not $taskEntry.Contains('await System.Windows.Threading.Dispatcher.Yield') -or -not $taskEntry.Contains('if (!CanExportCurrentSnapshot(snapshot))')) { throw 'Runner-triggered CT capture must enter and revalidate its dispatcher before reading native objects.' }
Write-Output 'PASS three-plane native navigation, shared report cache and owner-thread capture contracts'
