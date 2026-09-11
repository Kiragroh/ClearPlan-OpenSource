$ErrorActionPreference = 'Stop'
$taskSource = Get-Content -LiteralPath (Join-Path $PSScriptRoot '../ClearPlan.Script/Review/EsapiPlanImageBuilder.cs') -Raw
$taskFailures = [System.Collections.Generic.List[string]]::new()
function Require-CaptureContract([bool] $condition, [string] $message) {
    if (-not $condition) { $taskFailures.Add($message) }
}
Require-CaptureContract ($taskSource -match 'public static Task<List<ReviewPlanImage>> BuildAsync\(PlanningItem planningItem, string planKey,\s*CancellationToken cancellation, Func<bool> isCurrent\)') 'Async GUI capture needs cancellation and mandatory lifetime callback.'
Require-CaptureContract ($taskSource.Contains('BuildCoreAsync(planningItem, planKey, new CaptureContext(false, CancellationToken.None, () => true)).GetAwaiter().GetResult()')) 'Synchronous report capture must run the no-yield path without a dispatcher deadlock.'
foreach ($taskTerm in @('ApartmentState.STA', 'Dispatcher.FromThread(Thread.CurrentThread)', 'SynchronizationContext.Current is DispatcherSynchronizationContext', 'Thread.CurrentThread.ManagedThreadId != ownerThreadId', 'cancellation.ThrowIfCancellationRequested();', 'if (!isCurrent()) throw new OperationCanceledException', 'await Dispatcher.Yield(DispatcherPriority.Background);', 'NativeRowChunkSize', 'DvhSelectionPolicy.ClassifyTarget(structureId, dicomType)')) {
    Require-CaptureContract ($taskSource.Contains($taskTerm)) "Missing owner/cancellation/bounded capture contract: $taskTerm"
}
Require-CaptureContract (-not ($taskSource -match 'Task\.Run|ConfigureAwait\(false\)|BeginModifications|SaveModifications|File\.Write')) 'Native capture must stay read-only on the owning dispatcher, with no patient output.'
$taskAwaits = [regex]::Matches($taskSource, '(?m)^\s*await\s+[^;]+;\s*(?<next>[^\r\n]+)')
Require-CaptureContract ($taskAwaits.Count -ge 7) 'CT rows, structures, mesh planes and dose rows must yield cooperatively.'
foreach ($taskAwait in $taskAwaits) {
    Require-CaptureContract ($taskAwait.Groups['next'].Value.Trim() -in @('context.RequireCurrent();', 'RequireCurrent();')) 'Every await must immediately recheck cancellation, owner and active context before any native access.'
}
$taskBroadCatches = [regex]::Matches($taskSource, 'catch\s*\(Exception\)')
foreach ($taskCatch in $taskBroadCatches) {
    $taskPrefix = $taskSource.Substring(0, $taskCatch.Index)
    Require-CaptureContract ($taskPrefix -match 'catch\s*\(OperationCanceledException\)\s*\{\s*throw;\s*\}\s*$') 'Cancellation must propagate through every broad fallback catch.'
}
$taskScopeStart = $taskSource.IndexOf('private static T WithAbsoluteDosePresentation<T>')
$taskScopeEnd = $taskSource.IndexOf('private static int StructureRank', [Math]::Max(0, $taskScopeStart))
Require-CaptureContract ($taskScopeStart -ge 0 -and $taskScopeEnd -gt $taskScopeStart) 'Dose reads require an uninterrupted absolute-presentation scope.'
if ($taskScopeStart -ge 0 -and $taskScopeEnd -gt $taskScopeStart) {
    $taskScope = $taskSource.Substring($taskScopeStart, $taskScopeEnd - $taskScopeStart)
    Require-CaptureContract ($taskScope.Contains('finally') -and $taskScope.Contains('context.IsCurrent') -and $taskScope.Contains('plan.DoseValuePresentation = previousPresentation.Value')) 'Dose presentation must be restored before yielding and never on a closed/replaced plan.'
    Require-CaptureContract (-not $taskScope.Contains('await ')) 'Dose presentation may not remain modified across an await.'
}
if ($taskFailures.Count -gt 0) { throw ($taskFailures -join [Environment]::NewLine) }
Write-Output 'PASS native CT asynchronous owner, cancellation, row-chunk, lifetime, dose restoration and target ordering contracts'
