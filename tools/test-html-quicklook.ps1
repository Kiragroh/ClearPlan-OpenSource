[CmdletBinding()]
param([string]$AssemblyDirectory = 'artifacts\bin\Debug', [string]$ArtifactsDirectory = 'artifacts\html-quicklook-check')
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run using powershell.exe -STA -File.' }
$taskAssemblyRoot = [IO.Path]::GetFullPath($AssemblyDirectory)
$taskOutputRoot = [IO.Path]::GetFullPath($ArtifactsDirectory)
if ($taskOutputRoot.StartsWith('\\')) { throw 'Synthetic review artifacts must use an explicit local folder.' }
[void][IO.Directory]::CreateDirectory($taskOutputRoot)
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$taskResolver = [ResolveEventHandler]{ param($sender,$eventArgs)
    $taskCandidate = Join-Path $taskAssemblyRoot (([Reflection.AssemblyName]$eventArgs.Name).Name + '.dll')
    if (Test-Path -LiteralPath $taskCandidate) { return [Reflection.Assembly]::LoadFrom($taskCandidate) }
    return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($taskResolver)
foreach ($taskAssembly in @('ClearPlan.Core.dll','ClearPlan.Rendering.dll','ClearPlan.Reporting.dll','ClearPlan.Reporting.MigraDoc.dll','ClearPlan.Presentation.dll')) {
    [void][Reflection.Assembly]::LoadFrom((Join-Path $taskAssemblyRoot $taskAssembly))
}
$taskSnapshot = [ClearPlan.Core.Simulation.SyntheticScenarioFactory]::Create('mixed-review')
$taskSnapshot.PlanImages = [ClearPlan.Core.Simulation.SyntheticPlanImageFactory]::Create($taskSnapshot.ActivePlanKey)
$taskSnapshot.PlanAnalysis = [ClearPlan.Core.PlanAnalysis.SyntheticPlanAnalysisFactory]::Create($true,$true)
$taskBeam = $taskSnapshot.PlanAnalysis.Beams[0]
$taskCp = $taskBeam.ControlPoints[0]
$taskCp.BevImage = [ClearPlan.Core.Simulation.SyntheticDrrFactory]::Create($taskBeam,$taskCp)
$taskDocument = (New-Object ClearPlan.Reporting.ReviewSnapshotReportMapper).Map($taskSnapshot)
$taskTimer = [Diagnostics.Stopwatch]::StartNew()
$taskHtml = (New-Object ClearPlan.Reporting.MigraDoc.HtmlReviewReportRenderer).Render($taskDocument)
$taskRenderMs = $taskTimer.Elapsed.TotalMilliseconds
[IO.File]::WriteAllText((Join-Path $taskOutputRoot 'synthetic-plan-review.html'),$taskHtml,(New-Object Text.UTF8Encoding($false)))
$taskWindow = New-Object ClearPlan.Presentation.Views.HtmlReportPreviewWindow($taskHtml,$taskOutputRoot)
$taskWindow.ShowActivated = $false
$taskWindow.ShowInTaskbar = $false
$taskWindow.Show()
$taskFrame = New-Object Windows.Threading.DispatcherFrame
$taskDeadline = [DateTime]::UtcNow.AddSeconds(20)
$taskPump = New-Object Windows.Threading.DispatcherTimer
$taskPump.Interval = [TimeSpan]::FromMilliseconds(50)
$taskPump.add_Tick({ if ($taskWindow.PreviewLoaded -or $taskWindow.PreviewUnavailable -or [DateTime]::UtcNow -ge $taskDeadline) { $taskFrame.Continue = $false } })
$taskPump.Start()
try {
    [Windows.Threading.Dispatcher]::PushFrame($taskFrame)
    if (-not $taskWindow.PreviewLoaded -or $taskWindow.PreviewUnavailable) { throw 'Native in-app HTML viewer did not load.' }
    $taskBrowser = $taskWindow.GetType().GetField('browser',[Reflection.BindingFlags]'NonPublic,Instance').GetValue($taskWindow)
    $taskDom = $taskBrowser.Document
    $taskImages = @($taskDom.images | ForEach-Object { [pscustomobject]@{ Complete=$_.complete; Width=$_.width; Height=$_.height } })
    $taskResult = [pscustomobject]@{ Synthetic=$true; Loaded=$taskWindow.PreviewLoaded; RenderMilliseconds=$taskRenderMs; HtmlCharacters=$taskHtml.Length;
        DocumentMode=$taskDom.documentMode; ScrollWidth=$taskDom.body.scrollWidth; ClientWidth=$taskDom.body.clientWidth;
        Images=$taskImages; EmbeddedImages=([regex]::Matches($taskHtml,'data:image/png;base64,').Count) }
    $taskResult | ConvertTo-Json -Depth 5
    if ($taskImages.Count -lt 5 -or @($taskImages | Where-Object { -not $_.Complete -or $_.Width -lt 10 }).Count -gt 0) { throw 'Embedded report images did not render.' }
    if ($taskResult.ScrollWidth -gt $taskResult.ClientWidth + 2) { throw 'HTML report overflows the native viewer.' }
} finally { $taskPump.Stop(); $taskWindow.Close(); [AppDomain]::CurrentDomain.remove_AssemblyResolve($taskResolver) }
