[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$SimulatorDirectory,
      [Parameter(Mandatory=$true)][string]$ArtifactsDirectory)
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Use Windows PowerShell -STA.' }
[AppContext]::SetSwitch('Switch.System.Windows.Media.ShouldRenderEvenWhenNoDisplayDevicesAreAvailable', $true)
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$taskRuntime = [IO.Path]::GetFullPath($SimulatorDirectory)
$taskOutput = [IO.Path]::GetFullPath($ArtifactsDirectory)
if ($taskOutput.StartsWith('\\')) { throw 'Synthetic evidence requires a local output directory.' }
[void][IO.Directory]::CreateDirectory($taskOutput)
[void][Reflection.Assembly]::LoadFrom((Join-Path $taskRuntime 'ClearPlan.Simulator.exe'))
$taskApp = New-Object Windows.Application
$taskApp.ShutdownMode = 'OnExplicitShutdown'
$taskRepository = New-Object ClearPlan.Simulator.SimulatorScenarioRepository($taskRuntime)
$taskArguments = [ClearPlan.Simulator.SimulatorArguments]::Parse([string[]]@('--scenario','mixed-review'))
$taskWindow = New-Object ClearPlan.Simulator.MainWindow($taskRepository, $taskArguments,
    (New-Object ClearPlan.Simulator.SyntheticReportService), (New-Object ClearPlan.Simulator.VisualCaptureService),
    (New-Object ClearPlan.Simulator.SyntheticDvhPngService), (New-Object ClearPlan.Simulator.ResponsiveLayoutProbeService))
$taskWindow.ShowActivated = $false
$taskWindow.ShowInTaskbar = $false
$taskWindow.WindowStartupLocation = 'Manual'
$taskWindow.Left = -20000
$taskWindow.Top = 0
function Render-Review {
    $taskWindow.UpdateLayout()
    [void]$taskWindow.Dispatcher.Invoke([Action]{},[Windows.Threading.DispatcherPriority]::ApplicationIdle)
    $taskWindow.UpdateLayout()
}
try {
    $taskWindow.Show()
    $taskWorkspace = $taskWindow.FindName('ReviewWorkspace')
    $taskCapture = $taskWindow.FindName('SimulatorCaptureRoot')
    $taskVm = $taskWorkspace.DataContext
    if (-not $taskVm.IsSynthetic) { throw 'Synthetic-only verification.' }
    $taskVm.Comparison.LoadExampleCommand.Execute($null)
    if (-not $taskVm.HideUnmatched) { throw 'Unmatched rows must be hidden by default.' }
    $taskVm.IncludeBeamEyeViews = $false
    $taskRows = foreach ($taskSize in @(@(1600,1000),@(1180,720))) {
        $taskWindow.Width = $taskSize[0]; $taskWindow.Height = $taskSize[1]
        foreach ($taskName in @('DvhTab','PqmTab','PlanCheckTab','ComparisonTab','RateUnavailable','RatePlanned','RateEstimated')) {
            if ($taskName.StartsWith('Rate')) {
                $taskTab = $taskWorkspace.FindName('PlanParametersTab')
                $taskTab.IsSelected = $true
                $taskBeam = $taskVm.Analysis.Beams[0]
                foreach ($taskCp in $taskBeam.ControlPoints) {
                    $taskCp.PlannedDoseRateMuPerMin = $null
                    $taskCp.EstimatedDoseRateMuPerMin = $null
                    if ($taskName -eq 'RatePlanned') { $taskCp.PlannedDoseRateMuPerMin = 200 + ($taskCp.Index % 4) * 100 }
                }
                $taskBeam.DoseRateEstimateStatus = 'Unavailable'
                if ($taskName -eq 'RateEstimated') {
                    $taskBeam.Technique = 'ARC'; $taskBeam.GantryDirection = 'Clockwise'
                    foreach ($taskCp in $taskBeam.ControlPoints) { $taskCp.CumulativeMetersetWeight = [Math]::Pow($taskCp.Index / [double]($taskBeam.ControlPoints.Count - 1), 2) }
                    $taskProfile = New-Object ClearPlan.Core.PlanAnalysis.DoseRateEstimationProfile
                    $taskProfile.Id = 'Synthetic timing example'; $taskProfile.MachineIds = New-Object 'Collections.Generic.List[string]'
                    $taskProfile.MachineIds.Add($taskBeam.MachineId); $taskProfile.MaxGantrySpeedDegreesPerSecond = 6
                    $taskProfile.Assumption = 'Synthetic test only. No measured delivery or machine commissioning.'
                    [ClearPlan.Core.PlanAnalysis.DoseRateEstimator]::Apply($taskBeam, $taskProfile)
                }
                $taskVm.Analysis.SelectedBeam = $taskBeam
                Render-Review
                $taskParameters = $taskTab.Content
                $taskParameters.GetType().GetMethod('ShowParameterPlots',[Reflection.BindingFlags]'Instance,NonPublic').Invoke($taskParameters,@($null,[Windows.RoutedEventArgs]::new()))
                Render-Review
                $taskExpectedRate = $taskName -ne 'RateUnavailable'
                if ($taskVm.Analysis.HasDoseRateTrace -ne $taskExpectedRate) { throw 'Dose-rate availability state incorrect.' }
                if ($taskName -eq 'RateEstimated' -and (-not $taskVm.Analysis.HasEstimatedDoseRate -or $taskVm.Analysis.HasPlannedDoseRate)) { throw 'Estimated and supplied rates must stay separate.' }
                $taskRatePlot = $taskParameters.FindName('DoseRatePlot')
                $taskEmpty = $taskParameters.FindName('DoseRateUnavailable')
                if ($taskRatePlot.IsVisible -ne $taskExpectedRate -or $taskEmpty.IsVisible -eq $taskExpectedRate) { throw 'Dose-rate empty state did not replace the missing curve.' }
                $taskVisibleRate = if ($taskExpectedRate) { $taskRatePlot } else { $taskEmpty }
                if ($taskVisibleRate.ActualWidth -lt 300 -or $taskVisibleRate.ActualHeight -lt 200) { throw 'Dose-rate panel is clipped.' }
            } else { $taskWorkspace.FindName($taskName).IsSelected = $true }
            Render-Review
            if ($taskName -eq 'DvhTab') {
                $taskPlot = $taskWorkspace.FindName('DvhDetailPlot')
                $taskLegend = $taskWorkspace.FindName('DvhSeriesList')
                $taskPlotX = $taskPlot.TransformToAncestor($taskCapture).Transform([Windows.Point]::new(0,0)).X
                $taskLegendX = $taskLegend.TransformToAncestor($taskCapture).Transform([Windows.Point]::new(0,0)).X
                if ($taskLegendX -le $taskPlotX -or $taskPlot.ActualWidth -lt 500 -or $taskPlot.ActualHeight -lt 300) { throw 'Compressed DVH or misplaced legend.' }
                $taskSeries = $taskVm.DvhSeries[0]
                $taskSeries.IsSelected = $false
                if ($taskVm.PlanImages.IsStructureVisible($taskSeries.StructureId)) { throw 'DVH selection did not propagate.' }
                $taskSeries.IsSelected = $true
            }
            if ($taskName -eq 'PlanCheckTab') {
                $taskGrid = $taskWorkspace.FindName('PlanCheckDetailGrid')
                if ($taskGrid.Columns[0].Header -ne 'Meldung' -or $taskGrid.Columns[1].Header -ne 'Status') { throw 'Check column order incorrect.' }
            }
            $taskBitmap = New-Object Windows.Media.Imaging.RenderTargetBitmap([int][Math]::Ceiling($taskCapture.ActualWidth),[int][Math]::Ceiling($taskCapture.ActualHeight),96,96,[Windows.Media.PixelFormats]::Pbgra32)
            $taskBitmap.Render($taskCapture)
            [ClearPlan.Simulator.CapturePixelValidator]::AssertNonBlank($taskBitmap)
            $taskPng = Join-Path $taskOutput ($taskName+'-'+$taskSize[0]+'x'+$taskSize[1]+'.png')
            $taskEncoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
            $taskEncoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($taskBitmap))
            $taskStream = [IO.File]::Create($taskPng)
            try { $taskEncoder.Save($taskStream) } finally { $taskStream.Dispose() }
            [pscustomobject]@{tab=$taskName;width=$taskSize[0];height=$taskSize[1];synthetic=$true;png=$taskPng}
        }
    }
    $taskRows | Format-Table -AutoSize
    Write-Output 'PASS fourteen synthetic native-WPF captures; DVH legend; compact checks; default unmatched filter; planned/estimated/missing dose-rate states.'
} catch { Write-Output $_.Exception.ToString(); throw }
finally { $taskWindow.Close(); $taskApp.Shutdown() }
