[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$SimulatorDirectory,
      [Parameter(Mandatory=$true)][string]$OutputPath,
      [switch]$Estimated,
      [switch]$SingleLayer)
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Use Windows PowerShell -STA.' }
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$taskRuntime = [IO.Path]::GetFullPath($SimulatorDirectory)
[void][Reflection.Assembly]::LoadFrom((Join-Path $taskRuntime 'ClearPlan.Simulator.exe'))
foreach ($taskAssembly in @('ClearPlan.Core.dll','ClearPlan.Reporting.dll')) {
    [void][Reflection.Assembly]::LoadFrom((Join-Path $taskRuntime $taskAssembly))
}
$taskSnapshot = [ClearPlan.Core.Simulation.SyntheticScenarioFactory]::Create('mixed-review')
$taskSnapshot.PlanAnalysis = [ClearPlan.Core.PlanAnalysis.SyntheticPlanAnalysisFactory]::Create((-not $SingleLayer), $true)
foreach ($taskBeam in $taskSnapshot.PlanAnalysis.Beams) {
    foreach ($taskCp in $taskBeam.ControlPoints) { $taskCp.PlannedDoseRateMuPerMin = $null }
    if ($Estimated) {
        $taskBeam.Technique = 'ARC'; $taskBeam.GantryDirection = 'Clockwise'
        foreach ($taskCp in $taskBeam.ControlPoints) { $taskCp.CumulativeMetersetWeight = [Math]::Pow($taskCp.Index / [double]($taskBeam.ControlPoints.Count - 1), 2) }
        $taskProfile = New-Object ClearPlan.Core.PlanAnalysis.DoseRateEstimationProfile
        $taskProfile.Id = if ($SingleLayer) { 'Synthetic C-arm timing' } else { 'Synthetic dual-layer timing' }
        $taskProfile.MachineIds = New-Object 'Collections.Generic.List[string]'; $taskProfile.MachineIds.Add($taskBeam.MachineId)
        $taskProfile.MaxGantrySpeedDegreesPerSecond = if ($SingleLayer) { 6 } else { 12 }
        $taskProfile.Assumption = 'Synthetic example only; no measured delivery or commissioning.'
        [ClearPlan.Core.PlanAnalysis.DoseRateEstimator]::Apply($taskBeam, $taskProfile)
        if ($taskBeam.DoseRateEstimateStatus -ne 'Estimated') { throw 'Synthetic dose-rate estimate failed.' }
    }
}
if ($Estimated) { $taskSnapshot.PlanAnalysis = [ClearPlan.Core.PlanAnalysis.PlanAnalysisCalculator]::Calculate($taskSnapshot.PlanAnalysis) }
$taskSnapshot.PlanAnalysis.DoseRateProvenance = if ($Estimated) { 'Estimated synthetic plan trajectory. No measured delivery. Model omits acceleration, MLC/jaws, ramping and holds.' } else { 'Synthetic missing-rate fixture: no planned or measured time-resolved rates supplied.' }
$taskSnapshot.PqmRows[0].ResolvedStructureId = $null
$taskMapper = New-Object ClearPlan.Reporting.ReviewSnapshotReportMapper
$taskDocument = $taskMapper.Map($taskSnapshot)
if (-not $taskDocument.HideUnmatched) { throw 'Reports must hide unmatched rows by default.' }
$taskHidden = @($taskDocument.PqmRows | Where-Object { [string]::IsNullOrWhiteSpace($_.ResolvedStructureId) }).Count
if ($taskHidden -lt 1 -or $taskDocument.VisiblePqmRows().Count -ne ($taskDocument.PqmRows.Count - $taskHidden)) { throw 'Default report filtering is incorrect.' }
$taskExporter = New-Object ClearPlan.Simulator.SyntheticReportService
$taskFile = $taskExporter.Export($taskSnapshot, [IO.Path]::GetFullPath($OutputPath), $null)
Write-Output ('PASS synthetic PDF/HTML export; unmatched hidden: '+$taskHidden+'; estimated: '+$Estimated+'; nominal setting retained as scalar.')
Write-Output $taskFile
