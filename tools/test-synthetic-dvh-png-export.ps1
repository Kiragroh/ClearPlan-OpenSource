[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SimulatorDirectory,
    [Parameter(Mandatory = $true)][string]$ArtifactsDirectory
)

# Synthetic-only publication export verification. No TPS context, native capture,
# or clinical source is loaded. Existing publication artifacts are never inputs for writing.
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run with powershell.exe -STA.' }
[AppContext]::SetSwitch('Switch.System.Windows.Media.ShouldRenderEvenWhenNoDisplayDevicesAreAvailable', $true)
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Drawing
$runtime = [IO.Path]::GetFullPath($SimulatorDirectory)
$output = [IO.Path]::GetFullPath($ArtifactsDirectory)
if ($output.StartsWith('\\')) { throw 'Use an explicit local synthetic QA directory.' }
if ($output -match '[\\/]publication-release-3\.2\.0(?:[\\/]|$)') { throw 'Do not overwrite the publication QA inputs.' }
[void][Reflection.Assembly]::LoadFrom((Join-Path $runtime 'ClearPlan.Core.dll'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $runtime 'ClearPlan.Presentation.dll'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $runtime 'ClearPlan.Simulator.exe'))
[void][IO.Directory]::CreateDirectory($output)
$service = New-Object ClearPlan.Simulator.SyntheticDvhPngService
$results = foreach ($scenario in @('publication-single-layer', 'publication-dual-layer')) {
    $snapshot = [ClearPlan.Core.Simulation.SyntheticPublicationScenarioFactory]::Create($scenario)
    if (-not $snapshot.Synthetic) { throw 'The QA export refuses non-synthetic data.' }
    $vm = New-Object ClearPlan.Presentation.ViewModels.ReviewWorkspaceViewModel($snapshot)
    $plot = $vm.DetailPlotModel
    $initialLegend = $plot.IsLegendVisible
    $initialTitle = $plot.Title
    $initialPoints = @($plot.Series | ForEach-Object { $_.Points.Count })
    $path = Join-Path $output ($scenario + '-dvh.png')
    [void]$service.Export($plot, $path)
    $bitmap = [Drawing.Bitmap]::FromFile($path)
    try {
        if ($bitmap.Width -ne 1400 -or $bitmap.Height -ne 800) { throw 'Unexpected DVH export dimensions.' }
        $markerPixels = 0
        for ($y = 10; $y -lt 40; $y++) {
            for ($x = 230; $x -lt 1100; $x++) {
                $pixel = $bitmap.GetPixel($x, $y)
                if ($pixel.R -lt 120 -and $pixel.G -lt 120 -and $pixel.B -lt 120) { $markerPixels++ }
            }
        }
        if ($markerPixels -le 200) { throw 'DVH PNG lacks its own visible synthetic-use notice.' }
        $legendArea = $plot.LegendArea
        if ($legendArea.Width -le 0 -or $legendArea.Left -lt $plot.PlotArea.Right) { throw 'DVH legend must be outside the plot on the right.' }
        if ($plot.PlotArea.Left -lt 40 -or $plot.PlotArea.Bottom -gt ($bitmap.Height - 35)) { throw 'DVH axes do not fit fully within the export.' }
        $legendPixels = foreach ($line in $plot.Series) {
            if (-not $line.IsVisible) { continue }
            $count = 0
            for ($y = [int][Math]::Ceiling($legendArea.Top); $y -lt [Math]::Min($bitmap.Height, $legendArea.Bottom); $y++) {
                for ($x = [int][Math]::Ceiling($legendArea.Left); $x -lt [Math]::Min($bitmap.Width, $legendArea.Right); $x++) {
                    $pixel = $bitmap.GetPixel($x, $y)
                    if ($pixel.R -eq $line.Color.R -and $pixel.G -eq $line.Color.G -and $pixel.B -eq $line.Color.B) { $count++ }
                }
            }
            if ($count -lt 8) { throw ('Missing right-side legend key: ' + $line.Title) }
            [pscustomobject]@{ label = $line.Title; coloredPixels = $count }
        }
        if ($plot.IsLegendVisible -ne $initialLegend -or $plot.Title -ne $initialTitle) { throw 'Export changed GUI plot settings.' }
        for ($index = 0; $index -lt $plot.Series.Count; $index++) {
            if ($plot.Series[$index].Points.Count -ne $initialPoints[$index]) { throw 'Export changed DVH sample counts.' }
        }
        [pscustomobject]@{ synthetic = $true; scenario = $scenario; width = $bitmap.Width; height = $bitmap.Height; visibleSyntheticMarkerPixels = $markerPixels; legend = @($legendPixels); png = $path; sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash }
    }
    finally { $bitmap.Dispose() }
}
[IO.File]::WriteAllText((Join-Path $output 'dvh-png-export-verification.json'), (ConvertTo-Json -InputObject @($results) -Depth 5), (New-Object Text.UTF8Encoding($false)))
$results | Select-Object scenario, width, height, visibleSyntheticMarkerPixels, png | Format-Table -AutoSize
Write-Output 'PASS both synthetic publication DVH PNG exports: marker, selected-curve legend, dimensions and GUI-state restoration'
