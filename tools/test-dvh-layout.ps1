[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SimulatorDirectory,
    [Parameter(Mandatory = $true)]
    [string]$ArtifactsDirectory
)

# Run with Windows PowerShell -STA: the simulator targets .NET Framework WPF.
# Only the checked-in, validated synthetic scenario is loaded by this probe.
$ErrorActionPreference = "Stop"
if ([Threading.Thread]::CurrentThread.ApartmentState -ne "STA") {
    throw "Run this probe with powershell.exe -STA -File."
}
[AppContext]::SetSwitch(
    "Switch.System.Windows.Media.ShouldRenderEvenWhenNoDisplayDevicesAreAvailable",
    $true)
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$simulatorRoot = [IO.Path]::GetFullPath($SimulatorDirectory)
$outputRoot = [IO.Path]::GetFullPath($ArtifactsDirectory)
$simulatorPath = Join-Path $simulatorRoot "ClearPlan.Simulator.exe"
if (-not (Test-Path -LiteralPath $simulatorPath -PathType Leaf)) {
    throw "Simulator executable is missing: $simulatorPath"
}
if ($outputRoot.StartsWith("\\")) {
    throw "Use an explicit local directory for synthetic layout artifacts."
}
[void][Reflection.Assembly]::LoadFrom($simulatorPath)
[void][Reflection.Assembly]::LoadFrom(
    (Join-Path $simulatorRoot "ClearPlan.Presentation.dll"))
[void][IO.Directory]::CreateDirectory($outputRoot)
$application = New-Object System.Windows.Application
$application.ShutdownMode = "OnExplicitShutdown"
$repository = New-Object ClearPlan.Simulator.SimulatorScenarioRepository($simulatorRoot)
$arguments = [ClearPlan.Simulator.SimulatorArguments]::Parse(
    [string[]]@("--scenario", "field-and-mapping", "--tab", "dvh"))
$window = New-Object ClearPlan.Simulator.MainWindow(
    $repository,
    $arguments,
    (New-Object ClearPlan.Simulator.SyntheticReportService),
    (New-Object ClearPlan.Simulator.VisualCaptureService),
    (New-Object ClearPlan.Simulator.SyntheticDvhPngService),
    (New-Object ClearPlan.Simulator.ResponsiveLayoutProbeService))
$window.ShowActivated = $false
$window.ShowInTaskbar = $false
$window.WindowStartupLocation = "Manual"
$window.Left = -20000
$window.Top = 0
$window.Width = 1180
$window.Height = 720

function Invoke-RenderPass {
    $window.UpdateLayout()
    [void]$window.Dispatcher.Invoke(
        [Action]{}, [Windows.Threading.DispatcherPriority]::ApplicationIdle)
    $window.UpdateLayout()
}

try {
    $window.Show()
    $workspace = $window.FindName("ReviewWorkspace")
    $captureRoot = $window.FindName("SimulatorCaptureRoot")
    if (-not $workspace.DataContext.IsSynthetic) {
        throw "The DVH layout probe refuses non-synthetic data."
    }
    [ClearPlan.Simulator.VisualCaptureService]::SelectTab($workspace, "dvh")
    $results = foreach ($size in @(@(1180, 720), @(1600, 1000))) {
        $window.Width = $size[0]
        $window.Height = $size[1]
        Invoke-RenderPass
        $plot = $workspace.FindName("DvhDetailPlot")
        $content = $workspace.FindName("DvhContentGrid")
        $plot.Model.InvalidatePlot($true)
        Invoke-RenderPass
        $bounds = $plot.TransformToAncestor($captureRoot).TransformBounds(
            [Windows.Rect]::new([Windows.Point]::new(0, 0), $plot.RenderSize))
        if ($plot.ActualWidth -lt ($content.ActualWidth - 270)) {
            throw "DVH plot does not use the available width."
        }
        if ($plot.ActualHeight -lt 360) {
            throw "DVH plot is vertically compressed."
        }
        if ($bounds.Right -gt ($captureRoot.ActualWidth + 0.5) -or
            $bounds.Bottom -gt ($captureRoot.ActualHeight + 0.5)) {
            throw "DVH plot extends outside the rendered viewport."
        }
        if ($plot.Model.LegendPlacement -ne 'Outside' -or $plot.Model.LegendPosition -ne 'RightTop') {
            throw "DVH legend must be outside the plot on the right."
        }
        if ($plot.Model.LegendArea.Left -lt ($plot.Model.PlotArea.Right - 0.5)) {
            throw "DVH legend overlaps the dose-volume plotting area."
        }
        if ($plot.Model.PlotArea.Width -lt [Math]::Max(400, $plot.ActualWidth - 340)) {
            throw "DVH axes leave a compressed plotting area."
        }
        $sizeLabel = "{0}x{1}" -f $size[0], $size[1]
        $pngPath = Join-Path $outputRoot ("dvh-" + $sizeLabel + ".png")
        $bitmap = New-Object Windows.Media.Imaging.RenderTargetBitmap(
            [int][Math]::Ceiling($captureRoot.ActualWidth),
            [int][Math]::Ceiling($captureRoot.ActualHeight),
            96.0, 96.0, [Windows.Media.PixelFormats]::Pbgra32)
        $bitmap.Render($captureRoot)
        [ClearPlan.Simulator.CapturePixelValidator]::AssertNonBlank($bitmap)
        $encoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
        $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
        $stream = [IO.File]::Create($pngPath)
        try { $encoder.Save($stream) } finally { $stream.Dispose() }
        [pscustomobject]@{
            synthetic = $true
            requestedSize = $sizeLabel
            windowWidth = $window.ActualWidth
            windowHeight = $window.ActualHeight
            contentWidth = $content.ActualWidth
            plotWidth = $plot.ActualWidth
            plotHeight = $plot.ActualHeight
            plotAreaWidth = $plot.Model.PlotArea.Width
            plotAreaHeight = $plot.Model.PlotArea.Height
            legendLeft = $plot.Model.LegendArea.Left
            legendWidth = $plot.Model.LegendArea.Width
            legendPlacement = [string]$plot.Model.LegendPlacement
            legendPosition = [string]$plot.Model.LegendPosition
            plotRight = $bounds.Right
            plotBottom = $bounds.Bottom
            captureWidth = $captureRoot.ActualWidth
            captureHeight = $captureRoot.ActualHeight
            pngPath = $pngPath
        }
    }
    $jsonPath = Join-Path $outputRoot "dvh-layout.json"
    [IO.File]::WriteAllText($jsonPath, ($results | ConvertTo-Json -Depth 4),
        (New-Object Text.UTF8Encoding($false)))
    $results | Format-List
    Write-Host "PASS synthetic shared-workspace DVH layout and capture"
}
finally {
    $window.Close()
    $application.Shutdown()
}
