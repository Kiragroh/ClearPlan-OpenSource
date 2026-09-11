[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$SimulatorDirectory,
    [Parameter(Mandatory = $true)][string]$ArtifactsDirectory
)

# Synthetic WPF verification only. Ellipse contours and isodoses below are analytic
# test fixtures, not imported dose, structures, clinical CT, or production defaults.
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run with Windows PowerShell -STA.' }
[AppContext]::SetSwitch('Switch.System.Windows.Media.ShouldRenderEvenWhenNoDisplayDevicesAreAvailable', $true)
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$simulatorRoot = [IO.Path]::GetFullPath($SimulatorDirectory)
$outputRoot = [IO.Path]::GetFullPath($ArtifactsDirectory)
if ($outputRoot.StartsWith('\\')) { throw 'Use an explicit local synthetic artifact directory.' }
[void][Reflection.Assembly]::LoadFrom((Join-Path $simulatorRoot 'ClearPlan.Simulator.exe'))
[void][Reflection.Assembly]::LoadFrom((Join-Path $simulatorRoot 'ClearPlan.Presentation.dll'))
[void][IO.Directory]::CreateDirectory($outputRoot)
$application = New-Object Windows.Application
$application.ShutdownMode = 'OnExplicitShutdown'
$repository = New-Object ClearPlan.Simulator.SimulatorScenarioRepository($simulatorRoot)
$arguments = [ClearPlan.Simulator.SimulatorArguments]::Parse([string[]]@('--scenario', 'baseline-pass'))
$window = New-Object ClearPlan.Simulator.MainWindow($repository, $arguments,
    (New-Object ClearPlan.Simulator.SyntheticReportService), (New-Object ClearPlan.Simulator.VisualCaptureService),
    (New-Object ClearPlan.Simulator.SyntheticDvhPngService), (New-Object ClearPlan.Simulator.ResponsiveLayoutProbeService))
$window.ShowActivated = $false
$window.ShowInTaskbar = $false
$window.WindowStartupLocation = 'Manual'
$window.Left = -20000
$window.Top = 0

function Invoke-RenderPass {
    $window.UpdateLayout()
    [void]$window.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::ApplicationIdle)
    $window.UpdateLayout()
}

function Get-VisualDescendants([Windows.DependencyObject]$Element) {
    for ($index = 0; $index -lt [Windows.Media.VisualTreeHelper]::GetChildrenCount($Element); $index++) {
        $child = [Windows.Media.VisualTreeHelper]::GetChild($Element, $index)
        $child
        Get-VisualDescendants $child
    }
}

function Add-AnalyticOverlay($Image, [string]$Kind, [string]$Label, [string]$Color, [double]$CenterX, [double]$CenterY, [double]$RadiusX, [double]$RadiusY) {
    $overlay = New-Object ClearPlan.Core.Review.ReviewImageOverlay
    $overlay.Kind = $Kind
    $overlay.Label = $Label
    $overlay.ColorHex = $Color
    $overlay.SourceStatus = 'available'
    $overlay.Source = 'SYNTHETIC VERIFICATION ONLY: analytic ellipse; not calculated clinical dose or a native structure'
    $path = New-Object ClearPlan.Core.Review.ReviewImagePath
    $path.Closed = $true
    for ($sample = 0; $sample -lt 120; $sample++) {
        $angle = 2 * [Math]::PI * $sample / 120
        $point = New-Object ClearPlan.Core.Review.ReviewImagePoint
        $point.X = $CenterX + $RadiusX * [Math]::Cos($angle)
        $point.Y = $CenterY + $RadiusY * [Math]::Sin($angle)
        $path.Points.Add($point)
    }
    $overlay.Paths.Add($path)
    $Image.Overlays.Add($overlay)
}

try {
    $window.Width = 1600
    $window.Height = 1000
    $window.Show()
    $workspace = $window.FindName('ReviewWorkspace')
    $captureRoot = $window.FindName('SimulatorCaptureRoot')
    if (-not $workspace.DataContext.IsSynthetic) { throw 'This probe refuses non-synthetic data.' }
    $images = [ClearPlan.Core.Simulation.SyntheticPlanImageFactory]::Create($workspace.DataContext.ActivePlanKey)
    foreach ($image in $images) {
        if (-not $image.Synthetic) { throw 'Synthetic fixture contract failed.' }
        Add-AnalyticOverlay $image 'structure' 'DEMO target contour' '#FF6FCF' 160 160 26 35
        Add-AnalyticOverlay $image 'structure' 'DEMO organ contour' '#5CDFCC' 202 179 19 28
        Add-AnalyticOverlay $image 'isodose' 'DEMO 100% ellipse' '#FF6868' 160 160 34 42
        Add-AnalyticOverlay $image 'isodose' 'DEMO 50% ellipse' '#FFD166' 160 160 53 63
        Add-AnalyticOverlay $image 'isodose' 'DEMO 20% ellipse' '#65AFFF' 160 160 80 88
        $unavailableOverlay = New-Object ClearPlan.Core.Review.ReviewImageOverlay
        $unavailableOverlay.Kind = 'isodose'
        $unavailableOverlay.Label = 'DEMO unavailable source'
        $unavailableOverlay.SourceStatus = 'unavailable'
        $unavailableOverlay.UnavailableReason = 'Synthetic verification: deliberately unavailable source, not zero dose.'
        $image.Overlays.Add($unavailableOverlay)
        $image.OverlaySummary = 'SYNTHETIC VERIFICATION ONLY: analytic contours and dose-like ellipses. No clinical or calculated dose data.'
        $image.DoseFocusRegion = New-Object ClearPlan.Core.Review.ReviewImageDoseRegion
        $image.DoseFocusRegion.PrescriptionPercent = 2
        $image.DoseFocusRegion.ThresholdGy = 1
        $image.DoseFocusRegion.MinPixelX = 60
        $image.DoseFocusRegion.MaxPixelX = 260
        $image.DoseFocusRegion.MinPixelY = 60
        $image.DoseFocusRegion.MaxPixelY = 260
    }
    $vm = $workspace.DataContext.PlanImages
    $vm.ReplaceImages($images)
    if (-not $vm.Planes[0].AvailabilityText.Contains('1 Quelle fehlt')) { throw 'Unavailable overlay source is not visibly distinguished from zero contours.' }
    $tab = $workspace.FindName('PlanImagesTab')
    $tab.IsSelected = $true
    $view = $tab.Content
    Invoke-RenderPass
    $results = foreach ($size in @(@(1600, 1000), @(1180, 720))) {
        $window.Width = $size[0]
        $window.Height = $size[1]
        foreach ($mode in @('three-planes', 'enlarged-coronal')) {
            if ($mode -eq 'three-planes') { $vm.ShowAllPlanesCommand.Execute($null) }
            else { $vm.EnlargePlaneCommand.Execute('coronal') }
            Invoke-RenderPass
            $ctViewport = $view.FindName('CtScrollViewport')
            $ctViewport.ScrollToTop()
            Invoke-RenderPass
            $scrollDescendants = @(Get-VisualDescendants $ctViewport)
            $expectedCount = if ($mode -eq 'three-planes') { 3 } else { 1 }
            $visuals = @(Get-VisualDescendants $view | Where-Object { $_ -is [Windows.FrameworkElement] -and $_.IsVisible })
            $imageControls = @($visuals | Where-Object { $_ -is [Windows.Controls.Image] -and $null -ne $_.Source })
            if ($imageControls.Count -ne $expectedCount) { throw "Unexpected visible plane count in $mode : $($imageControls.Count)" }
            $measurements = foreach ($control in $visuals) {
                if ($control -isnot [Windows.Controls.Image] -and $control -isnot [Windows.Controls.Button] -and
                    $control -isnot [Windows.Controls.TextBlock] -and $control -isnot [Windows.Controls.CheckBox]) { continue }
                $bounds = $control.TransformToAncestor($captureRoot).TransformBounds([Windows.Rect]::new([Windows.Point]::new(0, 0), $control.RenderSize))
                $insideScrollableSinglePlane = $mode -eq 'enlarged-coronal' -and $scrollDescendants -contains $control
                if (-not $insideScrollableSinglePlane -and ($bounds.X -lt -0.5 -or $bounds.Y -lt -0.5 -or $bounds.Right -gt ($captureRoot.ActualWidth + 0.5) -or $bounds.Bottom -gt ($captureRoot.ActualHeight + 0.5))) {
                    throw "CT control outside the viewport at $($size[0]) x $($size[1]) in $mode : $($control.GetType().Name) $($control.Text)"
                }
                # Image.ActualWidth is the uniformly scaled raster, not the wider plane slot.
                # At 1180 x 720 the three thumbnails are height-limited; enlargement supplies detail.
                if ($control -is [Windows.Controls.Image] -and ($control.ActualHeight -lt 200 -or $control.ActualWidth -lt 200)) {
                    throw "CT plane is compressed at $($size[0]) x $($size[1]) in $mode : image=$($control.ActualWidth)x$($control.ActualHeight), view=$($view.ActualWidth)x$($view.ActualHeight)."
                }
                if ($control -is [Windows.Controls.Button] -and (-not $control.Focusable -or -not $control.IsTabStop)) {
                    throw 'A CT action is missing native keyboard focus.'
                }
                [pscustomobject]@{ type = $control.GetType().Name; text = [string]$control.Text; width = $control.ActualWidth; height = $control.ActualHeight; right = $bounds.Right; bottom = $bounds.Bottom }
            }
            $viewportBounds = $ctViewport.TransformToAncestor($captureRoot).TransformBounds([Windows.Rect]::new([Windows.Point]::new(0, 0), $ctViewport.RenderSize))
            if ($viewportBounds.Bottom -gt ($captureRoot.ActualHeight + 0.5) -or $viewportBounds.Right -gt ($captureRoot.ActualWidth + 0.5)) { throw 'The scrollable CT viewport overflows the window.' }
            if ($mode -eq 'three-planes' -and $ctViewport.ScrollableHeight -gt 1) { throw 'All three planes must fit simultaneously without scrolling.' }
            if ($mode -eq 'enlarged-coronal' -and $ctViewport.ScrollableHeight -gt 0) {
                $ctViewport.ScrollToEnd()
                Invoke-RenderPass
                if ($ctViewport.VerticalOffset -lt ($ctViewport.ScrollableHeight - 1)) { throw 'Enlarged CT bottom caption cannot be reached by native scrolling.' }
                $ctViewport.ScrollToVerticalOffset($ctViewport.ScrollableHeight / 2)
                Invoke-RenderPass
            }
            $focusButton = $visuals | Where-Object { $_ -is [Windows.Controls.Button] -and $_.Content -eq 'Gesamte CT' } | Select-Object -First 1
            [void][Windows.Input.Keyboard]::Focus($focusButton)
            Invoke-RenderPass
            if (-not $focusButton.IsKeyboardFocused) { throw 'Native CT keyboard-focus transfer failed.' }
            $label = '{0}x{1}' -f $size[0], $size[1]
            $png = Join-Path $outputRoot ('ct-' + $mode + '-' + $label + '.png')
            $bitmap = New-Object Windows.Media.Imaging.RenderTargetBitmap([int][Math]::Ceiling($captureRoot.ActualWidth), [int][Math]::Ceiling($captureRoot.ActualHeight), 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
            $bitmap.Render($captureRoot)
            [ClearPlan.Simulator.CapturePixelValidator]::AssertNonBlank($bitmap)
            $encoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
            $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
            $stream = [IO.File]::Create($png)
            try { $encoder.Save($stream) } finally { $stream.Dispose() }
            [pscustomobject]@{ synthetic = $true; analyticOverlayFixture = $true; mode = $mode; requestedSize = $label; imageCount = $imageControls.Count; rasterWidth = $imageControls[0].ActualWidth; scrollableHeight = $ctViewport.ScrollableHeight; hasKeyboardFocus = $focusButton.IsKeyboardFocused; controls = @($measurements); pngPath = $png }
        }
    }
    $compactOverview = $results | Where-Object { $_.requestedSize -eq '1180x720' -and $_.mode -eq 'three-planes' }
    $compactEnlarged = $results | Where-Object { $_.requestedSize -eq '1180x720' -and $_.mode -eq 'enlarged-coronal' }
    if ($compactEnlarged.rasterWidth -lt (1.5 * $compactOverview.rasterWidth)) {
        throw 'Enlarging one CT plane must enlarge the actual raster at compact window size, not only widen its background.'
    }
    # Real bound commands and toggles must preserve source data and restore the three-plane view.
    $vm.ShowDose = $false
    $vm.ShowStructures = $false
    $vm.FitCommand.Execute($null)
    Invoke-RenderPass
    if ($vm.LegendItems.Count -ne 0 -or $vm.FocusIsocenter) { throw 'CT-only fit mode failed.' }
    $vm.ShowDose = $true
    $vm.ShowStructures = $true
    $vm.FocusIsocenterCommand.Execute($null)
    $vm.ShowAllPlanesCommand.Execute($null)
    if ($vm.LegendItems.Count -ne 5 -or @($vm.VisiblePlanes).Count -ne 3) { throw 'CT overlay/three-plane restoration failed.' }
    if ($images[0].Overlays.Count -ne 6 -or $null -eq $images[0].DoseFocusRegion) { throw 'Display controls mutated synthetic source data.' }
    [IO.File]::WriteAllText((Join-Path $outputRoot 'ct-workspace-layout.json'), (ConvertTo-Json -InputObject @($results) -Depth 6), (New-Object Text.UTF8Encoding($false)))
    $results | Select-Object mode, requestedSize, imageCount, hasKeyboardFocus, pngPath | Format-Table -AutoSize
    Write-Output 'PASS four synthetic CT workspace captures, viewport/control bounds, keyboard focus, overlays, fit and enlarge/restore'
}
finally {
    $window.Close()
    $application.Shutdown()
}
