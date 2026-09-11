[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$RepositoryDirectory,
    [Parameter(Mandatory = $true)][string]$ArtifactsDirectory
)

# Measures the actual native header XAML and theme with synthetic text only.
# No ESAPI objects, patient context or native event handlers are instantiated.
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    throw 'Run with Windows PowerShell -STA.'
}
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
[AppContext]::SetSwitch('Switch.System.Windows.Media.ShouldRenderEvenWhenNoDisplayDevicesAreAvailable', $true)
$repo = [IO.Path]::GetFullPath($RepositoryDirectory)
$output = [IO.Path]::GetFullPath($ArtifactsDirectory)
if ($output.StartsWith('\\')) { throw 'Use a local explicit artifact directory.' }
[void][IO.Directory]::CreateDirectory($output)
[xml]$xaml = Get-Content -LiteralPath (Join-Path $repo 'ClearPlan.Script\MainView.xaml') -Raw -Encoding UTF8
$ns = New-Object Xml.XmlNamespaceManager($xaml.NameTable)
$ns.AddNamespace('p', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
$ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
$xaml.DocumentElement.RemoveAttribute('Class', 'http://schemas.microsoft.com/winfx/2006/xaml')
$rootGrid = $xaml.SelectSingleNode('/p:UserControl/p:Grid', $ns)
foreach ($node in @($rootGrid.SelectNodes('p:Grid', $ns))) { [void]$rootGrid.RemoveChild($node) }
$resource = $xaml.SelectSingleNode('//p:ResourceDictionary[@Source]', $ns)
$resource.SetAttribute('Source', (Join-Path $repo 'ClearPlan.Presentation\Styles\ClinicalBlueprint.xaml'))
foreach ($node in @($xaml.SelectNodes('//*'))) {
    foreach ($eventName in @('Click', 'SelectionChanged')) { $node.RemoveAttribute($eventName) }
}
$view = [Windows.Markup.XamlReader]::Parse($xaml.OuterXml)
$constraint = [pscustomobject]@{ ConstraintName = 'Synthetische Constraint-Tabelle mit langem Namen' }
$view.DataContext = [pscustomobject]@{
    ActivePlanningItem = [pscustomobject]@{ PlanningItemIdWithCourseAndType = 'SYNTHETIC-COURSE / SYNTHETIC-PLAN (Plan)' }
    ConstraintSourceStatus = 'Synthetischer Status: Tabelle muss vor der Auswertung erst noch freigegeben werden.'
    ActiveConstraintPath = $constraint
    ConstraintComboBoxList = @($constraint)
}
$application = New-Object Windows.Application
$application.ShutdownMode = 'OnExplicitShutdown'
$window = New-Object Windows.Window
$window.Content = $view
$window.ShowInTaskbar = $false
$window.ShowActivated = $false
$window.Left = -20000
$window.Top = 0
$window.Height = 720
$window.Width = 1180
$window.WindowStartupLocation = 'Manual'
try {
    $window.Show()
    $results = foreach ($width in @(1180, 1400, 1600)) {
        $window.Width = $width
        foreach ($mode in @('native-synthetic-fixture', 'sandbox')) {
            $sandbox = $mode -eq 'sandbox'
            foreach ($name in @('ClinicalContextText', 'ClinicalConstraintPanel', 'ClinicalConstraintStatus', 'HeaderConstraintRow', 'SwitchPlanButton', 'ClinicalExtrasMenuItem')) {
                $view.FindName($name).Visibility = if ($sandbox) { 'Collapsed' } else { 'Visible' }
            }
            $view.FindName('SyntheticContextText').Visibility = if ($sandbox) { 'Visible' } else { 'Collapsed' }
            $view.FindName('SyntheticDemoToggle').IsChecked = $sandbox
            $view.FindName('SharedWorkspaceStatusText').Text = if ($sandbox) {
                'Sandbox aktiv - GUI und PDF nur mit synthetischen Beispieldaten'
            } else { 'Synthetischer Hinweis: Die gesamte Statuszeile muss auch bei kleineren Fenstern sichtbar bleiben.' }
            $window.UpdateLayout()
            [void]$window.Dispatcher.Invoke([Action]{}, [Windows.Threading.DispatcherPriority]::ApplicationIdle)
            $window.UpdateLayout()
            $header = $view.FindName('NativeHeader')
            $measurements = foreach ($name in @('ClinicalContextText', 'SyntheticContextText', 'SharedWorkspaceStatusText', 'ConstraintTableLabel', 'SharedConstraintComboBox', 'ConfirmConstraintButton', 'SwitchPlanButton', 'SyntheticDemoToggle')) {
                $control = $view.FindName($name)
                if (-not $control.IsVisible) { continue }
                $bounds = $control.TransformToAncestor($header).TransformBounds([Windows.Rect]::new([Windows.Point]::new(0, 0), $control.RenderSize))
                if ($bounds.X -lt -0.5 -or $bounds.Y -lt -0.5 -or $bounds.Right -gt ($header.ActualWidth + 0.5) -or $bounds.Bottom -gt ($header.ActualHeight + 0.5)) {
                    throw ('Header clips ' + $name + ' at width ' + $width + ' in ' + $mode)
                }
                if ($control.ActualHeight + 0.5 -lt $control.DesiredSize.Height - $control.Margin.Top - $control.Margin.Bottom) {
                    throw ('Header compresses desired text/control height: ' + $name)
                }
                [pscustomobject]@{ control = $name; width = $control.ActualWidth; height = $control.ActualHeight; bottom = $bounds.Bottom }
            }
            $bitmap = New-Object Windows.Media.Imaging.RenderTargetBitmap([int][Math]::Ceiling($header.ActualWidth), [int][Math]::Ceiling($header.ActualHeight), 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
            $bitmap.Render($header)
            $encoder = New-Object Windows.Media.Imaging.PngBitmapEncoder
            $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
            $png = Join-Path $output ('native-header-' + $mode + '-' + $width + '.png')
            $stream = [IO.File]::Create($png)
            try { $encoder.Save($stream) } finally { $stream.Dispose() }
            [pscustomobject]@{ synthetic = $true; mode = $mode; windowWidth = $width; headerWidth = $header.ActualWidth; headerHeight = $header.ActualHeight; controls = @($measurements); pngPath = $png }
        }
    }
    [IO.File]::WriteAllText((Join-Path $output 'native-header-layout.json'), (ConvertTo-Json -InputObject @($results) -Depth 6), (New-Object Text.UTF8Encoding($false)))
    $results | Select-Object mode, windowWidth, headerWidth, headerHeight | Format-Table
    Write-Output 'PASS 6 native-header WPF measurements with synthetic fixtures'
}
finally {
    $window.Close()
    $application.Shutdown()
}
