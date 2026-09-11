[CmdletBinding()]
param([string]$AssemblyDirectory = '', [string]$ArtifactsDirectory = '')

# Detached synthetic WPF probe: no ESAPI, patient data or clinical application.
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run with powershell.exe -STA -File.' }
$repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $AssemblyDirectory) { $AssemblyDirectory = Join-Path $repo 'artifacts\bin\Debug' }
if (-not $ArtifactsDirectory) { $ArtifactsDirectory = Join-Path $repo 'artifacts\status-hover-check' }
$assemblyRoot = [IO.Path]::GetFullPath($AssemblyDirectory)
$output = [IO.Path]::GetFullPath($ArtifactsDirectory)
if ($output.StartsWith('\\')) { throw 'Use a local explicit artifact directory.' }
Add-Type -AssemblyName WindowsBase, PresentationCore, PresentationFramework
[AppContext]::SetSwitch('Switch.System.Windows.Media.ShouldRenderEvenWhenNoDisplayDevicesAreAvailable', $true)
$references = @('System.dll', 'System.Core.dll', 'System.Xaml.dll',
    [Windows.Threading.Dispatcher].Assembly.Location,
    [Windows.Application].Assembly.Location, [Windows.Media.Visual].Assembly.Location)
foreach ($name in @('Newtonsoft.Json.dll', 'OxyPlot.dll', 'OxyPlot.Wpf.dll', 'ClearPlan.Core.dll', 'ClearPlan.Rendering.dll', 'ClearPlan.Presentation.dll')) {
    $path = Join-Path $assemblyRoot $name
    [void][Reflection.Assembly]::LoadFrom($path)
    $references += $path
}
$source = @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Presentation.Views;
using OxyPlot;
using OxyPlot.Wpf;

public static class StatusDvhContrastProbe
{
    public static void Run(string output)
    {
        var source = new ReviewWorkspaceView();
        var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
        snapshot.DvhSeries = new List<ReviewDvhSeries> {
            new ReviewDvhSeries {
                StableId = "synthetic-tracker", StructureId = "SYNTHETIC PTV", DisplayName = "SYNTHETIC PTV",
                Selected = true, ColorHex = "#0F766E", Points = new List<ReviewDvhPoint> {
                    new ReviewDvhPoint { DoseGy = 0, VolumePercent = 100 },
                    new ReviewDvhPoint { DoseGy = 50, VolumePercent = 0 }
                }
            }
        };
        var workspace = new ReviewWorkspaceViewModel(snapshot);
        var grid = new Grid { Width = 1160, Height = 440, Background = Brushes.White };
        grid.ColumnDefinitions.Add(new ColumnDefinition());
        grid.ColumnDefinitions.Add(new ColumnDefinition());
        var plots = new List<PlotView>();
        for (int index = 0; index < 2; index++)
        {
            var original = (PlotView)source.FindName(index == 0 ? "OverviewDvhPlot" : "DvhDetailPlot");
            if (original.DefaultTrackerTemplate == null) throw new InvalidOperationException("Missing shared tracker template.");
            var plot = new PlotView {
                DefaultTrackerTemplate = original.DefaultTrackerTemplate,
                Model = index == 0 ? workspace.OverviewPlotModel : workspace.DetailPlotModel,
                Controller = index == 0 ? workspace.OverviewPlotController : workspace.DetailPlotController,
                Foreground = Brushes.White, // Reproduce hostile inherited text from the native host.
                Margin = new Thickness(14, 46, 14, 12)
            };
            var heading = new TextBlock {
                Text = index == 0 ? "SYNTHETIC - Overview DVH hover" : "SYNTHETIC - Detail DVH hover",
                FontFamily = new FontFamily("Segoe UI"), FontSize = 17, Foreground = Brushes.Black,
                VerticalAlignment = System.Windows.VerticalAlignment.Top, Margin = new Thickness(18, 12, 0, 0)
            };
            Grid.SetColumn(plot, index); Grid.SetColumn(heading, index);
            grid.Children.Add(plot); grid.Children.Add(heading); plots.Add(plot);
        }
        var app = Application.Current ?? new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
        var window = new Window { Content = grid, Width = 1180, Height = 480, ShowInTaskbar = false,
            ShowActivated = false, WindowStartupLocation = WindowStartupLocation.Manual, Left = -20000, Top = 0 };
        try
        {
            window.Show(); window.UpdateLayout();
            window.Dispatcher.Invoke(new Action(() => {}), DispatcherPriority.ApplicationIdle);
            foreach (var plot in plots)
            {
                var line = (OxyPlot.Series.LineSeries)plot.Model.Series[0];
                var position = line.Transform(20, 60);
                plot.Controller.HandleMouseEnter(plot, new OxyMouseEventArgs { Position = position });
                plot.Controller.HandleMouseMove(plot, new OxyMouseEventArgs { Position = position });
            }
            window.UpdateLayout();
            window.Dispatcher.Invoke(new Action(() => {}), DispatcherPriority.ApplicationIdle);
            foreach (var plot in plots)
            {
                var tracker = Descendants(plot).OfType<TrackerControl>().Single();
                var text = Descendants(tracker).OfType<TextBlock>().Single(t => t.Text.Contains("SYNTHETIC PTV"));
                var background = (SolidColorBrush)tracker.Background;
                var foreground = (SolidColorBrush)text.Foreground;
                double contrast = (Luminance(background.Color) + 0.05) / (Luminance(foreground.Color) + 0.05);
                if (background.Color.A != 255 || background.Opacity != 1 || contrast < 4.5)
                    throw new InvalidOperationException("Opaque readable tracker required; contrast=" + contrast);
                if (!text.Text.Contains("20.00") && !text.Text.Contains("20,00"))
                    throw new InvalidOperationException("Hover must interpolate the synthetic 20 Gy sample.");
                Console.WriteLine("PASS WPF hover: {0}; background={1}; foreground={2}; contrast={3:0.00}:1", text.Text.Replace("\n", " / "), background.Color, foreground.Color, contrast);
            }
            Directory.CreateDirectory(output);
            var bitmap = new RenderTargetBitmap(1160, 440, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(grid);
            var encoder = new PngBitmapEncoder(); encoder.Frames.Add(BitmapFrame.Create(bitmap));
            using (var stream = File.Create(Path.Combine(output, "synthetic-dvh-hover.png"))) encoder.Save(stream);
        }
        finally { window.Close(); }
    }

    private static IEnumerable<DependencyObject> Descendants(DependencyObject root)
    {
        for (int index = 0; index < VisualTreeHelper.GetChildrenCount(root); index++)
        {
            var child = VisualTreeHelper.GetChild(root, index); yield return child;
            foreach (var nested in Descendants(child)) yield return nested;
        }
    }
    private static double Luminance(Color color)
    {
        return Channel(color.R) * 0.2126 + Channel(color.G) * 0.7152 + Channel(color.B) * 0.0722;
    }
    private static double Channel(byte value)
    {
        double channel = value / 255.0;
        return channel <= 0.04045 ? channel / 12.92 : Math.Pow((channel + 0.055) / 1.055, 2.4);
    }
}
'@
Add-Type -TypeDefinition $source -ReferencedAssemblies $references
[StatusDvhContrastProbe]::Run($output)
Write-Host ('Synthetic tracker capture: ' + (Join-Path $output 'synthetic-dvh-hover.png'))
