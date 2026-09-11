using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClearPlan.Presentation.Views;

namespace ClearPlan.Simulator
{
    public sealed class ResponsiveLayoutProbeService
    {
        private const double VisibilityTolerance = 0.5;

        public async Task WriteAsync(
            Window window,
            FrameworkElement captureRoot,
            ReviewWorkspaceView workspace,
            double requestedWidth,
            double requestedHeight,
            string outputPath)
        {
            if (window == null)
            {
                throw new ArgumentNullException("window");
            }
            if (captureRoot == null)
            {
                throw new ArgumentNullException("captureRoot");
            }
            if (workspace == null)
            {
                throw new ArgumentNullException("workspace");
            }

            SimulatorArguments.ValidateLocalOutputFile(
                outputPath,
                ".json",
                "--output");
            VisualCaptureService.SelectTab(workspace, "dvh");
            captureRoot.UpdateLayout();
            workspace.UpdateLayout();
            await WaitForDispatcherAsync(
                window.Dispatcher,
                DispatcherPriority.Loaded);
            await WaitForDispatcherAsync(
                window.Dispatcher,
                DispatcherPriority.Render);
            await WaitForDispatcherAsync(
                window.Dispatcher,
                DispatcherPriority.ApplicationIdle);
            captureRoot.UpdateLayout();
            workspace.UpdateLayout();

            var navigation = FindRequiredElement<TabControl>(
                workspace,
                "WorkspaceNavigation");
            navigation.ApplyTemplate();
            var footer = FindRequiredTemplateElement<FrameworkElement>(
                navigation,
                "NavigationFooter");
            var sourceStatus =
                FindRequiredTemplateElement<FrameworkElement>(
                    navigation,
                    "NavigationSourceStatusText");
            var dvhContent = FindRequiredElement<Grid>(
                workspace,
                "DvhContentGrid");
            var dvhStructurePanel = FindRequiredElement<FrameworkElement>(
                workspace,
                "DvhStructurePanel");
            var dvhPlot = FindRequiredElement<FrameworkElement>(
                workspace,
                "DvhDetailPlot");

            Rect workspaceBounds = GetBounds(
                workspace,
                captureRoot);
            Rect footerBounds = GetBounds(
                footer,
                captureRoot);
            Rect sourceStatusBounds = GetBounds(
                sourceStatus,
                captureRoot);
            bool workspaceFitsViewport = IsFullyVisible(
                workspaceBounds,
                captureRoot);
            bool footerFullyVisible =
                IsFullyVisible(footerBounds, captureRoot) &&
                IsFullyVisibleWithin(
                    footerBounds,
                    workspaceBounds);
            bool sourceStatusFullyVisible =
                IsFullyVisible(sourceStatusBounds, captureRoot) &&
                IsFullyVisibleWithin(
                    sourceStatusBounds,
                    workspaceBounds);

            string fullPath = Path.GetFullPath(outputPath);
            string directory = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidOperationException(
                    "The layout-probe output directory could not be resolved.");
            }
            Directory.CreateDirectory(directory);
            File.WriteAllText(
                fullPath,
                SerializeResult(
                    requestedWidth,
                    requestedHeight,
                    window,
                    captureRoot,
                    workspace,
                    workspaceBounds,
                    footerBounds,
                    sourceStatusBounds,
                    dvhContent,
                    dvhStructurePanel,
                    dvhPlot,
                    workspaceFitsViewport,
                    footerFullyVisible,
                    sourceStatusFullyVisible),
                new UTF8Encoding(false));

            // Capture the new scrolling review surfaces at the real requested
            // workstation size, not the fixed publication capture canvas.
            foreach (string tab in new[] { "parameters", "comparison", "bev" })
            {
                VisualCaptureService.SelectTab(workspace, tab);
                var model = workspace.DataContext as ClearPlan.Presentation.ViewModels.ReviewWorkspaceViewModel;
                if (tab == "bev" && model != null) await model.Analysis.ActivateBevAsync();
                if (tab == "comparison" && model != null && model.IsSynthetic && !model.Comparison.HasReference)
                    model.Comparison.LoadExampleCommand.Execute(null);
                captureRoot.UpdateLayout();
                await WaitForDispatcherAsync(window.Dispatcher, DispatcherPriority.ApplicationIdle);
                captureRoot.UpdateLayout();
                var bitmap = new RenderTargetBitmap(
                    (int)Math.Ceiling(captureRoot.ActualWidth), (int)Math.Ceiling(captureRoot.ActualHeight),
                    96.0, 96.0, PixelFormats.Pbgra32);
                bitmap.Render(captureRoot);
                CapturePixelValidator.AssertNonBlank(bitmap);
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(bitmap));
                using (var stream = File.Create(Path.Combine(directory,
                    Path.GetFileNameWithoutExtension(fullPath) + "-" + tab + ".png")))
                    encoder.Save(stream);
            }
        }

        private static T FindRequiredElement<T>(
            FrameworkElement scope,
            string name)
            where T : FrameworkElement
        {
            T element = scope.FindName(name) as T;
            if (element == null)
            {
                throw new InvalidOperationException(
                    "The responsive layout probe cannot find: " + name);
            }

            return element;
        }

        private static T FindRequiredTemplateElement<T>(
            Control control,
            string name)
            where T : FrameworkElement
        {
            T element = control.Template.FindName(
                name,
                control) as T;
            if (element == null)
            {
                throw new InvalidOperationException(
                    "The responsive layout probe cannot find: " + name);
            }

            return element;
        }

        private static Rect GetBounds(
            FrameworkElement element,
            FrameworkElement ancestor)
        {
            return element
                .TransformToAncestor(ancestor)
                .TransformBounds(
                    new Rect(
                        new Point(0.0, 0.0),
                        element.RenderSize));
        }

        private static bool IsFullyVisible(
            Rect bounds,
            FrameworkElement viewport)
        {
            return bounds.Left >= -VisibilityTolerance &&
                   bounds.Top >= -VisibilityTolerance &&
                   bounds.Right <=
                       viewport.ActualWidth + VisibilityTolerance &&
                   bounds.Bottom <=
                       viewport.ActualHeight + VisibilityTolerance;
        }

        private static bool IsFullyVisibleWithin(
            Rect bounds,
            Rect containingBounds)
        {
            return bounds.Left >=
                       containingBounds.Left - VisibilityTolerance &&
                   bounds.Top >=
                       containingBounds.Top - VisibilityTolerance &&
                   bounds.Right <=
                       containingBounds.Right + VisibilityTolerance &&
                   bounds.Bottom <=
                       containingBounds.Bottom + VisibilityTolerance;
        }

        private static string SerializeResult(
            double requestedWidth,
            double requestedHeight,
            Window window,
            FrameworkElement captureRoot,
            ReviewWorkspaceView workspace,
            Rect workspaceBounds,
            Rect footerBounds,
            Rect sourceStatusBounds,
            FrameworkElement dvhContent,
            FrameworkElement dvhStructurePanel,
            FrameworkElement dvhPlot,
            bool workspaceFitsViewport,
            bool footerFullyVisible,
            bool sourceStatusFullyVisible)
        {
            var result = new
            {
                requestedWidth,
                requestedHeight,
                windowActualWidth = window.ActualWidth,
                windowActualHeight = window.ActualHeight,
                captureRootWidth = captureRoot.ActualWidth,
                captureRootHeight = captureRoot.ActualHeight,
                workspaceViewportWidth = workspace.ActualWidth,
                workspaceViewportHeight = workspace.ActualHeight,
                workspaceTop = workspaceBounds.Top,
                workspaceBottom = workspaceBounds.Bottom,
                workspaceFitsViewport,
                footerTop = footerBounds.Top,
                footerBottom = footerBounds.Bottom,
                footerFullyVisible,
                sourceStatusHeight = sourceStatusBounds.Height,
                sourceStatusBottom = sourceStatusBounds.Bottom,
                sourceStatusFullyVisible,
                dvhContentWidth = dvhContent.ActualWidth,
                dvhContentHeight = dvhContent.ActualHeight,
                dvhStructurePanelWidth = dvhStructurePanel.ActualWidth,
                dvhPlotWidth = dvhPlot.ActualWidth,
                dvhPlotHeight = dvhPlot.ActualHeight,
                dvhHorizontalOverflowAvailable = false,
                dvhVerticalOverflowAvailable = false
            };
            return new JavaScriptSerializer().Serialize(result);
        }

        private static Task WaitForDispatcherAsync(
            Dispatcher dispatcher,
            DispatcherPriority priority)
        {
            var completion = new TaskCompletionSource<bool>();
            dispatcher.BeginInvoke(
                new Action(() => completion.SetResult(true)),
                priority);
            return completion.Task;
        }
    }
}
