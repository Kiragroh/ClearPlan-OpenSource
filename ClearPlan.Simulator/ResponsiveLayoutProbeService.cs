using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
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
            var dvhScroller = FindRequiredElement<ScrollViewer>(
                workspace,
                "DvhOverflowScrollViewer");

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
                    dvhScroller,
                    workspaceFitsViewport,
                    footerFullyVisible,
                    sourceStatusFullyVisible),
                new UTF8Encoding(false));
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
            ScrollViewer dvhScroller,
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
                dvhViewportWidth = dvhScroller.ViewportWidth,
                dvhViewportHeight = dvhScroller.ViewportHeight,
                dvhExtentWidth = dvhScroller.ExtentWidth,
                dvhExtentHeight = dvhScroller.ExtentHeight,
                dvhScrollableWidth = dvhScroller.ScrollableWidth,
                dvhScrollableHeight = dvhScroller.ScrollableHeight,
                dvhHorizontalOverflowAvailable =
                    dvhScroller.ScrollableWidth > VisibilityTolerance,
                dvhVerticalOverflowAvailable =
                    dvhScroller.ScrollableHeight > VisibilityTolerance
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
