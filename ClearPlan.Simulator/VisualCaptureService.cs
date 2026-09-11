using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Presentation.Views;

namespace ClearPlan.Simulator
{
    public sealed class VisualCaptureService
    {
        public const int CaptureWidth = 1600;
        public const int CaptureHeight = 1000;

        private static readonly IDictionary<string, string> TabNames =
            new Dictionary<string, string>(
                StringComparer.Ordinal)
            {
                { "overview", "OverviewTab" },
                { "pqm", "PqmTab" },
                { "plancheck", "PlanCheckTab" },
                { "fields", "FieldsTab" },
                { "dvh", "DvhTab" },
                { "images", "PlanImagesTab" },
                { "parameters", "PlanParametersTab" },
                { "comparison", "ComparisonTab" },
                { "bev", "BevTab" },
                { "warnings", "WarningsTab" }
            };

        public async Task CaptureAsync(
            FrameworkElement captureRoot,
            ReviewWorkspaceView workspace,
            string tabId,
            string outputPath)
        {
            if (captureRoot == null)
            {
                throw new ArgumentNullException("captureRoot");
            }
            if (workspace == null)
            {
                throw new ArgumentNullException("workspace");
            }
            if (!TabNames.ContainsKey(tabId))
            {
                throw new ArgumentException(
                    "Unknown simulator tab: " + tabId,
                    "tabId");
            }

            SimulatorArguments.ValidateLocalOutputFile(
                outputPath,
                ".png",
                "PNG capture");
            SelectTab(workspace, tabId);
            var activeWorkspace = workspace.DataContext as ReviewWorkspaceViewModel;
            if (tabId == "bev" && activeWorkspace != null) await activeWorkspace.Analysis.ActivateBevAsync();
            if (tabId == "comparison" && activeWorkspace != null && activeWorkspace.IsSynthetic && !activeWorkspace.Comparison.HasReference)
                activeWorkspace.Comparison.LoadExampleCommand.Execute(null);
            Keyboard.ClearFocus();
            PreparePlotModels(workspace);
            captureRoot.UpdateLayout();
            await WaitForDispatcherAsync(
                captureRoot.Dispatcher,
                DispatcherPriority.Loaded);
            await WaitForDispatcherAsync(
                captureRoot.Dispatcher,
                DispatcherPriority.Render);
            await WaitForDispatcherAsync(
                captureRoot.Dispatcher,
                DispatcherPriority.ApplicationIdle);
            PreparePlotModels(workspace);
            captureRoot.UpdateLayout();
            if (string.Equals(
                tabId,
                "overview",
                StringComparison.Ordinal))
            {
                AssertOverviewTablesFillViewport(workspace);
            }

            string fullPath = Path.GetFullPath(outputPath);
            string parent = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrWhiteSpace(parent))
            {
                throw new InvalidOperationException(
                    "The PNG output directory could not be resolved.");
            }
            Directory.CreateDirectory(parent);

            var bitmap = new RenderTargetBitmap(
                CaptureWidth,
                CaptureHeight,
                96.0,
                96.0,
                PixelFormats.Pbgra32);
            bitmap.Render(captureRoot);
            CapturePixelValidator.AssertNonBlank(bitmap);

            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bitmap));
            using (var output = new FileStream(
                fullPath,
                FileMode.Create,
                FileAccess.Write,
                FileShare.None))
            {
                encoder.Save(output);
            }
        }

        public static void SelectTab(
            ReviewWorkspaceView workspace,
            string tabId)
        {
            var navigation = workspace.FindName(
                "WorkspaceNavigation") as TabControl;
            var target = workspace.FindName(
                TabNames[tabId]) as TabItem;
            if (navigation == null || target == null)
            {
                throw new InvalidOperationException(
                    "The shared workspace does not expose the required tab.");
            }

            navigation.SelectedItem = target;
            target.IsSelected = true;
            target.BringIntoView();
        }

        private static void PreparePlotModels(
            ReviewWorkspaceView workspace)
        {
            var viewModel =
                workspace.DataContext as ReviewWorkspaceViewModel;
            if (viewModel == null)
            {
                return;
            }

            viewModel.OverviewPlotModel.InvalidatePlot(true);
            viewModel.DetailPlotModel.InvalidatePlot(true);
            viewModel.Analysis.DoseRatePlotModel.InvalidatePlot(true);
            viewModel.Analysis.AperturePlotModel.InvalidatePlot(true);
            viewModel.Comparison.DvhPlotModel.InvalidatePlot(true);
        }

        private static void AssertOverviewTablesFillViewport(
            ReviewWorkspaceView workspace)
        {
            string[] gridNames =
            {
                "OverviewSourceGrid",
                "OverviewPlanGrid",
                "OverviewPqmGrid",
                "OverviewPlanCheckGrid",
                "OverviewFieldGrid",
                "OverviewStructureMappingGrid"
            };

            foreach (string gridName in gridNames)
            {
                var grid = workspace.FindName(gridName) as DataGrid;
                if (grid == null)
                {
                    throw new InvalidOperationException(
                        "Missing overview table: " + gridName);
                }

                double columnWidth = 0.0;
                foreach (DataGridColumn column in grid.Columns)
                {
                    columnWidth += column.ActualWidth;
                }

                if (grid.ActualWidth < workspace.ActualWidth * 0.60 ||
                    columnWidth < grid.ActualWidth * 0.80)
                {
                    throw new InvalidOperationException(
                        "Overview table did not fill the finite viewport: " +
                        gridName);
                }
            }
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
