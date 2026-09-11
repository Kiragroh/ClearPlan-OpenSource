using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;
using OxyPlot;
using OxyPlot.Series;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewDvhInteractionTests
    {
        public static void ClinicalResetUsesDetachedWorkspaceOnly()
        {
            string source = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.xaml.cs"));
            string body = Regex.Match(source,
                @"internal void HandleSharedDvhReset\([^)]*\)\s*\{(?<body>.*?)\n        \}",
                RegexOptions.Singleline).Groups["body"].Value;
            TestAssert.True(body.Contains("workspace.ResetDvhSelections()"),
                "DVH reset must restore detached selections without rebuilding native goals, DVHs or plan analysis.");
            TestAssert.False(body.Contains("RefreshClinicalReviewWorkspace") ||
                body.Contains("ApplyDefaultDvhSelections") || body.Contains("SyncSharedDvhSelections"),
                "Reset is a presentation action; native reads and export synchronization must not run here.");
        }

        public static void ResetBatchesRedrawAndRestoresAxesWithoutReplacingModels()
        {
            var workspace = CreateWorkspace();
            var overview = new TrackingView(workspace.OverviewPlotModel);
            var detail = new TrackingView(workspace.DetailPlotModel);
            workspace.ClearDvhSelectionCommand.Execute(null);
            foreach (var axis in workspace.OverviewPlotModel.Axes.Concat(workspace.DetailPlotModel.Axes))
                axis.Zoom(10, 20);
            overview.ClearCounts();
            detail.ClearCounts();
            var analysis = workspace.Analysis;
            var points = ((LineSeries)workspace.DetailPlotModel.Series[0]).Points;

            workspace.ResetDvhSelections();

            TestAssert.Equal(1, overview.Invalidations, "Bulk reset must invalidate the overview once, not once per structure.");
            TestAssert.Equal(1, detail.Invalidations, "Bulk reset must invalidate the detail view once, not once per structure.");
            TestAssert.False(overview.RequestedDataUpdate || detail.RequestedDataUpdate,
                "Reset does not change cached curve data and must not request data recalculation.");
            TestAssert.True(ReferenceEquals(analysis, workspace.Analysis));
            TestAssert.True(ReferenceEquals(overview.ActualModel, workspace.OverviewPlotModel));
            TestAssert.True(ReferenceEquals(detail.ActualModel, workspace.DetailPlotModel));
            TestAssert.False(ReferenceEquals(workspace.OverviewPlotModel, workspace.DetailPlotModel),
                "OxyPlot permits only one PlotView owner per model.");
            TestAssert.True(ReferenceEquals(points, ((LineSeries)workspace.DetailPlotModel.Series[0]).Points));
            foreach (var series in workspace.DvhSeries)
                TestAssert.Equal(series.IsInitiallySelected, series.IsSelected);
            AssertSeriesVisibility(workspace);
            overview.Render();
            detail.Render();
            foreach (var axis in workspace.OverviewPlotModel.Axes.Concat(workspace.DetailPlotModel.Axes))
                TestAssert.Equal(0.0, axis.ActualMinimum, "Reset must clear dose/volume zoom on both independent plots.");
        }

        public static void BulkSelectionRefreshesOnceAndKeepsRequiredTargets()
        {
            var workspace = CreateWorkspace();
            var overview = new TrackingView(workspace.OverviewPlotModel);
            var detail = new TrackingView(workspace.DetailPlotModel);
            workspace.ShowSelectedDvhOnly = true;
            int filterRefreshes = 0;
            workspace.FilteredDvhSeries.CollectionChanged += (sender, args) => filterRefreshes++;
            workspace.SelectAllDvhCommand.Execute(null);
            TestAssert.Equal(1, overview.Invalidations);
            TestAssert.Equal(1, detail.Invalidations);
            TestAssert.Equal(1, filterRefreshes, "Bulk selection must refresh the filtered structure list once.");
            AssertSeriesVisibility(workspace);
            overview.ClearCounts();
            detail.ClearCounts();
            filterRefreshes = 0;
            workspace.ClearDvhSelectionCommand.Execute(null);
            TestAssert.Equal(1, overview.Invalidations);
            TestAssert.Equal(1, detail.Invalidations);
            TestAssert.Equal(1, filterRefreshes);
            TestAssert.True(workspace.DvhSeries.Where(row => row.RequiredForTargetReview).All(row => row.IsSelected));
            TestAssert.True(workspace.DvhSeries.Where(row => !row.RequiredForTargetReview).All(row => !row.IsSelected));
            AssertSeriesVisibility(workspace);
        }

        public static void BothPlotViewsBindIndependentHoverControllers()
        {
            var workspace = CreateWorkspace();
            PlotController overview = GetController(workspace, "OverviewPlotController");
            PlotController detail = GetController(workspace, "DetailPlotController");
            TestAssert.False(ReferenceEquals(overview, detail),
                "Each plot needs an independent controller because hover manipulators own view-specific state.");
            string xaml = File.ReadAllText(Path.Combine("ClearPlan.Presentation", "Views", "ReviewWorkspaceView.xaml"));
            foreach (string name in new[] { "OverviewDvhPlot", "DvhDetailPlot" })
            {
                string element = Regex.Match(xaml, "<oxy:PlotView\\b[^>]*x:Name=\"" + name + "\"[^>]*>").Value;
                string expected = name == "OverviewDvhPlot" ? "OverviewPlotController" : "DetailPlotController";
                TestAssert.True(element.Contains("Controller=\"{Binding " + expected + "}\""),
                    name + " must bind the hover controller.");
            }
        }

        public static void HoverShowsStructureDoseAndVolumeWithoutMouseDown()
        {
            var workspace = CreateWorkspace();
            foreach (bool overview in new[] { true, false })
            {
                var model = overview ? workspace.OverviewPlotModel : workspace.DetailPlotModel;
                var controller = GetController(workspace, overview ? "OverviewPlotController" : "DetailPlotController");
                var view = new TrackingView(model) { ActualController = controller };
                view.Render();
                var line = (LineSeries)model.Series[0];
                var position = line.Transform(20, 60);
                controller.HandleMouseEnter(view, new OxyMouseEventArgs { Position = position });
                controller.HandleMouseMove(view, new OxyMouseEventArgs { Position = position });
                TestAssert.NotNull(view.Tracker, "Moving over a DVH curve must show the tracker without a mouse click.");
                TestAssert.True(view.Tracker.Text.Contains("Structure 0") &&
                    view.Tracker.Text.Contains("Dosis:") && view.Tracker.Text.Contains("Gy") &&
                    view.Tracker.Text.Contains("Volumen:") && view.Tracker.Text.Contains("%"),
                    "Tracker must identify structure/dose/volume: " + view.Tracker.Text);
                TestAssert.True(Math.Abs(view.Tracker.DataPoint.X - 20) < 0.001,
                    "Tracker must interpolate 20 Gy between dose bins, got " + view.Tracker.DataPoint.X);
                TestAssert.True(Math.Abs(view.Tracker.DataPoint.Y - 60) < 0.001,
                    "Tracker must interpolate 60 percent between dose bins, got " + view.Tracker.DataPoint.Y);
                controller.HandleMouseLeave(view, new OxyMouseEventArgs { Position = position });
                TestAssert.True(view.Tracker == null, "Leaving a plot must hide its tracker.");
            }
        }

        private static PlotController GetController(ReviewWorkspaceViewModel workspace, string name)
        {
            PropertyInfo property = workspace.GetType().GetProperty(name);
            TestAssert.NotNull(property, name + " is missing; the default OxyPlot controller requires a mouse click.");
            var controller = property.GetValue(workspace, null) as PlotController;
            TestAssert.NotNull(controller);
            return controller;
        }

        private static ReviewWorkspaceViewModel CreateWorkspace()
        {
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            snapshot.DvhSeries = Enumerable.Range(0, 12).Select(index => new ReviewDvhSeries
            {
                StableId = "dvh-" + index,
                StructureId = "Structure " + index,
                DisplayName = "Structure " + index,
                Selected = index % 2 == 0,
                RequiredForTargetReview = index == 0,
                Points = new List<ReviewDvhPoint>
                {
                    new ReviewDvhPoint { DoseGy = 0, VolumePercent = 100 },
                    new ReviewDvhPoint { DoseGy = 50 + index * 5, VolumePercent = 0 }
                }
            }).ToList();
            return new ReviewWorkspaceViewModel(snapshot);
        }

        private static void AssertSeriesVisibility(ReviewWorkspaceViewModel workspace)
        {
            foreach (var model in new[] { workspace.OverviewPlotModel, workspace.DetailPlotModel })
                foreach (var line in model.Series)
                    TestAssert.Equal(workspace.DvhSeries.Single(row => row.StableId == (string)line.Tag).IsSelected, line.IsVisible);
        }

        // Captures view notifications without requiring a desktop, mouse or clinical data.
        private sealed class TrackingView : IPlotView
        {
            public TrackingView(PlotModel model)
            {
                ActualModel = model;
                ((IPlotModel)model).AttachPlotView(this);
            }
            public PlotModel ActualModel { get; private set; }
            Model IView.ActualModel { get { return ActualModel; } }
            public IController ActualController { get; set; }
            public OxyRect ClientArea { get { return new OxyRect(0, 0, 800, 500); } }
            public int Invalidations { get; private set; }
            public bool RequestedDataUpdate { get; private set; }
            public TrackerHitResult Tracker { get; private set; }
            public void InvalidatePlot(bool updateData) { Invalidations++; RequestedDataUpdate |= updateData; }
            public void ShowTracker(TrackerHitResult result) { Tracker = result; }
            public void HideTracker() { Tracker = null; }
            public void SetClipboardText(string text) { }
            public void SetCursorType(CursorType cursorType) { }
            public void HideZoomRectangle() { }
            public void ShowZoomRectangle(OxyRect rectangle) { }
            public void ClearCounts() { Invalidations = 0; RequestedDataUpdate = false; }
            public void Render()
            {
                ((IPlotModel)ActualModel).Update(true);
                ((IPlotModel)ActualModel).Render(new MeasurementRenderContext(), 800, 500);
            }
        }

        private sealed class MeasurementRenderContext : RenderContextBase
        {
            public override void DrawLine(IList<ScreenPoint> points, OxyColor color, double thickness, double[] dashArray, LineJoin lineJoin, bool aliased) { }
            public override void DrawPolygon(IList<ScreenPoint> points, OxyColor fill, OxyColor stroke, double thickness, double[] dashArray, LineJoin lineJoin, bool aliased) { }
            public override void DrawText(ScreenPoint point, string text, OxyColor color, string fontFamily, double fontSize, double fontWeight, double rotation, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment, OxySize? maxSize) { }
            public override OxySize MeasureText(string text, string fontFamily, double fontSize, double fontWeight)
            { return new OxySize((text ?? string.Empty).Length * fontSize * 0.5, fontSize); }
        }
    }
}
