using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Reporting;
using ClearPlan.Simulator;
using OxyPlot.Series;

namespace ClearPlan.Core.Tests
{
    internal static class SimulatorVisibleActionTests
    {
        public static void CommandsRaiseEveryVisibleActionEvent()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("baseline-pass");
            var viewModel = new ReviewWorkspaceViewModel(snapshot);
            var actions = new List<string>();
            object planParameter = viewModel.Plans[0];
            object mappingParameter = viewModel.StructureMappings[0];

            viewModel.ReportRequested +=
                (sender, args) => actions.Add(args.ActionCode);
            viewModel.OpenPlanRequested +=
                (sender, args) => actions.Add(args.ActionCode);
            viewModel.DvhResetRequested +=
                (sender, args) => actions.Add(args.ActionCode);
            viewModel.DvhExportRequested +=
                (sender, args) => actions.Add(args.ActionCode);
            viewModel.StructureMappingRequested +=
                (sender, args) => actions.Add(args.ActionCode);

            viewModel.ReportCommand.Execute(null);
            viewModel.OpenPlanCommand.Execute(planParameter);
            viewModel.ResetDvhCommand.Execute(null);
            viewModel.ExportDvhCommand.Execute(null);
            viewModel.ApplyStructureMappingCommand.Execute(
                mappingParameter);

            TestAssert.Equal(
                string.Join(
                    "|",
                    new[]
                    {
                        ReviewWorkspaceActionCodes.Report,
                        ReviewWorkspaceActionCodes.OpenPlan,
                        ReviewWorkspaceActionCodes.ResetDvh,
                        ReviewWorkspaceActionCodes.ExportDvh,
                        ReviewWorkspaceActionCodes.ApplyStructureMapping
                    }),
                string.Join("|", actions));
        }

        public static void ResetRestoresInitialSelectionAndPlotVisibility()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("baseline-pass");
            var viewModel = new ReviewWorkspaceViewModel(snapshot);
            IDictionary<string, bool> initial =
                viewModel.DvhSeries.ToDictionary(
                    row => row.StableId,
                    row => row.IsInitiallySelected);

            viewModel.ClearDvhSelectionCommand.Execute(null);
            viewModel.ResetDvhSelections();

            foreach (ReviewDvhSeriesViewModel row in viewModel.DvhSeries)
            {
                TestAssert.Equal(
                    initial[row.StableId],
                    row.IsSelected,
                    "DVH reset did not restore " + row.StableId);
                Series plotted = viewModel.DetailPlotModel.Series
                    .Single(
                        series => string.Equals(
                            Convert.ToString(series.Tag),
                            row.StableId,
                            StringComparison.Ordinal));
                TestAssert.Equal(
                    initial[row.StableId],
                    plotted.IsVisible,
                    "DVH reset did not update the plot for " +
                    row.StableId);
            }
        }

        public static void OpenPlanStatusIdentifiesTheSyntheticActivePlan()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("baseline-pass");
            ReviewPlanRow activePlan = snapshot.Plans.Single(
                row => row.PlanKey == snapshot.ActivePlanKey);

            string status =
                SimulatorSnapshotActions.GetOpenPlanStatus(
                    snapshot,
                    activePlan.PlanKey);

            TestAssert.True(
                status.IndexOf(
                    activePlan.DisplayLabel,
                    StringComparison.Ordinal) >= 0);
            TestAssert.True(
                status.IndexOf(
                    "synthetisch",
                    StringComparison.OrdinalIgnoreCase) >= 0);
            TestAssert.True(
                status.IndexOf(
                    "aktiv",
                    StringComparison.OrdinalIgnoreCase) >= 0);
        }

        public static void MappingPersistsIntoSnapshotAndReport()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("field-and-mapping");
            ReviewStructureMapping mapping = snapshot.StructureMappings[2];
            string selected = mapping.AvailableStructureIds.Last();

            string status =
                SimulatorSnapshotActions.ApplyStructureMapping(
                    snapshot,
                    mapping.StableId,
                    selected);

            TestAssert.Equal(selected, mapping.SelectedStructureId);
            TestAssert.Equal(ReviewStatusCodes.Pass, mapping.Status);
            TestAssert.True(
                status.IndexOf(
                    selected,
                    StringComparison.Ordinal) >= 0);
            ReviewReportStructureMappingRow reportRow =
                new ReviewSnapshotReportMapper()
                    .Map(snapshot)
                    .StructureMappings
                    .Single(row => row.StableId == mapping.StableId);
            TestAssert.Equal(selected, reportRow.SelectedStructureId);
            TestAssert.Equal(ReviewStatusCodes.Pass, reportRow.Status);
        }

        public static void MappingRejectsUnknownChoicesWithoutMutation()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("field-and-mapping");
            ReviewStructureMapping mapping = snapshot.StructureMappings[2];
            string original = mapping.SelectedStructureId;

            TestAssert.Throws<ArgumentException>(
                () => SimulatorSnapshotActions.ApplyStructureMapping(
                    snapshot,
                    mapping.StableId,
                    "UnknownSyntheticStructure"));

            TestAssert.Equal(original, mapping.SelectedStructureId);
        }

        public static void DvhExportWritesALocalPng()
        {
            // Match the simulator's headless WPF capture environment; otherwise
            // RenderTargetBitmap can return an all-transparent PNG on CI sessions.
            AppContext.SetSwitch("Switch.System.Windows.Media.ShouldRenderEvenWhenNoDisplayDevicesAreAvailable", true);
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("baseline-pass");
            var viewModel = new ReviewWorkspaceViewModel(snapshot);
            string directory = Path.Combine(
                Path.GetTempPath(),
                "clearplan-simulator-tests",
                Guid.NewGuid().ToString("N"));
            string outputPath = Path.Combine(
                directory,
                "synthetic-dvh.png");

            try
            {
                var model = viewModel.DetailPlotModel;
                var lines = model.Series.OfType<LineSeries>().ToArray();
                // Exercise the real GUI state: its native checkbox legend disables
                // OxyPlot's internal legend, and an unselected curve stays unselected.
                TestAssert.False(model.IsLegendVisible);
                lines.Last().IsVisible = false;
                var pointsBefore = lines.Select(line => line.Points.ToArray()).ToArray();
                var visibleBefore = lines.Select(line => line.IsVisible).ToArray();
                var legendBefore = lines.Select(line => line.RenderInLegend).ToArray();
                string titleBefore = model.Title;
                double legendFontBefore = model.LegendFontSize;
                double titleFontBefore = model.TitleFontSize;
                double legendWidthBefore = model.LegendMaxWidth;
                var placementBefore = model.LegendPlacement;
                var positionBefore = model.LegendPosition;
                Directory.CreateDirectory(directory);
                using (var lockedOutput = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    TestAssert.Throws<IOException>(() => new SyntheticDvhPngService().Export(model, outputPath));
                    TestAssert.False(model.IsLegendVisible, "A failed export must restore GUI legend visibility.");
                    TestAssert.Equal(titleBefore, model.Title, "A failed export must restore the GUI title.");
                    TestAssert.Equal(legendFontBefore, model.LegendFontSize);
                    TestAssert.Equal(titleFontBefore, model.TitleFontSize);
                    TestAssert.Equal(legendWidthBefore, model.LegendMaxWidth);
                    TestAssert.Equal(placementBefore, model.LegendPlacement);
                    TestAssert.Equal(positionBefore, model.LegendPosition);
                    TestAssert.True(legendBefore.SequenceEqual(lines.Select(line => line.RenderInLegend)));
                }
                string actual = new SyntheticDvhPngService().Export(
                    model,
                    outputPath);

                TestAssert.Equal(
                    Path.GetFullPath(outputPath),
                    actual);
                byte[] header = File.ReadAllBytes(actual)
                    .Take(8)
                    .ToArray();
                TestAssert.Equal(
                    "89504E470D0A1A0A",
                    BitConverter.ToString(header).Replace("-", ""));
                using (var bitmap = new System.Drawing.Bitmap(actual))
                {
                    TestAssert.Equal(1400, bitmap.Width);
                    TestAssert.Equal(800, bitmap.Height);
                    var legendArea = model.LegendArea;
                    TestAssert.True(legendArea.Width > 0 && legendArea.Height > 0 && legendArea.Left >= model.PlotArea.Right,
                        "The exported legend must have its own right-side area outside the DVH axes.");
                    TestAssert.True(model.PlotArea.Left >= 40 && model.PlotArea.Bottom <= bitmap.Height - 35,
                        "Complete axis labels must fit inside the exported bitmap.");
                    foreach (var line in lines)
                    {
                        int coloredLegendPixels = 0;
                        for (int y = (int)Math.Ceiling(legendArea.Top); y < Math.Min(bitmap.Height, legendArea.Bottom); y++)
                        for (int x = (int)Math.Ceiling(legendArea.Left); x < Math.Min(bitmap.Width, legendArea.Right); x++)
                        {
                            var pixel = bitmap.GetPixel(x, y);
                            if (pixel.R == line.Color.R && pixel.G == line.Color.G && pixel.B == line.Color.B)
                                coloredLegendPixels++;
                        }
                        if (line.IsVisible)
                            TestAssert.True(coloredLegendPixels >= 8, "Exported DVH must identify every selected curve in its right-side legend: " + line.Title +
                                "; key pixels=" + coloredLegendPixels + "; canvas=" + bitmap.GetPixel(0, 0) + "; legend=" + legendArea);
                        else
                            TestAssert.Equal(0, coloredLegendPixels, "Unselected curves must not appear in the export legend.");
                    }
                    int titlePixels = 0;
                    for (int y = 10; y < 40; y++)
                    for (int x = 230; x < 1100; x++)
                    {
                        var pixel = bitmap.GetPixel(x, y);
                        if (pixel.R < 120 && pixel.G < 120 && pixel.B < 120) titlePixels++;
                    }
                    TestAssert.True(titlePixels > 200, "The exported bitmap needs its own visible synthetic-use notice, independent of GUI chrome.");
                }
                TestAssert.False(model.IsLegendVisible, "Export must restore the native GUI legend state.");
                TestAssert.Equal(titleBefore, model.Title);
                TestAssert.Equal(legendFontBefore, model.LegendFontSize);
                TestAssert.Equal(titleFontBefore, model.TitleFontSize);
                TestAssert.Equal(legendWidthBefore, model.LegendMaxWidth);
                TestAssert.Equal(placementBefore, model.LegendPlacement);
                TestAssert.Equal(positionBefore, model.LegendPosition);
                for (int index = 0; index < lines.Length; index++)
                {
                    TestAssert.True(pointsBefore[index].SequenceEqual(lines[index].Points), "Export must not recalculate detached DVH samples.");
                    TestAssert.Equal(visibleBefore[index], lines[index].IsVisible);
                    TestAssert.Equal(legendBefore[index], lines[index].RenderInLegend);
                }
            }
            finally
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, true);
                }
            }
        }

        public static void DvhExportRejectsNetworkOutput()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("baseline-pass");
            var viewModel = new ReviewWorkspaceViewModel(snapshot);

            TestAssert.Throws<ArgumentException>(
                () => new SyntheticDvhPngService().Export(
                    viewModel.DetailPlotModel,
                    @"\\server\share\synthetic-dvh.png"));
        }
    }
}
