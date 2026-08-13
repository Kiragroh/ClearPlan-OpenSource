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
                string actual = new SyntheticDvhPngService().Export(
                    viewModel.DetailPlotModel,
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
