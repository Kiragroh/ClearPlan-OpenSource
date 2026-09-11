using System;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Core.Tests
{
    internal static class BevWorkspaceTests
    {
        public static void SyntheticMetadataIsExplicitAndReported()
        {
            var snapshot = ClearPlan.Core.Simulation.SyntheticScenarioFactory.Create("baseline-pass");
            TestAssert.Equal(100.0, snapshot.PlanAnalysis.PlanNormalizationPercent.Value);
            foreach (var beam in snapshot.PlanAnalysis.Beams)
            {
                TestAssert.True(!string.IsNullOrWhiteSpace(beam.EnergyDisplay));
                TestAssert.True(!string.IsNullOrWhiteSpace(beam.Technique));
            }
            var report = new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(snapshot);
            var method = typeof(ClearPlan.Reporting.MigraDoc.ReportPdf).GetMethod("CreateReviewReport", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var document = (MigraDoc.DocumentObjectModel.Document)method.Invoke(new ClearPlan.Reporting.MigraDoc.ReportPdf(), new object[] { report });
            string text = MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToString(document);
            TestAssert.True(text.Contains("Plan normalization: 100 %"));
        }
        public static void RejectsClinicalDrrInSyntheticSnapshot()
        {
            var snapshot = ClearPlan.Core.Simulation.SyntheticScenarioFactory.Create("baseline-pass");
            snapshot.PlanAnalysis.Beams[0].ControlPoints[0].BevImage = new BeamEyeViewImage { Synthetic = false };
            TestAssert.False(ClearPlan.Core.Review.ReviewSnapshotValidator.ValidateForSimulator(snapshot).IsValid,
                "Synthetic workspaces must refuse clinical BEV images even though pixels are not serialized.");
        }

        public static void MachineProfilePathRoundTripsIni()
        {
            var property = typeof(ClearPlan.Core.Settings.ClearPlanPathOptions).GetProperty("MlcGeometryProfilesJsonPath");
            TestAssert.NotNull(property, "Machine geometry path must be configurable in settings.ini.");
            var settings = new ClearPlan.Core.Settings.ClearPlanSettingsModel();
            ClearPlan.Core.Settings.PathSettingsIni.Apply(settings, "[Paths]\nMlcGeometryProfilesJsonPath=MachineGeometry\\custom.json\n");
            TestAssert.Equal("MachineGeometry\\custom.json", (string)property.GetValue(settings.Paths));
            var copy = new ClearPlan.Core.Settings.ClearPlanSettingsModel();
            ClearPlan.Core.Settings.PathSettingsIni.Apply(copy, ClearPlan.Core.Settings.PathSettingsIni.Serialize(settings));
            TestAssert.Equal("MachineGeometry\\custom.json", (string)property.GetValue(copy.Paths));
        }
        public static void ExposesControlPointNavigationAndExplicitDrrAction()
        {
            var type = typeof(PlanAnalysisViewModel);
            foreach (string name in new[] { "SelectedControlPoint", "ControlPointPosition", "MaximumControlPointPosition", "BevImageSource", "GenerateDrrCommand", "PreviousControlPointCommand", "NextControlPointCommand" })
                TestAssert.NotNull(type.GetProperty(name), "Missing BEV workspace property: " + name);
        }

        public static void NormalizationIsDescriptiveAndMissingIsNotZero()
        {
            var property = typeof(ReviewPlanAnalysis).GetProperty("PlanNormalizationPercent");
            TestAssert.NotNull(property, "Expose TPS normalization for paired-plan review.");
            var analysis = new ReviewPlanAnalysis();
            TestAssert.True(property.GetValue(analysis) == null);
            property.SetValue(analysis, 98.5);
            var row = PlanAnalysisViewModel.CreateMetrics(analysis).SingleOrDefault(r => r.Key == "Normalization");
            TestAssert.NotNull(row, "Normalization must be included in descriptive comparison metrics.");
            TestAssert.Equal("%", row.Unit);
            TestAssert.True(Math.Abs(row.Value.Value - 98.5) < 1e-9);
        }
    }
}
