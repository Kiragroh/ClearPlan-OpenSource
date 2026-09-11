using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using ClearPlan.Core.Review;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Simulator;

namespace ClearPlan.Core.Tests
{
    internal static class SyntheticPublicationIntegrationTests
    {
        public static void ScenarioCatalogAndCliAgree()
        {
            var repository = new SimulatorScenarioRepository(Path.Combine(RepositoryRoot(), "ClearPlan.Simulator"));
            TestAssert.Equal(9, repository.ScenarioIds.Count, "GUI must expose the unchanged seven examples and two explicit publication fixtures.");
            foreach (string id in new[] { "publication-single-layer", "publication-dual-layer" })
            {
                var args = SimulatorArguments.Parse(new[] { "--scenario", id, "--tab", "dvh", "--capture", "fixture.png", "--report", "fixture.pdf" });
                TestAssert.Equal(id, args.ScenarioId);
                TestAssert.True(repository.ScenarioIds.Contains(id));
                var loaded = repository.Load(id);
                TestAssert.True(loaded.StructureMappings.All(mapping => mapping.Status == ReviewStatusCodes.Pass),
                    "Exact synthetic structure matches must carry a resolved mapping status, not generic data availability.");
                var workspace = new ReviewWorkspaceViewModel(loaded);
                TestAssert.True(workspace.StructureMappings.All(mapping => mapping.StatusText == "Zugeordnet"),
                    "The public fixture must not display a green unresolved mapping label for its exact matches.");
                var quality = loaded.PlanAnalysis.TargetQuality.Single();
                TestAssert.True(quality.PaddickCi.HasValue && quality.PlanCheckCi.HasValue && quality.GradientIndex.HasValue && quality.HomogeneityIndex.HasValue);
                var availabilityProperty = typeof(ReviewTargetQuality).GetProperty("AvailabilityScope");
                TestAssert.NotNull(availabilityProperty, "GUI and HTML need one value-derived availability display.");
                string availability = (string)availabilityProperty.GetValue(quality);
                TestAssert.True(!string.IsNullOrWhiteSpace(availability) && availability.StartsWith("Available", StringComparison.Ordinal),
                    "Calculated publication target-quality values need an explicit availability/scope label, not an empty GUI cell or HTML Unavailable fallback.");
                TestAssert.True(quality.Note.Contains("Synthetic") && quality.Note.Contains("EXTERNAL"), "Availability must retain synthetic provenance and whole-plan scope.");
                TestAssert.Equal(quality.Note, workspace.Analysis.TargetQuality.Single().Note);
                var report = new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(loaded);
                string html = new ClearPlan.Reporting.MigraDoc.HtmlReviewReportRenderer().Render(report);
                var targetRows = Regex.Matches(html, @"<tr>[\s\S]*?</tr>").Cast<Match>()
                    .Where(match => match.Value.Contains("PTV_60") && match.Value.Contains("External")).ToList();
                TestAssert.Equal(1, targetRows.Count);
                TestAssert.True(targetRows[0].Value.Contains(availability));
                TestAssert.False(targetRows[0].Value.Contains(">Unavailable<"), "Complete calculated target metrics must not be labelled unavailable in HTML.");
                TestAssert.Equal(id, loaded.ScenarioId);
                TestAssert.Equal(3, loaded.PlanImages.Count);
                TestAssert.True(loaded.PlanAnalysis.Beams.All(b => b.MlcLayerCount == (id.EndsWith("dual-layer", StringComparison.Ordinal) ? 2 : 1)));
                TestAssert.True(ReviewSnapshotValidator.ValidateForSimulator(loaded).IsValid);
            }
            TestAssert.Equal(7, SyntheticScenarioFactory.ScenarioIds.Count);
            var baseline = repository.Load("baseline-pass");
            TestAssert.Equal(5, baseline.DvhSeries.Count);
            TestAssert.Throws<ArgumentException>(() => repository.Load("publication-unknown"));
        }

        public static void SuppliedImagesSurviveGuiAndReport()
        {
            var snapshot = SyntheticPublicationScenarioFactory.Create(false);
            var suppliedImages = snapshot.PlanImages;
            var workspace = new ReviewWorkspaceViewModel(snapshot);
            TestAssert.True(workspace.PlanImages.HasImages);
            TestAssert.True(workspace.PlanImages.LegendItems.Count > 6, "GUI must preserve analytical structures and dose contours.");
            string directory = Path.Combine(Path.GetTempPath(), "ClearPlan-Publication-" + Guid.NewGuid().ToString("N"));
            try
            {
                bool inspected = false;
                TestAssert.Throws<InspectionCompleteException>(() => new SyntheticReportService().Export(snapshot,
                    Path.Combine(directory, "synthetic.pdf"), document =>
                    {
                        TestAssert.True(ReferenceEquals(suppliedImages, snapshot.PlanImages), "Report export must not overwrite supplied phantom images.");
                        TestAssert.Equal(3, document.PlanImages.Count);
                        TestAssert.True(document.PlanImages.All(i => i.Overlays.Any(o => o.Kind == "isodose")));
                        TestAssert.True(document.PlanAnalysis.Beams.All(b => b.ControlPoints[0].BevImage.ProjectionDescription.Contains("same analytical phantom")));
                        inspected = true;
                        throw new InspectionCompleteException(); // Stop before disk/PDF rendering; the mapper has been inspected.
                    }));
                TestAssert.True(inspected);
                var defaultSnapshot = SyntheticScenarioFactory.Create("baseline-pass");
                TestAssert.Throws<InspectionCompleteException>(() => new SyntheticReportService().Export(defaultSnapshot,
                    Path.Combine(directory, "legacy.pdf"), document =>
                    {
                        TestAssert.Equal(3, document.PlanImages.Count, "Existing image-less fixtures retain their synthetic fallback.");
                        throw new InspectionCompleteException();
                    }));
            }
            finally
            {
                if (Directory.Exists(directory)) Directory.Delete(directory, false); // Owned empty directory only.
            }
        }

        public static void RegeneratedDrrRemainsCoherent()
        {
            TestAssert.NotNull(typeof(PlanAnalysisViewModel).GetConstructor(new[] { typeof(ReviewPlanAnalysis), typeof(bool) }),
                "Keep the existing two-argument constructor binary-compatible for the native adapter.");
            var snapshot = SyntheticPublicationScenarioFactory.Create(true);
            var workspace = new ReviewWorkspaceViewModel(snapshot);
            var analysis = workspace.Analysis;
            var first = analysis.SelectedControlPoint;
            var expected = first.BevImage.GrayscalePixels.ToArray();
            analysis.GenerateDrrCommand.Execute(null);
            var regeneration = (Task)typeof(PlanAnalysisViewModel).GetField("syntheticDrrTask", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(analysis);
            TestAssert.NotNull(regeneration);
            regeneration.GetAwaiter().GetResult();
            TestAssert.True(first.BevImage.ProjectionDescription.Contains("same analytical phantom"), "DRR button must not switch to an unrelated phantom.");
            TestAssert.True(expected.SequenceEqual(first.BevImage.GrayscalePixels));
            analysis.ControlPointPosition = 15;
            analysis.ActivateBevAsync().GetAwaiter().GetResult();
            var middle = analysis.SelectedControlPoint;
            TestAssert.NotNull(middle.BevImage);
            TestAssert.True(middle.BevImage.ProjectionDescription.Contains("same analytical phantom"));
            TestAssert.Equal(15, middle.BevImage.ControlPointIndex);
            TestAssert.Equal(middle.GantryAngleDegrees, middle.BevImage.GantryAngleDegrees);
            TestAssert.False(expected.SequenceEqual(middle.BevImage.GrayscalePixels), "Different projection angles must not reuse the first-CP raster.");
            TestAssert.True(ReviewSnapshotValidator.ValidateForSimulator(snapshot).IsValid);
        }

        public static void CtTabIsAvailableToAutomation()
        {
            TestAssert.Equal("images", SimulatorArguments.Parse(new[] { "--scenario", "publication-single-layer", "--tab", "images", "--capture", "planes.png" }).TabId);
            string visual = File.ReadAllText(Path.Combine(RepositoryRoot(), "ClearPlan.Simulator", "VisualCaptureService.cs"));
            TestAssert.True(visual.Contains("{ \"images\", \"PlanImagesTab\" }"), "Capture tab must target the existing shared CT view.");
            string main = File.ReadAllText(Path.Combine(RepositoryRoot(), "ClearPlan.Simulator", "MainWindow.xaml.cs"));
            TestAssert.True(main.Substring(main.IndexOf("CaptureTabIds", StringComparison.Ordinal), 400).Contains("\"images\""));
            var exports = Regex.Matches(main, @"reportService\.Export\([\s\S]*?\);").Cast<Match>().ToList();
            TestAssert.True(exports.Count >= 4 && exports.All(m => m.Value.Contains("ApplyReportOptions")),
                "GUI and CLI exports must share the same structure-visibility, unmatched and BEV options.");
        }

        public static void ReportsDescribeSharedPhantomAccurately()
        {
            var snapshot = SyntheticPublicationScenarioFactory.Create(false);
            var report = new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(snapshot);
            var method = typeof(ClearPlan.Reporting.MigraDoc.ReportPdf).GetMethod("CreateReviewReport", BindingFlags.Instance | BindingFlags.NonPublic);
            var document = (MigraDoc.DocumentObjectModel.Document)method.Invoke(new ClearPlan.Reporting.MigraDoc.ReportPdf(), new object[] { report });
            string ddl = MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToString(document);
            string html = new ClearPlan.Reporting.MigraDoc.HtmlReviewReportRenderer().Render(report);
            foreach (string rendered in new[] { ddl, html })
            {
                TestAssert.False(rendered.Contains("not the source volume for the BEV DRRs"), "Publication CT and BEV share one analytical phantom, unlike the old fixtures.");
                TestAssert.True(rendered.Contains("Shared analytical phantom"));
                TestAssert.True(rendered.Contains("not dose calculated from apertures"));
            }
        }

        private static string RepositoryRoot()
        {
            var root = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (root != null && !File.Exists(Path.Combine(root.FullName, "ClearPlan.sln"))) root = root.Parent;
            TestAssert.NotNull(root);
            return root.FullName;
        }

        private sealed class InspectionCompleteException : Exception { }
    }
}
