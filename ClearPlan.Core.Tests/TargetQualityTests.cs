using System;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class TargetQualityTests
    {
        public static void AvailabilityReflectsValuesAndPreservesWarnings()
        {
            var property = typeof(ReviewTargetQuality).GetProperty("AvailabilityScope");
            TestAssert.NotNull(property, "Availability must derive from calculated values, not absence of warning text.");
            Func<ReviewTargetQuality, string> display = row => (string)property.GetValue(row);
            var complete = TargetQualityCalculator.Calculate("PTV_Test", 60, 100, 90, 120, 360, 63, 57);
            TestAssert.True(string.IsNullOrEmpty(complete.Note), "The calculator must keep success independent of an explanatory note.");
            TestAssert.Equal("Available", display(complete));
            complete.Note = "Multiple targets: verify reference-dose applicability; not an approval criterion.";
            TestAssert.Equal("Available. " + complete.Note, display(complete), "A complete numeric row must retain explicit warnings.");
            var partial = TargetQualityCalculator.Calculate("PTV_Test", 60, 100, 90, null, null, 63, 57);
            TestAssert.Equal("Partly available. " + partial.Note, display(partial));
            TestAssert.True(partial.PaddickCi == null && partial.PlanCheckCi == null && partial.GradientIndex == null && partial.HomogeneityIndex.HasValue);
            var missing = new ReviewTargetQuality();
            TestAssert.Equal("Unavailable", display(missing));
            missing.Note = "No current dose. Values not evaluated.";
            TestAssert.Equal("Unavailable. " + missing.Note, display(missing));
            missing.PaddickCi = double.NaN; missing.PlanCheckCi = double.PositiveInfinity; missing.GradientIndex = -1;
            TestAssert.Equal("Unavailable. " + missing.Note, display(missing), "Nonfinite/negative values must not be displayed as available.");
            TestAssert.True(Newtonsoft.Json.Linq.JObject.FromObject(complete)["AvailabilityScope"] == null,
                "Display-only availability must not change the persisted snapshot schema.");
            string xaml = System.IO.File.ReadAllText("ClearPlan.Presentation/Views/PlanParametersView.xaml");
            TestAssert.True(xaml.Contains("Header=\"Verfügbarkeit\" Binding=\"{Binding AvailabilityScope}\""), "Native and simulator GUI must use the shared display property.");
            foreach (var row in new[] { complete, partial, missing })
            {
                row.StructureId = "PTV_Test"; row.BodyStructureId = "External";
                var snapshot = new ReviewSnapshot { PlanAnalysis = new ReviewPlanAnalysis() };
                snapshot.PlanAnalysis.TargetQuality.Add(row);
                var report = new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(snapshot);
                string html = new ClearPlan.Reporting.MigraDoc.HtmlReviewReportRenderer().Render(report);
                TestAssert.True(html.Contains(display(row)), "HTML must retain the same full/partial/missing availability and warnings as GUI.");
            }
        }

        public static void TargetRecognitionIsConsistentAndPamFailsClosed()
        {
            var structures = new[] {
                new TargetReviewStructure { StructureId = "PTV60", DicomType = "CONTROL" },
                new TargetReviewStructure { StructureId = "PTV_BODY", DicomType = "EXTERNAL" },
                new TargetReviewStructure { StructureId = "PTV_SUPPORT", DicomType = "SUPPORT" },
                new TargetReviewStructure { StructureId = "Lung", DicomType = "ORGAN" },
                new TargetReviewStructure { StructureId = "APTVALUE", DicomType = "CONTROL" }
            };
            var ptvs = structures.Where(s => DvhSelectionPolicy.ClassifyTarget(s.StructureId, s.DicomType) == "PTV")
                .Select(s => s.StructureId).ToArray();
            TestAssert.Equal("PTV60", string.Join("|", ptvs), "Only bounded PTV tokens may supplement native type; no arbitrary OAR fallback.");
            TestAssert.Equal("PTV60", string.Join("|", DvhSelectionPolicy.SelectTargetStructureIds(structures)));
            string provenance, reason;
            string[] eligible = structures.Select(s => s.StructureId).ToArray();
            TestAssert.Equal("PTV60", PamTargetSelection.Resolve(null, null, eligible, ptvs, out provenance, out reason));
            TestAssert.True(reason == null);
            TestAssert.Equal("Lung", PamTargetSelection.Resolve("lung", null, eligible, ptvs, out provenance, out reason),
                "A valid explicit choice retains precedence over target naming.");
            TestAssert.True(provenance.Contains("Explicit"));
            TestAssert.Equal("Lung", PamTargetSelection.Resolve(null, "Lung", eligible, ptvs, out provenance, out reason),
                "A valid native plan target retains precedence over target naming.");
            TestAssert.True(provenance.Contains("TargetVolumeID"));
            TestAssert.Equal(null, PamTargetSelection.Resolve("missing", null, eligible, ptvs, out provenance, out reason));
            TestAssert.True(reason.Contains("explicit"));
            TestAssert.Equal(null, PamTargetSelection.Resolve(null, null, new[] { "PTV60", "PTV70" },
                new[] { "PTV60", "PTV70" }, out provenance, out reason));
            TestAssert.True(reason.Contains("Multiple"));
            TestAssert.Equal(null, PamTargetSelection.Resolve(null, null, new[] { "Lung", "APTVALUE" },
                new string[0], out provenance, out reason));
            TestAssert.Equal(null, PamTargetSelection.Resolve(null, null, new[] { "PTV60", "ptv60" },
                new[] { "PTV60" }, out provenance, out reason), "Case-ambiguous identities must remain unavailable.");

            foreach (var file in new[] { "EsapiTargetQualityBuilder.cs", "EsapiPlanAnalysisBuilder.cs" })
            {
                string source = System.IO.File.ReadAllText("ClearPlan.Script/Review/" + file);
                TestAssert.True(source.Contains("DvhSelectionPolicy.ClassifyTarget(s.Id, s.DicomType) == \"PTV\""),
                    file + " must recognize exactly the same PTV kinds as the visible DVH.");
                TestAssert.True(source.Contains("!s.IsEmpty && s.HasSegment"), "An alias never replaces nonempty segmented geometry.");
                foreach (var mutation in new[] { "BeginModifications(", "SaveModifications(", "AddStructure(", "ConvertDoseLevelToStructure(", "SegmentVolume =" })
                    TestAssert.False(source.Contains(mutation), "Shared target recognition must remain read-only.");
            }
        }

        public static void CoverageMustBeFiniteAndWithinNativeRange()
        {
            var method = typeof(TargetQualityCalculator).GetMethod("HasSufficientCoverage");
            TestAssert.NotNull(method, "Dose and sampling coverage need a bounded, testable guard.");
            Func<double?, double?, bool> valid = (dose, sampling) => (bool)method.Invoke(null, new object[] { dose, sampling });
            TestAssert.True(valid(1, 0.99));
            TestAssert.True(valid(1, 1.0000376055430387), "Native sampling coverage can slightly exceed one; it is not normalized dose coverage.");
            TestAssert.False(valid(double.PositiveInfinity, 1));
            TestAssert.False(valid(1, double.PositiveInfinity));
            TestAssert.False(valid(double.NaN, 1));
            TestAssert.False(valid(1, 0.89));
            TestAssert.False(valid(null, 1));
        }
        public static void SyntheticGuiAndReportShareTheSameMetrics()
        {
            var snapshot = ClearPlan.Core.Simulation.SyntheticScenarioFactory.Create("baseline-pass");
            TestAssert.True(snapshot.PlanAnalysis.TargetQuality.Count > 0, "The synthetic workspace must demonstrate target quality.");
            var row = snapshot.PlanAnalysis.TargetQuality[0];
            TestAssert.Equal("PTV_60", row.StructureId);
            TestAssert.True(row.PaddickCi.HasValue && row.GradientIndex.HasValue && row.HomogeneityIndex.HasValue);
            var vm = new ClearPlan.Presentation.ViewModels.PlanAnalysisViewModel(snapshot.PlanAnalysis, true);
            TestAssert.NotNull(vm.GetType().GetProperty("TargetQuality"), "Target quality must be visible in the plan parameter panel.");
            var report = new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(snapshot);
            Near(row.PaddickCi.Value, report.PlanAnalysis.TargetQuality[0].PaddickCi);
            var document = new MigraDoc.DocumentObjectModel.Document(); var section = document.AddSection();
            var method = typeof(ClearPlan.Reporting.MigraDoc.ReportPdf).GetMethod("AddTargetQuality",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            TestAssert.NotNull(method, "The PDF needs its target quality section.");
            method.Invoke(new ClearPlan.Reporting.MigraDoc.ReportPdf(), new object[] { section, report.PlanAnalysis });
            var ddl = MigraDoc.DocumentObjectModel.IO.DdlWriter.WriteToString(document);
            TestAssert.True(ddl.Contains("Paddick CI") && ddl.Contains("HI (PlanCheck)") && ddl.Contains("PTV_60"));
            var native = System.IO.File.ReadAllText("ClearPlan.Script/Review/EsapiTargetQualityBuilder.cs");
            foreach (var mutation in new[] { "BeginModifications(", "SaveModifications(", "AddStructure(", "ConvertDoseLevelToStructure(", "SegmentVolume =" })
                TestAssert.False(native.Contains(mutation), "Target quality must not request or mutate TPS write state.");
            TestAssert.True(native.Contains("VolumePresentation.AbsoluteCm3") && native.Contains("DoseValue.DoseUnit.cGy"));
            TestAssert.True(native.Contains("plan.IsDoseValid"), "Target indices require valid current dose, not cached invalid DVHs.");
        }
        public static void ScalarDefinitionsAndMissingValues()
        {
            var type = typeof(ReviewPlanAnalysis).Assembly.GetType("ClearPlan.Core.PlanAnalysis.TargetQualityCalculator");
            TestAssert.NotNull(type, "A detached read-only target-quality calculator is required.");
            var method = type.GetMethod("Calculate");
            Func<double?, double?, double?, double?, double?, double?, double?, object> calc =
                (rx, tv, covered, v100, v50, d2, d98) => method.Invoke(null,
                    new object[] { "PTV_Test", rx, tv, covered, v100, v50, d2, d98 });
            var row = calc(60, 100, 90, 120, 360, 63, 57);
            Near(0.675, Value(row, "PaddickCi"));
            Near(1 / 0.675, Value(row, "PlanCheckCi"));
            Near(3, Value(row, "GradientIndex"));
            Near(0.1, Value(row, "HomogeneityIndex"));
            row = calc(60, 100, 0, 120, 360, 63, 57);
            Near(0, Value(row, "PaddickCi"));
            TestAssert.Equal(null, Value(row, "PlanCheckCi"));
            row = calc(60, 100, 90, null, null, 63, 57);
            TestAssert.Equal(null, Value(row, "PaddickCi"));
            Near(0.1, Value(row, "HomogeneityIndex"));
            row = calc(60, 100, 0, 0, 360, 63, 57);
            TestAssert.Equal(null, Value(row, "GradientIndex"));
            TestAssert.True(((string)row.GetType().GetProperty("Note").GetValue(row, null)).Contains("V100 = 0"),
                "Undefined indices at zero prescription-isodose volume need a specific explanation.");
            row = calc(60, 100, 110, 120, 60, 57, 63);
            TestAssert.Equal(null, Value(row, "PaddickCi"));
            TestAssert.Equal(null, Value(row, "GradientIndex"));
            TestAssert.Equal(null, Value(row, "HomogeneityIndex"));
            row = calc(0, 100, 90, 120, 360, 63, 57);
            TestAssert.Equal(null, Value(row, "PaddickCi"));
            TestAssert.Equal(null, Value(row, "GradientIndex"));
            TestAssert.Equal(null, Value(row, "HomogeneityIndex"));
            row = calc(60, 100, 90, double.NaN, double.PositiveInfinity, 63, 57);
            TestAssert.Equal(null, Value(row, "PaddickCi"));
            TestAssert.Equal(null, Value(row, "GradientIndex"));
            Near(0.1, Value(row, "HomogeneityIndex"));
        }
        private static double? Value(object row, string name) { return (double?)row.GetType().GetProperty(name).GetValue(row, null); }
        private static void Near(double expected, double? actual)
        { TestAssert.True(actual.HasValue && Math.Abs(expected - actual.Value) < 1e-9, "Target quality metric differs from analytical reference."); }
    }
}
