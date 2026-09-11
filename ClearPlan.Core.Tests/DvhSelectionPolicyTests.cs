using System.Collections.Generic;
using System;
using System.Linq;
using System.IO;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class DvhSelectionPolicyTests
    {
        public static void RequiredTargetsHaveASeparateSelectionContract()
        {
            var type = typeof(DvhSelectionPolicy);
            TestAssert.NotNull(type.GetMethod("SelectTargetStructureIds"),
                "Required PTVs and the single largest contained inner target must have a shared selection policy.");
            TestAssert.NotNull(type.GetMethod("ClassifyTarget"),
                "Native DICOM type and target aliases must be recognized centrally.");
            TestAssert.NotNull(type.Assembly.GetType("ClearPlan.Core.Review.UserTargetCoverageRule"),
                "User target D98 > Rx rule must exist independently of stock constraints and native Clinical Goals.");
        }

        public static void NativeTargetKindsAndAliasesAreRecognized()
        {
            TestAssert.Equal("PTV", DvhSelectionPolicy.ClassifyTarget("Boost", "PTV"));
            TestAssert.Equal("GTV", DvhSelectionPolicy.ClassifyTarget("GVT_Tm", "GTV"));
            TestAssert.Equal("", DvhSelectionPolicy.ClassifyTarget("GVT_Tm", "CONTROL"));
            TestAssert.Equal("PTV", DvhSelectionPolicy.ClassifyTarget("PTV70", ""));
            TestAssert.Equal("CTV", DvhSelectionPolicy.ClassifyTarget("eval_CTV", ""));
            TestAssert.Equal("", DvhSelectionPolicy.ClassifyTarget("PTV_SUPPORT", "SUPPORT"));
            TestAssert.Equal("", DvhSelectionPolicy.ClassifyTarget("ACTVITY", ""));
        }

        public static void KeepsAllPtvsAndOnlyLargestContainedInnerTargetOverall()
        {
            var structures = new[] {
                Target("PTV50", "PTV", 300, null), Target("Boost", "PTV", 100, null),
                Target("CTV", "CTV", 180, true), Target("ITV", "ITV", 120, true),
                Target("GVT_Tm", "GTV", 80, true), Target("CTV_outside", "CTV", 200, false)
            };
            TestAssert.Equal("PTV50|Boost|CTV", string.Join("|", DvhSelectionPolicy.SelectTargetStructureIds(structures)));
            structures[2].FullyContainedInPtv = false;
            TestAssert.Equal("PTV50|Boost|ITV", string.Join("|", DvhSelectionPolicy.SelectTargetStructureIds(structures)));
            structures[2].FullyContainedInPtv = null;
            TestAssert.Equal("PTV50|Boost", string.Join("|", DvhSelectionPolicy.SelectTargetStructureIds(structures)),
                "Larger unknown geometry must not make the smaller verified inner target look like the largest.");
            TestAssert.Equal(0, DvhSelectionPolicy.SelectTargetStructureIds(structures.Skip(2)).Count);
        }

        public static void TargetRulesKeepStrictBoundaryAndUnknownPrescriptionVisible()
        {
            var prescriptions = new[] { new TargetPrescriptionDose { TargetId = "PTV", TargetType = "Volume", DosePerFractionGy = 2, FractionCount = 30 } };
            var pass = UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, 30, 60.0001, true);
            TestAssert.Equal("pass", pass.Status);
            TestAssert.Equal(60.0, pass.Goal.Value);
            TestAssert.Equal(">", pass.Comparator);
            TestAssert.Equal("fail", UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, 30, 60, true).Status);
            TestAssert.Equal("fail", UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, 30, 59.9999, true).Status);
            foreach (var row in new[] {
                UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, 10, 60.1, true),
                UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, null, 60.1, true),
                UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions.Concat(prescriptions), 30, 60.1, true),
                UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV_boost", prescriptions, 30, 60.1, true),
                UserTargetCoverageRule.Evaluate(DefaultRule(), "CTV", prescriptions, 30, 60.1, true),
                UserTargetCoverageRule.Evaluate(DefaultRule(), "ptv", prescriptions, 30, 60.1, true),
                UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, 30, null, true),
                UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, 30, double.NaN, true),
                UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, 30, 60.1, false),
                UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, 30, 60.1, null),
                UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", null, 30, 60.1, true)
            })
            {
                TestAssert.Equal("not-evaluated", row.Status);
                TestAssert.True(row.Explanation.Contains("Not evaluated:"));
                TestAssert.Equal(UserTargetCoverageRule.SourceLabel, row.SourceLabel);
            }
            prescriptions[0].TargetType = "Isocenter";
            TestAssert.Equal("not-evaluated", UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, 30, 60.1, true).Status);
            prescriptions[0].TargetType = "Volume";
            prescriptions[0].DosePerFractionGy = null;
            TestAssert.Equal("not-evaluated", UserTargetCoverageRule.Evaluate(DefaultRule(), "PTV", prescriptions, 30, 60.1, true).Status);
        }

        public static void NativeTargetReviewIsReadOnlyAndSharedWithReport()
        {
            var root = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (root != null && !File.Exists(Path.Combine(root.FullName, "ClearPlan.sln"))) root = root.Parent;
            TestAssert.NotNull(root, "Repository root required for native integration contract.");
            string builderPath = Path.Combine(root.FullName, "ClearPlan.Script", "Review", "EsapiTargetReviewBuilder.cs");
            TestAssert.True(File.Exists(builderPath), "Native target-selection/Rx adapter must exist.");
            string native = File.ReadAllText(builderPath);
            foreach (string required in new[] { "GetContoursOnImagePlane", "GetDoseAtVolume", "DoseValuePresentation.Absolute", "VolumePresentation.Relative", "prescription.Targets", "target.TargetId", "target.DosePerFraction", "target.NumberOfFractions" })
                TestAssert.True(native.Contains(required), "Native target review must use " + required);
            foreach (string forbidden in new[] { "BeginModifications", "AddStructure(", "SaveModifications", "Task.Run(", "plan.TotalDose", "SegmentVolume =" })
                TestAssert.False(native.Contains(forbidden), "Native target review must not use " + forbidden);
            string snapshot = File.ReadAllText(Path.Combine(root.FullName, "ClearPlan.Script", "Review", "EsapiReviewSnapshotBuilder.cs"));
            TestAssert.True(snapshot.Contains("BuildGoals(") && snapshot.Contains("RequiredForTargetReview"), "Shared GUI/report snapshot must add user goals and required curves.");
            TestAssert.True(snapshot.Contains("StructuresSelectedForDvh"), "Native Eclipse DVH selections must be read through the documented ESAPI API.");
            TestAssert.True(snapshot.Contains("snapshot.PqmRows,") && snapshot.Contains("requestedStructureIds"), "Native goals and every resolved constraint must reach DVH selection, including unevaluated rows.");
            TestAssert.False(snapshot.Contains("targetKind.Length == 0 && item.IsSelected"), "An explicitly selected inner target must not disappear from the detached snapshot.");
            TestAssert.True(snapshot.Contains("not a general StructureSet visibility flag"), "Eclipse DVH activation must not be misrepresented as generic StructureSet display visibility.");
            var requiredIds = DvhSelectionPolicy.SelectTargetStructureIds(new[] {
                Target("PTV50", "PTV", 300, null), Target("PTV60", "PTV", 100, null),
                Target("CTV", "CTV", 180, true), Target("GTV", "GTV", 80, true)
            });
            var detached = new ReviewSnapshot();
            foreach (string id in new[] { "PTV50", "PTV60", "CTV", "GTV", "SpinalCord", "Helper" })
                detached.DvhSeries.Add(new ReviewDvhSeries {
                    StableId = "dvh-" + id, StructureId = id,
                    RequiredForTargetReview = requiredIds.Contains(id),
                    Selected = DvhSelectionPolicy.ShouldSelect(id, new[] { "SpinalCord" })
                });
            var report = new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(detached);
            TestAssert.Equal("PTV50|PTV60|CTV|SpinalCord", string.Join("|",
                report.DvhSeries.Where(row => row.Selected).Select(row => row.StructureId)),
                "An OAR-only request must retain every required PTV and the largest contained inner target in the shared report.");
            TestAssert.Equal(6, report.DvhSeries.Count, "Unselected curves remain cached for later manual selection.");
            report.HiddenStructureIds.Add("CTV");
            TestAssert.False(report.VisibleDvhSeries().Any(row => row.StructureId == "CTV"),
                "Explicit report visibility can hide a default target without changing its source curve.");
            TestAssert.True(detached.DvhSeries.Single(row => row.StructureId == "CTV").RequiredForTargetReview);
            TestAssert.False(detached.DvhSeries.Single(row => row.StructureId == "CTV").Selected,
                "Report defaults must not mutate the detached source selections.");
        }

        public static void DefaultRulesHaveAnEditableExternalConfiguration()
        {
            TestAssert.NotNull(typeof(ClearPlan.Core.Settings.ClearPlanPathOptions).GetProperty("DefaultReviewRulesJsonPath"),
                "Default rules need an editable configured path.");
            TestAssert.NotNull(typeof(UserTargetCoverageRule).Assembly.GetType("ClearPlan.Core.Review.DefaultReviewRuleConfiguration"),
                "Default rules must be deletable/disableable data, not an unconditional built-in evaluation.");
        }

        public static void DefaultRuleEditsDisablesAndDeletionAreHonoredWithoutFallback()
        {
            var rule = DefaultRule();
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(new { schemaVersion = 1, rules = new[] { rule } });
            var loaded = DefaultReviewRuleConfiguration.Parse(json);
            TestAssert.Equal("available", loaded.Status);
            TestAssert.Equal(1, loaded.Rules.Count);
            TestAssert.Equal(">", loaded.Rules[0].Comparator);
            rule.Comparator = ">=";
            rule.Metric = "D95%";
            rule.PrescriptionPercent = 95;
            loaded = DefaultReviewRuleConfiguration.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(new { schemaVersion = 1, rules = new[] { rule } }));
            double volume;
            TestAssert.True(loaded.Rules[0].TryGetVolumePercent(out volume));
            TestAssert.Equal(95.0, volume);
            var rx = new[] { new TargetPrescriptionDose { TargetId = "PTV", TargetType = "Volume", DosePerFractionGy = 2, FractionCount = 30 } };
            var result = UserTargetCoverageRule.Evaluate(loaded.Rules[0], "PTV", rx, 30, 57, true);
            TestAssert.Equal("pass", result.Status);
            TestAssert.Equal(57.0, result.Goal.Value);
            TestAssert.Equal("D95%", result.Objective);
            TestAssert.Equal("Default", result.SourceLabel);
            rule.Enabled = false;
            loaded = DefaultReviewRuleConfiguration.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(new { schemaVersion = 1, rules = new[] { rule } }));
            TestAssert.Equal(0, loaded.Rules.Count);
            TestAssert.Equal("not-configured", loaded.Status);
            TestAssert.Equal(0, DefaultReviewRuleConfiguration.Parse("{\"schemaVersion\":1,\"rules\":[]}").Rules.Count);
            var missingPath = Path.Combine(Path.GetTempPath(), "clearplan-missing-rules-" + Guid.NewGuid().ToString("N") + ".json");
            TestAssert.Equal("not-configured", DefaultReviewRuleConfiguration.Load(missingPath).Status);
            TestAssert.False(File.Exists(missingPath), "Missing/deleted rules must not be recreated.");
            TestAssert.Equal("not-configured", DefaultReviewRuleConfiguration.Load("").Status);
            foreach (string invalid in new[] { "{}", "not-json", "{\"schemaVersion\":1,\"rules\":null}", json.Replace("D98%", "D101%"), json.Replace("target-prescription", "plan-total"), json.Replace("target-d98-rx", "bad/id") })
            {
                var rejected = DefaultReviewRuleConfiguration.Parse(invalid);
                TestAssert.Equal("unavailable", rejected.Status);
                TestAssert.Equal(0, rejected.Rules.Count);
            }
            var settings = new ClearPlan.Core.Settings.ClearPlanSettingsModel();
            settings.Paths.DefaultReviewRulesJsonPath = "rules\\custom-defaults.json";
            var roundtrip = new ClearPlan.Core.Settings.ClearPlanSettingsModel();
            ClearPlan.Core.Settings.PathSettingsIni.Apply(roundtrip, ClearPlan.Core.Settings.PathSettingsIni.Serialize(settings));
            TestAssert.Equal(settings.Paths.DefaultReviewRulesJsonPath, roundtrip.Paths.DefaultReviewRulesJsonPath);
        }

        public static void RequiredTargetsSurviveUiAndReportSelectionAndNativeD98()
        {
            var series = new ReviewDvhSeries { StableId = "dvh-PTV", StructureId = "PTV", TargetKind = "PTV", Role = "target", RequiredForTargetReview = true, Selected = false, D98DoseGy = 60.1234 };
            var vm = new ClearPlan.Presentation.ViewModels.ReviewDvhSeriesViewModel(series);
            TestAssert.True(vm.IsSelected, "Required targets are selected by default.");
            vm.IsSelected = false;
            TestAssert.False(vm.IsSelected, "Explicit legend selection controls display, not capture or evaluation.");
            var snapshot = new ReviewSnapshot { Synthetic = false };
            snapshot.DvhSeries.Add(series);
            var report = new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(snapshot);
            TestAssert.True(report.DvhSeries[0].Selected && report.DvhSeries[0].RequiredForTargetReview);
            TestAssert.Equal(60.1234, report.DvhSeries[0].Statistics.D98Gy.Value, "Native D98 is retained without curve or resolved Rx.");
            var copy = Newtonsoft.Json.JsonConvert.DeserializeObject<ReviewDvhSeries>(Newtonsoft.Json.JsonConvert.SerializeObject(series));
            TestAssert.True(copy.RequiredForTargetReview);
            TestAssert.Equal("PTV", copy.TargetKind);
        }

        private static TargetReviewStructure Target(string id, string type, double volume, bool? contained)
        {
            return new TargetReviewStructure { StructureId = id, DicomType = type, VolumeCc = volume, FullyContainedInPtv = contained };
        }

        private static DefaultTargetReviewRule DefaultRule()
        {
            return new DefaultTargetReviewRule { Id = "target-d98-rx", Enabled = true,
                TargetTypes = new List<string> { "PTV", "CTV", "GTV", "ITV" }, Metric = "D98%", Comparator = ">",
                Reference = "target-prescription", PrescriptionPercent = 100 };
        }

        public static void SelectsExactMappedStructures()
        {
            var requested = new[] { "Heart", "PTV_20" };

            TestAssert.True(DvhSelectionPolicy.ShouldSelect("Heart", requested));
            TestAssert.True(DvhSelectionPolicy.ShouldSelect("ptv_20", requested));
            TestAssert.False(DvhSelectionPolicy.ShouldSelect("Heart_PRV", requested));
            TestAssert.False(DvhSelectionPolicy.ShouldSelect("", requested));
            TestAssert.False(DvhSelectionPolicy.ShouldSelect("Heart", null));
        }

        public static void ExcludesExternalAndBodyContours()
        {
            string source = File.ReadAllText(Path.Combine("ClearPlan.Script", "ViewModels", "MainViewModel.cs"));
            int start = source.IndexOf("public void AddDvhCurve(", StringComparison.Ordinal);
            int end = source.IndexOf("public void RemoveDvhCurve(", start, StringComparison.Ordinal);
            string addCurve = source.Substring(start, end - start);
            TestAssert.True(addCurve.Contains("dvh == null") && addCurve.Contains("dvh.CurveData == null") &&
                addCurve.Contains("catch (Exception)"), "Missing native DVH data must not abort default selection before snapshot diagnostics.");
            var requested = new[] { "External", "BODY", "Körper", "Lung_L" };

            TestAssert.True(DvhSelectionPolicy.ShouldSelect("External", requested), "An explicit constraint or native/manual selection must not be vetoed by a structure name.");
            TestAssert.True(DvhSelectionPolicy.ShouldSelect("BODY", requested));
            TestAssert.True(DvhSelectionPolicy.ShouldSelect("Körper", requested));
            TestAssert.True(DvhSelectionPolicy.ShouldSelect("Lung_L", requested));
        }

        public static void ExcludesPartialHelpersButKeepsTargets()
        {
            var requested = new[] { "Lunge Teil", "Heart_TL", "PTV_TL" };

            TestAssert.True(DvhSelectionPolicy.ShouldSelect("Lunge Teil", requested), "Mapped helper constraints are still constraints and require their own DVH.");
            TestAssert.True(DvhSelectionPolicy.ShouldSelect("Heart_TL", requested));
            TestAssert.True(DvhSelectionPolicy.ShouldSelect("PTV_TL", requested));
        }
    }
}
