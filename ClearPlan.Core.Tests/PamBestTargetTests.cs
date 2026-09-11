using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class PamBestTargetTests
    {
        public static int Main()
        {
            int failures = 0;
            foreach (var test in new Action[] { DefaultPtvRuleUsesStrict95PercentOnly,
                AutomaticSelectionWaitsForAllPtvCandidates, LowestCompletePamWinsAndTiesAreDeterministic,
                MissingNonfiniteAndIncompletePamNeverWin, ExplicitSelectionRemainsAuthoritative,
                NativeAutomaticWorkflowPreservesDetachedMlcGeometry })
            {
                try { test(); Console.WriteLine("PASS " + test.Method.Name); }
                catch (Exception error) { failures++; Console.Error.WriteLine("FAIL " + test.Method.Name + ": " + error.Message); }
            }
            Console.WriteLine("6 PAM/default-rule tests, " + failures + " failures");
            return failures == 0 ? 0 : 1;
        }

        public static void DefaultPtvRuleUsesStrict95PercentOnly()
        {
            var config = DefaultReviewRuleConfiguration.Load("ClearPlan.Script/Distribution/DefaultReviewRules.json");
            TestAssert.Equal("available", config.Status);
            var rule = config.Rules.Single(r => r.TargetTypes.Contains("PTV"));
            TestAssert.Equal(95.0, rule.PrescriptionPercent, "The shipped PTV rule must use 95%, not 100% of its assigned Rx.");
            TestAssert.Equal("D98%", rule.Metric);
            TestAssert.Equal(">", rule.Comparator);
            var rx = new[] { new TargetPrescriptionDose { TargetId = "PTV", TargetType = "Volume", DosePerFractionGy = 2, FractionCount = 30 } };
            TestAssert.Equal("fail", UserTargetCoverageRule.Evaluate(rule, "PTV", rx, 30, 57, true).Status);
            TestAssert.Equal("pass", UserTargetCoverageRule.Evaluate(rule, "PTV", rx, 30, 57.0001, true).Status);
            TestAssert.Equal("fail", UserTargetCoverageRule.Evaluate(rule, "PTV", rx, 30, 56.9999, true).Status);
            foreach (string targetType in new[] { "CTV", "GTV", "ITV" })
                TestAssert.Equal(100.0, config.Rules.Single(r => r.TargetTypes.Contains(targetType)).PrescriptionPercent,
                    "The PTV-only request must not silently change other target goals.");
        }

        public static void AutomaticSelectionWaitsForAllPtvCandidates()
        {
            string provenance, reason;
            TestAssert.Equal(null, PamTargetSelection.Resolve(null, "PTV_A", new[] { "PTV_A", "PTV_B" },
                new[] { "PTV_B", "PTV_A" }, out provenance, out reason),
                "A native assigned target must not short-circuit lowest-PAM comparison of multiple PTVs.");
            TestAssert.True(reason.Contains("lowest") && reason.Contains("2"), "Pending comparison must explain the rule and candidate count.");
            TestAssert.Equal("PTV_A", PamTargetSelection.Resolve("PTV_A", "PTV_B", new[] { "PTV_A", "PTV_B" },
                new[] { "PTV_B", "PTV_A" }, out provenance, out reason));
        }

        public static void LowestCompletePamWinsAndTiesAreDeterministic()
        {
            var worse = Candidate("PTV_B", 20);
            var better = Candidate("PTV_C", 40);
            var tied = Candidate("PTV_A", 40);
            var selected = Select(new[] { worse, better, tied }, new ReviewPlanAnalysis());
            TestAssert.Equal("PTV_A", selected.TargetStructureId);
            TestAssert.Equal(0.0, selected.Pam.Value, "Valid zero PAM is allowed; missing PAM is not zero.");
            TestAssert.Equal("PTV_A", Select(new[] { tied, better, worse }, new ReviewPlanAnalysis()).TargetStructureId);
            TestAssert.Equal(3, Read<int>(selected, "PamTargetCandidateCount"));
            TestAssert.Equal(3, Read<int>(selected, "PamValidTargetCandidateCount"));
            TestAssert.True(selected.TargetSelectionProvenance.Contains("lowest") && selected.TargetSelectionProvenance.Contains("clinical"));
            TestAssert.Equal("AutomaticLowestPam", Read<string>(selected, "TargetSelectionMode"));
        }

        public static void MissingNonfiniteAndIncompletePamNeverWin()
        {
            var valid = Candidate("PTV_VALID", 20);
            var missing = Candidate("PTV_MISSING", 40); missing.Pam = null;
            var nan = Candidate("PTV_NAN", 40); nan.Pam = double.NaN;
            var infinity = Candidate("PTV_INFINITY", 40); infinity.Pam = double.PositiveInfinity;
            var unavailable = Candidate("PTV_UNAVAILABLE", 40); unavailable.PamStatus = "unavailable";
            var partial = Candidate("PTV_PARTIAL", 40); partial.Beams[0].Pam = null;
            var missingPoint = Candidate("PTV_MISSING_CP", 40); missingPoint.Beams[0].ControlPoints[1].BlockedTargetFraction = null;
            var unknownWeighting = Candidate("PTV_UNKNOWN_WEIGHTING", 40); unknownWeighting.PamWeightingMode = "Unknown";
            var selected = Select(new[] { missing, nan, infinity, unavailable, partial, missingPoint, unknownWeighting, valid }, new ReviewPlanAnalysis());
            TestAssert.Equal("PTV_VALID", selected.TargetStructureId);
            TestAssert.Equal(8, Read<int>(selected, "PamTargetCandidateCount"));
            TestAssert.Equal(1, Read<int>(selected, "PamValidTargetCandidateCount"));
            var fallback = Candidate("OLD_TARGET", 40);
            selected = Select(new[] { missing, nan, infinity, unavailable, partial }, fallback);
            TestAssert.Equal(null, selected.Pam);
            TestAssert.Equal(null, selected.TargetStructureId);
            TestAssert.Equal("unavailable", selected.PamStatus);
            TestAssert.Equal(0, Read<int>(selected, "PamValidTargetCandidateCount"));
            TestAssert.True(selected.PamReason.Contains("No eligible PTV"));
        }

        public static void ExplicitSelectionRemainsAuthoritative()
        {
            string provenance, reason;
            TestAssert.Equal("BODY", PamTargetSelection.Resolve("BODY", null, new[] { "BODY", "PTV_A", "PTV_B" },
                new[] { "PTV_A", "PTV_B" }, out provenance, out reason));
            TestAssert.True(provenance.Contains("Explicit"));
            TestAssert.Equal(null, PamTargetSelection.Resolve("MISSING", null, new[] { "PTV_A", "PTV_B" },
                new[] { "PTV_A", "PTV_B" }, out provenance, out reason));
            TestAssert.True(reason.Contains("No fallback"));
        }

        public static void NativeAutomaticWorkflowPreservesDetachedMlcGeometry()
        {
            string code = File.ReadAllText("ClearPlan.Script/Review/EsapiPlanAnalysisBuilder.cs");
            TestAssert.True(code.Contains("CaptureAutomaticCandidates") && code.Contains("SelectLowestAvailable"),
                "Native automatic analysis must capture all eligible PTVs and select the lowest complete result.");
            TestAssert.True(code.Contains("PlanAnalysisSnapshot.Copy(analysis)"), "Retargeting must preserve imported physical MLC layers in detached copies.");
            TestAssert.True(code.Contains("AutomaticLowestPam") && code.Contains("Explicit"));
            foreach (string mutation in new[] { "BeginModifications(", "SaveModifications(", "AddStructure(", "TargetVolumeID =" })
                TestAssert.False(code.Contains(mutation));
        }

        private static ReviewPlanAnalysis Candidate(string id, double openWidth)
        {
            var plan = new ReviewPlanAnalysis { TargetStructureId = id };
            var beam = new ReviewBeamAnalysis { BeamId = "SYNTHETIC", MetersetMu = 100 };
            foreach (int index in new[] { 0, 1 }) beam.ControlPoints.Add(new ReviewControlPointSample {
                Index = index, CumulativeMetersetWeight = index,
                Aperture = new ApertureGeometry { Jaws = new ApertureRectangle(-20, -10, -20 + openWidth, 10) },
                TargetOutlines = new List<List<BeamPoint>> { new List<BeamPoint> {
                    new BeamPoint(-20,-10), new BeamPoint(20,-10), new BeamPoint(20,10), new BeamPoint(-20,10) } } });
            plan.Beams.Add(beam);
            return PlanAnalysisCalculator.Calculate(plan);
        }

        private static ReviewPlanAnalysis Select(IEnumerable<ReviewPlanAnalysis> candidates, ReviewPlanAnalysis fallback)
        {
            var method = typeof(PamTargetSelection).GetMethod("SelectLowestAvailable");
            TestAssert.NotNull(method, "A testable detached lowest-complete-PAM selector is required.");
            return (ReviewPlanAnalysis)method.Invoke(null, new object[] { candidates, fallback });
        }
        private static T Read<T>(object source, string property)
        {
            var info = source.GetType().GetProperty(property);
            TestAssert.NotNull(info, "Missing transparent PAM selection field " + property);
            return (T)info.GetValue(source, null);
        }
    }
}
