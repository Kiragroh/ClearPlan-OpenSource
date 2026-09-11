using System;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class TargetClassificationTests
    {
        // Also compile with the two policy sources and TestAssert for an isolated vendor-free regression run.
        public static int Main()
        {
            int failures = 0;
            foreach (var test in new Action[] { DerivedOrganNamesAreNotAutomaticTargets, IntentionalNamesRemainRecognized,
                NativeTypesRemainAuthoritative, DvhAndPamDoNotPromoteOrganExclusions })
            {
                try { test(); Console.WriteLine("PASS " + test.Method.Name); }
                catch (Exception error) { failures++; Console.Error.WriteLine("FAIL " + test.Method.Name + ": " + error.Message); }
            }
            Console.WriteLine("4 target classification tests, " + failures + " failures");
            return failures == 0 ? 0 : 1;
        }

        public static void DerivedOrganNamesAreNotAutomaticTargets()
        {
            foreach (var id in new[] { "Brain-PTV", "Brain minus PTV", "Brain_minus_PTV", "Brain (PTV excluded)",
                "Body-PTV", "Ring_PTV", "Avoid_PTV", "Not PTV", "eval_Brain-PTV", "01_PTV", "APTVALUE",
                "PTV minus Brain", "PTV_without_Brain", "PTV_ohne_Brain", "PTV-Brain", "eval_PTV-Brain",
                "PTV - Brain", "PTV60-Brain", "eval_PTV60 - Brain", "PTV60−Brain" })
                foreach (var type in new[] { "ORGAN", "CONTROL", "" })
                    TestAssert.Equal("", DvhSelectionPolicy.ClassifyTarget(id, type), id + " must not become an automatic target.");
        }

        public static void IntentionalNamesRemainRecognized()
        {
            foreach (var id in new[] { "PTV", "PTV70", "PTV_60", "PTV-60", "PTV_Boost", " eval_PTV60 ", "eval-PTV60" })
                foreach (var type in new[] { "ORGAN", "CONTROL", "" })
                    TestAssert.Equal("PTV", DvhSelectionPolicy.ClassifyTarget(id, type), id + " is an intentional leading target name.");
            TestAssert.Equal("CTV", DvhSelectionPolicy.ClassifyTarget("eval_CTV", ""));
            TestAssert.Equal("GTV", DvhSelectionPolicy.ClassifyTarget("gtv_1", ""));
            TestAssert.Equal("ITV", DvhSelectionPolicy.ClassifyTarget("ITV50", ""));
        }

        public static void NativeTypesRemainAuthoritative()
        {
            foreach (var type in new[] { "PTV", "CTV", "GTV", "ITV" })
                TestAssert.Equal(type, DvhSelectionPolicy.ClassifyTarget("Brain minus PTV", type));
            TestAssert.Equal("CTV", DvhSelectionPolicy.ClassifyTarget("PTV70", "CTV"));
            TestAssert.Equal("", DvhSelectionPolicy.ClassifyTarget("PTV60", "EXTERNAL"));
            TestAssert.Equal("", DvhSelectionPolicy.ClassifyTarget("PTV60", "SUPPORT"));
        }

        public static void DvhAndPamDoNotPromoteOrganExclusions()
        {
            var structures = new[] {
                new TargetReviewStructure { StructureId = "Brain-PTV", DicomType = "ORGAN" },
                new TargetReviewStructure { StructureId = "Brain minus PTV", DicomType = "CONTROL" },
                new TargetReviewStructure { StructureId = "PTV60", DicomType = "ORGAN" }
            };
            TestAssert.Equal("PTV60", string.Join("|", DvhSelectionPolicy.SelectTargetStructureIds(structures)));
            var ptvs = structures.Where(s => DvhSelectionPolicy.ClassifyTarget(s.StructureId, s.DicomType) == "PTV")
                .Select(s => s.StructureId).ToArray();
            string provenance, reason;
            string[] eligible = structures.Select(s => s.StructureId).ToArray();
            TestAssert.Equal("PTV60", PamTargetSelection.Resolve(null, null, eligible, ptvs, out provenance, out reason));
            TestAssert.Equal("Brain-PTV", PamTargetSelection.Resolve("Brain-PTV", null, eligible, ptvs, out provenance, out reason),
                "Explicit selection is not vetoed by automatic name recognition.");
            TestAssert.True(DvhSelectionPolicy.ShouldSelect("Brain-PTV", new[] { "Brain-PTV" }),
                "An explicitly requested OAR DVH must not disappear when target classification is corrected.");
        }
    }
}
